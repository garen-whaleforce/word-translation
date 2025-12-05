"""
DOCX translation module using OpenAI API.
Reads a Word document, translates content to Traditional Chinese,
and writes the result to a new document.
"""

import asyncio
import os
from dataclasses import dataclass, field
from enum import Enum
from typing import Callable, Optional, Awaitable, Union, Tuple

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from openai import AsyncOpenAI, AsyncAzureOpenAI, APIError, RateLimitError

# Type alias for the client (can be either OpenAI or Azure OpenAI)
OpenAIClient = Union[AsyncOpenAI, AsyncAzureOpenAI]


class TranslationError(Exception):
    """Raised when translation fails."""
    pass


class ElementType(Enum):
    """Type of document element."""
    PARAGRAPH = "paragraph"
    TABLE_CELL = "table_cell"


@dataclass
class TextElement:
    """Represents a text element in the document with its location."""
    element_type: ElementType
    paragraph_index: Optional[int] = None
    table_index: Optional[int] = None
    row_index: Optional[int] = None
    cell_index: Optional[int] = None
    cell_paragraph_index: Optional[int] = None
    original_text: str = ""
    translated_text: str = ""


@dataclass
class TranslationStats:
    """Statistics for translation process."""
    original_chars: int = 0
    translated_chars: int = 0
    prompt_tokens: int = 0
    completion_tokens: int = 0
    total_tokens: int = 0
    api_calls: int = 0

    @property
    def estimated_cost(self) -> float:
        """Estimate cost based on GPT-4o-mini pricing (as of 2025)."""
        # GPT-4o-mini: $0.15/1M input, $0.60/1M output
        input_cost = (self.prompt_tokens / 1_000_000) * 0.15
        output_cost = (self.completion_tokens / 1_000_000) * 0.60
        return input_cost + output_cost


# System prompt for translation
SYSTEM_PROMPT = """You are a professional technical document translator specializing in translating electrical safety certification (CB) test reports from English to Traditional Chinese (Taiwan).

Rules you must follow:
1. Translate the given text from English to Traditional Chinese accurately and naturally.
2. Do NOT summarize, omit, or add any content. Translate everything faithfully.
3. Preserve all numbers, units, proper nouns, technical terms, and formatting markers (such as headings, bullet points, etc.).
4. Maintain the original paragraph structure and line breaks.
5. Do NOT add any notes, explanations, or comments. Output ONLY the translated text.
6. If the input contains multiple paragraphs separated by special markers like "|||", preserve these markers in your output.
7. If text appears to be a heading or title, keep it as a heading in Traditional Chinese.
8. For technical terms that are commonly kept in English (like API, HTTP, etc.), you may keep them in English.

IMPORTANT - Industry-specific terminology (MUST follow these translations):

Circuit/Winding Terms:
- "primary" / "primary circuit" / "primary winding" / "primary side" → 一次側 (NOT 初級/一次測)
- "secondary" / "secondary circuit" / "secondary winding" / "Sec." → 二次側 (NOT 次級/二次測)
- "primary wire" → 一次側引線
- "winding" → 繞線
- "core" (transformer/magnetic) → 鐵芯 (NOT 核心)
- "trace" (PCB) → 銅箔

Component Terms:
- "fuse" → 保險絲 (NOT 熔絲)
- "varistor" / "MOV" → 突波吸收器 (NOT 壓敏電阻)
- "bleeding resistor" → 洩放電阻
- "current limit resistor" → 限流電阻
- "electrolytic capacitor" → 電解電容
- "MOSFET" → 電晶體
- "line choke" → 電感
- "bobbin" → 線架
- "triple insulated wire" → 三層絕緣線 (NOT 三重絕緣線)
- "AC connector" → AC連接器

Plug/Enclosure Terms:
- "plug holder" / "blade holder" / "插頭座" / "針套材料" → 刃片插座塑膠材質 / AC刃片插座塑膠材質
- "plug" / "blade" (electrical) → 刀刃座 (NOT 插座頭)
- "plastic enclosure outside near" → 塑膠外殼內側靠近

Test/Measurement Terms:
- "ambient" (temperature/condition) → 室溫 (NOT 環境)
- "unit shutdown immediately" → 設備立即中斷
- "unit shutdown" → 設備中斷
- "for model" → 適用型號

General Terms:
- "interchangeable" → 不限 (NOT 可互換)
- "minimum" / "at least" → 至少 (NOT 最小/最低)
- "optional" → 可選

IMPORTANT - Table cell formatting rules:
- Flammability rating cells: When you see "UL 94, UL 746C" or similar, output ONLY "UL 94" (remove UL 746C)
- Empty or blank cells: Keep them empty/blank, do NOT add any content
- Certification/approval cells with file numbers: Remove file numbers, keep ONLY the certification standard names
  Example: "VDE↓40029550↓UL E249609" → "VDE" (remove all file numbers like 40029550, E249609, E121562, etc.)
  Example: "UL 94, UL 746C↓UL E121562" → "UL 94" (keep only the flammability standard)"""


CHUNK_SIZE = 2000  # characters per chunk (increased for efficiency)
MAX_RETRIES = 3
RETRY_DELAY = 2  # seconds
MAX_CONCURRENT_CHUNKS = 15  # max concurrent API calls per file
MAX_CONCURRENT_FILES = 2  # max concurrent file translations

# Global semaphore for file-level concurrency control
_file_semaphore: Optional[asyncio.Semaphore] = None


def get_file_semaphore() -> asyncio.Semaphore:
    """Get or create the global file semaphore."""
    global _file_semaphore
    if _file_semaphore is None:
        _file_semaphore = asyncio.Semaphore(MAX_CONCURRENT_FILES)
    return _file_semaphore


# Special translations for test result indicators
SPECIAL_TRANSLATIONS = {
    "P": "符合",
    "N/A": "不適用",
    "--": "--",
}


def get_special_translation(text: str) -> Optional[str]:
    """
    Check if text matches a special translation case.

    Args:
        text: The text to check.

    Returns:
        The special translation if matched, None otherwise.
    """
    stripped = text.strip()
    return SPECIAL_TRANSLATIONS.get(stripped)


# Patterns that indicate LLM refusal/error messages
REFUSAL_PATTERNS = [
    "抱歉，我無法",
    "抱歉，我不能",
    "對不起，我無法",
    "對不起，我不能",
    "很抱歉，我無法",
    "很抱歉，我不能",
    "I'm sorry",
    "I cannot",
    "I can't",
]


def filter_refusal_message(text: str) -> str:
    """
    Filter LLM refusal messages and replace with '--'.

    Args:
        text: The translated text to check.

    Returns:
        '--' if text is a refusal message, otherwise the original text.
    """
    if not text:
        return text

    stripped = text.strip()
    for pattern in REFUSAL_PATTERNS:
        if stripped.startswith(pattern):
            return "--"

    return text


# Type alias for progress callback: (stage: str, current: int, total: int) -> Awaitable[None]
ProgressCallback = Callable[[str, int, int], Awaitable[None]]


def get_openai_client() -> Tuple[OpenAIClient, str]:
    """
    Create and return an OpenAI client (Azure or standard).

    Returns:
        Tuple of (client, model_or_deployment_name)

    Raises:
        TranslationError: If required API keys are not set.
    """
    # Check for Azure OpenAI first
    azure_api_key = os.environ.get("AZURE_OPENAI_API_KEY")
    azure_endpoint = os.environ.get("AZURE_OPENAI_ENDPOINT")
    azure_deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT")
    azure_api_version = os.environ.get("AZURE_OPENAI_API_VERSION", "2024-12-01-preview")

    if azure_api_key and azure_endpoint and azure_deployment:
        client = AsyncAzureOpenAI(
            api_key=azure_api_key,
            azure_endpoint=azure_endpoint,
            api_version=azure_api_version
        )
        return client, azure_deployment

    # Fallback to standard OpenAI
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise TranslationError(
            "No OpenAI API key configured. "
            "Set either AZURE_OPENAI_* or OPENAI_API_KEY environment variables."
        )
    return AsyncOpenAI(api_key=api_key), "gpt-4o-mini"


async def translate_text_chunk(
    client: OpenAIClient,
    model: str,
    text: str,
    stats: TranslationStats
) -> str:
    """
    Translate a chunk of text using OpenAI API.

    Args:
        client: The OpenAI async client (Azure or standard).
        model: The model or deployment name to use.
        text: The text to translate.
        stats: TranslationStats object to update.

    Returns:
        The translated text in Traditional Chinese.

    Raises:
        TranslationError: If translation fails after retries.
    """
    if not text.strip():
        return text

    last_error = None
    for attempt in range(MAX_RETRIES):
        try:
            response = await client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": text}
                ],
                max_completion_tokens=4096
            )

            # Update stats
            stats.api_calls += 1
            if response.usage:
                stats.prompt_tokens += response.usage.prompt_tokens
                stats.completion_tokens += response.usage.completion_tokens
                stats.total_tokens += response.usage.total_tokens

            # Get content and filter refusal messages
            content = response.choices[0].message.content or ""
            return filter_refusal_message(content)

        except RateLimitError as e:
            last_error = e
            if attempt < MAX_RETRIES - 1:
                await asyncio.sleep(RETRY_DELAY * (attempt + 1))
            continue

        except APIError as e:
            if e.status_code and e.status_code >= 500:
                last_error = e
                if attempt < MAX_RETRIES - 1:
                    await asyncio.sleep(RETRY_DELAY * (attempt + 1))
                continue
            raise TranslationError(f"OpenAI API error: {str(e)}")

        except Exception as e:
            raise TranslationError(f"Unexpected error during translation: {str(e)}")

    raise TranslationError(
        f"Translation failed after {MAX_RETRIES} attempts: {str(last_error)}"
    )


def extract_text_elements(doc: Document) -> list[TextElement]:
    """
    Extract all text elements from a Word document.

    Args:
        doc: The python-docx Document object.

    Returns:
        A list of TextElement objects with location and text information.
    """
    elements: list[TextElement] = []

    # Extract paragraphs (not in tables)
    for para_idx, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if text:
            elements.append(TextElement(
                element_type=ElementType.PARAGRAPH,
                paragraph_index=para_idx,
                original_text=text
            ))

    # Extract table cells
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for cell_idx, cell in enumerate(row.cells):
                for para_idx, paragraph in enumerate(cell.paragraphs):
                    text = paragraph.text.strip()
                    if text:
                        elements.append(TextElement(
                            element_type=ElementType.TABLE_CELL,
                            table_index=table_idx,
                            row_index=row_idx,
                            cell_index=cell_idx,
                            cell_paragraph_index=para_idx,
                            original_text=text
                        ))

    return elements


def create_chunks(elements: list[TextElement]) -> list[list[int]]:
    """
    Group text elements into chunks for batch translation.

    Args:
        elements: List of TextElement objects.

    Returns:
        List of lists, where each inner list contains indices of elements in that chunk.
    """
    chunks: list[list[int]] = []
    current_chunk: list[int] = []
    current_length = 0

    for idx, element in enumerate(elements):
        text_length = len(element.original_text)

        if current_length + text_length > CHUNK_SIZE and current_chunk:
            chunks.append(current_chunk)
            current_chunk = []
            current_length = 0

        current_chunk.append(idx)
        current_length += text_length

    if current_chunk:
        chunks.append(current_chunk)

    return chunks


async def translate_elements(
    client: OpenAIClient,
    model: str,
    elements: list[TextElement],
    chunks: list[list[int]],
    stats: TranslationStats,
    progress_callback: Optional[ProgressCallback] = None
) -> None:
    """
    Translate all text elements in chunks with concurrent processing.

    Args:
        client: The OpenAI async client (Azure or standard).
        model: The model or deployment name to use.
        elements: List of TextElement objects to translate.
        chunks: List of index groups for batch translation.
        stats: TranslationStats object to update.
        progress_callback: Optional callback for progress updates.
    """
    # First pass: handle special translations (P, N/A, --, etc.)
    for element in elements:
        special = get_special_translation(element.original_text)
        if special is not None:
            element.translated_text = special

    separator = " ||| "
    total_chunks = len(chunks)
    completed_count = 0
    count_lock = asyncio.Lock()
    semaphore = asyncio.Semaphore(MAX_CONCURRENT_CHUNKS)

    async def process_chunk(chunk_indices: list[int]) -> None:
        """Process a single chunk with semaphore control."""
        nonlocal completed_count

        async with semaphore:
            # Filter out elements that already have special translations
            indices_to_translate = [
                idx for idx in chunk_indices
                if not elements[idx].translated_text
            ]

            if not indices_to_translate:
                # Update progress even for skipped chunks
                async with count_lock:
                    completed_count += 1
                    if progress_callback:
                        await progress_callback("translating", completed_count, total_chunks)
                return

            # Combine texts in this chunk
            texts = [elements[i].original_text for i in indices_to_translate]
            combined_text = separator.join(texts)

            # Translate the combined text
            translated_combined = await translate_text_chunk(client, model, combined_text, stats)

            # Split the translation back
            translated_texts = translated_combined.split(separator)

            # Handle case where separator might be translated or modified
            if len(translated_texts) != len(indices_to_translate):
                # Fallback: translate each element individually
                for idx in indices_to_translate:
                    translated = await translate_text_chunk(
                        client, model, elements[idx].original_text, stats
                    )
                    elements[idx].translated_text = translated
            else:
                # Assign translations to elements
                for i, idx in enumerate(indices_to_translate):
                    elements[idx].translated_text = translated_texts[i].strip()

            # Update progress
            async with count_lock:
                completed_count += 1
                if progress_callback:
                    await progress_callback("translating", completed_count, total_chunks)

    # Process all chunks concurrently
    await asyncio.gather(*[process_chunk(chunk) for chunk in chunks])


def write_translations_to_doc(doc: Document, elements: list[TextElement]) -> None:
    """
    Write translated text back to the document.

    Args:
        doc: The python-docx Document object.
        elements: List of TextElement objects with translations.
    """
    for element in elements:
        if not element.translated_text:
            continue

        if element.element_type == ElementType.PARAGRAPH:
            if element.paragraph_index is not None:
                paragraph = doc.paragraphs[element.paragraph_index]
                # Preserve runs structure but replace text
                if paragraph.runs:
                    # Clear all runs except first, put all text in first run
                    first_run = paragraph.runs[0]
                    for run in paragraph.runs[1:]:
                        run.text = ""
                    first_run.text = element.translated_text
                else:
                    paragraph.text = element.translated_text

        elif element.element_type == ElementType.TABLE_CELL:
            if (element.table_index is not None and
                element.row_index is not None and
                element.cell_index is not None and
                element.cell_paragraph_index is not None):

                table = doc.tables[element.table_index]
                cell = table.rows[element.row_index].cells[element.cell_index]
                paragraph = cell.paragraphs[element.cell_paragraph_index]

                if paragraph.runs:
                    first_run = paragraph.runs[0]
                    for run in paragraph.runs[1:]:
                        run.text = ""
                    first_run.text = element.translated_text
                else:
                    paragraph.text = element.translated_text


async def translate_docx_to_zh_hant(
    src_docx_path: str,
    dst_docx_path: str,
    progress_callback: Optional[ProgressCallback] = None
) -> TranslationStats:
    """
    Read a DOCX file, translate its content to Traditional Chinese,
    and save the result to a new file.

    Uses file-level semaphore to limit concurrent file translations.

    Args:
        src_docx_path: Path to the source DOCX file.
        dst_docx_path: Path where the translated DOCX will be saved.
        progress_callback: Optional callback for progress updates.

    Returns:
        TranslationStats with usage statistics.

    Raises:
        TranslationError: If translation fails.
    """
    stats = TranslationStats()
    file_semaphore = get_file_semaphore()

    async with file_semaphore:
        try:
            # Report: extracting text
            if progress_callback:
                await progress_callback("extracting", 0, 0)

            # Load the document
            doc = Document(src_docx_path)

            # Extract all text elements
            elements = extract_text_elements(doc)

            if not elements:
                # No text to translate, just save a copy
                doc.save(dst_docx_path)
                return stats

            # Calculate original character count
            stats.original_chars = sum(len(e.original_text) for e in elements)

            # Create chunks for batch translation
            chunks = create_chunks(elements)

            # Get OpenAI client and translate
            client, model = get_openai_client()
            await translate_elements(client, model, elements, chunks, stats, progress_callback)

            # Calculate translated character count
            stats.translated_chars = sum(len(e.translated_text) for e in elements)

            # Report: saving document
            if progress_callback:
                await progress_callback("saving", 0, 0)

            # Write translations back to document
            write_translations_to_doc(doc, elements)

            # Save the translated document
            doc.save(dst_docx_path)

            return stats

        except TranslationError:
            raise

        except Exception as e:
            raise TranslationError(f"Failed to process document: {str(e)}")
