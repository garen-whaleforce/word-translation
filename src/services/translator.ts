/**
 * Azure OpenAI translation service
 */

import { AzureOpenAI } from "openai";
import { JobState, updateJob } from "../jobs";
import { DocxSegment } from "./docx";

// Initialize Azure OpenAI client
let client: AzureOpenAI | null = null;

function getClient(): AzureOpenAI {
  if (!client) {
    const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
    const apiKey = process.env.AZURE_OPENAI_API_KEY;
    const apiVersion = process.env.AZURE_OPENAI_API_VERSION || "2024-10-01-preview";

    if (!endpoint || !apiKey) {
      throw new Error("AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_API_KEY must be set");
    }

    client = new AzureOpenAI({
      endpoint,
      apiKey,
      apiVersion,
      timeout: 120000, // 120 seconds timeout for large translation batches
    });
  }
  return client;
}

export interface TranslateOptions {
  chunkSize?: number; // default: 20 (from env CHUNK_SIZE)
  parallelChunks?: number; // default: 2 (from env PARALLEL_CHUNKS)
  sourceLang?: string; // default: "English"
  targetLang?: string; // default: "Traditional Chinese"
}

/**
 * Heuristic to check if text contains English that needs translation
 * Returns true if:
 * 1. Ratio of A-Z letters to non-space chars is > 0.3 (mostly English text), OR
 * 2. Contains English words (3+ consecutive letters) that are not in the skip list
 */
export function looksLikeEnglish(text: string): boolean {
  if (!text || text.trim().length === 0) return false;

  // Skip list: terms that should remain in English
  const skipTerms = [
    // Standards
    /\b(IEC|EN|UL|CSA|CB|ISO|IEEE|ANSI)\b/gi,
    /\bIEC\s*\d+/gi,
    /\bEN\s*\d+/gi,
    /\bUL\s*\d+/gi,
    // PCB designators
    /\b[A-Z]{1,2}\d+\b/g,  // R1, C2, T1, Q1, etc.
    // Units and symbols
    /\b(mm|cm|kg|Hz|kHz|MHz|GHz|mA|A|V|W|kW|MW|°C|°F)\b/g,
    // Common abbreviations that stay in English
    /\b(AC|DC|USB|LED|LCD|PCB|IC|CPU|GPU|RAM|ROM|SSD|HDD|PDF|DOCX)\b/gi,
  ];

  // Remove skip terms from text for checking
  let textToCheck = text;
  for (const pattern of skipTerms) {
    textToCheck = textToCheck.replace(pattern, "");
  }

  // Check 1: Ratio-based detection (for mostly-English text)
  const letterCount = (text.match(/[A-Za-z]/g) || []).length;
  const nonSpaceCount = text.replace(/\s/g, "").length;
  if (nonSpaceCount > 0 && letterCount / nonSpaceCount > 0.3) {
    return true;
  }

  // Check 2: Find English words (3+ consecutive letters) not in skip list
  const englishWords = textToCheck.match(/[A-Za-z]{3,}/g) || [];
  if (englishWords.length > 0) {
    // Filter out common OK words
    const okWords = ["the", "and", "for", "with", "from", "that", "this", "are", "was", "were", "been", "have", "has", "had", "not", "but", "can", "may", "will", "shall"];
    const remainingWords = englishWords.filter(
      (w) => !okWords.includes(w.toLowerCase())
    );
    return remainingWords.length > 0;
  }

  return false;
}

/**
 * Translate a batch of segments using Azure OpenAI
 */
async function translateBatch(
  segments: DocxSegment[],
  sourceLang: string,
  targetLang: string,
  signal: AbortSignal
): Promise<{ index: number; translated: string }[]> {
  const openai = getClient();
  const deployment = process.env.AZURE_OPENAI_DEPLOYMENT_NAME;

  if (!deployment) {
    throw new Error("AZURE_OPENAI_DEPLOYMENT_NAME must be set");
  }

  // Build payload
  const payload = segments.map((seg) => ({
    index: seg.id,
    text: seg.text,
  }));

  const systemPrompt = `You are a senior bilingual technical translator. Your ONLY task is to translate from **English to Traditional Chinese (Taiwan)**.

The documents are CB / IEC safety test reports and power electronics specifications. Your translation MUST sound like it was written by an experienced compliance engineer familiar with IEC/EN standards and safety reports used in Taiwan.

### Core rules
1. **Direction:** Always translate **from English to Traditional Chinese**. Never translate Chinese back to English.
2. **Style:**
   - Use formal, concise wording suitable for test reports, specifications, and certification documents.
   - Use clear engineering wording, not marketing language.
   - Keep sentence structure close to the source when it improves traceability in audits or cross-checking.
3. **Formatting & layout:**
   - Preserve tables, item numbers, headings, clause numbers, units, symbols, and values.
   - Do NOT change numbers, limits, dates, test results, verdicts, or standard identifiers.
   - Keep IEC / EN / UL standard codes (e.g., "IEC 62368-1") in English.

4. **What must remain in English:**
   - Standard names and numbers (IEC/EN/UL/CSA, etc.).
   - Trade names, model names, company names, PCB designators (R1, C2, T1, etc.).
   - Keep abbreviations like "CB", "ICT", "AV" if they are part of standard terminology in the report.

5. **Do NOT leave English untranslated**
   - Except for items listed above, **everything else must be translated into Traditional Chinese**.
   - If you must keep a term in English for technical accuracy, add a clear Traditional Chinese explanation on first occurrence.

### Terminology – MANDATORY glossary (English → Traditional Chinese)
When these English terms or phrases appear, you MUST use EXACTLY the following translations.
Always match the **longest phrase first**.

Parts / components:
- Bleeding resistor → 洩放電阻
- Electrolytic capacitor → 電解電容
- MOSFET → 電晶體
- Current limit resistor → 限流電阻
- Varistor → 突波吸收器
- Primary wire → 一次側引線
- Line chock / Line choke → 電感
- Bobbin → 線架
- Plug holder → 刃片插座塑膠材質
- AC connector → AC 連接器

Circuit sides & windings:
- primary winding → 一次側繞線
- primary circuit → 一次側電路
- primary (alone, referring to primary side) → 一次側
- secondary → 二次側
- Sec. (abbreviation) → 二次側
- winding (general) → 繞線
- core (magnetic core) → 鐵芯

Test conditions, environment, status:
- Unit shutdown immediately → 設備立即中斷
- Unit shutdown / Unit shutdownr → 設備中斷
- Ambient (temperature, condition) → 室溫
- Plastic enclosure outside near → 塑膠外殼內側靠近
- For model → 適用型號
- Optional → 可選

Verdict / Result (in tables):
- P (Pass) → 符合
- N/A (Not applicable) → 不適用
- F (Fail) → 不符合
- Pass → 符合
- Fail → 不符合
- Not applicable → 不適用

Additional wording constraints:
- NEVER translate "primary" as "初級" or "一次測" or "一次"; always use **一次側**.
- NEVER translate "secondary" as "次級"; always use **二次側**.
- Use **Traditional Chinese** characters only.

### Output format
Return ONLY a JSON object with a "translations" array.
Each item in the array: {"index": <number>, "translated": "<text>"}.
Do NOT add explanations or any other text outside the JSON.`;

  const userContent = JSON.stringify(payload);

  let response: any;
  try {
    response = await openai.chat.completions.create(
      {
        model: deployment,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userContent },
        ],
        response_format: { type: "json_object" },
      },
      { signal }
    );
  } catch (apiError: any) {
    console.error("Azure OpenAI API call failed:", apiError.message);
    console.error("Full error:", JSON.stringify(apiError, null, 2));
    throw new Error(`Azure OpenAI API error: ${apiError.message}`);
  }

  // Parse response - check if choices array exists
  if (!response || !response.choices || response.choices.length === 0) {
    console.error("Azure OpenAI response has no choices:", JSON.stringify(response));
    throw new Error("Azure OpenAI returned empty response");
  }

  const content = response.choices[0]?.message?.content;
  if (!content) {
    throw new Error("No content in Azure OpenAI response");
  }

  let parsed: any;
  try {
    parsed = JSON.parse(content);
  } catch (e) {
    console.error("Failed to parse LLM response:", content);
    throw new Error("Failed to parse translation response as JSON");
  }

  // Extract translations array
  const translations = parsed.translations || parsed;
  if (!Array.isArray(translations)) {
    throw new Error("Expected translations array in response");
  }

  // Extract usage info
  const usage = response.usage;

  return {
    translations,
    usage: {
      prompt: usage?.prompt_tokens || 0,
      completion: usage?.completion_tokens || 0,
      reasoning: (usage as any)?.reasoning_tokens || 0,
    },
  } as any;
}

/**
 * Translate all segments in chunks with parallel processing
 */
export async function translateSegments(
  job: JobState,
  segments: DocxSegment[],
  options?: TranslateOptions
): Promise<void> {
  const chunkSize = options?.chunkSize ?? parseInt(process.env.CHUNK_SIZE || "20");
  const parallelChunks = options?.parallelChunks ?? parseInt(process.env.PARALLEL_CHUNKS || "2");
  const sourceLang = options?.sourceLang ?? "English";
  const targetLang = options?.targetLang ?? "Traditional Chinese";

  // Filter segments that need translation (only English-looking ones)
  const toTranslate = segments.filter((seg) => looksLikeEnglish(seg.text));

  if (toTranslate.length === 0) {
    console.log("No English segments found to translate");
    return;
  }

  job.totalSegments = toTranslate.length;
  job.doneSegments = 0;

  // Split into chunks
  const chunks: DocxSegment[][] = [];
  for (let i = 0; i < toTranslate.length; i += chunkSize) {
    chunks.push(toTranslate.slice(i, i + chunkSize));
  }

  console.log(
    `Translating ${toTranslate.length} segments in ${chunks.length} batches (parallel: ${parallelChunks})`
  );

  // Process chunks in parallel batches
  for (let i = 0; i < chunks.length; i += parallelChunks) {
    // Check for cancellation
    if (job.cancelled) {
      throw new Error("Job cancelled");
    }

    // Get the next batch of chunks to process in parallel
    const parallelBatch = chunks.slice(i, i + parallelChunks);
    const batchStart = i + 1;
    const batchEnd = Math.min(i + parallelChunks, chunks.length);

    updateJob(job, {
      stepMessage: `翻譯中 批次 ${batchStart}-${batchEnd}/${chunks.length}...`,
    });

    try {
      // Process chunks in parallel
      const results = await Promise.all(
        parallelBatch.map((chunk) =>
          translateBatch(chunk, sourceLang, targetLang, job.abortController.signal)
        )
      );

      // Process all results
      for (let j = 0; j < results.length; j++) {
        const result = results[j] as any;
        const chunk = parallelBatch[j];
        const { translations, usage } = result;

        // Map translations back to segments
        for (const item of translations) {
          const segment = segments.find((s) => s.id === item.index);
          if (segment) {
            segment.translated = item.translated;
          }
        }

        // Update usage
        job.usage.prompt += usage.prompt;
        job.usage.completion += usage.completion;
        job.usage.reasoning += usage.reasoning;

        // Update progress
        job.doneSegments += chunk.length;
      }

      const progressRatio = job.doneSegments / job.totalSegments;
      // Translation phase is 15% to 70%
      updateJob(job, {
        progress: Math.round(15 + progressRatio * 55),
      });
    } catch (error: any) {
      if (error.name === "AbortError" || job.cancelled) {
        throw new Error("Job cancelled");
      }
      throw error;
    }
  }

  console.log(`Translation complete: ${job.doneSegments} segments translated`);
}

/**
 * Stricter check for QA - only flag segments that are MOSTLY English
 * This is used after initial translation to avoid re-translating segments
 * that just have some technical terms in English (which is expected)
 */
function needsRetranslation(text: string): boolean {
  if (!text || text.trim().length === 0) return false;

  // Skip list: terms that should remain in English
  const skipTerms = [
    // Standards
    /\b(IEC|EN|UL|CSA|CB|ISO|IEEE|ANSI)\s*\d*[-\d]*/gi,
    // PCB designators
    /\b[A-Z]{1,2}\d+\b/g,
    // Units and symbols
    /\b(mm|cm|kg|Hz|kHz|MHz|GHz|mA|A|V|W|kW|MW|°C|°F)\b/gi,
    // Common abbreviations that stay in English
    /\b(AC|DC|USB|LED|LCD|PCB|IC|CPU|GPU|RAM|ROM|SSD|HDD|PDF|DOCX|RMS|EMC|EMI|ESD|RF|IO|I\/O)\b/gi,
    // Model numbers, part numbers
    /\b[A-Z]{2,}-?\d+[A-Z]?\b/g,
  ];

  // Remove skip terms from text for checking
  let textToCheck = text;
  for (const pattern of skipTerms) {
    textToCheck = textToCheck.replace(pattern, "");
  }

  // Check if there are Chinese characters
  const chineseChars = textToCheck.match(/[\u4e00-\u9fff]/g) || [];
  const letterCount = (textToCheck.match(/[A-Za-z]/g) || []).length;

  // If there are Chinese characters and they outnumber English letters, it's translated
  if (chineseChars.length > 0 && chineseChars.length >= letterCount) {
    return false;
  }

  // If > 60% of non-space chars are A-Z letters, needs retranslation
  const nonSpaceCount = textToCheck.replace(/\s/g, "").length;
  if (nonSpaceCount > 0 && letterCount / nonSpaceCount > 0.6) {
    return true;
  }

  return false;
}

/**
 * QA check and retranslate any remaining English segments
 */
export async function qaAndRetranslate(
  job: JobState,
  segments: DocxSegment[],
  options?: TranslateOptions
): Promise<void> {
  updateJob(job, {
    status: "qa-check",
    stepMessage: "QA 檢查中...",
    progress: 80,
  });

  // Find segments that still look like English (using strict looksLikeEnglish check)
  const pending = segments.filter((seg) => {
    // Only check segments that were translated
    if (!seg.translated) return false;
    // Use strict check - if translation still looks like English, retranslate
    return looksLikeEnglish(seg.translated);
  });

  if (pending.length === 0) {
    updateJob(job, {
      stepMessage: "QA 完成：無殘留英文",
      progress: 85,
    });
    console.log("QA done: no English remained");
    return;
  }

  console.log(`QA found ${pending.length} segments still in English`);

  updateJob(job, {
    status: "retranslating",
    stepMessage: `重新翻譯 ${pending.length} 個區段...`,
  });

  // Re-translate only the pending segments with smaller chunk size
  const oldDone = job.doneSegments;
  job.doneSegments = 0;
  job.totalSegments = pending.length;

  // Use smaller chunk size for QA retranslation (default: 5)
  const qaChunkSize = parseInt(process.env.QA_CHUNK_SIZE || "5");
  await translateSegments(job, pending, {
    ...options,
    chunkSize: qaChunkSize,
  });

  // Restore done count for reporting
  job.doneSegments = oldDone + pending.length;

  updateJob(job, {
    stepMessage: "QA 重新翻譯完成",
    progress: 90,
  });
}
