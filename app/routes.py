"""
API routes for PDF upload, conversion, and translation.
"""

import asyncio
import json
import time
import uuid
from pathlib import Path
from typing import AsyncGenerator
from urllib.parse import quote

from fastapi import APIRouter, File, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse

from app.converter import pdf_to_docx_async, ConversionError
from app.docx_translate import translate_docx_to_zh_hant, TranslationError

router = APIRouter()

UPLOAD_DIR = Path("/tmp/uploads")
EXPORT_DIR = Path("/tmp/exports")
DEBUG_DIR = Path("/tmp/debug")  # For intermediate files

MAX_FILE_SIZE = 20 * 1024 * 1024  # 20MB

# Ensure directories exist
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
EXPORT_DIR.mkdir(parents=True, exist_ok=True)
DEBUG_DIR.mkdir(parents=True, exist_ok=True)


@router.get("/healthz")
async def health_check() -> dict:
    """Health check endpoint."""
    return {"status": "ok"}


@router.post("/upload")
async def upload_and_translate(file: UploadFile = File(...)) -> StreamingResponse:
    """
    Upload a PDF file, convert to Word, translate to Traditional Chinese,
    and stream progress updates via Server-Sent Events.

    Args:
        file: The uploaded PDF file.

    Returns:
        StreamingResponse with SSE progress updates ending with result/error.

    Raises:
        HTTPException: If validation fails before processing starts.
    """
    # Validate file extension
    if not file.filename:
        raise HTTPException(status_code=400, detail="No filename provided")

    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(
            status_code=400,
            detail="Invalid file type. Only PDF files are accepted."
        )

    # Read file content and check size
    content = await file.read()
    input_file_size = len(content)

    if input_file_size > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=400,
            detail=f"File too large. Maximum size is {MAX_FILE_SIZE // (1024 * 1024)}MB."
        )

    # Generate unique ID for this job
    file_id = str(uuid.uuid4())

    # Get original filename without extension
    original_name = Path(file.filename).stem

    async def event_stream() -> AsyncGenerator[str, None]:
        """Generate SSE events for progress updates."""
        start_time = time.time()
        progress_queue: asyncio.Queue = asyncio.Queue()

        async def progress_callback(stage: str, current: int, total: int) -> None:
            """Put progress updates into the queue."""
            await progress_queue.put({"stage": stage, "current": current, "total": total})

        def send_event(event_type: str, data: dict) -> str:
            """Format an SSE event."""
            return f"event: {event_type}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"

        # Save uploaded PDF
        pdf_path = UPLOAD_DIR / f"{file_id}.pdf"
        pdf_path.write_bytes(content)

        try:
            # Send: upload complete
            yield send_event("progress", {"stage": "uploaded", "message": "上傳完成"})

            # Step 1: Convert PDF to DOCX
            yield send_event("progress", {"stage": "converting", "message": "PDF 轉換中..."})

            docx_path = await pdf_to_docx_async(str(pdf_path), str(UPLOAD_DIR))

            # Save converted DOCX for debugging
            debug_converted_path = DEBUG_DIR / f"{file_id}_converted.docx"
            import shutil
            shutil.copy(docx_path, debug_converted_path)

            yield send_event("progress", {"stage": "converted", "message": "PDF 轉換完成"})

            # Step 2: Translate DOCX with progress updates
            translated_docx_path = EXPORT_DIR / f"{file_id}.docx"
            first_pass_path = DEBUG_DIR / f"{file_id}_first_pass.docx"

            # Start translation in background task (with first_pass_path for debugging)
            translation_task = asyncio.create_task(
                translate_docx_to_zh_hant(docx_path, str(translated_docx_path), progress_callback, str(first_pass_path))
            )

            # Stream progress updates
            while not translation_task.done():
                try:
                    progress = await asyncio.wait_for(progress_queue.get(), timeout=0.5)
                    if progress["stage"] == "extracting":
                        yield send_event("progress", {"stage": "extracting", "message": "擷取文字中..."})
                    elif progress["stage"] == "translating":
                        yield send_event("progress", {
                            "stage": "translating",
                            "current": progress["current"],
                            "total": progress["total"],
                            "message": f"翻譯中... ({progress['current']}/{progress['total']})"
                        })
                    elif progress["stage"] == "saving":
                        yield send_event("progress", {"stage": "saving", "message": "儲存文件中..."})
                    elif progress["stage"] == "second_pass":
                        yield send_event("progress", {"stage": "second_pass", "message": "掃描漏翻文字..."})
                    elif progress["stage"] == "translating_pass2":
                        yield send_event("progress", {
                            "stage": "translating_pass2",
                            "current": progress["current"],
                            "total": progress["total"],
                            "message": f"補翻中... ({progress['current']}/{progress['total']})"
                        })
                except asyncio.TimeoutError:
                    continue

            # Get translation result
            stats = await translation_task

            # Get output file size
            output_file_size = translated_docx_path.stat().st_size

            # Clean up source PDF (keep debug files)
            pdf_path.unlink(missing_ok=True)
            Path(docx_path).unlink(missing_ok=True)

            # Calculate processing time
            processing_time = time.time() - start_time

            # Send final result with debug download URLs
            yield send_event("complete", {
                "file_id": file_id,
                "download_url": f"/api/download/{file_id}",
                "debug_converted_url": f"/api/debug/{file_id}/converted",
                "debug_first_pass_url": f"/api/debug/{file_id}/first_pass",
                "original_name": original_name,
                "stats": {
                    "processing_time_seconds": round(processing_time, 2),
                    "input_file_size_bytes": input_file_size,
                    "output_file_size_bytes": output_file_size,
                    "original_chars": stats.original_chars,
                    "translated_chars": stats.translated_chars,
                    "prompt_tokens": stats.prompt_tokens,
                    "completion_tokens": stats.completion_tokens,
                    "total_tokens": stats.total_tokens,
                    "api_calls": stats.api_calls,
                    "estimated_cost_usd": round(stats.estimated_cost, 4)
                }
            })

        except ConversionError as e:
            pdf_path.unlink(missing_ok=True)
            yield send_event("error", {"detail": f"PDF conversion failed: {str(e)}"})

        except TranslationError as e:
            pdf_path.unlink(missing_ok=True)
            yield send_event("error", {"detail": f"Translation failed: {str(e)}"})

        except Exception as e:
            pdf_path.unlink(missing_ok=True)
            yield send_event("error", {"detail": f"An unexpected error occurred: {str(e)}"})

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no"
        }
    )


@router.get("/download/{file_id}")
async def download_translated_file(
    file_id: str,
    filename: str = Query(default="translated", description="Original filename for download")
) -> FileResponse:
    """
    Download the translated Word document.

    Args:
        file_id: The unique identifier for the translated file.
        filename: Original filename to use for the download.

    Returns:
        The translated DOCX file.

    Raises:
        HTTPException: If the file is not found.
    """
    file_path = EXPORT_DIR / f"{file_id}.docx"

    if not file_path.exists():
        raise HTTPException(
            status_code=404,
            detail="File not found. It may have expired or never existed."
        )

    # Create download filename: original_translated.docx
    download_filename = f"{filename}_translated.docx"

    # Encode filename for Content-Disposition (RFC 5987) to handle Chinese characters
    encoded_filename = quote(download_filename)

    return FileResponse(
        path=str(file_path),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
        }
    )


@router.get("/debug/{file_id}/{debug_type}")
async def download_debug_file(
    file_id: str,
    debug_type: str,
    filename: str = Query(default="debug", description="Original filename for download")
) -> FileResponse:
    """
    Download debug files (converted or first_pass).

    Args:
        file_id: The unique identifier for the file.
        debug_type: Type of debug file ('converted' or 'first_pass').
        filename: Original filename to use for the download.

    Returns:
        The requested debug DOCX file.

    Raises:
        HTTPException: If the file is not found.
    """
    if debug_type == "converted":
        file_path = DEBUG_DIR / f"{file_id}_converted.docx"
        suffix = "_converted"
    elif debug_type == "first_pass":
        file_path = DEBUG_DIR / f"{file_id}_first_pass.docx"
        suffix = "_first_pass"
    else:
        raise HTTPException(status_code=400, detail="Invalid debug type")

    if not file_path.exists():
        raise HTTPException(
            status_code=404,
            detail="Debug file not found. It may have expired or never existed."
        )

    download_filename = f"{filename}{suffix}.docx"
    encoded_filename = quote(download_filename)

    return FileResponse(
        path=str(file_path),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
        }
    )
