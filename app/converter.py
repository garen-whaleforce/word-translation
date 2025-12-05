"""
PDF to DOCX converter using Adobe PDF Services API.
"""

import asyncio
import os
from pathlib import Path
from typing import Optional

import aiohttp
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class ConversionError(Exception):
    """Raised when PDF to DOCX conversion fails."""
    pass


# Adobe PDF Services API endpoints
ADOBE_AUTH_URL = "https://pdf-services.adobe.io/token"
ADOBE_ASSETS_URL = "https://pdf-services.adobe.io/assets"
ADOBE_EXPORT_URL = "https://pdf-services.adobe.io/operation/exportpdf"

# Polling configuration
POLL_INTERVAL = 2  # seconds
MAX_POLL_ATTEMPTS = 60  # max 2 minutes


def get_adobe_credentials() -> tuple[str, str]:
    """
    Get Adobe API credentials from environment variables.

    Returns:
        Tuple of (client_id, client_secret)

    Raises:
        ConversionError: If credentials are not configured.
    """
    client_id = os.environ.get("ADOBE_CLIENT_ID")
    client_secret = os.environ.get("ADOBE_CLIENT_SECRET")

    if not client_id or not client_secret:
        raise ConversionError(
            "Adobe API credentials not configured. "
            "Set ADOBE_CLIENT_ID and ADOBE_CLIENT_SECRET environment variables."
        )

    return client_id, client_secret


async def get_access_token(
    session: aiohttp.ClientSession,
    client_id: str,
    client_secret: str
) -> str:
    """
    Get Adobe API access token.

    Args:
        session: aiohttp client session.
        client_id: Adobe client ID.
        client_secret: Adobe client secret.

    Returns:
        Access token string.

    Raises:
        ConversionError: If authentication fails.
    """
    data = {
        "client_id": client_id,
        "client_secret": client_secret
    }

    async with session.post(ADOBE_AUTH_URL, data=data) as response:
        if response.status != 200:
            text = await response.text()
            raise ConversionError(f"Adobe authentication failed: {text}")

        result = await response.json()
        return result["access_token"]


async def upload_pdf(
    session: aiohttp.ClientSession,
    access_token: str,
    client_id: str,
    pdf_path: str
) -> str:
    """
    Upload PDF to Adobe and get asset ID.

    Args:
        session: aiohttp client session.
        access_token: Adobe access token.
        client_id: Adobe client ID.
        pdf_path: Path to the PDF file.

    Returns:
        Asset ID for the uploaded file.

    Raises:
        ConversionError: If upload fails.
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "X-API-Key": client_id,
        "Content-Type": "application/json"
    }

    # Step 1: Get upload URL
    payload = {
        "mediaType": "application/pdf"
    }

    async with session.post(ADOBE_ASSETS_URL, headers=headers, json=payload) as response:
        if response.status != 200:
            text = await response.text()
            raise ConversionError(f"Failed to get upload URL: {text}")

        result = await response.json()
        upload_uri = result["uploadUri"]
        asset_id = result["assetID"]

    # Step 2: Upload the PDF file
    with open(pdf_path, "rb") as f:
        pdf_data = f.read()

    upload_headers = {
        "Content-Type": "application/pdf"
    }

    async with session.put(upload_uri, headers=upload_headers, data=pdf_data) as response:
        if response.status not in (200, 201):
            text = await response.text()
            raise ConversionError(f"Failed to upload PDF: {text}")

    return asset_id


async def start_export_job(
    session: aiohttp.ClientSession,
    access_token: str,
    client_id: str,
    asset_id: str
) -> str:
    """
    Start PDF to DOCX export job.

    Args:
        session: aiohttp client session.
        access_token: Adobe access token.
        client_id: Adobe client ID.
        asset_id: Asset ID of the uploaded PDF.

    Returns:
        Job polling URL.

    Raises:
        ConversionError: If job creation fails.
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "X-API-Key": client_id,
        "Content-Type": "application/json"
    }

    payload = {
        "assetID": asset_id,
        "targetFormat": "docx"
    }

    async with session.post(ADOBE_EXPORT_URL, headers=headers, json=payload) as response:
        if response.status not in (200, 201):
            text = await response.text()
            raise ConversionError(f"Failed to start export job: {text}")

        # Get polling URL from Location header
        location = response.headers.get("Location") or response.headers.get("location")
        if not location:
            # Try to get from x-request-id
            result = await response.json()
            if "location" in result:
                location = result["location"]
            else:
                raise ConversionError("No polling URL returned from export job")

        return location


async def poll_job_status(
    session: aiohttp.ClientSession,
    access_token: str,
    client_id: str,
    poll_url: str
) -> str:
    """
    Poll job status until completion and return download URL.

    Args:
        session: aiohttp client session.
        access_token: Adobe access token.
        client_id: Adobe client ID.
        poll_url: URL to poll for job status.

    Returns:
        Download URL for the converted file.

    Raises:
        ConversionError: If job fails or times out.
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "X-API-Key": client_id
    }

    for attempt in range(MAX_POLL_ATTEMPTS):
        async with session.get(poll_url, headers=headers) as response:
            if response.status != 200:
                text = await response.text()
                raise ConversionError(f"Failed to poll job status: {text}")

            result = await response.json()
            status = result.get("status", "").lower()

            if status == "done":
                # Get download URL from asset
                asset = result.get("asset", {})
                download_uri = asset.get("downloadUri")
                if not download_uri:
                    raise ConversionError("No download URL in completed job")
                return download_uri

            elif status == "failed":
                error = result.get("error", {})
                raise ConversionError(f"Export job failed: {error}")

            elif status in ("in progress", "in_progress", "pending"):
                await asyncio.sleep(POLL_INTERVAL)

            else:
                # Unknown status, wait and retry
                await asyncio.sleep(POLL_INTERVAL)

    raise ConversionError(f"Export job timed out after {MAX_POLL_ATTEMPTS * POLL_INTERVAL} seconds")


async def download_docx(
    session: aiohttp.ClientSession,
    download_url: str,
    output_path: str
) -> None:
    """
    Download the converted DOCX file.

    Args:
        session: aiohttp client session.
        download_url: URL to download the file from.
        output_path: Path to save the downloaded file.

    Raises:
        ConversionError: If download fails.
    """
    async with session.get(download_url) as response:
        if response.status != 200:
            raise ConversionError(f"Failed to download converted file: {response.status}")

        with open(output_path, "wb") as f:
            async for chunk in response.content.iter_chunked(8192):
                f.write(chunk)


def _remove_blank_pages(docx_path: str) -> None:
    """
    Post-process the DOCX to remove blank pages caused by section breaks.
    Changes all section types to 'continuous' to prevent forced page breaks.
    """
    doc = Document(docx_path)

    for i, section in enumerate(doc.sections):
        if i == 0:
            # Keep first section as-is
            continue

        # Change section type to continuous (no page break)
        sectPr = section._sectPr
        type_elem = sectPr.find(qn('w:type'))
        if type_elem is None:
            type_elem = OxmlElement('w:type')
            sectPr.insert(0, type_elem)
        type_elem.set(qn('w:val'), 'continuous')

    doc.save(docx_path)


async def pdf_to_docx_async(pdf_path: str, output_dir: str) -> str:
    """
    Convert a PDF file to DOCX format using Adobe PDF Services API.

    Args:
        pdf_path: Absolute path to the input PDF file.
        output_dir: Directory where the output DOCX will be saved.

    Returns:
        Absolute path to the generated DOCX file.

    Raises:
        ConversionError: If the conversion fails.
    """
    pdf_file = Path(pdf_path)
    output_directory = Path(output_dir)

    if not pdf_file.exists():
        raise ConversionError(f"Input PDF file not found: {pdf_path}")

    if not output_directory.exists():
        output_directory.mkdir(parents=True, exist_ok=True)

    # Determine output file path
    docx_filename = pdf_file.stem + ".docx"
    docx_path = output_directory / docx_filename

    # Get credentials
    client_id, client_secret = get_adobe_credentials()

    try:
        async with aiohttp.ClientSession() as session:
            # Step 1: Get access token
            access_token = await get_access_token(session, client_id, client_secret)

            # Step 2: Upload PDF
            asset_id = await upload_pdf(session, access_token, client_id, str(pdf_file))

            # Step 3: Start export job
            poll_url = await start_export_job(session, access_token, client_id, asset_id)

            # Step 4: Poll for completion
            download_url = await poll_job_status(session, access_token, client_id, poll_url)

            # Step 5: Download result
            await download_docx(session, download_url, str(docx_path))

    except aiohttp.ClientError as e:
        raise ConversionError(f"Network error during conversion: {str(e)}")

    if not docx_path.exists():
        raise ConversionError(
            f"Conversion completed but output file not found: {docx_path}"
        )

    # Post-process to remove blank pages
    _remove_blank_pages(str(docx_path))

    return str(docx_path)


def pdf_to_docx(pdf_path: str, output_dir: str) -> str:
    """
    Convert a PDF file to DOCX format using Adobe PDF Services API.

    Synchronous wrapper for pdf_to_docx_async.

    Args:
        pdf_path: Absolute path to the input PDF file.
        output_dir: Directory where the output DOCX will be saved.

    Returns:
        Absolute path to the generated DOCX file.

    Raises:
        ConversionError: If the conversion fails.
    """
    return asyncio.run(pdf_to_docx_async(pdf_path, output_dir))
