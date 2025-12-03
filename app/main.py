"""
FastAPI application entry point.
PDF to Traditional Chinese Word translation service.
"""

from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI

# Load environment variables from .env file
load_dotenv()
from fastapi.staticfiles import StaticFiles

from app.routes import router

# Create necessary directories
UPLOAD_DIR = Path("/tmp/uploads")
EXPORT_DIR = Path("/tmp/exports")
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
EXPORT_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(
    title="PDF to Traditional Chinese Word Translator",
    description="Upload English PDF, convert to Word, and translate to Traditional Chinese",
    version="1.0.0",
)

# Include API routes
app.include_router(router, prefix="/api")

# Serve static files (frontend)
static_dir = Path(__file__).parent.parent / "static"
app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")
