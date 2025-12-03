"""
PDF to DOCX converter using pdf2docx library.
"""

from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pdf2docx import Converter


class ConversionError(Exception):
    """Raised when PDF to DOCX conversion fails."""
    pass


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


def pdf_to_docx(pdf_path: str, output_dir: str) -> str:
    """
    Convert a PDF file to DOCX format using pdf2docx.

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

    try:
        cv = Converter(str(pdf_file))
        cv.convert(str(docx_path))
        cv.close()

        # Post-process to remove blank pages
        _remove_blank_pages(str(docx_path))

    except Exception as e:
        raise ConversionError(f"PDF conversion failed: {str(e)}")

    if not docx_path.exists():
        raise ConversionError(
            f"Conversion completed but output file not found: {docx_path}"
        )

    return str(docx_path)
