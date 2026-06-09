"""
UNGREED-PDF — Core Conversion Engine
Converts any PDF (text-based or scanned/image) into an editable Word .docx.
"""

from __future__ import annotations

import os
import io
import fitz  # PyMuPDF
import pytesseract
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
from typing import Optional, Callable


# Minimum character count to consider a page as having extractable text
MIN_TEXT_CHARS = 30


def _is_text_page(page: fitz.Page) -> bool:
    """Check if a page has enough extractable text to skip OCR."""
    text = page.get_text("text").strip()
    # Filter out whitespace-only or very short extractions (likely artefacts)
    return len(text) >= MIN_TEXT_CHARS


def _extract_text_from_page(page: fitz.Page) -> str:
    """Extract text from a text-based PDF page using PyMuPDF."""
    return page.get_text("text")


def _ocr_page_image(pil_image: Image.Image) -> str:
    """Run Tesseract OCR on a PIL image and return the extracted text."""
    return pytesseract.image_to_string(pil_image)


def _extract_images_from_page(page: fitz.Page) -> list:
    """Extract embedded images from a PDF page as (image_bytes, ext) tuples."""
    images = []
    for img_info in page.get_images(full=True):
        xref = img_info[0]
        try:
            base_image = page.parent.extract_image(xref)
            if base_image:
                images.append((base_image["image"], base_image["ext"]))
        except Exception:
            continue
    return images


def convert_pdf_to_docx(
    pdf_path: str,
    output_path: str | None = None,
    progress_callback=None,
    ocr_lang: str = "eng",
) -> str:
    """
    Convert a PDF file to an editable Word .docx.

    Parameters
    ----------
    pdf_path : str
        Path to the input PDF.
    output_path : str | None
        Path for the output .docx. Defaults to same name/location as PDF.
    progress_callback : callable | None
        Called with (current_page, total_pages) for progress updates.
    ocr_lang : str
        Tesseract language code(s), e.g. "eng", "eng+fra".

    Returns
    -------
    str
        Path to the generated .docx file.
    """
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    if output_path is None:
        base = os.path.splitext(pdf_path)[0]
        output_path = base + ".docx"

    doc = Document()

    # Style setup
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(11)

    pdf_doc = fitz.open(pdf_path)
    total_pages = len(pdf_doc)

    # Pre-render all pages for OCR fallback
    pil_images = None  # Lazy-loaded only if needed

    for page_num in range(total_pages):
        page = pdf_doc[page_num]

        if page_num > 0:
            doc.add_page_break()

        if _is_text_page(page):
            # --- Text-based page ---
            text = _extract_text_from_page(page)
            _add_text_to_doc(doc, text)

            # Also embed any inline images from the page
            images = _extract_images_from_page(page)
            for img_bytes, ext in images:
                _add_image_to_doc(doc, img_bytes, ext)
        else:
            # --- Image / scanned page → OCR ---
            if pil_images is None:
                pil_images = convert_from_path(pdf_path, dpi=300)

            if page_num < len(pil_images):
                ocr_text = _ocr_page_image(pil_images[page_num])
                if ocr_text.strip():
                    _add_text_to_doc(doc, ocr_text)
                else:
                    # If OCR returns nothing, embed the page as an image
                    _add_pil_image_to_doc(doc, pil_images[page_num])

        if progress_callback:
            progress_callback(page_num + 1, total_pages)

    pdf_doc.close()
    doc.save(output_path)
    return output_path


def _add_text_to_doc(doc: Document, text: str):
    """Add text to a Word document, preserving paragraph breaks."""
    paragraphs = text.split("\n")
    for para_text in paragraphs:
        stripped = para_text.strip()
        if stripped:
            doc.add_paragraph(stripped)


def _add_image_to_doc(doc: Document, img_bytes: bytes, ext: str):
    """Add an image from bytes to the Word document."""
    try:
        stream = io.BytesIO(img_bytes)
        doc.add_picture(stream, width=Inches(5.5))
        last_para = doc.paragraphs[-1]
        last_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception:
        pass


def _add_pil_image_to_doc(doc: Document, pil_image: Image.Image):
    """Add a PIL image to the Word document."""
    try:
        stream = io.BytesIO()
        pil_image.save(stream, format="PNG")
        stream.seek(0)
        doc.add_picture(stream, width=Inches(6.0))
        last_para = doc.paragraphs[-1]
        last_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception:
        pass
