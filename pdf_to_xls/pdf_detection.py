"""PDF type detection utilities."""

import pdfplumber
import fitz  # PyMuPDF


def pdf_has_text(pdf_path):
    """Check if PDF has extractable text or is image-based.

    Args:
        pdf_path: Path to the PDF file

    Returns:
        bool: True if PDF has extractable text, False otherwise
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:3]:  # Check first 3 pages
                text = page.extract_text()
                if text and len(text.strip()) > 50:
                    return True
        return False
    except Exception:
        return False


def pdf_is_image_based(pdf_path):
    """Check if PDF is image-based (contains images but may also have OCR'd text).

    Args:
        pdf_path: Path to the PDF file

    Returns:
        bool: True if PDF is image-based, False otherwise
    """
    try:
        doc = fitz.open(pdf_path)
        for page_num in range(min(3, len(doc))):  # Check first 3 pages
            page = doc[page_num]
            # Check if page has images
            image_list = page.get_images()
            if image_list:
                # Has images - likely scanned/image-based
                doc.close()
                return True
        doc.close()
        return False
    except Exception:
        return False
