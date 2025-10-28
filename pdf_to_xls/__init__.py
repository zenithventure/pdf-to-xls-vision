"""PDF to XLS Vision - Convert PDF and image tables to Excel using Claude Vision API.

This package provides intelligent PDF and image to Excel conversion with automatic detection
of text-based vs image-based PDFs, rotation correction, and quality validation.

Supports: PDF, JPG, JPEG, PNG, TIFF, TIF

Example usage:
    from pdf_to_xls import convert_pdf_to_excel, batch_convert_directory

    # Convert a single PDF or image
    convert_pdf_to_excel('input.pdf', 'output.xlsx')
    convert_pdf_to_excel('table.jpg', 'output.xlsx')

    # Batch convert a directory (finds all PDFs and images)
    batch_convert_directory('files/', 'output/', recursive=True)

Main functions:
    - convert_pdf_to_excel: Convert a single PDF or image to Excel
    - batch_convert_directory: Batch convert multiple PDFs and images
"""

from .converter import convert_pdf_to_excel, batch_convert_directory
from .config import get_api_key, get_model_name
from .pdf_detection import pdf_has_text, pdf_is_image_based
from .quality_check import detect_quality_issues

__version__ = '1.0.2'

__all__ = [
    'convert_pdf_to_excel',
    'batch_convert_directory',
    'get_api_key',
    'get_model_name',
    'pdf_has_text',
    'pdf_is_image_based',
    'detect_quality_issues',
]
