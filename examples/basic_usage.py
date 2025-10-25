#!/usr/bin/env python3
"""Basic usage examples for pdf-to-xls-vision library."""

from pdf_to_xls import convert_pdf_to_excel

# Example 1: Convert a single PDF file
# The simplest usage - auto-detects PDF type and uses the best extraction method
result = convert_pdf_to_excel('input.pdf')
if result:
    print(f"Successfully converted to: {result}")

# Example 2: Specify output path
convert_pdf_to_excel('input.pdf', output_path='output.xlsx')

# Example 3: Specify output directory
convert_pdf_to_excel('input.pdf', output_dir='output_folder/')

# Example 4: Force Vision API for complex tables
# Useful when text extraction doesn't capture table structure well
convert_pdf_to_excel('complex_table.pdf', force_vision=True)

# Example 5: Use custom API key and model
convert_pdf_to_excel(
    'input.pdf',
    api_key='your-api-key-here',
    model_name='claude-3-5-sonnet-20241022'
)

# Example 6: Save progress for large PDFs
# Saves progress every 5 pages instead of default 10
convert_pdf_to_excel('large_document.pdf', save_every=5)
