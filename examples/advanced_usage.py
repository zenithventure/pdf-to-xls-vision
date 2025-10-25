#!/usr/bin/env python3
"""Advanced usage examples for pdf-to-xls-vision library."""

from pathlib import Path
from pdf_to_xls import (
    convert_pdf_to_excel,
    pdf_is_image_based,
    pdf_has_text,
    detect_quality_issues
)

# Example 1: Pre-check PDF type before conversion
pdf_file = 'document.pdf'

if pdf_is_image_based(pdf_file):
    print(f"{pdf_file} is image-based (scanned), will use Vision API")
elif pdf_has_text(pdf_file):
    print(f"{pdf_file} has extractable text, will use fast text extraction")

# Convert
convert_pdf_to_excel(pdf_file)

# Example 2: Error handling and validation
try:
    result = convert_pdf_to_excel('input.pdf', output_path='output.xlsx')
    if result:
        print(f"Conversion successful: {result}")
    else:
        print("No tables found in PDF")
except FileNotFoundError:
    print("PDF file not found")
except ValueError as e:
    print(f"Configuration error: {e}")
except Exception as e:
    print(f"Conversion error: {e}")

# Example 3: Process multiple files with custom logic
pdf_files = Path('pdfs/').glob('*.pdf')

for pdf_file in pdf_files:
    print(f"\nProcessing: {pdf_file}")

    # Skip image-based PDFs (to save on API costs)
    if pdf_is_image_based(pdf_file):
        print(f"  Skipping image-based PDF: {pdf_file}")
        continue

    # Convert text-based PDFs only
    try:
        result = convert_pdf_to_excel(pdf_file, output_dir='excel_output/')
        if result:
            print(f"  ✓ Converted: {result}")
    except Exception as e:
        print(f"  ✗ Failed: {e}")

# Example 4: Working with extracted data programmatically
# For more advanced use cases, you can use the lower-level APIs
from pdf_to_xls.table_extraction import extract_tables_from_text_pdf
import pandas as pd

pdf_path = 'data.pdf'
tables, quality_issues = extract_tables_from_text_pdf(pdf_path)

for idx, table_data in enumerate(tables):
    df = table_data['dataframe']
    page_num = table_data['page']

    print(f"\nTable from page {page_num}:")
    print(f"  Shape: {df.shape}")
    print(f"  Columns: {list(df.columns)}")

    # Do custom processing on the dataframe
    # For example, filter rows, calculate summaries, etc.

    # Save to custom format
    df.to_csv(f'table_page_{page_num}.csv', index=False)
