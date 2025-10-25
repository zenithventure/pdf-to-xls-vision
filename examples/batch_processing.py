#!/usr/bin/env python3
"""Batch processing examples for pdf-to-xls-vision library."""

from pdf_to_xls import batch_convert_directory

# Example 1: Convert all PDFs in a directory
# Converts all PDFs in the directory, outputs to same location
batch_convert_directory('pdfs/')

# Example 2: Convert with separate output directory
batch_convert_directory('pdfs/', output_dir='excel_files/')

# Example 3: Recursive search through subdirectories
# Searches all subdirectories and maintains folder structure in output
batch_convert_directory('pdfs/', output_dir='excel_files/', recursive=True)

# Example 4: Force Vision API for all files
# Useful for directories with complex table layouts
batch_convert_directory('pdfs/', force_vision=True)

# Example 5: Handle results
# Get information about successful and failed conversions
results = batch_convert_directory('pdfs/', output_dir='excel_files/', recursive=True)

print(f"\nSuccessfully converted {len(results['success'])} files:")
for file_path in results['success']:
    print(f"  ✓ {file_path}")

if results['failed']:
    print(f"\nFailed to convert {len(results['failed'])} files:")
    for file_path in results['failed']:
        print(f"  ✗ {file_path}")
