"""Validation utilities for verifying extracted data accuracy."""

import re
from pathlib import Path
import pdfplumber
import pandas as pd
from collections import Counter


def extract_numbers_from_text(text):
    """Extract all numbers from text string.

    Args:
        text: String to extract numbers from

    Returns:
        list: List of number strings (preserves formatting like commas, parentheses)
    """
    # Pattern to match numbers with optional commas, decimals, parentheses, dollar signs
    # Examples: 1,234.56, (123.45), $1,234, 50%, etc.
    pattern = r'\$?\(?\d{1,3}(?:,\d{3})*(?:\.\d+)?\)?%?'
    numbers = re.findall(pattern, text)

    # Clean up numbers for comparison
    cleaned = []
    for num in numbers:
        # Remove formatting but preserve negative sign (parentheses)
        cleaned_num = num.replace('$', '').replace(',', '').replace('%', '')
        # Convert (123) to -123 for comparison
        if cleaned_num.startswith('(') and cleaned_num.endswith(')'):
            cleaned_num = '-' + cleaned_num[1:-1]
        cleaned.append(cleaned_num)

    return cleaned


def extract_numbers_from_pdf(pdf_path):
    """Extract all numbers from PDF text.

    Args:
        pdf_path: Path to PDF file

    Returns:
        dict: Dictionary mapping page numbers to lists of numbers
    """
    pdf_path = Path(pdf_path)
    page_numbers = {}

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if text:
                    numbers = extract_numbers_from_text(text)
                    page_numbers[page_num] = numbers
    except Exception as e:
        print(f"Warning: Could not extract text from PDF for validation: {e}")
        return {}

    return page_numbers


def extract_numbers_from_dataframe(df):
    """Extract all numbers from a dataframe.

    Args:
        df: pandas DataFrame

    Returns:
        list: List of number strings
    """
    numbers = []

    for col in df.columns:
        # Skip Row_Type and Category columns
        if col in ['Row_Type', 'Category', 'Notes']:
            continue

        for val in df[col]:
            if pd.notna(val):
                val_str = str(val)
                # Extract numbers from the value
                extracted = extract_numbers_from_text(val_str)
                numbers.extend(extracted)

    return numbers


def validate_extracted_data(pdf_path, tables, output_path=None):
    """Validate extracted table data against source PDF.

    Compares numbers extracted from PDF text with numbers in the extracted tables
    and generates a validation report of potential discrepancies.

    Args:
        pdf_path: Path to source PDF file
        tables: List of table dictionaries with 'dataframe', 'page', and 'table' keys
        output_path: Optional path to save validation report (text file)

    Returns:
        dict: Validation report with statistics and potential issues
    """
    pdf_path = Path(pdf_path)

    # Extract numbers from PDF text
    pdf_numbers_by_page = extract_numbers_from_pdf(pdf_path)

    if not pdf_numbers_by_page:
        return {
            'status': 'skipped',
            'message': 'Could not extract text from PDF for validation'
        }

    # Collect all PDF numbers
    all_pdf_numbers = []
    for page_nums in pdf_numbers_by_page.values():
        all_pdf_numbers.extend(page_nums)

    # Extract numbers from tables
    all_table_numbers = []
    for table in tables:
        df = table['dataframe']
        table_nums = extract_numbers_from_dataframe(df)
        all_table_numbers.extend(table_nums)

    # Convert to float for comparison (where possible)
    pdf_number_set = Counter()
    for num in all_pdf_numbers:
        try:
            # Normalize to float
            float_val = float(num)
            pdf_number_set[float_val] += 1
        except ValueError:
            # Keep as string if can't convert
            pdf_number_set[num] += 1

    table_number_set = Counter()
    for num in all_table_numbers:
        try:
            float_val = float(num)
            table_number_set[float_val] += 1
        except ValueError:
            table_number_set[num] += 1

    # Find discrepancies
    missing_in_tables = []
    extra_in_tables = []

    # Numbers in PDF but not in tables (or wrong count)
    for num, pdf_count in pdf_number_set.items():
        table_count = table_number_set.get(num, 0)
        if table_count < pdf_count:
            missing_in_tables.append({
                'number': num,
                'pdf_count': pdf_count,
                'table_count': table_count
            })

    # Numbers in tables but not in PDF (or wrong count)
    for num, table_count in table_number_set.items():
        pdf_count = pdf_number_set.get(num, 0)
        if table_count > pdf_count:
            extra_in_tables.append({
                'number': num,
                'pdf_count': pdf_count,
                'table_count': table_count
            })

    # Calculate accuracy metrics
    total_pdf_numbers = sum(pdf_number_set.values())
    total_table_numbers = sum(table_number_set.values())

    # Count matches (numbers that appear same number of times in both)
    matches = sum(min(pdf_number_set[num], table_number_set[num])
                  for num in set(pdf_number_set.keys()) | set(table_number_set.keys()))

    accuracy = (matches / total_pdf_numbers * 100) if total_pdf_numbers > 0 else 0

    # Generate report
    report = {
        'status': 'completed',
        'statistics': {
            'total_pdf_numbers': total_pdf_numbers,
            'total_table_numbers': total_table_numbers,
            'matches': matches,
            'accuracy_percent': round(accuracy, 2)
        },
        'discrepancies': {
            'missing_in_tables': missing_in_tables,  # Show all discrepancies
            'extra_in_tables': extra_in_tables
        }
    }

    # Generate Markdown report if output path provided
    if output_path:
        output_path = Path(output_path)
        # Change extension to .md
        output_path = output_path.parent / f"{output_path.stem}.md"

        report_lines = []
        report_lines.append("# Data Validation Report")
        report_lines.append("")
        report_lines.append(f"**Source PDF:** `{pdf_path.name}`")
        report_lines.append("")
        report_lines.append("## Summary")
        report_lines.append("")
        report_lines.append(f"| Metric | Count |")
        report_lines.append(f"|--------|-------|")
        report_lines.append(f"| Total numbers in PDF | {total_pdf_numbers:,} |")
        report_lines.append(f"| Total numbers in tables | {total_table_numbers:,} |")
        report_lines.append(f"| Matching numbers | {matches:,} |")
        report_lines.append(f"| **Accuracy** | **{accuracy:.2f}%** |")
        report_lines.append("")

        if missing_in_tables:
            report_lines.append("## ⚠️ Numbers in PDF but Missing/Undercounted in Tables")
            report_lines.append("")
            report_lines.append(f"**Found {len(missing_in_tables)} discrepancies**")
            report_lines.append("")
            report_lines.append("| Number | PDF Count | Table Count | Difference |")
            report_lines.append("|--------|-----------|-------------|------------|")
            for item in missing_in_tables:
                diff = item['pdf_count'] - item['table_count']
                report_lines.append(f"| {item['number']:>15} | {item['pdf_count']:>9} | {item['table_count']:>11} | {diff:>10} |")
            report_lines.append("")

        if extra_in_tables:
            report_lines.append("## ⚠️ Numbers in Tables but Missing/Undercounted in PDF")
            report_lines.append("")
            report_lines.append(f"**Found {len(extra_in_tables)} discrepancies**")
            report_lines.append("")
            report_lines.append("| Number | PDF Count | Table Count | Difference |")
            report_lines.append("|--------|-----------|-------------|------------|")
            for item in extra_in_tables:
                diff = item['table_count'] - item['pdf_count']
                report_lines.append(f"| {item['number']:>15} | {item['pdf_count']:>9} | {item['table_count']:>11} | {diff:>10} |")
            report_lines.append("")

        if not missing_in_tables and not extra_in_tables:
            report_lines.append("## ✅ Validation Passed")
            report_lines.append("")
            report_lines.append("No discrepancies detected! All numbers match.")
            report_lines.append("")
        else:
            report_lines.append("## 📋 Recommendation")
            report_lines.append("")
            report_lines.append("Please manually verify the flagged numbers in the Excel output.")
            report_lines.append("Cross-reference with the source PDF to correct any errors.")
            report_lines.append("")
            report_lines.append("### How to Use This Report")
            report_lines.append("")
            report_lines.append("1. **Missing in Tables**: These numbers appear in the PDF but are missing or appear fewer times in your extracted tables. Check if they were:")
            report_lines.append("   - Misread by OCR (e.g., 6 read as 8)")
            report_lines.append("   - Skipped during extraction")
            report_lines.append("   - Part of headers or text that shouldn't be in tables")
            report_lines.append("")
            report_lines.append("2. **Extra in Tables**: These numbers appear in your tables but don't match the PDF. Check if they are:")
            report_lines.append("   - OCR errors (incorrect readings)")
            report_lines.append("   - Duplicates")
            report_lines.append("   - Numbers from headers mistakenly included")
            report_lines.append("")

        report_lines.append("---")
        report_lines.append("")
        report_lines.append(f"*Report generated automatically by pdf-to-xls-vision*")

        # Write report
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(report_lines))

        print(f"  ✓ Validation report saved: {output_path}")

    return report
