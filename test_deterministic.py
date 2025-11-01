#!/usr/bin/env python3
"""
Test script to validate deterministic extraction from Claude Vision API.

This script runs extraction N times on the same PDF and validates that:
1. All runs produce the same number of sheets
2. All runs produce the same row counts per sheet
3. All runs produce the same column counts per sheet
4. All runs produce identical data content (MD5 hash comparison)

Usage:
    python test_deterministic.py <pdf_path> [--runs N] [--api-key KEY] [--model MODEL]

Example:
    python test_deterministic.py sample.pdf --runs 5
"""

import argparse
import hashlib
import sys
from pathlib import Path
from collections import Counter
import anthropic

from pdf_to_xls.table_extraction import extract_table_with_claude_vision
from pdf_to_xls.config import get_api_key, get_model_name


def hash_dataframe(df):
    """Generate MD5 hash of dataframe content for comparison."""
    # Convert to CSV string and hash it
    csv_content = df.to_csv(index=False)
    return hashlib.md5(csv_content.encode()).hexdigest()


def extract_run(pdf_path, client, model_name, run_number):
    """Run extraction once and return results summary."""
    print(f"\n{'='*60}")
    print(f"RUN #{run_number}")
    print(f"{'='*60}")

    tables = extract_table_with_claude_vision(pdf_path, client, model_name)

    # Collect metrics
    result = {
        'run': run_number,
        'sheet_count': len(tables),
        'sheets': []
    }

    for i, table in enumerate(tables, 1):
        df = table['dataframe']
        sheet_info = {
            'sheet_num': i,
            'page': table['page'],
            'row_count': len(df),
            'col_count': len(df.columns),
            'columns': list(df.columns),
            'hash': hash_dataframe(df)
        }
        result['sheets'].append(sheet_info)
        print(f"  Sheet {i} (Page {table['page']}): {len(df)} rows x {len(df.columns)} cols")

    return result


def compare_results(results):
    """Compare all run results and report differences."""
    print(f"\n{'='*60}")
    print("COMPARISON RESULTS")
    print(f"{'='*60}\n")

    all_passed = True

    # Check 1: Sheet counts
    sheet_counts = [r['sheet_count'] for r in results]
    sheet_count_freq = Counter(sheet_counts)

    print(f"1. Sheet Count Consistency:")
    if len(sheet_count_freq) == 1:
        print(f"   ✓ PASS - All runs produced {sheet_counts[0]} sheets")
    else:
        print(f"   ✗ FAIL - Inconsistent sheet counts detected!")
        for count, freq in sheet_count_freq.items():
            print(f"     {count} sheets: {freq} runs")
        all_passed = False

    # Check 2: Row/Column counts per sheet
    print(f"\n2. Row/Column Count Consistency per Sheet:")
    max_sheets = max(sheet_counts)

    for sheet_idx in range(max_sheets):
        sheet_num = sheet_idx + 1
        row_counts = []
        col_counts = []

        for result in results:
            if sheet_idx < len(result['sheets']):
                sheet = result['sheets'][sheet_idx]
                row_counts.append(sheet['row_count'])
                col_counts.append(sheet['col_count'])

        row_count_freq = Counter(row_counts)
        col_count_freq = Counter(col_counts)

        print(f"\n   Sheet {sheet_num}:")

        # Check rows
        if len(row_count_freq) == 1:
            print(f"     ✓ Row count: {row_counts[0]} (consistent)")
        else:
            print(f"     ✗ Row count: INCONSISTENT!")
            for count, freq in row_count_freq.items():
                print(f"       {count} rows: {freq} runs")
            all_passed = False

        # Check columns
        if len(col_count_freq) == 1:
            print(f"     ✓ Column count: {col_counts[0]} (consistent)")
        else:
            print(f"     ✗ Column count: INCONSISTENT!")
            for count, freq in col_count_freq.items():
                print(f"       {count} columns: {freq} runs")
            all_passed = False

    # Check 3: Data content (MD5 hash)
    print(f"\n3. Data Content Consistency (MD5 hash):")

    for sheet_idx in range(max_sheets):
        sheet_num = sheet_idx + 1
        hashes = []

        for result in results:
            if sheet_idx < len(result['sheets']):
                sheet = result['sheets'][sheet_idx]
                hashes.append(sheet['hash'])

        hash_freq = Counter(hashes)

        if len(hash_freq) == 1:
            print(f"   ✓ Sheet {sheet_num}: Identical content across all runs")
        else:
            print(f"   ✗ Sheet {sheet_num}: DIFFERENT content detected!")
            for hash_val, freq in hash_freq.items():
                print(f"     Hash {hash_val[:8]}...: {freq} runs")
            all_passed = False

    # Final verdict
    print(f"\n{'='*60}")
    if all_passed:
        print("✓ ALL CHECKS PASSED - Extraction is deterministic!")
        print(f"{'='*60}\n")
        return True
    else:
        print("✗ SOME CHECKS FAILED - Extraction is non-deterministic!")
        print(f"{'='*60}\n")
        return False


def main():
    parser = argparse.ArgumentParser(
        description='Test deterministic extraction from PDF files',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python test_deterministic.py input.pdf
  python test_deterministic.py input.pdf --runs 10
  python test_deterministic.py input.pdf --runs 5 --model claude-3-5-sonnet-20241022
        """
    )

    parser.add_argument('pdf_path', type=str, help='Path to PDF file to test')
    parser.add_argument('--runs', type=int, default=3, help='Number of test runs (default: 3)')
    parser.add_argument('--api-key', type=str, help='Anthropic API key (uses env var if not provided)')
    parser.add_argument('--model', type=str, help='Claude model name (uses env var if not provided)')

    args = parser.parse_args()

    # Validate input file
    pdf_path = Path(args.pdf_path)
    if not pdf_path.exists():
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)

    if args.runs < 2:
        print("Error: --runs must be at least 2")
        sys.exit(1)

    # Get API configuration
    api_key = args.api_key or get_api_key()
    model_name = args.model or get_model_name()

    print(f"{'='*60}")
    print("DETERMINISTIC EXTRACTION TEST")
    print(f"{'='*60}")
    print(f"PDF File: {pdf_path}")
    print(f"Test Runs: {args.runs}")
    print(f"Model: {model_name}")
    print(f"{'='*60}")

    # Initialize client
    client = anthropic.Anthropic(api_key=api_key)

    # Run extraction N times
    results = []
    for i in range(1, args.runs + 1):
        try:
            result = extract_run(pdf_path, client, model_name, i)
            results.append(result)
        except Exception as e:
            print(f"\n✗ Run #{i} failed with error: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)

    # Compare results
    all_passed = compare_results(results)

    # Exit with appropriate code
    sys.exit(0 if all_passed else 1)


if __name__ == '__main__':
    main()
