# Deterministic Extraction Test

This test script validates that the Claude Vision API produces consistent, deterministic results when extracting tables from the same PDF file multiple times.

## Purpose

Validates the fix for [Issue #23](https://github.com/zenithventure/FinancialProcessing/issues/23) - non-deterministic table extraction.

## What It Tests

The script runs extraction N times on the same PDF and validates:

1. **Sheet Count Consistency** - All runs produce the same number of sheets
2. **Row/Column Count Consistency** - Each sheet has the same dimensions across all runs
3. **Data Content Consistency** - All runs produce identical data (verified via MD5 hash)

## Usage

### Basic Usage

```bash
python test_deterministic.py <pdf_path>
```

This runs 3 extraction attempts by default.

### Advanced Usage

```bash
# Run 10 times for more confidence
python test_deterministic.py input.pdf --runs 10

# Use specific model
python test_deterministic.py input.pdf --runs 5 --model claude-3-5-sonnet-20241022

# Use specific API key
python test_deterministic.py input.pdf --api-key YOUR_API_KEY
```

## Command Line Options

- `pdf_path` - Path to the PDF file to test (required)
- `--runs N` - Number of test runs (default: 3, minimum: 2)
- `--api-key KEY` - Anthropic API key (uses environment variable if not provided)
- `--model MODEL` - Claude model name (uses environment variable if not provided)

## Example Output

```
============================================================
DETERMINISTIC EXTRACTION TEST
============================================================
PDF File: sample.pdf
Test Runs: 3
Model: claude-3-5-sonnet-20241022
============================================================

============================================================
RUN #1
============================================================
  Processing page 1/2 with Claude Vision...
    ✓ Extracted table: 32 rows x 4 columns
  Processing page 2/2 with Claude Vision...
    ✓ Extracted table: 30 rows x 4 columns

  Sheet 1 (Page 1): 32 rows x 4 cols
  Sheet 2 (Page 2): 30 rows x 4 cols

============================================================
RUN #2
============================================================
  Processing page 1/2 with Claude Vision...
    ✓ Extracted table: 32 rows x 4 columns
  Processing page 2/2 with Claude Vision...
    ✓ Extracted table: 30 rows x 4 columns

  Sheet 1 (Page 1): 32 rows x 4 cols
  Sheet 2 (Page 2): 30 rows x 4 cols

============================================================
RUN #3
============================================================
  Processing page 1/2 with Claude Vision...
    ✓ Extracted table: 32 rows x 4 columns
  Processing page 2/2 with Claude Vision...
    ✓ Extracted table: 30 rows x 4 columns

  Sheet 1 (Page 1): 32 rows x 4 cols
  Sheet 2 (Page 2): 30 rows x 4 cols

============================================================
COMPARISON RESULTS
============================================================

1. Sheet Count Consistency:
   ✓ PASS - All runs produced 2 sheets

2. Row/Column Count Consistency per Sheet:

   Sheet 1:
     ✓ Row count: 32 (consistent)
     ✓ Column count: 4 (consistent)

   Sheet 2:
     ✓ Row count: 30 (consistent)
     ✓ Column count: 4 (consistent)

3. Data Content Consistency (MD5 hash):
   ✓ Sheet 1: Identical content across all runs
   ✓ Sheet 2: Identical content across all runs

============================================================
✓ ALL CHECKS PASSED - Extraction is deterministic!
============================================================
```

## Exit Codes

- `0` - All tests passed (extraction is deterministic)
- `1` - Some tests failed (extraction is non-deterministic) or error occurred

## Testing the Fix

Before fix (temperature not set):
- Different runs would produce different sheet counts
- Row counts would vary between runs
- Data content would differ

After fix (temperature=0):
- All runs produce identical results
- Same sheet count, row count, column count
- Identical data content (verified via MD5 hash)

## Requirements

- Python 3.7+
- Anthropic API key (in environment variable `ANTHROPIC_API_KEY` or via `--api-key`)
- pdf-to-xls-vision package installed
