#!/bin/bash

# Script to create GitHub issues for code review findings
# Requires: gh CLI (GitHub CLI) to be installed and authenticated
# Usage: ./create-github-issues.sh

set -e  # Exit on error

echo "Creating GitHub issues for code review findings..."
echo ""

# Check if gh is installed
if ! command -v gh &> /dev/null; then
    echo "Error: GitHub CLI (gh) is not installed."
    echo "Please install it from: https://cli.github.com/"
    exit 1
fi

# Check if authenticated
if ! gh auth status &> /dev/null; then
    echo "Error: Not authenticated with GitHub CLI."
    echo "Please run: gh auth login"
    exit 1
fi

echo "Creating issues..."
echo ""

# Issue 1: Replace Bare Exception Handlers
echo "Creating Issue #1: Replace Bare Exception Handlers..."
gh issue create \
  --title "Replace Bare Exception Handlers" \
  --label "bug,priority: critical" \
  --body "## Priority
**Critical**

## Description
The codebase contains multiple instances of bare \`except:\` clauses that catch all exceptions without specifying the exception type. This is a dangerous anti-pattern that can hide serious errors and make debugging extremely difficult.

## Problem
Bare exception handlers catch ALL exceptions, including:
- \`SystemExit\` (preventing clean program shutdown)
- \`KeyboardInterrupt\` (preventing Ctrl+C from working)
- \`MemoryError\` and other critical system exceptions
- Making debugging difficult as errors are silently swallowed

## Affected Code Locations

### pdf_to_xls_vision.py:55-56
\`\`\`python
except:
    return False
\`\`\`

### pdf_to_xls_vision.py:73-74
\`\`\`python
except:
    return False
\`\`\`

### pdf_to_xls_vision.py:245-247
\`\`\`python
except Exception as e:
    # If OSD fails, return 0 (no rotation)
    return 0, 0
\`\`\`
*Note: This one is better as it catches \`Exception\`, but could be more specific*

### pdf_to_xls_vision.py:377, 381, 384
Multiple bare except clauses in CSV parsing fallback logic

### pdf_to_xls_vision.py:395-397, 402-404
Bare except in parenthesis cleaning functions

## Recommended Fix

Replace bare exception handlers with specific exception types:

\`\`\`python
# Bad - current code
try:
    with pdfplumber.open(pdf_path) as pdf:
        # ... code ...
except:
    return False

# Good - specify exceptions
try:
    with pdfplumber.open(pdf_path) as pdf:
        # ... code ...
except (PDFSyntaxError, PDFPageCountError, FileNotFoundError, PermissionError) as e:
    logger.warning(f\"Could not open PDF {pdf_path}: {e}\")
    return False
except Exception as e:
    logger.error(f\"Unexpected error reading PDF {pdf_path}: {e}\")
    raise
\`\`\`

## Benefits
1. Allows \`KeyboardInterrupt\` and \`SystemExit\` to work correctly
2. Makes debugging easier by seeing actual errors
3. Documents what errors are expected
4. Prevents hiding unexpected bugs

## Implementation Steps
1. Identify the specific exceptions that can be raised in each try block
2. Replace \`except:\` with \`except (ExceptionType1, ExceptionType2) as e:\`
3. Add proper logging of the caught exceptions
4. Re-raise unexpected exceptions
5. Add tests to verify exception handling works correctly

## Testing
- Test that Ctrl+C (KeyboardInterrupt) still works during processing
- Test that invalid PDFs raise appropriate errors
- Test that corrupted files are handled gracefully
- Verify error messages are informative"

echo "✓ Issue #1 created"
echo ""

# Issue 2: Add Proper Logging
echo "Creating Issue #2: Add Proper Logging..."
gh issue create \
  --title "Add Proper Logging" \
  --label "enhancement,priority: high" \
  --body "## Priority
**High**

## Description
The codebase currently uses \`print()\` statements throughout for user feedback and debugging. This approach lacks flexibility and makes it difficult to control output verbosity, redirect logs to files, or debug production issues.

## Problem
Current issues with using \`print()\`:
1. No log levels (INFO, WARNING, ERROR, DEBUG)
2. Can't control verbosity (e.g., verbose vs quiet mode)
3. Can't redirect output to log files
4. Difficult to distinguish between user messages and debug output
5. No timestamps or context information
6. Hard to filter or search logs
7. Makes testing difficult (can't easily verify log messages)

## Current Code Pattern
\`\`\`python
print(f\"Converting: {pdf_path}\")
print(f\"  Processing page {page_num}/{num_pages} with Claude Vision...\")
print(f\"    ✓ Extracted table: {len(df)} rows x {len(df.columns)} columns\")
print(f\"Error converting {pdf_path}: {e}\")
\`\`\`

## Recommended Solution

### 1. Import and Configure Logging
\`\`\`python
import logging

# Configure logger
logger = logging.getLogger(__name__)

def setup_logging(verbose=False):
    \"\"\"Configure logging with appropriate level.\"\"\"
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler('pdf_to_xls.log')
        ]
    )
\`\`\`

### 2. Replace Print Statements
\`\`\`python
# Instead of:
print(f\"Converting: {pdf_path}\")

# Use:
logger.info(f\"Converting: {pdf_path}\")

# Instead of:
print(f\"Error converting {pdf_path}: {e}\")

# Use:
logger.error(f\"Error converting {pdf_path}: {e}\", exc_info=True)
\`\`\`

### 3. Log Levels to Use

- **DEBUG**: Detailed progress (page-by-page processing, API calls)
- **INFO**: General progress (file conversion started/completed, major steps)
- **WARNING**: Recoverable issues (page skipped, parsing fallback used)
- **ERROR**: Serious errors (API failures, file errors)
- **CRITICAL**: System-level failures (API key missing, out of memory)

## CLI Integration

\`\`\`python
parser.add_argument('-v', '--verbose', action='store_true',
                   help='Enable verbose debug output')
parser.add_argument('-q', '--quiet', action='store_true',
                   help='Suppress informational messages')
parser.add_argument('--log-file', type=str,
                   help='Path to log file (default: pdf_to_xls.log)')
\`\`\`

## Benefits

1. **Debugging**: Enable verbose mode to see detailed processing
2. **Production**: Log errors to file for troubleshooting
3. **Testing**: Assert on log messages in tests
4. **Performance**: Disable debug logs in production
5. **Integration**: Other tools can parse structured logs
6. **Monitoring**: Track processing stats and errors over time

## Example Output

### Normal Mode (INFO level)
\`\`\`
2025-10-21 10:30:15 - pdf_to_xls_vision - INFO - Converting: input.pdf
2025-10-21 10:30:15 - pdf_to_xls_vision - INFO - Image-based PDF detected, using Vision API
2025-10-21 10:30:45 - pdf_to_xls_vision - INFO - Successfully created: output.xlsx
\`\`\`

### Verbose Mode (DEBUG level)
\`\`\`
2025-10-21 10:30:15 - pdf_to_xls_vision - INFO - Converting: input.pdf
2025-10-21 10:30:15 - pdf_to_xls_vision - DEBUG - Checking PDF type...
2025-10-21 10:30:15 - pdf_to_xls_vision - DEBUG - Found images on page 1
2025-10-21 10:30:15 - pdf_to_xls_vision - INFO - Image-based PDF detected, using Vision API
\`\`\`"

echo "✓ Issue #2 created"
echo ""

# Issue 3: Extract Configuration Constants
echo "Creating Issue #3: Extract Configuration Constants..."
gh issue create \
  --title "Extract Configuration Constants" \
  --label "enhancement,priority: high,refactoring" \
  --body "## Priority
**High**

## Description
The codebase contains numerous \"magic numbers\" and hardcoded values scattered throughout the code. These values lack context, make the code hard to understand, and are difficult to tune or modify.

## Problem
Magic numbers make code:
1. **Hard to understand**: What does \`50\` mean? Why \`3\` pages?
2. **Hard to maintain**: Need to search entire codebase to change a threshold
3. **Hard to test**: Can't easily test with different configurations
4. **Hard to document**: No central place to explain parameter choices
5. **Error-prone**: Risk of inconsistent values across the codebase

## Current Magic Numbers

### pdf_to_xls_vision.py:50
\`\`\`python
for page in pdf.pages[:3]:  # Check first 3 pages
\`\`\`
**Issue**: Why 3? What if PDF structure varies?

### pdf_to_xls_vision.py:52
\`\`\`python
if text and len(text.strip()) > 50:
\`\`\`
**Issue**: Why 50 characters? Is this enough to determine text-based PDF?

### pdf_to_xls_vision.py:260
\`\`\`python
mat = fitz.Matrix(2.0, 2.0)  # 2x zoom for better quality
\`\`\`
**Issue**: Hardcoded zoom factor - what if we need higher/lower quality?

### pdf_to_xls_vision.py:274
\`\`\`python
if detected_rotation != 0 and confidence > 1.0:
\`\`\`
**Issue**: Why confidence > 1.0? This seems arbitrary

### pdf_to_xls_vision.py:331
\`\`\`python
max_tokens=4096,
\`\`\`
**Issue**: Hardcoded token limit - may need adjustment for different models

## Recommended Solution

Create Configuration Section:
\`\`\`python
# ============================================================================
# Configuration Constants
# ============================================================================

# PDF Type Detection
PDF_SAMPLE_PAGES = 3  # Number of pages to sample for type detection
MIN_TEXT_LENGTH = 50  # Minimum chars to consider PDF as text-based
MIN_CSV_LENGTH = 50   # Minimum chars to consider valid CSV response

# Image Processing
IMAGE_ZOOM_FACTOR = 2.0  # Multiplier for PDF to image conversion
ROTATION_CONFIDENCE_THRESHOLD = 1.0  # Minimum OSD confidence to apply rotation

# API Configuration
CLAUDE_MAX_TOKENS = 4096  # Maximum tokens for Claude API response
CLAUDE_DEFAULT_MODEL = 'claude-sonnet-4-5-20250929'

# Processing
SAVE_PROGRESS_INTERVAL = 10  # Save progress every N pages for large PDFs
\`\`\`

### Usage in Code
\`\`\`python
# Before:
for page in pdf.pages[:3]:
    text = page.extract_text()
    if text and len(text.strip()) > 50:
        return True

# After:
for page in pdf.pages[:PDF_SAMPLE_PAGES]:
    text = page.extract_text()
    if text and len(text.strip()) > MIN_TEXT_LENGTH:
        return True
\`\`\`

## Benefits

1. **Maintainability**: Easy to find and update configuration
2. **Testability**: Can test with different configurations
3. **Documentation**: Constants serve as self-documentation
4. **Flexibility**: Easy to tune for different use cases
5. **Consistency**: Avoid duplicate/inconsistent values
6. **Environment-specific**: Can override per deployment"

echo "✓ Issue #3 created"
echo ""

# Issue 4: Add Input Validation
echo "Creating Issue #4: Add Input Validation..."
gh issue create \
  --title "Add Input Validation Function" \
  --label "enhancement,priority: high" \
  --body "## Priority
**High**

## Description
The codebase currently lacks comprehensive input validation. PDFs are processed without first verifying they are valid, readable, or not corrupted. This leads to cryptic error messages and wasted processing time when issues are discovered late in the pipeline.

## Problem

### Current Behavior
1. **Late failure detection**: Errors discovered only during processing
2. **Cryptic error messages**: Users see library-specific errors, not user-friendly messages
3. **Wasted resources**: Invalid files processed through expensive operations before failing
4. **Poor user experience**: Unclear what went wrong or how to fix it
5. **API key check too late**: Only validated in main(), after some work may be done

## Current Validation Gaps

### pdf_to_xls_vision.py:522-523
\`\`\`python
if not pdf_path.exists():
    raise FileNotFoundError(f\"PDF file not found: {pdf_path}\")
\`\`\`
**Only checks**: File exists
**Missing**: File is readable, not empty, valid PDF format, not corrupted, not password-protected

### No validation for:
- PDF file size (empty files, unreasonably large files)
- PDF page count (0 pages, corrupted structure)
- PDF permissions (password-protected, restricted)
- Output path writability
- Disk space availability
- PDF corruption

## Recommended Solution

Create comprehensive validation:

\`\`\`python
def validate_pdf_file(pdf_path: Path) -> None:
    \"\"\"Validate that PDF file is readable and processable.

    Raises:
        FileNotFoundError: If file doesn't exist
        PDFValidationError: If PDF is invalid, corrupted, or can't be processed
    \"\"\"
    # Check existence
    if not pdf_path.exists():
        raise FileNotFoundError(f\"PDF file not found: {pdf_path}\")

    # Check file size
    if pdf_path.stat().st_size == 0:
        raise PDFValidationError(f\"PDF file is empty: {pdf_path}\")

    # Validate PDF structure
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                raise PDFValidationError(f\"PDF has no pages: {pdf_path}\")
    except Exception as e:
        raise PDFValidationError(f\"Cannot read PDF: {pdf_path}\\nError: {e}\")


def validate_output_path(output_path: Path) -> None:
    \"\"\"Validate that output path is writable.\"\"\"
    # Check parent directory exists
    if not output_path.parent.exists():
        raise OutputPathError(f\"Output directory does not exist: {output_path.parent}\")

    # Check write permissions
    if not os.access(output_path.parent, os.W_OK):
        raise OutputPathError(f\"No write permission: {output_path.parent}\")
\`\`\`

## Benefits

1. **Early failure detection**: Catch issues before wasting time/money
2. **Clear error messages**: User-friendly explanations with solutions
3. **Better UX**: Users know exactly what's wrong and how to fix it
4. **Resource efficiency**: Don't process invalid files
5. **Fail fast**: Exit quickly on configuration issues
6. **Helpful warnings**: Alert on potential issues (low disk space, large files)

## Implementation Steps

1. Create validation functions module section
2. Add custom exception classes for validation errors
3. Implement \`validate_pdf_file()\` function
4. Implement \`validate_output_path()\` function
5. Implement \`validate_api_key_if_needed()\` function
6. Integrate validation into main processing flow
7. Add tests for all validation scenarios"

echo "✓ Issue #4 created"
echo ""

# Issue 5: Improve Resource Management
echo "Creating Issue #5: Improve Resource Management..."
gh issue create \
  --title "Improve Resource Management" \
  --label "bug,priority: high" \
  --body "## Priority
**High**

## Description
The codebase has several resource management issues that could lead to resource leaks, file handle exhaustion, or incomplete cleanup when errors occur. While some resources are properly closed, the pattern is not consistent and lacks fail-safe mechanisms.

## Problem

Resource leaks can cause:
1. **File handle exhaustion**: Too many open files
2. **Memory leaks**: Objects not garbage collected
3. **Disk space waste**: Temporary files not cleaned up
4. **Lock conflicts**: Files locked preventing other operations
5. **Production issues**: Problems that only appear after long running processes

## Current Issues

### 1. PyMuPDF Documents Not Always Closed

#### pdf_to_xls_vision.py:62-72 (pdf_is_image_based)
\`\`\`python
try:
    doc = fitz.open(pdf_path)
    for page_num in range(min(3, len(doc))):
        # ... process ...
        if image_list:
            doc.close()  # ✓ Closed on early return
            return True
    doc.close()  # ✓ Closed on normal path
    return False
except:
    return False  # ✗ Doc not closed if exception occurs!
\`\`\`

**Issue**: If an exception occurs, \`doc.close()\` is never called, leaving file handle open.

### 2. Excel Workbooks Not Closed Explicitly

#### pdf_to_xls_vision.py:588
\`\`\`python
wb.save(output_path)
# ✗ Workbook not explicitly closed
\`\`\`

While Python's garbage collector will eventually clean this up, explicit cleanup is better practice.

## Recommended Solution

Use context managers for guaranteed cleanup:

\`\`\`python
# Before (pdf_is_image_based):
try:
    doc = fitz.open(pdf_path)
    for page_num in range(min(3, len(doc))):
        page = doc[page_num]
        image_list = page.get_images()
        if image_list:
            doc.close()
            return True
    doc.close()
    return False
except:
    return False

# After:
try:
    with fitz.open(pdf_path) as doc:
        for page_num in range(min(3, len(doc))):
            page = doc[page_num]
            image_list = page.get_images()
            if image_list:
                return True  # Context manager handles close
        return False
except (fitz.FileNotFoundError, fitz.PDFError) as e:
    logger.error(f\"Error checking PDF type: {e}\")
    return False
\`\`\`

### Explicitly Close Excel Workbooks

\`\`\`python
# Before:
wb = Workbook()
# ... populate workbook ...
wb.save(output_path)
# Workbook left open

# After:
wb = Workbook()
try:
    # ... populate workbook ...
    wb.save(output_path)
finally:
    wb.close()  # Explicit cleanup
\`\`\`

## Benefits

1. **Reliability**: No resource leaks even when errors occur
2. **Performance**: Resources released promptly, better memory usage
3. **Scalability**: Can process more files in batch without hitting limits
4. **Maintainability**: Clear resource ownership and cleanup
5. **Production ready**: Handles interrupts and errors gracefully

## Implementation Steps

1. Audit all resource acquisitions (file opens, API clients, etc.)
2. Replace manual close with context managers where possible
3. Add try-finally blocks where context managers not available
4. Add signal handlers for graceful shutdown
5. Add tests that verify cleanup on errors
6. Test with interrupts (Ctrl+C) to verify cleanup"

echo "✓ Issue #5 created"
echo ""

echo "=========================================="
echo "✓ All 5 GitHub issues created successfully!"
echo "=========================================="
echo ""
echo "View your issues at: https://github.com/$(gh repo view --json nameWithOwner -q .nameWithOwner)/issues"
