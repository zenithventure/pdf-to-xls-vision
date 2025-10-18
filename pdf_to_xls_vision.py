#!/usr/bin/env python3
"""
PDF to XLS Converter using Claude Vision API
Handles both text-based and image-based (scanned) PDFs using AI.
"""

import os
import sys
import argparse
from pathlib import Path
import pdfplumber
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import anthropic
import base64
from io import BytesIO
import json
from dotenv import load_dotenv
import fitz  # PyMuPDF
from PIL import Image, ImageOps
import pytesseract

# Load environment variables from .env file
load_dotenv()


def get_api_key():
    """Get Anthropic API key from environment."""
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key or api_key == 'your-api-key-here':
        raise ValueError(
            "ANTHROPIC_API_KEY not found or not set.\n"
            "Please edit the .env file and add your API key.\n"
            "Get your API key from: https://console.anthropic.com/"
        )
    return api_key


def get_model_name():
    """Get Claude model name from environment."""
    model = os.environ.get('CLAUDE_MODEL', 'claude-sonnet-4-5-20250929')
    return model


def pdf_has_text(pdf_path):
    """Check if PDF has extractable text or is image-based."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:3]:  # Check first 3 pages
                text = page.extract_text()
                if text and len(text.strip()) > 50:
                    return True
        return False
    except:
        return False


def pdf_is_image_based(pdf_path):
    """Check if PDF is image-based (contains images but may also have OCR'd text)."""
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
    except:
        return False


def _fix_cell_parens(value):
    """Fix common parenthesis issues in a single cell.

    Handles:
    - Spaces inside parens: "( 297)" -> "(297)"
    - Duplicate opening parens: "((123)" -> "(123)"
    - Missing closing paren: "( 4410" -> "(4410)"
    - Orphaned closing paren: "123)" -> "(123)"
    """
    import re

    if not isinstance(value, str):
        return value

    val = str(value).strip()

    # Pattern 1: Remove spaces after opening paren: "( 123)" -> "(123)"
    val = re.sub(r'\(\s+', '(', val)

    # Pattern 2: Remove spaces before closing paren: "(123 )" -> "(123)"
    val = re.sub(r'\s+\)', ')', val)

    # Pattern 3: Remove duplicate opening parens: "((123)" -> "(123)"
    val = re.sub(r'\(+', '(', val)

    # Pattern 4: Fix missing closing paren for negative numbers: "( 123" -> "(123)"
    # Only if it starts with ( and contains digits
    if val.startswith('(') and not val.endswith(')'):
        if re.search(r'[\d,.-]+$', val):
            val = val + ')'

    # Pattern 5: Fix orphaned closing paren: "123)" -> "(123)" if it looks like a negative number
    if val.endswith(')') and not val.startswith('('):
        # Check if it's a number followed by )
        if re.match(r'^[\d,.-]+\)$', val):
            val = '(' + val

    return val


def clean_malformed_parentheses(df):
    """Clean malformed parentheses within individual cells.

    This is a post-processing step that fixes common OCR errors from Claude Vision API:
    - Spaces inside parentheses
    - Double opening parentheses
    - Missing closing parentheses
    - Orphaned closing parentheses
    """
    for col in df.columns:
        df[col] = df[col].apply(lambda x: _fix_cell_parens(x) if pd.notna(x) else x)

    return df


def clean_dataframe_parentheses(df):
    """Clean misplaced parentheses in entire dataframe.

    Handles cascading typewriter artifacts where parentheses are shifted across multiple cells.
    A cell ending with '(' means that '(' belongs to the next cell. This can cascade across
    multiple cells in a row.

    Example cascade:
    Before: ["10,947 (", "3,094)(", "578)(", "173"]
    After:  ["10,947", "(3,094)", "(578)", "(173"]

    The logic:
    - Scan left to right looking for cells ending with '('
    - Move that '(' to the next cell
    - If next cell has ')(' pattern, split it: ) stays with current, number gets wrapped, ( cascades
    - Continue until no more trailing '(' found
    """
    import re

    # Process each row
    for idx in df.index:
        # Keep processing until no more changes are made
        changed = True
        while changed:
            changed = False

            for col_idx in range(len(df.columns) - 1):
                curr_col = df.columns[col_idx]
                next_col = df.columns[col_idx + 1]

                curr_val = df.at[idx, curr_col]
                next_val = df.at[idx, next_col]

                # Check if current cell ends with '('
                if pd.notna(curr_val):
                    curr_str = str(curr_val).strip()

                    if curr_str.endswith('('):
                        # Remove trailing '(' from current cell
                        curr_str = curr_str[:-1].strip()

                        # Process next cell
                        if pd.notna(next_val):
                            next_str = str(next_val).strip()

                            # Check if next cell has ')(' pattern
                            match = re.search(r'^([\d,.-]+)\)\($', next_str)
                            if match:
                                # Pattern: curr="X (" + next="123)("
                                # With incoming (, next becomes (123) with trailing ( to cascade
                                # Result: curr="X" + next="(123)("
                                number = match.group(1)
                                df.at[idx, curr_col] = curr_str if curr_str else None
                                df.at[idx, next_col] = f'({number})('
                                changed = True
                            elif next_str.endswith(')') and not next_str.startswith('('):
                                # Pattern: curr="X (" + next="123)"
                                # Result: curr="X" + next="(123)"
                                df.at[idx, curr_col] = curr_str if curr_str else None
                                df.at[idx, next_col] = f'({next_str}'
                                changed = True
                            else:
                                # Just move ( to beginning of next cell
                                df.at[idx, curr_col] = curr_str if curr_str else None
                                df.at[idx, next_col] = '(' + next_str
                                changed = True
                        else:
                            # Next cell is empty
                            df.at[idx, curr_col] = curr_str if curr_str else None
                            df.at[idx, next_col] = '('
                            changed = True

                # Also check if next cell has ')(' pattern without incoming '('
                # This handles cells like "3,094)(" where the ) should go to previous cell
                if pd.notna(next_val):
                    next_str = str(next_val).strip()
                    match = re.search(r'^([\d,.-]+)\)\($', next_str)
                    if match:
                        # Check if current cell doesn't end with '('
                        curr_str = str(curr_val).strip() if pd.notna(curr_val) else ''
                        if not curr_str.endswith('('):
                            # Pattern: curr="X" + next="123)("
                            # Move ) to curr, wrap number, keep trailing (
                            # Result: curr="X)" + next="(123)("
                            number = match.group(1)
                            df.at[idx, curr_col] = (curr_str + ')') if curr_str else ')'
                            df.at[idx, next_col] = f'({number})('
                            changed = True

    # Second pass: Clean up any remaining percentage artifacts
    # Pattern: "-3.34% (" should be "-3.34%"
    for col in df.columns:
        df[col] = df[col].apply(lambda x: re.sub(r'(%)\s*\($', r'\1', str(x).strip()) if pd.notna(x) and isinstance(x, str) else x)

    return df


def detect_orientation(img):
    """Detect image orientation using pytesseract OSD."""
    try:
        # Use pytesseract to detect orientation
        osd = pytesseract.image_to_osd(img)

        # Parse the OSD output to get rotation angle
        rotation = 0
        confidence = 0
        for line in osd.split('\n'):
            if 'Rotate:' in line:
                rotation = int(line.split(':')[1].strip())
            if 'Orientation confidence:' in line:
                confidence = float(line.split(':')[1].strip())

        return rotation, confidence
    except Exception as e:
        # If OSD fails, return 0 (no rotation)
        return 0, 0


def convert_pdf_page_to_image(pdf_path, page_num):
    """Convert a PDF page to base64-encoded image using PyMuPDF, with automatic orientation detection."""
    try:
        # Open PDF with PyMuPDF
        doc = fitz.open(pdf_path)

        # Get the page (0-indexed)
        page = doc[page_num - 1]

        # Render page to image at high resolution
        mat = fitz.Matrix(2.0, 2.0)  # 2x zoom for better quality
        pix = page.get_pixmap(matrix=mat)

        # Convert to PIL Image
        img_bytes = pix.tobytes("png")
        img = Image.open(BytesIO(img_bytes))

        # Close document
        doc.close()

        # Detect actual visual orientation using OCR
        detected_rotation, confidence = detect_orientation(img)

        # Apply rotation correction if needed (only if confidence is reasonable)
        if detected_rotation != 0 and confidence > 1.0:
            # Rotation direction conversion:
            # - Tesseract OSD "Rotate" value = degrees to rotate CLOCKWISE to correct orientation
            # - PIL's rotate() method rotates COUNTER-CLOCKWISE by default
            # - Therefore: PIL_angle = -Tesseract_angle to convert conventions
            # - expand=True ensures the canvas expands to fit the rotated image without cropping
            # Example: If text is 90° clockwise (sideways right), Tesseract returns 270,
            #          and rotate(-270) = rotate 90° clockwise = corrects the orientation
            img = img.rotate(-detected_rotation, expand=True)
            print(f"    Detected rotation {detected_rotation}° (confidence: {confidence:.1f}) - correcting")

        # Convert PIL Image to PNG bytes
        output = BytesIO()
        img.save(output, format='PNG')
        final_img_data = output.getvalue()

        # Encode to base64
        return base64.standard_b64encode(final_img_data).decode('utf-8')

    except Exception as e:
        print(f"    Error converting page to image: {e}")
        import traceback
        traceback.print_exc()
        return None


def extract_table_with_claude_vision(pdf_path, client, model_name, output_path=None, save_every=10):
    """Extract tables from PDF using Claude Vision API with incremental saving.

    Args:
        pdf_path: Path to PDF file
        client: Anthropic API client
        model_name: Claude model name
        output_path: Optional path to save incremental progress
        save_every: Save progress every N pages (default: 10)
    """
    print(f"  Using Claude Vision API ({model_name}) for extraction...")
    tables = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            num_pages = len(pdf.pages)

            for page_num in range(1, num_pages + 1):
                print(f"  Processing page {page_num}/{num_pages} with Claude Vision...")

                # Convert page to image
                image_data = convert_pdf_page_to_image(pdf_path, page_num)

                if not image_data:
                    print(f"    Could not convert page {page_num} to image")
                    continue

                # Call Claude Vision API
                try:
                    message = client.messages.create(
                        model=model_name,
                        max_tokens=4096,
                        messages=[
                            {
                                "role": "user",
                                "content": [
                                    {
                                        "type": "image",
                                        "source": {
                                            "type": "base64",
                                            "media_type": "image/png",
                                            "data": image_data,
                                        },
                                    },
                                    {
                                        "type": "text",
                                        "text": """Extract all tabular data from this image and return it as a CSV format.

Requirements:
1. Preserve all rows and columns exactly as they appear
2. Keep all numbers, text, and formatting characters
3. Use commas to separate columns
4. Put values with commas inside quotes
5. Include column headers if present
6. Return ONLY the CSV data, no explanation

If there are multiple tables, extract the largest/main table."""
                                    }
                                ],
                            }
                        ],
                    )

                    # Extract CSV from response
                    csv_content = message.content[0].text.strip()

                    # Remove markdown code blocks if present
                    if csv_content.startswith('```'):
                        lines = csv_content.split('\n')
                        csv_content = '\n'.join(lines[1:-1]) if len(lines) > 2 else csv_content

                    if csv_content and len(csv_content) > 50:
                        # Parse CSV into DataFrame with error handling
                        from io import StringIO
                        try:
                            # Try standard CSV parsing
                            df = pd.read_csv(StringIO(csv_content))
                        except Exception as e:
                            # If CSV parsing fails, try with error_bad_lines=False
                            try:
                                df = pd.read_csv(StringIO(csv_content), on_bad_lines='skip')
                            except:
                                # Last resort: try reading as TSV or with different settings
                                try:
                                    df = pd.read_csv(StringIO(csv_content), sep=None, engine='python')
                                except:
                                    print(f"    CSV parsing error on page {page_num}: {e}")
                                    continue

                        # Clean up
                        df = df.dropna(how='all').dropna(axis=1, how='all')

                        # Fix typewriter parenthesis artifacts (only after successful parsing)
                        # This fixes values like "3,094)(" -> "(3,094)"
                        try:
                            df = clean_dataframe_parentheses(df)
                        except Exception as e:
                            print(f"    Warning: Could not clean cascading parentheses: {e}")

                        # Fix malformed parentheses within individual cells
                        # This fixes OCR errors like "( 297)" -> "(297)" and "( 4410" -> "(4410)"
                        try:
                            df = clean_malformed_parentheses(df)
                        except Exception as e:
                            print(f"    Warning: Could not clean malformed parentheses: {e}")

                        if not df.empty and len(df) > 0:
                            tables.append({
                                'dataframe': df,
                                'page': page_num,
                                'table': 1
                            })
                            print(f"    ✓ Extracted table: {len(df)} rows x {len(df.columns)} columns")
                        else:
                            print(f"    No valid table data on page {page_num}")
                    else:
                        print(f"    No table content found on page {page_num}")

                except Exception as e:
                    print(f"    API error on page {page_num}: {e}")
                    continue

                # Save progress incrementally every N pages
                if output_path and save_every > 0 and len(tables) > 0 and len(tables) % save_every == 0:
                    _save_excel_incremental(tables, output_path, page_num, num_pages)

    except Exception as e:
        print(f"  Vision extraction failed: {e}")
        import traceback
        traceback.print_exc()

    return tables


def _save_excel_incremental(tables, output_path, current_page, total_pages):
    """Save progress to Excel file incrementally."""
    try:
        print(f"  💾 Saving progress... ({current_page}/{total_pages} pages processed)")
        wb = Workbook()
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])

        for idx, table_data in enumerate(tables, start=1):
            df = table_data['dataframe']
            page_num = table_data['page']
            table_num = table_data['table']

            if len(tables) == 1:
                sheet_name = "Sheet1"
            else:
                sheet_name = f"Page{page_num}_Table{table_num}"

            if len(sheet_name) > 31:
                sheet_name = f"P{page_num}_T{table_num}"

            if sheet_name in wb.sheetnames:
                sheet_name = f"{sheet_name}_{idx}"

            ws = wb.create_sheet(title=sheet_name)

            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
                for c_idx, value in enumerate(row, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=value)

        wb.save(output_path)
        print(f"  ✓ Progress saved: {len(tables)} tables")
    except Exception as e:
        print(f"  Warning: Could not save progress: {e}")


def extract_tables_from_text_pdf(pdf_path):
    """Extract tables from text-based PDF using pdfplumber."""
    tables = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            # Try default extraction
            page_tables = page.extract_tables()

            # Try with text-based settings if nothing found
            if not page_tables:
                page_tables = page.extract_tables(table_settings={
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                    "edge_min_length": 3,
                    "min_words_vertical": 3,
                    "min_words_horizontal": 1,
                })

            if page_tables:
                for table_num, table in enumerate(page_tables, start=1):
                    if table and len(table) > 0:
                        try:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            df = df.dropna(how='all').dropna(axis=1, how='all')

                            if not df.empty:
                                tables.append({
                                    'dataframe': df,
                                    'page': page_num,
                                    'table': table_num
                                })
                                print(f"  Page {page_num}, Table {table_num}: {len(df)} rows x {len(df.columns)} columns")
                        except Exception as e:
                            print(f"  Warning: Could not process table on page {page_num}: {e}")

    return tables


def convert_pdf_to_xls(pdf_path, output_path=None, output_dir=None, save_every=10):
    """Convert PDF to Excel. Uses text extraction for text-based PDFs, Vision API with rotation detection for image-based PDFs.

    Args:
        pdf_path: Path to PDF file
        output_path: Optional output Excel file path
        output_dir: Optional output directory
        save_every: For large PDFs, save progress every N pages (default: 10)
    """
    pdf_path = Path(pdf_path)

    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    if output_path:
        output_path = Path(output_path)
    else:
        if output_dir:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
        else:
            output_dir = pdf_path.parent
        output_path = output_dir / f"{pdf_path.stem}.xlsx"

    print(f"Converting: {pdf_path}")
    print(f"Output: {output_path}")

    try:
        # Check if PDF is image-based (scanned/photos)
        is_image_based = pdf_is_image_based(pdf_path)

        if is_image_based:
            # Image-based PDF: use Vision API with rotation detection
            print("  Image-based PDF detected, using Vision API with rotation detection...")
            api_key = get_api_key()
            model_name = get_model_name()
            client = anthropic.Anthropic(api_key=api_key)
            tables = extract_table_with_claude_vision(pdf_path, client, model_name, output_path, save_every)
        else:
            # Text-based PDF: use direct extraction (fast, no API needed)
            print("  Text-based PDF, using direct extraction...")
            tables = extract_tables_from_text_pdf(pdf_path)

        if not tables:
            print(f"Warning: No tables found in {pdf_path}")
            return None

        print(f"\nCreating Excel file with {len(tables)} table(s)...")

        wb = Workbook()
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])

        for idx, table_data in enumerate(tables, start=1):
            df = table_data['dataframe']
            page_num = table_data['page']
            table_num = table_data['table']

            if len(tables) == 1:
                sheet_name = "Sheet1"
            else:
                sheet_name = f"Page{page_num}_Table{table_num}"

            if len(sheet_name) > 31:
                sheet_name = f"P{page_num}_T{table_num}"

            if sheet_name in wb.sheetnames:
                sheet_name = f"{sheet_name}_{idx}"

            ws = wb.create_sheet(title=sheet_name)

            for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
                for c_idx, value in enumerate(row, start=1):
                    ws.cell(row=r_idx, column=c_idx, value=value)

            print(f"  Sheet {idx}: {sheet_name} (Page {page_num}, {len(df)} rows x {len(df.columns)} columns)")

        wb.save(output_path)
        print(f"✓ Successfully created: {output_path}\n")

        return output_path

    except Exception as e:
        print(f"Error converting {pdf_path}: {e}")
        import traceback
        traceback.print_exc()
        raise


def batch_convert_directory(input_dir, output_dir=None, recursive=False):
    """Batch convert PDFs in directory. Auto-detects text vs image-based PDFs."""
    input_dir = Path(input_dir)

    if not input_dir.exists():
        raise FileNotFoundError(f"Directory not found: {input_dir}")

    if recursive:
        pdf_files = list(input_dir.rglob("*.pdf"))
    else:
        pdf_files = list(input_dir.glob("*.pdf"))

    pdf_files = [f for f in pdf_files if ':Zone.Identifier' not in str(f)]

    if not pdf_files:
        print(f"No PDF files found in {input_dir}")
        return

    print(f"Found {len(pdf_files)} PDF file(s)")
    print("=" * 70)

    success_count = 0
    failed_files = []

    for pdf_path in pdf_files:
        try:
            if output_dir and recursive:
                rel_path = pdf_path.relative_to(input_dir)
                out_dir = Path(output_dir) / rel_path.parent
            else:
                out_dir = output_dir or pdf_path.parent

            result = convert_pdf_to_xls(pdf_path, output_dir=out_dir)
            if result:
                success_count += 1
            print("=" * 70)

        except Exception as e:
            print(f"Failed to convert {pdf_path}: {e}")
            failed_files.append(pdf_path)
            print("=" * 70)

    print(f"\n✓ Conversion complete!")
    print(f"  Successful: {success_count}/{len(pdf_files)}")

    if failed_files:
        print(f"\n✗ Failed files ({len(failed_files)}):")
        for f in failed_files:
            print(f"  - {f}")


def main():
    parser = argparse.ArgumentParser(
        description="PDF to Excel converter with auto-detection (text extraction or Vision API with rotation)"
    )
    parser.add_argument("input", help="PDF file or directory")
    parser.add_argument("-o", "--output", help="Output file or directory")
    parser.add_argument("-r", "--recursive", action="store_true",
                       help="Recursively search subdirectories")

    args = parser.parse_args()
    input_path = Path(args.input)

    if not input_path.exists():
        print(f"Error: Path not found: {input_path}")
        sys.exit(1)

    # Check for API key
    try:
        get_api_key()
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    try:
        if input_path.is_file():
            convert_pdf_to_xls(input_path, output_path=args.output)
        elif input_path.is_dir():
            batch_convert_directory(input_path, output_dir=args.output,
                                   recursive=args.recursive)
        else:
            print(f"Error: Invalid path: {input_path}")
            sys.exit(1)

    except KeyboardInterrupt:
        print("\nConversion cancelled by user")
        sys.exit(1)
    except Exception as e:
        print(f"\nError: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
