"""Table extraction utilities using vision API and text-based methods."""

from io import StringIO
import pandas as pd
import pdfplumber

from .image_processing import convert_pdf_page_to_image, convert_image_file_to_base64
from .data_cleaning import clean_dataframe_parentheses, clean_malformed_parentheses
from .quality_check import detect_quality_issues


def extract_table_with_claude_vision(pdf_path, client, model_name, output_path=None, save_every=10):
    """Extract tables from PDF using Claude Vision API with incremental saving.

    Args:
        pdf_path: Path to PDF file
        client: Anthropic API client
        model_name: Claude model name
        output_path: Optional path to save incremental progress
        save_every: Save progress every N pages (default: 10)

    Returns:
        list: List of table dictionaries with 'dataframe', 'page', and 'table' keys
    """
    from .excel_writer import save_excel_incremental

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
1. IDENTIFY THE TABLE STRUCTURE:
   - Ignore marginal note references (like "Note 14.", "Note 3.", etc.) that appear in the left margin - these are NOT part of the table columns
   - Focus on the actual table columns that contain categories, line items, and numeric values
   - The main table structure typically has: Category/Section headers, Line item descriptions, and Numeric columns (years, amounts)

2. Preserve all rows and columns exactly as they appear, including:
   - Total/summary rows with their FULL labels (e.g., "Total Other Income", "Gross Potential Income", "Effective Gross Income (EGI)", "Utilities Total", "Total Expenses", "Net Operating Income (NOI)")
   - ALL breakdown/sub-item rows (e.g., "Parking", "Utility Reimbursement", "Pet Fee")
   - Indented or hierarchical items must be included as separate rows

3. HANDLE HIERARCHICAL/CATEGORY STRUCTURE:
   - When you see a category header (e.g., "Administrative Expenses", "Utility Expenses"), note it as the current category
   - ALL line items that follow that category header belong to that category UNTIL a new category header appears
   - If line items appear under a category without the category name being repeated, still include the category name for those items
   - Create a "Category" column that contains the category name for each detail and rollup row

4. Keep all numbers, text, and formatting characters

5. Use commas to separate columns

6. Put values with commas inside quotes

7. Include column headers if present

8. Add a "Row_Type" column as the FIRST column to indicate the type of each row:
   - Use "ROLLUP" for rows that contain words like "Total", "Gross", "Net", "Effective" and represent sums (e.g., "Total Other Income", "Total Expenses", "Net Operating Income")
   - Use "DETAIL" for individual line items that are not totals
   - Use "HEADER" for header/section title rows (main table title, category headers)

9. CRITICAL: Look for notes, annotations, or text outside/beside the main table columns:
   - If you see a "NOTES AND ASSUMPTIONS" section or numbered notes on the side, create a "Notes" column as the LAST column
   - Add the full text of each note to its corresponding row ONLY if the note specifically references that row
   - If a note is general context (not tied to a specific row), leave the Notes column empty for that row
   - Include ALL text content visible in the image, not just the numeric table data

10. Return ONLY the CSV data, no explanation

IMPORTANT:
- Do NOT include marginal note references (like "Note 14.") as table columns
- Do NOT skip breakdown items or sub-categories. Every line item visible in the table must appear in the output.
- Do NOT skip total/rollup rows. These are CRITICAL and must include their full labels with all numbers.
- Do NOT skip text annotations, notes, or explanatory text. All text content should be captured.
- DO carry forward the category name to all items under that category, even if the category name doesn't appear on every line
- Clearly mark which rows are ROLLUP totals vs DETAIL items using the Row_Type column.

If there are multiple tables, extract the largest/main table and any associated notes."""
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

                    # Improved validation: Check if content is valid CSV with actual table data
                    if csv_content and csv_content.strip():
                        # Try to parse as CSV to validate it's actual table data
                        df = None
                        try:
                            # Try standard CSV parsing
                            df = pd.read_csv(StringIO(csv_content))
                        except Exception as e:
                            # If CSV parsing fails, try with on_bad_lines='skip'
                            try:
                                df = pd.read_csv(StringIO(csv_content), on_bad_lines='skip')
                            except Exception:
                                # Last resort: try reading as TSV or with different settings
                                try:
                                    df = pd.read_csv(StringIO(csv_content), sep=None, engine='python')
                                except Exception:
                                    print(f"    CSV parsing error on page {page_num}: {e}")
                                    continue

                        # Check if it has at least one cell of data
                        if df is not None and not df.empty and df.shape[0] > 0 and df.shape[1] > 0:
                            # Valid table - proceed with data cleaning

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
                    else:
                        print(f"    No table content found on page {page_num}")

                except Exception as e:
                    print(f"    API error on page {page_num}: {e}")
                    continue

                # Save progress incrementally every N pages
                if output_path and save_every > 0 and len(tables) > 0 and len(tables) % save_every == 0:
                    save_excel_incremental(tables, output_path, page_num, num_pages)

    except Exception as e:
        print(f"  Vision extraction failed: {e}")
        import traceback
        traceback.print_exc()

    return tables


def extract_table_from_image(image_path, client, model_name):
    """Extract table from image file using Claude Vision API.

    Args:
        image_path: Path to image file (.jpg, .jpeg, .png, .tiff, .tif)
        client: Anthropic API client
        model_name: Claude model name

    Returns:
        list: List of table dictionaries with 'dataframe', 'page', and 'table' keys
    """
    print(f"  Using Claude Vision API ({model_name}) for extraction...")
    tables = []

    try:
        print(f"  Processing image with Claude Vision...")

        # Convert image to base64
        image_data = convert_image_file_to_base64(image_path)

        if not image_data:
            print(f"    Could not convert image to base64")
            return tables

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
1. IDENTIFY THE TABLE STRUCTURE:
   - Ignore marginal note references (like "Note 14.", "Note 3.", etc.) that appear in the left margin - these are NOT part of the table columns
   - Focus on the actual table columns that contain categories, line items, and numeric values
   - The main table structure typically has: Category/Section headers, Line item descriptions, and Numeric columns (years, amounts)

2. Preserve all rows and columns exactly as they appear, including:
   - Total/summary rows with their FULL labels (e.g., "Total Other Income", "Gross Potential Income", "Effective Gross Income (EGI)", "Utilities Total", "Total Expenses", "Net Operating Income (NOI)")
   - ALL breakdown/sub-item rows (e.g., "Parking", "Utility Reimbursement", "Pet Fee")
   - Indented or hierarchical items must be included as separate rows

3. HANDLE HIERARCHICAL/CATEGORY STRUCTURE:
   - When you see a category header (e.g., "Administrative Expenses", "Utility Expenses"), note it as the current category
   - ALL line items that follow that category header belong to that category UNTIL a new category header appears
   - If line items appear under a category without the category name being repeated, still include the category name for those items
   - Create a "Category" column that contains the category name for each detail and rollup row

4. Keep all numbers, text, and formatting characters

5. Use commas to separate columns

6. Put values with commas inside quotes

7. Include column headers if present

8. Add a "Row_Type" column as the FIRST column to indicate the type of each row:
   - Use "ROLLUP" for rows that contain words like "Total", "Gross", "Net", "Effective" and represent sums (e.g., "Total Other Income", "Total Expenses", "Net Operating Income")
   - Use "DETAIL" for individual line items that are not totals
   - Use "HEADER" for header/section title rows (main table title, category headers)

9. CRITICAL: Look for notes, annotations, or text outside/beside the main table columns:
   - If you see a "NOTES AND ASSUMPTIONS" section or numbered notes on the side, create a "Notes" column as the LAST column
   - Add the full text of each note to its corresponding row ONLY if the note specifically references that row
   - If a note is general context (not tied to a specific row), leave the Notes column empty for that row
   - Include ALL text content visible in the image, not just the numeric table data

10. Return ONLY the CSV data, no explanation

IMPORTANT:
- Do NOT include marginal note references (like "Note 14.") as table columns
- Do NOT skip breakdown items or sub-categories. Every line item visible in the table must appear in the output.
- Do NOT skip total/rollup rows. These are CRITICAL and must include their full labels with all numbers.
- Do NOT skip text annotations, notes, or explanatory text. All text content should be captured.
- DO carry forward the category name to all items under that category, even if the category name doesn't appear on every line
- Clearly mark which rows are ROLLUP totals vs DETAIL items using the Row_Type column.

If there are multiple tables, extract the largest/main table and any associated notes."""
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

            # Improved validation: Check if content is valid CSV with actual table data
            if csv_content and csv_content.strip():
                # Try to parse as CSV to validate it's actual table data
                df = None
                try:
                    # Try standard CSV parsing
                    df = pd.read_csv(StringIO(csv_content))
                except Exception as e:
                    # If CSV parsing fails, try with on_bad_lines='skip'
                    try:
                        df = pd.read_csv(StringIO(csv_content), on_bad_lines='skip')
                    except Exception:
                        # Last resort: try reading as TSV or with different settings
                        try:
                            df = pd.read_csv(StringIO(csv_content), sep=None, engine='python')
                        except Exception:
                            print(f"    CSV parsing error: {e}")
                            return tables

                # Check if it has at least one cell of data
                if df is not None and not df.empty and df.shape[0] > 0 and df.shape[1] > 0:
                    # Valid table - proceed with data cleaning

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
                            'page': 1,  # Image files are single "page"
                            'table': 1
                        })
                        print(f"    ✓ Extracted table: {len(df)} rows x {len(df.columns)} columns")
                    else:
                        print(f"    No valid table data in image")
                else:
                    print(f"    No table content found in image")
            else:
                print(f"    No table content found in image")

        except Exception as e:
            print(f"    API error: {e}")
            return tables

    except Exception as e:
        print(f"  Vision extraction failed: {e}")
        import traceback
        traceback.print_exc()

    return tables


def extract_tables_from_text_pdf(pdf_path):
    """Extract tables from text-based PDF using pdfplumber with quality validation.

    Args:
        pdf_path: Path to the PDF file

    Returns:
        tuple: (tables, quality_issues_detected)
            - tables: List of extracted table dictionaries
            - quality_issues_detected: Boolean indicating if any quality issues were found
    """
    tables = []
    pages_with_issues = []
    all_quality_issues = []

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
                                # Check for quality issues
                                issues = detect_quality_issues(df)

                                if issues:
                                    pages_with_issues.append(page_num)
                                    all_quality_issues.extend(issues)
                                    print(
                                        f"  Page {page_num}, Table {table_num}: "
                                        f"{len(df)} rows x {len(df.columns)} columns "
                                        "⚠️  Quality issues detected"
                                    )
                                    for issue in issues:
                                        print(f"    - {issue}")
                                else:
                                    print(
                                        f"  Page {page_num}, Table {table_num}: "
                                        f"{len(df)} rows x {len(df.columns)} columns"
                                    )

                                tables.append({
                                    'dataframe': df,
                                    'page': page_num,
                                    'table': table_num
                                })
                        except Exception as e:
                            print(f"  Warning: Could not process table on page {page_num}: {e}")

    quality_issues_detected = len(pages_with_issues) > 0

    if quality_issues_detected:
        print(f"\n  ⚠️  Quality issues detected on {len(set(pages_with_issues))} page(s)")

    return tables, quality_issues_detected
