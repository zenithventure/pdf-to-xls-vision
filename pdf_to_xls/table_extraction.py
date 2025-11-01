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
                        temperature=0,  # Ensures deterministic output for consistent table extraction
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

CRITICAL ACCURACY REQUIREMENTS:
- Read each character VERY CAREFULLY - verify every letter and digit
- Pay special attention to similar-looking characters: 6 vs 8, O vs 0, l vs I, etc.
- Double-check all numbers for accuracy - transcription errors are NOT acceptable
- Verify text spelling character-by-character - do not guess or autocorrect
- If text is unclear, examine it closely before transcribing

Requirements:
1. IDENTIFY THE TABLE STRUCTURE:
   - Ignore marginal note references (like "Note 14.", "Note 3.", etc.) that appear in the left margin - these are NOT part of the table columns
   - Focus on the actual table columns that contain line items/categories and numeric values
   - The main table structure has: A single column for all categories and line items, followed by numeric columns (years, amounts)
   - CRITICAL: Watch for MULTIPLE SUB-COLUMNS per year/period:
     * Some tables have 2+ columns under each year header (e.g., percentage + amount, budget + actual, quantity + price)
     * Each sub-column MUST be a separate column in the CSV output
     * Create descriptive column names that identify BOTH the period AND the type
     * Examples: "2022_Percent,2022_Amount" or "2023_Budget,2023_Actual" or "Q1_Units,Q1_Price"
     * Look for sub-headers, data patterns, or $ signs to identify column types
     * If no sub-header exists, use descriptive names based on the data (e.g., "2022_Col1", "2022_Col2")

2. OUTPUT STRUCTURE:
   - Add a "Row_Type" column as the FIRST column to indicate the type of each row:
     * Use "HEADER" for section/category headers (e.g., "REVENUES", "EXPENSES", "Administrative Expenses", "Utility Expenses")
     * Use "DETAIL" for individual line items (e.g., "Gross rental income", "Manager's salary", "Electricity")
     * Use "ROLLUP" for total rows (e.g., "Total revenues", "Total expenses", "Net Operating Income")

   - Add a "Category" column as the SECOND column containing:
     * For HEADER rows: The section/category name (e.g., "REVENUES", "Administrative Expenses")
     * For DETAIL rows: The line item name (e.g., "Gross rental income", "Manager's salary")
     * For ROLLUP rows: The total label (e.g., "Total revenues", "Total expenses")

   - DO NOT create separate columns for categories and line items - everything goes in the single "Category" column

   - Follow with the numeric data columns (e.g., "2020", "2019")

3. Preserve all rows exactly as they appear:
   - Section headers (REVENUES, EXPENSES, etc.) → Row_Type: HEADER
   - Category headers (Administrative Expenses, Utility Expenses, etc.) → Row_Type: HEADER
   - Line items (Gross rental income, Manager's salary, etc.) → Row_Type: DETAIL
   - Total rows (Total revenues, Total expenses, etc.) → Row_Type: ROLLUP

4. Keep all numbers, text, and formatting characters (parentheses for negative numbers)

5. Use commas to separate columns

6. Put values with commas inside quotes

7. Include column headers if present

8. CRITICAL: Look for notes, annotations, or text outside/beside the main table columns:
   - If you see a "NOTES AND ASSUMPTIONS" section or numbered notes on the side, create a "Notes" column as the LAST column
   - Add the full text of each note to its corresponding row ONLY if the note specifically references that row
   - If a note is general context (not tied to a specific row), leave the Notes column empty for that row

9. Return ONLY the CSV data, no explanation

IMPORTANT:
- Do NOT include marginal note references (like "Note 14.") as table columns or data
- Do NOT create separate columns for categories vs line items - use ONE "Category" column for all text
- Do NOT skip breakdown items or sub-categories. Every line item visible in the table must appear in the output.
- Do NOT skip total/rollup rows. These are CRITICAL and must include their full labels with all numbers.
- Clearly mark which rows are HEADER, DETAIL, or ROLLUP using the Row_Type column.

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

CRITICAL ACCURACY REQUIREMENTS:
- Read each character VERY CAREFULLY - verify every letter and digit
- Pay special attention to similar-looking characters: 6 vs 8, O vs 0, l vs I, etc.
- Double-check all numbers for accuracy - transcription errors are NOT acceptable
- Verify text spelling character-by-character - do not guess or autocorrect
- If text is unclear, examine it closely before transcribing

Requirements:
1. IDENTIFY THE TABLE STRUCTURE:
   - Ignore marginal note references (like "Note 14.", "Note 3.", etc.) that appear in the left margin - these are NOT part of the table columns
   - Focus on the actual table columns that contain line items/categories and numeric values
   - The main table structure has: A single column for all categories and line items, followed by numeric columns (years, amounts)
   - CRITICAL: Watch for MULTIPLE SUB-COLUMNS per year/period:
     * Some tables have 2+ columns under each year header (e.g., percentage + amount, budget + actual, quantity + price)
     * Each sub-column MUST be a separate column in the CSV output
     * Create descriptive column names that identify BOTH the period AND the type
     * Examples: "2022_Percent,2022_Amount" or "2023_Budget,2023_Actual" or "Q1_Units,Q1_Price"
     * Look for sub-headers, data patterns, or $ signs to identify column types
     * If no sub-header exists, use descriptive names based on the data (e.g., "2022_Col1", "2022_Col2")

2. OUTPUT STRUCTURE:
   - Add a "Row_Type" column as the FIRST column to indicate the type of each row:
     * Use "HEADER" for section/category headers (e.g., "REVENUES", "EXPENSES", "Administrative Expenses", "Utility Expenses")
     * Use "DETAIL" for individual line items (e.g., "Gross rental income", "Manager's salary", "Electricity")
     * Use "ROLLUP" for total rows (e.g., "Total revenues", "Total expenses", "Net Operating Income")

   - Add a "Category" column as the SECOND column containing:
     * For HEADER rows: The section/category name (e.g., "REVENUES", "Administrative Expenses")
     * For DETAIL rows: The line item name (e.g., "Gross rental income", "Manager's salary")
     * For ROLLUP rows: The total label (e.g., "Total revenues", "Total expenses")

   - DO NOT create separate columns for categories and line items - everything goes in the single "Category" column

   - Follow with the numeric data columns (e.g., "2020", "2019")

3. Preserve all rows exactly as they appear:
   - Section headers (REVENUES, EXPENSES, etc.) → Row_Type: HEADER
   - Category headers (Administrative Expenses, Utility Expenses, etc.) → Row_Type: HEADER
   - Line items (Gross rental income, Manager's salary, etc.) → Row_Type: DETAIL
   - Total rows (Total revenues, Total expenses, etc.) → Row_Type: ROLLUP

4. Keep all numbers, text, and formatting characters (parentheses for negative numbers)

5. Use commas to separate columns

6. Put values with commas inside quotes

7. Include column headers if present

8. CRITICAL: Look for notes, annotations, or text outside/beside the main table columns:
   - If you see a "NOTES AND ASSUMPTIONS" section or numbered notes on the side, create a "Notes" column as the LAST column
   - Add the full text of each note to its corresponding row ONLY if the note specifically references that row
   - If a note is general context (not tied to a specific row), leave the Notes column empty for that row

9. Return ONLY the CSV data, no explanation

IMPORTANT:
- Do NOT include marginal note references (like "Note 14.") as table columns or data
- Do NOT create separate columns for categories vs line items - use ONE "Category" column for all text
- Do NOT skip breakdown items or sub-categories. Every line item visible in the table must appear in the output.
- Do NOT skip total/rollup rows. These are CRITICAL and must include their full labels with all numbers.
- Clearly mark which rows are HEADER, DETAIL, or ROLLUP using the Row_Type column.

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
