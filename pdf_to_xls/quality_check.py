"""Quality checking utilities for extracted table data."""

import re
import pandas as pd


def detect_quality_issues(table_data):
    """Detect signs of poor table extraction quality.

    Returns a list of detected quality issues (empty list if no issues).

    Quality heuristics:
    1. Single column trap: Flag if table has only 1 column with multiple rows
    2. Row explosion: Flag if row count is suspiciously high (>50 rows with other issues)
    3. Column consistency: Flag if >30% of rows have different column counts
    4. Empty cells ratio: Flag if >50% of cells are empty
    5. Duplicate rows: Flag if >20% of rows are duplicates
    6. Garbled text: Flag if >10% of cells contain encoding/parsing errors

    Args:
        table_data: DataFrame or raw table list to check

    Returns:
        list: List of quality issue descriptions (empty if no issues)
    """
    issues = []

    if table_data is None or (isinstance(table_data, list) and len(table_data) == 0):
        return issues

    # Convert table_data to dataframe if it's a raw table
    if isinstance(table_data, list):
        try:
            df = pd.DataFrame(table_data[1:], columns=table_data[0])
        except Exception:
            return ["Could not parse table data"]
    else:
        df = table_data

    if df.empty:
        return issues

    num_rows = len(df)
    num_cols = len(df.columns)

    # Heuristic 1: Single column trap with multiple rows (likely parsing failure)
    # Single column is only suspicious if there are multiple rows (>3)
    if num_cols == 1 and num_rows > 3:
        issues.append(f"Single column table with {num_rows} rows (likely parsing error)")

    # Heuristic 2: Row explosion (suspiciously high row count)
    # More nuanced: only flag if BOTH high row count AND other suspicious patterns
    if num_rows > 70:
        # Very high threshold - definitely suspicious
        issues.append(f"Excessive row count ({num_rows} rows, likely incorrect parsing)")
    elif num_rows > 50:
        # Medium threshold - only flag if combined with narrow columns or high emptiness
        # Check if columns are suspiciously narrow (suggests over-splitting)
        if num_cols > 12:
            issues.append(
                f"Excessive row count ({num_rows} rows) with many columns ({num_cols}), "
                "likely incorrect parsing"
            )

    # Heuristic 3: Check column count consistency across rows
    # Count non-null values in each row as a proxy for effective column count
    row_col_counts = df.notna().sum(axis=1)
    if len(row_col_counts) > 0:
        most_common_count = row_col_counts.mode()[0] if len(row_col_counts.mode()) > 0 else num_cols
        inconsistent_rows = (row_col_counts != most_common_count).sum()
        inconsistency_ratio = inconsistent_rows / len(row_col_counts)

        if inconsistency_ratio > 0.3:
            issues.append(f"Inconsistent column counts ({inconsistency_ratio:.1%} of rows differ)")

    # Heuristic 4: Empty cells ratio
    # Only flag if table is large and mostly empty (suggests phantom rows/columns)
    total_cells = num_rows * num_cols
    empty_cells = df.isna().sum().sum()
    empty_ratio = empty_cells / total_cells if total_cells > 0 else 0

    # Stricter threshold for smaller tables, looser for larger ones
    empty_threshold = 0.6 if num_rows < 20 else 0.5
    if empty_ratio > empty_threshold:
        issues.append(f"High empty cell ratio ({empty_ratio:.1%} empty cells)")

    # Heuristic 5: Duplicate rows detection
    # pdfplumber often repeats rows when parsing fails
    if num_rows > 5:  # Only check tables with enough rows
        # Convert to string representation for comparison (handles NaN properly)
        df_str = df.astype(str)
        duplicated_rows = df_str.duplicated(keep='first').sum()
        duplicate_ratio = duplicated_rows / num_rows if num_rows > 0 else 0

        if duplicate_ratio > 0.2:  # More than 20% duplicates
            issues.append(
                f"High duplicate row ratio ({duplicate_ratio:.1%} of rows are duplicates, "
                f"{duplicated_rows}/{num_rows} rows)"
            )

    # Heuristic 6: Check for garbled text patterns (encoding issues, random characters)
    # Look for cells with unusual character patterns that suggest OCR/parsing errors
    garbled_count = 0
    sample_size = min(100, total_cells)  # Sample up to 100 cells
    cells_checked = 0

    for col in df.columns:
        for val in df[col].head(20):  # Check first 20 values per column
            if pd.notna(val) and isinstance(val, str):
                cells_checked += 1
                # Check for: excessive special chars, mixed scripts, control chars
                # Non-printable chars
                if re.search(r'[^\x20-\x7E\u00A0-\u024F\u20A0-\u20CF]{3,}', str(val)):
                    garbled_count += 1
                # Excessive special chars
                elif len(val) > 5 and re.search(r'[^\w\s$,.%()\-\'/]{3,}', str(val)):
                    garbled_count += 1

            if cells_checked >= sample_size:
                break
        if cells_checked >= sample_size:
            break

    if cells_checked > 0 and garbled_count / cells_checked > 0.1:
        issues.append(f"Garbled text detected ({garbled_count}/{cells_checked} cells)")

    return issues
