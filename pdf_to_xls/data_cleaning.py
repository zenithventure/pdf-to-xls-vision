"""Data cleaning utilities for extracted table data."""

import re
import pandas as pd


def _fix_cell_parens(value):
    """Fix common parenthesis issues in a single cell.

    Handles:
    - Spaces inside parens: "( 297)" -> "(297)"
    - Duplicate opening parens: "((123)" -> "(123)"
    - Missing closing paren: "( 4410" -> "(4410)"
    - Orphaned closing paren: "123)" -> "(123)"

    Args:
        value: Cell value to fix

    Returns:
        Fixed cell value
    """
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

    Args:
        df: pandas DataFrame to clean

    Returns:
        pandas DataFrame with cleaned data
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

    Args:
        df: pandas DataFrame to clean

    Returns:
        pandas DataFrame with cleaned data
    """
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
        df[col] = df[col].apply(
            lambda x: re.sub(r'(%)\s*\($', r'\1', str(x).strip())
            if pd.notna(x) and isinstance(x, str) else x
        )

    return df
