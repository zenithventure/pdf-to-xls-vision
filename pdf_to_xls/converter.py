"""Main PDF to Excel conversion functions."""

from pathlib import Path
import anthropic

from .config import get_api_key, get_model_name
from .pdf_detection import pdf_is_image_based
from .table_extraction import extract_table_with_claude_vision, extract_tables_from_text_pdf, extract_table_from_image
from .excel_writer import create_excel_file, merge_continuation_tables
from .validation import validate_extracted_data


# Supported image file extensions
IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.tiff', '.tif'}


def is_image_file(file_path):
    """Check if file is a supported image format.

    Args:
        file_path: Path to file (str or Path)

    Returns:
        bool: True if file has a supported image extension
    """
    file_path = Path(file_path)
    return file_path.suffix.lower() in IMAGE_EXTENSIONS


def convert_pdf_to_excel(
    pdf_path,
    output_path=None,
    output_dir=None,
    save_every=10,
    force_vision=False,
    api_key=None,
    model_name=None
):
    """Convert PDF or image file to Excel file.

    Uses text extraction for text-based PDFs, Vision API with rotation detection
    for image-based PDFs and image files (.jpg, .jpeg, .png, .tiff, .tif).

    Args:
        pdf_path: Path to PDF or image file (str or Path)
        output_path: Optional output Excel file path (str or Path)
        output_dir: Optional output directory (str or Path)
        save_every: For large PDFs, save progress every N pages (default: 10, ignored for images)
        force_vision: Force Vision API extraction even for text-based PDFs (default: False)
        api_key: Optional Anthropic API key (uses env var if not provided)
        model_name: Optional Claude model name (uses env var if not provided)

    Returns:
        Path: Path to the created Excel file, or None if no tables found

    Raises:
        FileNotFoundError: If file does not exist
        ValueError: If API key is required but not found
    """
    pdf_path = Path(pdf_path)

    if not pdf_path.exists():
        raise FileNotFoundError(f"File not found: {pdf_path}")

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
        # Check if input is an image file
        if is_image_file(pdf_path):
            # Image file: use Vision API directly
            print("  Image file detected, using Vision API...")

            # Get API configuration
            if not api_key:
                api_key = get_api_key()
            if not model_name:
                model_name = get_model_name()

            client = anthropic.Anthropic(api_key=api_key)
            tables = extract_table_from_image(pdf_path, client, model_name)
        else:
            # PDF file: check if it's image-based or text-based
            is_image_based = pdf_is_image_based(pdf_path)

            if force_vision or is_image_based:
                # Image-based PDF or forced Vision API: use Vision API with rotation detection
                if force_vision and not is_image_based:
                    print("  Text-based PDF, using Vision API (forced)...")
                else:
                    print("  Image-based PDF detected, using Vision API with rotation detection...")

                # Get API configuration
                if not api_key:
                    api_key = get_api_key()
                if not model_name:
                    model_name = get_model_name()

                client = anthropic.Anthropic(api_key=api_key)
                tables = extract_table_with_claude_vision(pdf_path, client, model_name, output_path, save_every)
            else:
                # Text-based PDF: use direct extraction (fast, no API needed)
                print("  Text-based PDF, using direct extraction...")
                tables, quality_issues_detected = extract_tables_from_text_pdf(pdf_path)

                # Auto-retry with Vision API if quality issues detected OR no tables found
                if quality_issues_detected or not tables:
                    if quality_issues_detected:
                        print("\n  ⚠️  Quality issues detected in text extraction!")
                        print("  🔄 Retrying with Vision API for better accuracy...\n")
                    else:
                        print("\n  ⚠️  No tables found with text extraction!")
                        print("  🔄 Retrying with Vision API...\n")

                    # Get API configuration
                    if not api_key:
                        api_key = get_api_key()
                    if not model_name:
                        model_name = get_model_name()

                    client = anthropic.Anthropic(api_key=api_key)
                    tables = extract_table_with_claude_vision(pdf_path, client, model_name, output_path, save_every)

        if not tables:
            print(f"Warning: No tables found in {pdf_path}")
            return None

        # Merge continuation tables (tables that span multiple pages)
        tables = merge_continuation_tables(tables)

        # Validate extracted data (only for PDF files, not images)
        if not is_image_file(pdf_path):
            print("\nValidating extracted data...")
            validation_report_path = output_path.parent / f"{output_path.stem}_validation.md"
            validation_result = validate_extracted_data(pdf_path, tables, validation_report_path)

            if validation_result['status'] == 'completed':
                accuracy = validation_result['statistics']['accuracy_percent']
                if accuracy >= 95:
                    print(f"  ✓ Validation complete: {accuracy:.2f}% accuracy")
                else:
                    print(f"  ⚠️  Validation complete: {accuracy:.2f}% accuracy - please review report")

                discrepancies = validation_result['discrepancies']
                if discrepancies['missing_in_tables'] or discrepancies['extra_in_tables']:
                    print(f"  ⚠️  Found discrepancies - see validation report for details")

        # Create Excel file
        return create_excel_file(tables, output_path)

    except Exception as e:
        print(f"Error converting {pdf_path}: {e}")
        import traceback
        traceback.print_exc()
        raise


def batch_convert_directory(
    input_dir,
    output_dir=None,
    recursive=False,
    force_vision=False,
    api_key=None,
    model_name=None
):
    """Batch convert PDFs and image files in directory.

    Auto-detects text vs image-based PDFs. Processes all supported image formats
    (.jpg, .jpeg, .png, .tiff, .tif) using Vision API.

    Args:
        input_dir: Directory containing PDF and/or image files (str or Path)
        output_dir: Optional output directory (str or Path)
        recursive: Recursively search subdirectories (default: False)
        force_vision: Force Vision API extraction for all PDFs (default: False)
        api_key: Optional Anthropic API key (uses env var if not provided)
        model_name: Optional Claude model name (uses env var if not provided)

    Returns:
        dict: Dictionary with 'success' and 'failed' lists of file paths

    Raises:
        FileNotFoundError: If input directory does not exist
    """
    input_dir = Path(input_dir)

    if not input_dir.exists():
        raise FileNotFoundError(f"Directory not found: {input_dir}")

    # Find all PDF files
    if recursive:
        pdf_files = list(input_dir.rglob("*.pdf"))
    else:
        pdf_files = list(input_dir.glob("*.pdf"))

    # Find all image files
    image_files = []
    for ext in IMAGE_EXTENSIONS:
        if recursive:
            image_files.extend(input_dir.rglob(f"*{ext}"))
        else:
            image_files.extend(input_dir.glob(f"*{ext}"))

    # Combine all files
    all_files = pdf_files + image_files

    # Filter out zone identifier files
    all_files = [f for f in all_files if ':Zone.Identifier' not in str(f)]

    if not all_files:
        print(f"No PDF or image files found in {input_dir}")
        return {'success': [], 'failed': []}

    print(f"Found {len(all_files)} file(s) ({len(pdf_files)} PDF, {len(image_files)} image)")
    print("=" * 70)

    success_list = []
    failed_list = []

    for file_path in all_files:
        try:
            if output_dir and recursive:
                rel_path = file_path.relative_to(input_dir)
                out_dir = Path(output_dir) / rel_path.parent
            else:
                out_dir = output_dir or file_path.parent

            result = convert_pdf_to_excel(
                file_path,
                output_dir=out_dir,
                force_vision=force_vision,
                api_key=api_key,
                model_name=model_name
            )
            if result:
                success_list.append(file_path)
            print("=" * 70)

        except Exception as e:
            print(f"Failed to convert {file_path}: {e}")
            failed_list.append(file_path)
            print("=" * 70)

    print(f"\n✓ Conversion complete!")
    print(f"  Successful: {len(success_list)}/{len(all_files)}")

    if failed_list:
        print(f"\n✗ Failed files ({len(failed_list)}):")
        for f in failed_list:
            print(f"  - {f}")

    return {'success': success_list, 'failed': failed_list}
