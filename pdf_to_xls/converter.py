"""Main PDF to Excel conversion functions."""

from pathlib import Path
import anthropic

from .config import get_api_key, get_model_name
from .pdf_detection import pdf_is_image_based
from .table_extraction import extract_table_with_claude_vision, extract_tables_from_text_pdf
from .excel_writer import create_excel_file


def convert_pdf_to_excel(
    pdf_path,
    output_path=None,
    output_dir=None,
    save_every=10,
    force_vision=False,
    api_key=None,
    model_name=None
):
    """Convert PDF to Excel file.

    Uses text extraction for text-based PDFs, Vision API with rotation detection
    for image-based PDFs.

    Args:
        pdf_path: Path to PDF file (str or Path)
        output_path: Optional output Excel file path (str or Path)
        output_dir: Optional output directory (str or Path)
        save_every: For large PDFs, save progress every N pages (default: 10)
        force_vision: Force Vision API extraction even for text-based PDFs (default: False)
        api_key: Optional Anthropic API key (uses env var if not provided)
        model_name: Optional Claude model name (uses env var if not provided)

    Returns:
        Path: Path to the created Excel file, or None if no tables found

    Raises:
        FileNotFoundError: If PDF file does not exist
        ValueError: If API key is required but not found
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
    """Batch convert PDFs in directory.

    Auto-detects text vs image-based PDFs.

    Args:
        input_dir: Directory containing PDF files (str or Path)
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

    if recursive:
        pdf_files = list(input_dir.rglob("*.pdf"))
    else:
        pdf_files = list(input_dir.glob("*.pdf"))

    # Filter out zone identifier files
    pdf_files = [f for f in pdf_files if ':Zone.Identifier' not in str(f)]

    if not pdf_files:
        print(f"No PDF files found in {input_dir}")
        return {'success': [], 'failed': []}

    print(f"Found {len(pdf_files)} PDF file(s)")
    print("=" * 70)

    success_list = []
    failed_list = []

    for pdf_path in pdf_files:
        try:
            if output_dir and recursive:
                rel_path = pdf_path.relative_to(input_dir)
                out_dir = Path(output_dir) / rel_path.parent
            else:
                out_dir = output_dir or pdf_path.parent

            result = convert_pdf_to_excel(
                pdf_path,
                output_dir=out_dir,
                force_vision=force_vision,
                api_key=api_key,
                model_name=model_name
            )
            if result:
                success_list.append(pdf_path)
            print("=" * 70)

        except Exception as e:
            print(f"Failed to convert {pdf_path}: {e}")
            failed_list.append(pdf_path)
            print("=" * 70)

    print(f"\n✓ Conversion complete!")
    print(f"  Successful: {len(success_list)}/{len(pdf_files)}")

    if failed_list:
        print(f"\n✗ Failed files ({len(failed_list)}):")
        for f in failed_list:
            print(f"  - {f}")

    return {'success': success_list, 'failed': failed_list}
