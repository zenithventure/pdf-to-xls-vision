#!/usr/bin/env python3
"""Command-line interface for PDF to XLS converter."""

import sys
import argparse
from pathlib import Path

from pdf_to_xls import convert_pdf_to_excel, batch_convert_directory
from pdf_to_xls.config import get_api_key


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="PDF/Image to Excel converter with auto-detection (text extraction or Vision API with rotation)"
    )
    parser.add_argument("input", help="PDF/image file or directory (supports .pdf, .jpg, .jpeg, .png, .tiff, .tif)")
    parser.add_argument("-o", "--output", help="Output file or directory")
    parser.add_argument(
        "-r", "--recursive",
        action="store_true",
        help="Recursively search subdirectories"
    )
    parser.add_argument(
        "--force-vision",
        action="store_true",
        help="Force Vision API extraction even for text-based PDFs (useful for complex table layouts)"
    )

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
            convert_pdf_to_excel(input_path, output_path=args.output, force_vision=args.force_vision)
        elif input_path.is_dir():
            batch_convert_directory(
                input_path,
                output_dir=args.output,
                recursive=args.recursive,
                force_vision=args.force_vision
            )
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
