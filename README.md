# PDF to XLS Converter

An intelligent Python utility to convert PDF files containing tables into Excel (XLSX) files using Claude Vision API with automatic rotation detection. Each table found in the PDF becomes a separate sheet in the output Excel file.

## Features

- **Automatic PDF type detection** - Intelligently detects text-based vs image-based PDFs
- **Rotation detection & correction** - Automatically detects and corrects rotated pages (90°, 180°, 270°)
- **Dual extraction modes:**
  - Text-based PDFs: Fast, direct extraction (free, no API needed)
  - Image-based PDFs: Claude Vision API with superior accuracy
- **Incremental saving** - Saves progress every 10 pages for large PDFs
- **Batch processing** - Process entire directories
- **Recursive scanning** - Search subdirectories

## Requirements

- Python 3.7+
- Anthropic API key (for image-based PDFs)

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Set up your Anthropic API key:
   - Get your API key from: https://console.anthropic.com/
   - Edit the `.env` file and replace `your-api-key-here` with your actual API key:
     ```
     ANTHROPIC_API_KEY=sk-ant-your-actual-key-here
     ```

## Usage

### Convert a Single PDF File

```bash
python3 pdf_to_xls_vision.py input.pdf
```

Output will be saved as `input.xlsx` in the same directory.

### Specify Output Path

```bash
python3 pdf_to_xls_vision.py input.pdf -o output.xlsx
```

### Convert All PDFs in a Directory

```bash
python3 pdf_to_xls_vision.py /path/to/pdfs
```

### Batch Convert with Recursive Scanning

```bash
python3 pdf_to_xls_vision.py /path/to/pdfs -r -o /path/to/output
```

## Examples

Convert all PDFs in a directory:
```bash
python3 pdf_to_xls_vision.py "pdfs/Op Stmts from Rob 10.14.2025" -r
```

Convert a single file:
```bash
python3 pdf_to_xls_vision.py "pdfs/Op Stmts from Rob 10.14.2025/Meridian TTM 1206.pdf"
```

## How It Works

1. **Detection Phase**: Analyzes the PDF to determine if it's text-based or image-based
2. **Text-based PDFs**: Uses fast, free pdfplumber extraction
3. **Image-based PDFs**:
   - Converts each page to an image
   - Detects rotation using Tesseract OSD
   - Corrects rotation if needed
   - Extracts tables using Claude Vision API
   - Saves progress every 10 pages
4. **Output**: Creates an Excel file with one sheet per page/table

## Rotation Detection

The converter automatically detects and corrects rotated pages:
- Supports 90°, 180°, 270° rotations
- Uses Tesseract OSD (Orientation and Script Detection)
- Only corrects when confidence > 1.0
- Logs each rotation correction

Example output:
```
Processing page 2/31 with Claude Vision...
  Detected rotation 270° (confidence: 5.3) - correcting
  ✓ Extracted table: 23 rows x 15 columns
```

## Large PDF Support

For PDFs with 30+ pages:
- Progress is saved incrementally every 10 pages
- If interrupted, partial results are preserved
- Visual progress indicators show completion status

Example:
```
Processing page 10/31...
💾 Saving progress... (10/31 pages processed)
✓ Progress saved: 10 tables
```

## Cost Information

- **Text-based PDFs**: Free (no API calls)
- **Image-based PDFs**: ~$0.01-0.05 per page
- The tool automatically chooses the most cost-effective method

## Troubleshooting

**PDF not converting properly?**
- The tool automatically detects and uses the best method
- Check that your `.env` file has a valid API key for image-based PDFs
- Make sure the PDF isn't password-protected

**Process taking too long?**
- Large image-based PDFs (30+ pages) may take 15-25 minutes
- Progress is saved every 10 pages
- Check for incremental save messages

**Rotation issues?**
- Rotation detection requires Tesseract OCR to be installed
- Install via: `brew install tesseract` (Mac) or `apt-get install tesseract-ocr` (Linux)

## Technical Details

- Uses `pdfplumber` for text extraction
- Uses `pytesseract` for rotation detection
- Uses Claude Vision API (Sonnet 4.5) for image-based extraction
- Uses `openpyxl` for Excel file generation
- Supports incremental saving for large files
