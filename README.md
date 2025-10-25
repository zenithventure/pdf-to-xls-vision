# PDF to XLS Vision

An intelligent Python library to convert PDF files containing tables into Excel (XLSX) files using Claude Vision API with automatic rotation detection. Each table found in the PDF becomes a separate sheet in the output Excel file.

## Features

- **Automatic PDF type detection** - Intelligently detects text-based vs image-based PDFs
- **Rotation detection & correction** - Automatically detects and corrects rotated pages (90°, 180°, 270°)
- **Dual extraction modes:**
  - Text-based PDFs: Fast, direct extraction (free, no API needed)
  - Image-based PDFs: Claude Vision API with superior accuracy
- **Quality validation** - Automatically detects poor extraction quality and retries with Vision API
- **Incremental saving** - Saves progress every 10 pages for large PDFs
- **Batch processing** - Process entire directories with recursive scanning
- **Python library & CLI** - Use as a library in your code or as a command-line tool

## Requirements

- Python 3.7+
- Anthropic API key (for image-based PDFs)

## Installation

### Install as a Python package

```bash
# Clone the repository
git clone https://github.com/yourusername/pdf-to-xls-vision.git
cd pdf-to-xls-vision

# Install in development mode
pip install -e .

# Or install from PyPI (when published)
pip install pdf-to-xls-vision
```

### Configuration

Set up your configuration:

1. Copy `.env.sample` to `.env`:
   ```bash
   cp .env.sample .env
   ```

2. Get your API key from: https://console.anthropic.com/

3. Edit the `.env` file and replace `your-api-key-here` with your actual API key:
   ```
   ANTHROPIC_API_KEY=sk-ant-your-actual-key-here
   ```

4. (Optional) Choose a different Claude model:
   ```
   CLAUDE_MODEL=claude-sonnet-4-5-20250929
   ```

   Available models:
   - `claude-sonnet-4-5-20250929` (default, most accurate)
   - `claude-3-5-sonnet-20241022` (fast, cost-effective)
   - `claude-3-5-sonnet-20240620` (balanced)
   - `claude-3-opus-20240229` (highest quality, slower)

## Usage

### As a Python Library

```python
from pdf_to_xls import convert_pdf_to_excel, batch_convert_directory

# Convert a single PDF
convert_pdf_to_excel('input.pdf', output_path='output.xlsx')

# Batch convert a directory
batch_convert_directory('pdfs/', output_dir='excel_files/', recursive=True)

# Force Vision API for complex tables
convert_pdf_to_excel('complex_table.pdf', force_vision=True)

# Use custom API key and model
convert_pdf_to_excel(
    'input.pdf',
    api_key='your-api-key',
    model_name='claude-3-5-sonnet-20241022'
)
```

See the [examples/](examples/) directory for more usage examples:
- [basic_usage.py](examples/basic_usage.py) - Simple conversion examples
- [batch_processing.py](examples/batch_processing.py) - Batch processing examples
- [advanced_usage.py](examples/advanced_usage.py) - Advanced features and error handling

### As a Command-Line Tool

After installation, you can use the `pdf-to-xls` command:

#### Convert a Single PDF File

```bash
pdf-to-xls input.pdf
```

Output will be saved as `input.xlsx` in the same directory.

#### Specify Output Path

```bash
pdf-to-xls input.pdf -o output.xlsx
```

#### Convert All PDFs in a Directory

```bash
pdf-to-xls /path/to/pdfs
```

#### Batch Convert with Recursive Scanning

```bash
pdf-to-xls /path/to/pdfs -r -o /path/to/output
```

#### Force Vision API

```bash
pdf-to-xls input.pdf --force-vision
```

### CLI Examples

Convert all PDFs in a directory:
```bash
pdf-to-xls "pdfs/OpStmts" -r
```

Convert a single file:
```bash
pdf-to-xls "pdfs/OpStmts/1206.pdf"
```

## How It Works

1. **Detection Phase**: Analyzes the PDF to determine if it's text-based or image-based
2. **Text-based PDFs**: Uses fast, free pdfplumber extraction with quality validation
3. **Image-based PDFs**:
   - Converts each page to an image
   - Detects rotation using Tesseract OSD
   - Corrects rotation if needed
   - Extracts tables using Claude Vision API
   - Saves progress every 10 pages
4. **Quality Check**: If text extraction has quality issues, automatically retries with Vision API
5. **Output**: Creates an Excel file with one sheet per page/table

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

## API Reference

### Main Functions

#### `convert_pdf_to_excel(pdf_path, output_path=None, output_dir=None, save_every=10, force_vision=False, api_key=None, model_name=None)`

Convert a single PDF to Excel.

**Parameters:**
- `pdf_path` (str|Path): Path to PDF file
- `output_path` (str|Path, optional): Output Excel file path
- `output_dir` (str|Path, optional): Output directory
- `save_every` (int): Save progress every N pages (default: 10)
- `force_vision` (bool): Force Vision API even for text PDFs (default: False)
- `api_key` (str, optional): Anthropic API key (uses env var if not provided)
- `model_name` (str, optional): Claude model name (uses env var if not provided)

**Returns:** Path to created Excel file, or None if no tables found

**Raises:**
- `FileNotFoundError`: If PDF file does not exist
- `ValueError`: If API key is required but not found

#### `batch_convert_directory(input_dir, output_dir=None, recursive=False, force_vision=False, api_key=None, model_name=None)`

Batch convert PDFs in a directory.

**Parameters:**
- `input_dir` (str|Path): Directory containing PDF files
- `output_dir` (str|Path, optional): Output directory
- `recursive` (bool): Recursively search subdirectories (default: False)
- `force_vision` (bool): Force Vision API for all PDFs (default: False)
- `api_key` (str, optional): Anthropic API key
- `model_name` (str, optional): Claude model name

**Returns:** Dictionary with 'success' and 'failed' lists of file paths

**Raises:**
- `FileNotFoundError`: If input directory does not exist

### Utility Functions

#### `pdf_is_image_based(pdf_path)`

Check if PDF is image-based (contains images).

**Parameters:**
- `pdf_path` (str|Path): Path to PDF file

**Returns:** bool - True if PDF is image-based

#### `pdf_has_text(pdf_path)`

Check if PDF has extractable text.

**Parameters:**
- `pdf_path` (str|Path): Path to PDF file

**Returns:** bool - True if PDF has extractable text

#### `detect_quality_issues(table_data)`

Detect quality issues in extracted table data.

**Parameters:**
- `table_data`: DataFrame or raw table data

**Returns:** list - List of quality issue descriptions

## Cost Information

- **Text-based PDFs**: Free (no API calls)
- **Image-based PDFs**: ~$0.01-0.05 per page with Claude Vision API
- The tool automatically chooses the most cost-effective method

## Troubleshooting

**PDF not converting properly?**
- The tool automatically detects and uses the best method
- Check that your `.env` file has a valid API key for image-based PDFs
- Make sure the PDF isn't password-protected
- Try `--force-vision` flag for complex table layouts

**Process taking too long?**
- Large image-based PDFs (30+ pages) may take 15-25 minutes
- Progress is saved every 10 pages
- Check for incremental save messages

**Rotation issues?**
- Rotation detection requires Tesseract OCR to be installed
- Install via: `brew install tesseract` (Mac) or `apt-get install tesseract-ocr` (Linux)

**Import errors?**
- Make sure you installed the package: `pip install -e .`
- Check that all dependencies are installed: `pip install -r requirements.txt`

## Development

### Project Structure

```
pdf-to-xls-vision/
├── pdf_to_xls/              # Main package
│   ├── __init__.py         # Public API
│   ├── config.py           # Configuration management
│   ├── converter.py        # Main conversion functions
│   ├── data_cleaning.py    # Data cleaning utilities
│   ├── excel_writer.py     # Excel generation
│   ├── image_processing.py # Image conversion and rotation
│   ├── pdf_detection.py    # PDF type detection
│   ├── quality_check.py    # Quality validation
│   └── table_extraction.py # Table extraction (vision & text)
├── examples/               # Usage examples
│   ├── basic_usage.py
│   ├── batch_processing.py
│   └── advanced_usage.py
├── pdf_to_xls_cli.py      # CLI entry point
├── setup.py               # Package setup
├── pyproject.toml         # Modern Python packaging
├── requirements.txt       # Dependencies
├── README.md             # This file
└── LICENSE               # License file
```

### Running Tests

```bash
# Install development dependencies
pip install -e ".[dev]"

# Run tests (when test suite is added)
pytest
```

## Technical Details

- Uses `pdfplumber` for text extraction
- Uses `pytesseract` for rotation detection
- Uses Claude Vision API for image-based extraction
- Uses `openpyxl` for Excel file generation
- Supports incremental saving for large files
- Automatic quality validation and retry logic

## License

MIT License - see [LICENSE](LICENSE) file for details

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Changelog

### Version 1.0.0
- Initial release with library structure
- Modular package design
- Python library API
- Command-line interface
- Automatic PDF type detection
- Vision API with rotation correction
- Quality validation and auto-retry
- Batch processing support
- Comprehensive examples
