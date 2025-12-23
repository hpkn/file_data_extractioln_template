# File Data Extraction

A Python tool for extracting text from various document formats including PDF, HWP (Korean Hangul), HWPX, Excel, Word, and ZIP archives.

## Supported Formats

| Format | Extensions | Library Used |
|--------|------------|--------------|
| PDF | `.pdf` | pdfplumber, PyPDF2 |
| HWP (Korean) | `.hwp` | olefile |
| HWPX (Korean) | `.hwpx` | xml parser |
| Excel | `.xlsx`, `.xls`, `.xlsm` | openpyxl, xlrd |
| Word | `.docx` | python-docx |
| ZIP | `.zip` | zipfile (built-in) |

## Requirements

- Python 3.8 or higher
- Windows/Linux/macOS

## Installation

### 1. Clone or Download the Project

```bash
cd file_data_extraction
```

### 2. Create Virtual Environment (Recommended)

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Linux/macOS
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

## Usage

### Command Line Interface

#### Extract Single File

```bash
# Basic extraction (prints summary to console)
python main.py document.pdf

# Save extracted text to a file
python main.py document.pdf -o output.txt

# Print extracted text to console
python main.py document.hwp --print-text

# Verbose mode for debugging
python main.py document.xlsx -v
```

#### Extract Multiple Files (Directory)

```bash
# Process all supported files in a directory
python main.py ./documents/

# Process directory and save outputs
python main.py ./documents/ -o ./extracted/

# Quiet mode (only show errors)
python main.py ./documents/ -q
```

### Command Line Options

| Option | Description |
|--------|-------------|
| `input` | Input file or directory path (required) |
| `-o, --output` | Output file (for single file) or directory (for batch) |
| `-v, --verbose` | Enable detailed logging |
| `-q, --quiet` | Suppress all output except errors |
| `--print-text` | Print extracted text to console |

### Python API

You can also use the extractor directly in your Python code:

```python
from pathlib import Path
from text_extraction import TextExtractor

# Initialize extractor
extractor = TextExtractor(data_lake_root=Path("."))

# Extract from a single file
result = extractor.extract_text(Path("document.pdf"))

if result.get('error'):
    print(f"Error: {result['error']}")
else:
    print(f"Extracted {result['char_count']} characters")
    print(result['text'])

# Access sections (pages, sheets, etc.)
for section in result.get('sections', []):
    print(f"Section: {section['section_title']}")
    print(f"Text: {section['text'][:100]}...")
```

## Project Structure

```
file_data_extraction/
├── main.py              # CLI entry point
├── text_extraction.py   # Main extraction module
├── hwp_extraction.py    # HWP format parser
├── hwpx_extraction.py   # HWPX format parser
├── requirements.txt     # Python dependencies
├── .gitignore          # Git ignore rules
└── README.md           # This file
```

## Output Format

The extractor returns a dictionary with the following structure:

```python
{
    'text': str,           # Full extracted text
    'method': str,         # Extraction method used
    'sections': [          # List of document sections
        {
            'section_type': str,    # 'page', 'sheet', 'document'
            'section_title': str,   # Section name
            'text': str,            # Section text
            'section_order': int,   # Order in document
            'word_count': int,      # Words in section
            'char_count': int       # Characters in section
        }
    ],
    'word_count': int,     # Total word count
    'char_count': int,     # Total character count
    'error': str           # Error message (if failed)
}
```

## Features

### PDF Extraction
- Uses pdfplumber for better table and layout handling
- Falls back to PyPDF2 if pdfplumber fails
- Extracts text page by page

### HWP/HWPX Extraction
- Detects encrypted files before processing
- Handles password-protected and DRM files gracefully
- Removes Chinese characters and control characters for clean output

### Excel Extraction
- Supports both `.xlsx` (openpyxl) and `.xls` (xlrd) formats
- Extracts all sheets with sheet names
- Handles corrupted cells gracefully

### Word Extraction
- Extracts paragraphs and tables
- Preserves table structure with tab-separated values

### ZIP Extraction
- Automatically extracts and processes inner files
- Supports nested documents of all supported formats
- Detects password-protected archives

## Error Handling

The extractor handles various error conditions:

- **Encrypted files**: Returns `encrypted: True` with encryption type
- **Corrupted files**: Returns error message with details
- **Unsupported formats**: Returns clear error message
- **Missing dependencies**: Warns and continues with available extractors

## Database Integration (Optional)

For ETL pipeline integration with PostgreSQL:

```python
import asyncio
import asyncpg
from pathlib import Path
from text_extraction import extract_and_save_text

async def run_pipeline():
    pool = await asyncpg.create_pool(
        host='localhost',
        database='your_db',
        user='your_user',
        password='your_password'
    )

    result = await extract_and_save_text(
        db_pool=pool,
        data_lake_root=Path('./data_lake'),
        limit=100  # Optional: limit number of documents
    )

    print(f"Processed: {result['extracted']}, Errors: {result['errors']}")

    await pool.close()

asyncio.run(run_pipeline())
```

## Troubleshooting

### Missing Dependencies

If you see import warnings, install the specific library:

```bash
# PDF support
pip install PyPDF2 pdfplumber

# HWP support
pip install olefile

# Excel support
pip install openpyxl xlrd

# Word support
pip install python-docx
```

### Encrypted Files

Encrypted files cannot be processed without the password. The extractor will:
1. Detect encryption before attempting extraction
2. Return an error with `encrypted: True`
3. Specify the encryption type when possible

### Large Files

For very large files or directories:
- Use `-q` flag to reduce console output
- Process in batches if memory is limited
- Consider increasing system memory for large PDFs

## License

MIT License
