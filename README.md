# EPUB/PDF to Markdown Converter

A simple Python script to convert EPUB and PDF files to Markdown format.

## Features

- **Unified Converter**: Handles both EPUB and PDF files in a single script
- **Batch Processing**: Converts all EPUB and PDF files in the current directory
- **Simple Output**: Creates .md files with the same base filename as the source
- **No Chunking**: Outputs complete files (no splitting into chunks)
- **Metadata Extraction**: Includes title and author information in the output

## Installation

### Requirements

Install the required Python packages:

```bash
# For EPUB support
pip install ebooklib html2text

# For PDF support
pip install pymupdf

# Or install both
pip install ebooklib html2text pymupdf
```

## Usage

### Method 1: Copy the script to your folder

1. Copy `convert_to_md.py` to the folder containing your EPUB/PDF files
2. Open a command prompt in that folder
3. Run: `python convert_to_md.py`

### Method 2: Run from this directory

1. Place your EPUB/PDF files in this directory
2. Double-click `convert.bat` (Windows) or run `python convert_to_md.py`

### What it does

- Scans the current directory for all `.epub` and `.pdf` files
- Converts each file to Markdown format
- Outputs files with the same name but `.md` extension
- Example: `MyBook.epub` → `MyBook.md`

## Output Format

Each converted file includes:
- Title (from metadata or filename)
- Author (from metadata)
- Full content in Markdown format

## Notes

- The script will skip file types it doesn't have libraries for (will show warnings)
- Original files are not modified or deleted
- Conversion happens in the current working directory
- Special characters in filenames are handled automatically

## Legacy Scripts

This folder also contains older scripts with chunking functionality:
- `epub_to_md.py` - EPUB to MD (no chunking)
- `epub_to_chunks_complete.py` - EPUB to MD with chunking
- `pdf_to_chunks_complete.py` - PDF to MD with chunking

## Troubleshooting

### "No conversion libraries available"
Install the required packages using pip (see Installation section above)

### "No EPUB or PDF files found"
Make sure you're running the script in a directory that contains .epub or .pdf files

### Permission errors
Make sure you have write permissions in the directory
