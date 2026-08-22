# ebook2md

Convert EPUB, PDF, FB2, DJVU, DOC, and DOCX files to clean Markdown.

## Why

LLMs like Claude can process ebooks directly, but you pay for every token of raw EPUB/PDF markup, embedded metadata, and formatting overhead. By pre-converting to Markdown you feed the model only the actual text content, which significantly reduces token usage and cost. This script does that conversion locally, in batch, for free.

## Supported Formats

| Format | Library / Tool | Install |
|--------|---------------|---------|
| EPUB | `ebooklib` + `html2text` | `pip install ebooklib html2text` |
| PDF | `pymupdf` (PyMuPDF) | `pip install pymupdf` |
| FB2 | `lxml` | `pip install lxml` |
| DOCX | `python-docx` | `pip install python-docx` |
| DOC | `python-docx` + MS Word (or LibreOffice) | `pip install python-docx pywin32` |
| DJVU | DjVuLibre (`djvutxt`) | `winget install DjVuLibre.DjView` |

## Features

- **Batch Processing**: Converts all supported files in the current directory in one run
- **Graceful Degradation**: Missing libraries are detected at startup; unsupported formats are skipped with a warning
- **Metadata Extraction**: Title and author are pulled from file metadata (EPUB, PDF, FB2) or parsed from the filename (DJVU)
- **Markdown Cleanup**: Strips excessive whitespace, duplicate headers/footers, page numbers, and internal navigation links, without destroying numeric content (see [Page numbers and numeric content](#page-numbers-and-numeric-content))
- **Heading Detection**: Automatically promotes ALL-CAPS lines and numbered sections to Markdown headings (PDF, DJVU)
- **FB2 Structure Preservation**: Converts sections, epigraphs, citations, poems, emphasis, and strong markup to proper Markdown
- **DOCX Structure Preservation**: Converts headings, bold/italic, lists, block quotes, and tables to proper Markdown
- **DOC via MS Word or LibreOffice**: Legacy `.doc` files are converted to `.docx` via MS Word COM automation (preferred) or LibreOffice headless (fallback)
- **DJVU Non-ASCII Handling**: Transparently copies files with non-ASCII filenames to a temp path so `djvutxt` can process them
- **Simple Output**: Creates `.md` files with the same base filename as the source -- no chunking, no subdirectories

## Installation

Install only the libraries you need:

```bash
# All formats
pip install ebooklib html2text pymupdf lxml python-docx

# Or pick and choose
pip install ebooklib html2text   # EPUB
pip install pymupdf              # PDF
pip install lxml                 # FB2
pip install python-docx          # DOCX (and DOC with LibreOffice)
```

For legacy `.doc` support, install LibreOffice (so `soffice` is available):

```bash
winget install TheDocumentFoundation.LibreOffice
```

For DJVU support, install DjVuLibre so that `djvutxt` is on your PATH (or in `C:\Program Files\DjVuLibre\`):

```bash
winget install DjVuLibre.DjView
```

## Usage

### Method 1: Copy the script to your folder

1. Copy `ebook2md.py` to the folder containing your files
2. Open a terminal in that folder
3. Run: `python ebook2md.py`

### Method 2: Run from this directory

1. Place your files in this directory
2. Double-click `ebook2md.bat` (Windows) or run `python ebook2md.py`

### What it does

- Scans the current directory for `.epub`, `.pdf`, `.fb2`, `.djvu`, `.doc`, and `.docx` files
- Converts each file to Markdown
- Outputs files with the same name but `.md` extension
- Example: `MyBook.epub` -> `MyBook.md`

## Output Format

Each converted file includes:
- Title (from metadata or filename)
- Author (from metadata or filename pattern `Author - Title`)
- Full content in Markdown format

## Page numbers and numeric content

Removing running page numbers is easy to get wrong. Deleting every line that
contains only digits also deletes real values, because plenty of documents put
a bare number on its own line as content: lab results, statistical tables,
numbered lists, years.

This script only removes a bare integer when it can show the number is
pagination:

- decorated forms are always removed: `- 12 -`, `[12]`, `Page 12`,
  `Page 12 of 340`, `12 / 340`;
- a bare integer is removed only if it equals the number of the page it was
  found on **and** sits within the first or last couple of non-empty lines of
  that page, where running heads and feet live.

Anything else is kept.

If you want the old, indiscriminate behaviour, pass
`--aggressive-page-numbers`. It deletes every digits-only line and will remove
real data from any document where numbers are content. It is off by default
for that reason.

> Earlier versions applied the indiscriminate rule unconditionally. On a
> corpus of 18 lab-report PDFs it silently deleted 246 measurements, whole
> results among them. If you converted numeric documents with an older
> version, reconvert them.

## Limitations

This script extracts **existing text** from source files. It does not perform OCR. Scanned PDFs and image-only DJVUs that contain no embedded text layer will produce empty output and be skipped.

For scanned PDFs that require OCR, see [pdf2md](https://github.com/vkorost/pdf2md).

## Notes

- The script skips formats whose libraries are not installed (warnings are shown at startup)
- Original files are never modified or deleted
- Conversion runs in the current working directory
- Special characters in filenames are handled automatically
- DJVU files without a text layer (image-only scans) will be skipped with a warning

## Troubleshooting

### "No conversion libraries available"
Install at least one set of dependencies (see Installation above).

### "No supported files found"
Make sure the script is running in a directory that contains `.epub`, `.pdf`, `.fb2`, `.djvu`, `.doc`, or `.docx` files.

### DJVU: "No text layer found"
The DJVU file is an image-only scan with no embedded OCR text. You will need to OCR it first.

### Permission errors
Make sure you have write permissions in the directory.

## Building a Standalone Executable

You can package the script as a single `.exe` using PyInstaller so it runs on machines without Python installed.

### Prerequisites

```bash
pip install pyinstaller ebooklib html2text pymupdf lxml python-docx
```

### Build

```bash
pyinstaller --onefile --name ebook2md ebook2md.py
```

The executable will be in the `dist/` folder. DJVU support still requires DjVuLibre to be installed separately (the `.exe` calls `djvutxt` as an external process).
