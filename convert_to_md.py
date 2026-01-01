#!/usr/bin/env python3
"""
Unified EPUB and PDF to Markdown Converter
Converts all EPUB and PDF files in the current directory to Markdown format
"""

import os
import re
import sys
from pathlib import Path

# Try importing EPUB libraries
try:
    from ebooklib import epub
    from ebooklib import ITEM_DOCUMENT
    import html2text
    EPUB_AVAILABLE = True
except ImportError:
    EPUB_AVAILABLE = False
    print("WARNING: EPUB libraries not found. EPUB conversion will be skipped.")
    print("To enable EPUB support, run: pip install ebooklib html2text")

# Try importing PDF libraries
try:
    import pymupdf  # PyMuPDF
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False
    print("WARNING: PDF libraries not found. PDF conversion will be skipped.")
    print("To enable PDF support, run: pip install pymupdf")

def clean_filename(filename):
    """Clean filename for use as markdown file name"""
    # Remove file extension
    name = Path(filename).stem
    # Replace problematic characters
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    # Remove extra spaces and parentheses content for cleaner names
    name = re.sub(r'\s*\([^)]*\)', '', name)
    name = re.sub(r'\s+', '_', name.strip())
    return name + '.md'

def extract_text_from_epub(epub_path):
    """Extract text content from EPUB file"""
    if not EPUB_AVAILABLE:
        return None

    try:
        book = epub.read_epub(epub_path)

        # Get book metadata
        title = book.get_metadata('DC', 'title')[0][0] if book.get_metadata('DC', 'title') else "Unknown Title"
        author = book.get_metadata('DC', 'creator')[0][0] if book.get_metadata('DC', 'creator') else "Unknown Author"

        # Initialize HTML to text converter
        h = html2text.HTML2Text()
        h.ignore_links = False
        h.ignore_images = True
        h.ignore_emphasis = False
        h.body_width = 0  # No line wrapping

        content = []
        content.append(f"# {title}")
        content.append(f"**Author:** {author}")
        content.append("")

        # Extract text from all document items
        for item in book.get_items():
            if item.get_type() == ITEM_DOCUMENT:
                # Convert HTML to markdown
                html_content = item.get_content().decode('utf-8')
                markdown_text = h.handle(html_content)

                # Clean up the markdown
                markdown_text = re.sub(r'\n\s*\n\s*\n+', '\n\n', markdown_text)  # Remove excessive newlines
                markdown_text = re.sub(r'^\s+|\s+$', '', markdown_text, flags=re.MULTILINE)  # Trim whitespace

                # Remove table of contents links and improve formatting
                markdown_text = re.sub(r'\[.*?\]\(.*?\.html.*?\)', '', markdown_text)  # Remove internal links
                markdown_text = re.sub(r'\n\s*\d+\.\s*$', '', markdown_text, flags=re.MULTILINE)  # Remove hanging numbers

                if markdown_text.strip():
                    content.append(markdown_text)

        return '\n\n'.join(content)

    except Exception as e:
        print(f"  [ERROR] Failed to process EPUB: {str(e)}")
        return None

def extract_text_from_pdf(pdf_path):
    """Extract text content from PDF file"""
    if not PDF_AVAILABLE:
        return None

    try:
        doc = pymupdf.open(pdf_path)

        # Try to get metadata
        metadata = doc.metadata
        title = metadata.get('title', pdf_path.stem)
        author = metadata.get('author', 'Unknown Author')

        # If title is empty or just the filename, use the filename
        if not title or title == pdf_path.stem:
            title = pdf_path.stem.replace('_', ' ')

        content = []
        content.append(f"# {title}")
        content.append(f"**Author:** {author}")
        content.append("")

        # Extract text from all pages
        full_text = ""
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text()
            if text.strip():
                full_text += text + "\n\n"

        doc.close()

        # Clean up the text
        # Remove excessive whitespace
        full_text = re.sub(r'\n\s*\n\s*\n+', '\n\n', full_text)
        # Remove leading/trailing whitespace from lines
        full_text = re.sub(r'^\s+|\s+$', '', full_text, flags=re.MULTILINE)
        # Remove page numbers that appear alone on lines
        full_text = re.sub(r'\n\s*\d+\s*\n', '\n', full_text)
        # Remove headers/footers that repeat
        lines = full_text.split('\n')
        cleaned_lines = []
        prev_line = ""

        for line in lines:
            # Skip if this line is identical to the previous non-empty line
            if line.strip() and line.strip() != prev_line:
                cleaned_lines.append(line)
                if line.strip():
                    prev_line = line.strip()
            elif not line.strip():
                cleaned_lines.append(line)

        full_text = '\n'.join(cleaned_lines)

        # Try to detect and format headings
        # Common patterns for headings in academic papers
        full_text = re.sub(r'\n([A-Z][A-Z\s]{2,})\n', r'\n## \1\n', full_text)  # ALL CAPS headings
        full_text = re.sub(r'\n(\d+\.?\s+[A-Z][^.!?]*[^.!?\s])\n', r'\n## \1\n', full_text)  # Numbered sections

        if full_text.strip():
            content.append(full_text)

        return '\n\n'.join(content)

    except Exception as e:
        print(f"  [ERROR] Failed to process PDF: {str(e)}")
        return None

def convert_file_to_md(file_path, output_dir):
    """Convert a single file (EPUB or PDF) to Markdown"""
    file_name = file_path.name
    file_ext = file_path.suffix.lower()

    print(f"Converting: {file_name}")

    # Extract content based on file type
    if file_ext == '.epub':
        markdown_content = extract_text_from_epub(file_path)
    elif file_ext == '.pdf':
        markdown_content = extract_text_from_pdf(file_path)
    else:
        print(f"  [SKIP] Unsupported file type: {file_ext}")
        return False

    if markdown_content:
        # Create output filename (same base name, .md extension)
        md_filename = file_path.stem + '.md'
        output_path = output_dir / md_filename

        # Write markdown file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)

        file_size = len(markdown_content.encode('utf-8'))
        print(f"  [OK] -> {md_filename} ({file_size:,} bytes)")
        return True
    else:
        print(f"  [ERROR] Failed to convert: {file_name}")
        return False

def main():
    """Main conversion function"""
    current_dir = Path.cwd()

    # Check if at least one library is available
    if not EPUB_AVAILABLE and not PDF_AVAILABLE:
        print("\nERROR: No conversion libraries available!")
        print("Install at least one of the following:")
        print("  - For EPUB: pip install ebooklib html2text")
        print("  - For PDF: pip install pymupdf")
        sys.exit(1)

    print("=" * 70)
    print("EPUB/PDF to Markdown Converter")
    print("=" * 70)
    print(f"Working directory: {current_dir}")
    print(f"EPUB support: {'YES' if EPUB_AVAILABLE else 'NO'}")
    print(f"PDF support: {'YES' if PDF_AVAILABLE else 'NO'}")
    print("-" * 70)

    # Find all EPUB and PDF files
    files_to_convert = []

    if EPUB_AVAILABLE:
        epub_files = list(current_dir.glob("*.epub"))
        files_to_convert.extend(epub_files)
        if epub_files:
            print(f"Found {len(epub_files)} EPUB file(s)")

    if PDF_AVAILABLE:
        pdf_files = list(current_dir.glob("*.pdf"))
        files_to_convert.extend(pdf_files)
        if pdf_files:
            print(f"Found {len(pdf_files)} PDF file(s)")

    if not files_to_convert:
        print("\nNo EPUB or PDF files found in current directory")
        return

    print(f"\nTotal files to convert: {len(files_to_convert)}")
    print("-" * 70)

    # Convert each file
    successful = 0
    failed = 0

    for file_path in sorted(files_to_convert):
        if convert_file_to_md(file_path, current_dir):
            successful += 1
        else:
            failed += 1

    print("-" * 70)
    print(f"Conversion complete!")
    print(f"  Successful: {successful}")
    print(f"  Failed: {failed}")
    print(f"  Total: {len(files_to_convert)}")
    print("=" * 70)

if __name__ == "__main__":
    main()
