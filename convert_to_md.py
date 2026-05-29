#!/usr/bin/env python3
"""
Unified EPUB, PDF, FB2, and DJVU to Markdown Converter
Converts all EPUB, PDF, FB2, and DJVU files in the current directory to Markdown format
"""

import os
import re
import shutil
import subprocess
import sys
import tempfile
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

# Try importing FB2 libraries
try:
    from lxml import etree
    FB2_AVAILABLE = True
except ImportError:
    FB2_AVAILABLE = False
    print("WARNING: lxml not found. FB2 conversion will be skipped.")
    print("To enable FB2 support, run: pip install lxml")

# Check for DjVuLibre (djvutxt command)
DJVUTXT_PATH = None
_djvu_candidates = [
    r"C:\Program Files\DjVuLibre\djvutxt.exe",
    r"C:\Program Files (x86)\DjVuLibre\djvutxt.exe",
]
for _p in _djvu_candidates:
    if os.path.isfile(_p):
        DJVUTXT_PATH = _p
        break
if DJVUTXT_PATH is None:
    DJVUTXT_PATH = shutil.which("djvutxt")
DJVU_AVAILABLE = DJVUTXT_PATH is not None
if not DJVU_AVAILABLE:
    print("WARNING: djvutxt not found. DJVU conversion will be skipped.")
    print("To enable DJVU support, install DjVuLibre: winget install DjVuLibre.DjView")


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


def _fb2_element_text(element, ns):
    """Recursively extract text from an FB2 element, converting inline markup to markdown."""
    parts = []
    tag = etree.QName(element.tag).localname if isinstance(element.tag, str) else ''

    # Opening markup
    if tag == 'emphasis':
        parts.append('*')
    elif tag == 'strong':
        parts.append('**')

    # Element's own text
    if element.text:
        parts.append(element.text)

    # Process children
    for child in element:
        parts.append(_fb2_element_text(child, ns))
        # Tail text (text after the child but before the next sibling)
        if child.tail:
            parts.append(child.tail)

    # Closing markup
    if tag == 'emphasis':
        parts.append('*')
    elif tag == 'strong':
        parts.append('**')

    return ''.join(parts)


def _fb2_process_section(section, ns, level=2):
    """Recursively process an FB2 section, returning markdown lines."""
    lines = []

    for child in section:
        tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ''

        if tag == 'title':
            # Title contains <p> elements
            title_parts = []
            for p in child:
                t = _fb2_element_text(p, ns).strip()
                if t:
                    title_parts.append(t)
            if title_parts:
                heading = ' '.join(title_parts)
                lines.append(f"{'#' * level} {heading}")
                lines.append("")

        elif tag == 'subtitle':
            text = _fb2_element_text(child, ns).strip()
            if text:
                lines.append(f"{'#' * (level + 1)} {text}")
                lines.append("")

        elif tag == 'p':
            text = _fb2_element_text(child, ns).strip()
            if text:
                lines.append(text)
                lines.append("")

        elif tag == 'empty-line':
            lines.append("")

        elif tag == 'epigraph':
            for ep_child in child:
                ep_tag = etree.QName(ep_child.tag).localname if isinstance(ep_child.tag, str) else ''
                if ep_tag == 'p':
                    text = _fb2_element_text(ep_child, ns).strip()
                    if text:
                        lines.append(f"> {text}")
                elif ep_tag == 'text-author':
                    text = _fb2_element_text(ep_child, ns).strip()
                    if text:
                        lines.append(f"> -- {text}")
            lines.append("")

        elif tag == 'cite':
            for cite_child in child:
                cite_tag = etree.QName(cite_child.tag).localname if isinstance(cite_child.tag, str) else ''
                if cite_tag == 'p':
                    text = _fb2_element_text(cite_child, ns).strip()
                    if text:
                        lines.append(f"> {text}")
                elif cite_tag == 'text-author':
                    text = _fb2_element_text(cite_child, ns).strip()
                    if text:
                        lines.append(f"> -- {text}")
            lines.append("")

        elif tag == 'poem':
            for poem_child in child:
                poem_tag = etree.QName(poem_child.tag).localname if isinstance(poem_child.tag, str) else ''
                if poem_tag == 'title':
                    title_parts = []
                    for p in poem_child:
                        t = _fb2_element_text(p, ns).strip()
                        if t:
                            title_parts.append(t)
                    if title_parts:
                        lines.append(f"**{' '.join(title_parts)}**")
                        lines.append("")
                elif poem_tag == 'stanza':
                    for v in poem_child:
                        v_tag = etree.QName(v.tag).localname if isinstance(v.tag, str) else ''
                        if v_tag == 'v':
                            text = _fb2_element_text(v, ns).strip()
                            lines.append(text)
                    lines.append("")
                elif poem_tag == 'text-author':
                    text = _fb2_element_text(poem_child, ns).strip()
                    if text:
                        lines.append(f"-- {text}")
                        lines.append("")

        elif tag == 'section':
            # Nested section
            lines.extend(_fb2_process_section(child, ns, level + 1))

    return lines


def extract_text_from_fb2(fb2_path):
    """Extract text content from FB2 file"""
    if not FB2_AVAILABLE:
        return None

    try:
        tree = etree.parse(open(fb2_path, 'rb'))
        root = tree.getroot()

        # Detect namespace
        ns_map = root.nsmap
        fb_ns = ns_map.get(None, 'http://www.gribuser.ru/xml/fictionbook/2.0')
        ns = {'fb': fb_ns}

        # Extract metadata
        title_el = root.find('.//fb:description/fb:title-info/fb:book-title', ns)
        title = title_el.text.strip() if title_el is not None and title_el.text else fb2_path.stem.replace('_', ' ')

        author_parts = []
        for author_el in root.findall('.//fb:description/fb:title-info/fb:author', ns):
            fn = author_el.find('fb:first-name', ns)
            mn = author_el.find('fb:middle-name', ns)
            ln = author_el.find('fb:last-name', ns)
            parts = []
            if fn is not None and fn.text:
                parts.append(fn.text.strip())
            if mn is not None and mn.text:
                parts.append(mn.text.strip())
            if ln is not None and ln.text:
                parts.append(ln.text.strip())
            if parts:
                author_parts.append(' '.join(parts))
        author = ', '.join(author_parts) if author_parts else 'Unknown Author'

        content = []
        content.append(f"# {title}")
        content.append(f"**Author:** {author}")
        content.append("")

        # Process all body elements
        for body in root.findall('.//fb:body', ns):
            # Body-level title (e.g., "Notes")
            body_title = body.find('fb:title', ns)
            body_name = body.get('name', '')
            if body_name == 'notes':
                content.append("## Notes")
                content.append("")

            for section in body.findall('fb:section', ns):
                section_lines = _fb2_process_section(section, ns, level=2)
                content.extend(section_lines)

        result = '\n'.join(content)
        # Clean up excessive blank lines
        result = re.sub(r'\n{3,}', '\n\n', result)
        return result

    except Exception as e:
        print(f"  [ERROR] Failed to process FB2: {str(e)}")
        return None


def extract_text_from_djvu(djvu_path):
    """Extract text content from DJVU file using djvutxt"""
    if not DJVU_AVAILABLE:
        return None

    try:
        # djvutxt cannot handle non-ASCII filenames; copy to a temp file if needed
        need_temp = False
        try:
            str(djvu_path).encode('ascii')
        except UnicodeEncodeError:
            need_temp = True

        if need_temp:
            tmp_dir = tempfile.mkdtemp()
            tmp_file = os.path.join(tmp_dir, 'input.djvu')
            shutil.copy2(str(djvu_path), tmp_file)
            source_file = tmp_file
        else:
            source_file = str(djvu_path)
            tmp_dir = None

        try:
            result = subprocess.run(
                [DJVUTXT_PATH, source_file],
                capture_output=True,
                timeout=300,
            )
            # djvutxt outputs UTF-8
            full_text = result.stdout.decode('utf-8', errors='replace')
        finally:
            if tmp_dir:
                shutil.rmtree(tmp_dir, ignore_errors=True)

        if not full_text.strip():
            print("  [WARN] No text layer found in DJVU (may be image-only scan)")
            return None

        # Use filename as title for DJVU (no reliable metadata)
        title = djvu_path.stem.replace('_', ' ')
        author = 'Unknown Author'
        # Try to extract author from filename pattern "Author - Title"
        name_match = re.match(r'^(.+?)\s*[-\u2013\u2014]\s*(.+)$', title)
        if name_match:
            author = name_match.group(1).strip()
            title = name_match.group(2).strip()

        content = []
        content.append(f"# {title}")
        content.append(f"**Author:** {author}")
        content.append("")

        # Clean up OCR text
        # Collapse multiple spaces (common in DJVU OCR)
        full_text = re.sub(r'  +', ' ', full_text)
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
            if line.strip() and line.strip() != prev_line:
                cleaned_lines.append(line)
                if line.strip():
                    prev_line = line.strip()
            elif not line.strip():
                cleaned_lines.append(line)

        full_text = '\n'.join(cleaned_lines)

        # Try to detect and format headings
        full_text = re.sub(r'\n([A-ZА-ЯЁ][A-ZА-ЯЁ\s]{2,})\n', r'\n## \1\n', full_text)  # ALL CAPS headings
        full_text = re.sub(r'\n(\d+\.?\s+[A-ZА-ЯЁ][^.!?]*[^.!?\s])\n', r'\n## \1\n', full_text)  # Numbered sections

        if full_text.strip():
            content.append(full_text)

        return '\n\n'.join(content)

    except subprocess.TimeoutExpired:
        print(f"  [ERROR] djvutxt timed out processing DJVU file")
        return None
    except Exception as e:
        print(f"  [ERROR] Failed to process DJVU: {str(e)}")
        return None


def convert_file_to_md(file_path, output_dir):
    """Convert a single file (EPUB, PDF, FB2, or DJVU) to Markdown"""
    file_name = file_path.name
    file_ext = file_path.suffix.lower()

    print(f"Converting: {file_name}")

    # Extract content based on file type
    if file_ext == '.epub':
        markdown_content = extract_text_from_epub(file_path)
    elif file_ext == '.pdf':
        markdown_content = extract_text_from_pdf(file_path)
    elif file_ext == '.fb2':
        markdown_content = extract_text_from_fb2(file_path)
    elif file_ext == '.djvu':
        markdown_content = extract_text_from_djvu(file_path)
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
    if not EPUB_AVAILABLE and not PDF_AVAILABLE and not FB2_AVAILABLE and not DJVU_AVAILABLE:
        print("\nERROR: No conversion libraries available!")
        print("Install at least one of the following:")
        print("  - For EPUB: pip install ebooklib html2text")
        print("  - For PDF:  pip install pymupdf")
        print("  - For FB2:  pip install lxml")
        print("  - For DJVU: winget install DjVuLibre.DjView")
        sys.exit(1)

    print("=" * 70)
    print("EPUB/PDF/FB2/DJVU to Markdown Converter")
    print("=" * 70)
    print(f"Working directory: {current_dir}")
    print(f"EPUB support: {'YES' if EPUB_AVAILABLE else 'NO'}")
    print(f"PDF support:  {'YES' if PDF_AVAILABLE else 'NO'}")
    print(f"FB2 support:  {'YES' if FB2_AVAILABLE else 'NO'}")
    print(f"DJVU support: {'YES' if DJVU_AVAILABLE else 'NO'}")
    print("-" * 70)

    # Find all supported files
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

    if FB2_AVAILABLE:
        fb2_files = list(current_dir.glob("*.fb2"))
        files_to_convert.extend(fb2_files)
        if fb2_files:
            print(f"Found {len(fb2_files)} FB2 file(s)")

    if DJVU_AVAILABLE:
        djvu_files = list(current_dir.glob("*.djvu"))
        files_to_convert.extend(djvu_files)
        if djvu_files:
            print(f"Found {len(djvu_files)} DJVU file(s)")

    if not files_to_convert:
        print("\nNo supported files found in current directory")
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
