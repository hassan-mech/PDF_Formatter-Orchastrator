"""
pdf_extract_text.py — Extract text from a PDF to a .txt file (reusable, CLI-driven)

Merges: dump_pages_1_9.py + dump_all_paragraphs.py + list_paragraphs.py

Usage:
    python pdf_extract_text.py -i input/file.pdf -o extracted/text_dump.txt --format blocks
"""

import argparse
import sys
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    print("[ERROR] PyMuPDF (fitz) is not installed. Run: pip install pymupdf")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_page_range(pages_str: str, total_pages: int) -> list[int]:
    """Return 0-based page indices from a range string like '1-9' or 'all'."""
    if pages_str.strip().lower() == "all":
        return list(range(total_pages))
    indices = []
    for part in pages_str.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            indices.extend(range(int(start) - 1, int(end)))
        else:
            indices.append(int(part) - 1)
    return [i for i in indices if 0 <= i < total_pages]


def extract_raw(page, page_num: int) -> str:
    """Extract raw text from a page."""
    text = page.get_text("text")
    return f"\n{'='*60}\nPAGE {page_num}\n{'='*60}\n{text}"


def extract_blocks(page, page_num: int) -> str:
    """Extract text blocks with bbox coordinates."""
    lines = [f"\n{'='*60}\nPAGE {page_num} — BLOCKS\n{'='*60}"]
    blocks = page.get_text("blocks")
    for idx, blk in enumerate(blocks):
        x0, y0, x1, y1, text, block_no, block_type = blk
        type_label = "TEXT" if block_type == 0 else "IMAGE"
        lines.append(
            f"\n  Block {idx+1} [{type_label}] bbox=({x0:.1f},{y0:.1f},{x1:.1f},{y1:.1f})\n"
            f"  {repr(text[:200])}"
        )
    return "\n".join(lines)


def extract_words(page, page_num: int) -> str:
    """Extract individual words with bbox coordinates."""
    lines = [f"\n{'='*60}\nPAGE {page_num} — WORDS\n{'='*60}"]
    words = page.get_text("words")
    for w in words:
        x0, y0, x1, y1, word, block_no, line_no, word_no = w
        lines.append(f"  [{word_no:>3}] ({x0:.1f},{y0:.1f}) {repr(word)}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Extract text from a PDF to a .txt file."
    )
    parser.add_argument("-i", "--input", required=True, help="Path to PDF file")
    parser.add_argument(
        "-o", "--output", default="extracted/text_dump.txt",
        help="Output .txt file path (default: extracted/text_dump.txt)"
    )
    parser.add_argument(
        "--pages", default="all",
        help="Page range, e.g. '1-9' or 'all' (default: all)"
    )
    parser.add_argument(
        "--format",
        choices=["raw", "blocks", "words"],
        default="raw",
        help="Extraction format: raw | blocks | words (default: raw)"
    )
    args = parser.parse_args()

    pdf_path = Path(args.input)
    if not pdf_path.exists():
        print(f"[ERROR] File not found: {pdf_path}")
        sys.exit(1)

    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as exc:
        print(f"[ERROR] Cannot open PDF: {exc}")
        sys.exit(1)

    page_indices = parse_page_range(args.pages, doc.page_count)
    print(f"📄 Extracting text from '{pdf_path.name}' ({len(page_indices)} page(s), format={args.format})")

    chunks = []
    for i in page_indices:
        page = doc[i]
        page_num = i + 1
        try:
            if args.format == "raw":
                chunk = extract_raw(page, page_num)
            elif args.format == "blocks":
                chunk = extract_blocks(page, page_num)
            else:
                chunk = extract_words(page, page_num)
            chunks.append(chunk)
            print(f"  ✅ Page {page_num}: extracted {len(page.get_text('text'))} chars")
        except Exception as exc:
            chunks.append(f"\n[ERROR on page {page_num}: {exc}]")
            print(f"  ❌ Page {page_num}: {exc}")

    doc.close()

    full_text = "\n".join(chunks)
    try:
        out_path.write_text(full_text, encoding="utf-8")
        print(f"\n✅ Saved to: {out_path}  ({out_path.stat().st_size:,} bytes)")
    except Exception as exc:
        print(f"[ERROR] Could not write output: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
