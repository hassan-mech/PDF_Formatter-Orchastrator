"""
pdf_inspect.py — PDF structure inspector (reusable, CLI-driven)

Merges: analyze_all.py + inspect_pdf.py + inspect_layout.py + inspect_styles.py

Usage:
    python pdf_inspect.py -i input/file.pdf --mode all --pages 1-5
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


def is_image_only_page(page) -> bool:
    """Return True when a page has no selectable text (likely scanned/image-only)."""
    text = page.get_text("text").strip()
    return len(text) == 0


# ---------------------------------------------------------------------------
# Mode handlers
# ---------------------------------------------------------------------------

def mode_summary(doc, page_indices: list[int]):
    print(f"\n📄 Document Summary")
    print(f"   Total pages : {doc.page_count}")
    print(f"   Metadata    : {doc.metadata}\n")
    for i in page_indices:
        page = doc[i]
        text = page.get_text("text").strip()
        img_list = page.get_images(full=True)
        image_only = is_image_only_page(page)
        warn = " ⚠️  IMAGE-ONLY (OCR needed)" if image_only else ""
        print(
            f"  Page {i+1:>3}: "
            f"size=({page.rect.width:.0f}x{page.rect.height:.0f})pt  "
            f"chars={len(text):>6}  images={len(img_list):>3}{warn}"
        )


def mode_layout(doc, page_indices: list[int]):
    print(f"\n🗂  Layout Analysis")
    for i in page_indices:
        page = doc[i]
        blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        print(f"\n  --- Page {i+1} ---")
        for blk in blocks:
            btype = blk.get("type", -1)
            bbox = tuple(round(v, 1) for v in blk["bbox"])
            if btype == 0:  # text block
                lines = blk.get("lines", [])
                sample = ""
                if lines and lines[0].get("spans"):
                    sample = lines[0]["spans"][0].get("text", "")[:60]
                print(f"    TEXT  bbox={bbox}  lines={len(lines)}  sample={repr(sample)}")
            elif btype == 1:  # image block
                print(f"    IMAGE bbox={bbox}")
            else:
                print(f"    OTHER type={btype} bbox={bbox}")


def mode_styles(doc, page_indices: list[int]):
    print(f"\n🎨 Font / Style Analysis")
    font_stats: dict[str, int] = {}
    for i in page_indices:
        page = doc[i]
        blocks = page.get_text("dict")["blocks"]
        for blk in blocks:
            if blk.get("type") != 0:
                continue
            for line in blk.get("lines", []):
                for span in line.get("spans", []):
                    key = f"{span.get('font', '?')} | size={span.get('size', 0):.1f} | flags={span.get('flags', 0)}"
                    font_stats[key] = font_stats.get(key, 0) + len(span.get("text", ""))
    print(f"  {'Font | Size | Flags':<60} {'Chars':>8}")
    print("  " + "-" * 70)
    for k, v in sorted(font_stats.items(), key=lambda x: -x[1])[:40]:
        print(f"  {k:<60} {v:>8}")


def mode_tables(doc, page_indices: list[int]):
    print(f"\n📊 Table Detection")
    for i in page_indices:
        page = doc[i]
        try:
            tabs = page.find_tables()
            if tabs.tables:
                print(f"  Page {i+1}: {len(tabs.tables)} table(s) found")
                for t_idx, t in enumerate(tabs.tables):
                    print(f"    Table {t_idx+1}: rows={t.row_count} cols={t.col_count} bbox={tuple(round(v,1) for v in t.bbox)}")
            else:
                print(f"  Page {i+1}: no tables detected")
        except Exception as exc:
            print(f"  Page {i+1}: table detection failed — {exc}")


def mode_images(doc, page_indices: list[int]):
    print(f"\n🖼  Embedded Images")
    for i in page_indices:
        page = doc[i]
        img_list = page.get_images(full=True)
        print(f"  Page {i+1}: {len(img_list)} image(s)")
        for img in img_list:
            xref = img[0]
            info = doc.extract_image(xref)
            print(
                f"    xref={xref} ext={info.get('ext','?')} "
                f"size={info.get('width','?')}x{info.get('height','?')} "
                f"colorspace={info.get('colorspace','?')}"
            )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="PDF structure inspector — summary, layout, styles, tables, images."
    )
    parser.add_argument("-i", "--input", required=True, help="Path to PDF file")
    parser.add_argument(
        "--pages", default="all",
        help="Page range, e.g. '1-9' or 'all' (default: all)"
    )
    parser.add_argument(
        "--mode",
        choices=["summary", "layout", "tables", "images", "styles", "all"],
        default="summary",
        help="Inspection mode (default: summary)"
    )
    args = parser.parse_args()

    pdf_path = Path(args.input)
    if not pdf_path.exists():
        print(f"[ERROR] File not found: {pdf_path}")
        sys.exit(1)

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as exc:
        print(f"[ERROR] Cannot open PDF: {exc}")
        sys.exit(1)

    page_indices = parse_page_range(args.pages, doc.page_count)
    print(f"🔍 Inspecting: {pdf_path.name}  ({len(page_indices)} page(s), mode={args.mode})")

    mode = args.mode
    if mode in ("summary", "all"):
        mode_summary(doc, page_indices)
    if mode in ("layout", "all"):
        mode_layout(doc, page_indices)
    if mode in ("styles", "all"):
        mode_styles(doc, page_indices)
    if mode in ("tables", "all"):
        mode_tables(doc, page_indices)
    if mode in ("images", "all"):
        mode_images(doc, page_indices)

    doc.close()
    print("\n✅ Inspection complete.")


if __name__ == "__main__":
    main()
