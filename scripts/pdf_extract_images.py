"""
pdf_extract_images.py — Render PDF pages as PNG images (and optionally stitch slices)

Merges: extract_all_stitched.py + find_table_lines.py

Usage:
    python pdf_extract_images.py -i input/file.pdf -o renders/ --dpi 200 --stitch
"""

import argparse
import sys
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    print("[ERROR] PyMuPDF (fitz) is not installed. Run: pip install pymupdf")
    sys.exit(1)

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


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
    return len(page.get_text("text").strip()) == 0


def render_page(page, dpi: int, out_path: Path) -> Path:
    """Render a single fitz page to a PNG file and return the path."""
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    pix.save(str(out_path))
    return out_path


def stitch_page_images(page, doc, dpi: int, out_path: Path) -> Path | None:
    """
    Extract embedded image slices from a page and stitch them vertically.
    Returns the stitched image path, or None if fewer than 2 images found.
    """
    if not PIL_AVAILABLE:
        print("  ⚠️  Pillow not installed — skipping stitch (pip install pillow)")
        return None

    img_list = page.get_images(full=True)
    if len(img_list) < 2:
        return None

    pil_images = []
    for img_info in img_list:
        xref = img_info[0]
        try:
            img_data = doc.extract_image(xref)
            import io
            pil_img = Image.open(io.BytesIO(img_data["image"]))
            pil_images.append(pil_img.convert("RGB"))
        except Exception as exc:
            print(f"    ⚠️  Could not extract image xref={xref}: {exc}")

    if not pil_images:
        return None

    total_w = max(im.width for im in pil_images)
    total_h = sum(im.height for im in pil_images)
    stitched = Image.new("RGB", (total_w, total_h), (255, 255, 255))
    y_offset = 0
    for im in pil_images:
        stitched.paste(im, (0, y_offset))
        y_offset += im.height

    stitched.save(str(out_path))
    return out_path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Render PDF pages as PNG images, optionally stitching embedded image slices."
    )
    parser.add_argument("-i", "--input", required=True, help="Path to PDF file")
    parser.add_argument(
        "-o", "--output", default="renders/",
        help="Output directory for PNG files (default: renders/)"
    )
    parser.add_argument(
        "--pages", default="all",
        help="Page range, e.g. '1-9' or 'all' (default: all)"
    )
    parser.add_argument(
        "--stitch", action="store_true",
        help="Stitch embedded image slices vertically per page"
    )
    parser.add_argument(
        "--dpi", type=int, default=150,
        help="Resolution for page rendering (default: 150)"
    )
    args = parser.parse_args()

    pdf_path = Path(args.input)
    if not pdf_path.exists():
        print(f"[ERROR] File not found: {pdf_path}")
        sys.exit(1)

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        doc = fitz.open(str(pdf_path))
    except Exception as exc:
        print(f"[ERROR] Cannot open PDF: {exc}")
        sys.exit(1)

    page_indices = parse_page_range(args.pages, doc.page_count)
    stem = pdf_path.stem
    image_only_pages = []

    print(f"🖼  Rendering {len(page_indices)} page(s) at {args.dpi} DPI → {out_dir}/")

    for i in page_indices:
        page = doc[i]
        page_num = i + 1

        # Detect image-only page
        if is_image_only_page(page):
            image_only_pages.append(page_num)
            print(f"  ⚠️  Page {page_num}: IMAGE-ONLY (no selectable text — OCR may be needed)")

        # Render full page
        png_path = out_dir / f"{stem}_page_{page_num:03d}.png"
        try:
            render_page(page, args.dpi, png_path)
            print(f"  ✅ Page {page_num}: rendered → {png_path.name}")
        except Exception as exc:
            print(f"  ❌ Page {page_num}: render failed — {exc}")
            continue

        # Optionally stitch embedded image slices
        if args.stitch:
            stitch_path = out_dir / f"{stem}_page_{page_num:03d}_stitched.png"
            try:
                result = stitch_page_images(page, doc, args.dpi, stitch_path)
                if result:
                    print(f"     🧵 Stitched → {stitch_path.name}")
                else:
                    print(f"     ℹ️  Stitch skipped (< 2 embedded images on page {page_num})")
            except Exception as exc:
                print(f"     ⚠️  Stitch failed on page {page_num}: {exc}")

    doc.close()

    if image_only_pages:
        print(
            f"\n⚠️  IMAGE-ONLY pages detected (no selectable text): {image_only_pages}"
            f"\n   → Run OCR with: python ocr_pages.py -i {out_dir} --engine windows"
        )
    print(f"\n✅ Done. Output saved to: {out_dir}")


if __name__ == "__main__":
    main()
