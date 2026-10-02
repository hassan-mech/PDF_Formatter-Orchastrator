"""
compare_visual.py — Visual comparison of original vs converted PDF pages (reusable, CLI-driven)

Replaces: compare_pages.py + verify_page4.py

Usage:
    python compare_visual.py -a input/original.pdf -b output/converted.pdf --dpi 100
    python compare_visual.py -a original.pdf -b converted.pdf --pages 1-5 --threshold 0.03
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
    from PIL import Image, ImageChops, ImageEnhance
    import PIL
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("[WARNING] Pillow not installed — diff images will not be saved. Run: pip install pillow")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_page_range(pages_str: str, total_pages: int) -> list[int]:
    """Return 0-based page indices from '1-9', '2,4,6', or 'all'."""
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


def render_page_to_pil(doc, page_index: int, dpi: int):
    """Render a fitz page to a Pillow RGBA image."""
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = doc[page_index].get_pixmap(matrix=mat, alpha=False)
    from io import BytesIO
    return Image.open(BytesIO(pix.tobytes("png"))).convert("RGB")


def compute_similarity(img_a: "Image.Image", img_b: "Image.Image") -> tuple[float, "Image.Image | None"]:
    """
    Compare two PIL images pixel-by-pixel.

    Returns (similarity_ratio, diff_image).
    similarity_ratio is 1.0 for identical images, lower for more differences.
    diff_image is the highlighted difference image (or None if Pillow unavailable).
    """
    # Resize to smallest common size
    w = min(img_a.width, img_b.width)
    h = min(img_a.height, img_b.height)
    a = img_a.resize((w, h), Image.LANCZOS)
    b = img_b.resize((w, h), Image.LANCZOS)

    diff = ImageChops.difference(a, b)
    # Sum of all pixel differences (per channel, 0-255)
    import struct
    stat = diff.histogram()
    total_pixels = w * h * 3  # RGB channels
    # Count non-zero diff values: histogram bins 1-255 per channel
    nonzero = sum(stat[i] for i in range(1, 256)) + \
              sum(stat[256 + i] for i in range(1, 256)) + \
              sum(stat[512 + i] for i in range(1, 256))
    similarity = 1.0 - (nonzero / total_pixels)

    # Enhance diff for visibility
    enhanced = ImageEnhance.Brightness(diff).enhance(5.0)
    return max(0.0, similarity), enhanced


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Visual pixel-by-pixel comparison of two PDFs."
    )
    parser.add_argument("-a", "--original",  required=True, help="Original PDF path")
    parser.add_argument("-b", "--converted", required=True, help="Converted PDF (or DOCX-exported PDF) path")
    parser.add_argument(
        "--pages", default="all",
        help="Page range to compare, e.g. '1-5' or 'all' (default: all)"
    )
    parser.add_argument("--dpi", type=int, default=100, help="Render DPI (default: 100)")
    parser.add_argument(
        "-o", "--output", default="renders/diffs/",
        help="Output directory for diff images (default: renders/diffs/)"
    )
    parser.add_argument(
        "--threshold", type=float, default=0.05,
        help="Max acceptable difference ratio 0-1 (default: 0.05 = 5%%)"
    )
    args = parser.parse_args()

    orig_path = Path(args.original)
    conv_path = Path(args.converted)

    for p in (orig_path, conv_path):
        if not p.exists():
            print(f"[ERROR] File not found: {p}")
            sys.exit(1)

    if not PIL_AVAILABLE:
        print("[ERROR] Pillow required for visual comparison. Run: pip install pillow")
        sys.exit(1)

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        doc_a = fitz.open(str(orig_path))
        doc_b = fitz.open(str(conv_path))
    except Exception as exc:
        print(f"[ERROR] Cannot open PDF(s): {exc}")
        sys.exit(1)

    pages_a = parse_page_range(args.pages, doc_a.page_count)
    pages_b = parse_page_range(args.pages, doc_b.page_count)
    compare_pages = list(zip(pages_a, pages_b))

    print(f"🔍 Comparing {len(compare_pages)} page(s) at {args.dpi} DPI")
    print(f"   Original : {orig_path.name} ({doc_a.page_count}p)")
    print(f"   Converted: {conv_path.name} ({doc_b.page_count}p)")
    print(f"   Threshold: {args.threshold:.1%}\n")

    scores = []
    pass_count = 0

    for i_a, i_b in compare_pages:
        page_label = f"Page {i_a+1}"
        try:
            img_a = render_page_to_pil(doc_a, i_a, args.dpi)
            img_b = render_page_to_pil(doc_b, i_b, args.dpi)
            sim, diff_img = compute_similarity(img_a, img_b)
            diff_ratio = 1.0 - sim
            passed = diff_ratio <= args.threshold
            status = "✅ PASS" if passed else "❌ FAIL"
            if passed:
                pass_count += 1
            scores.append(sim)

            diff_path = out_dir / f"diff_p{i_a+1:03d}.png"
            if diff_img is not None:
                diff_img.save(str(diff_path))

            print(
                f"  {page_label}: similarity={sim:.2%}  diff={diff_ratio:.2%}  {status}"
                + (f"  → {diff_path.name}" if not passed else "")
            )
        except Exception as exc:
            print(f"  {page_label}: ❌ ERROR — {exc}")

    doc_a.close()
    doc_b.close()

    if scores:
        avg_sim = sum(scores) / len(scores)
        overall = "✅ PASS" if pass_count == len(scores) else "❌ FAIL"
        print(f"\n{'='*55}")
        print(f"  Overall: {pass_count}/{len(scores)} pages passed  |  avg similarity={avg_sim:.2%}  |  {overall}")
        print(f"  Diff images saved to: {out_dir}")
        print(f"{'='*55}")
        sys.exit(0 if pass_count == len(scores) else 1)


if __name__ == "__main__":
    main()
