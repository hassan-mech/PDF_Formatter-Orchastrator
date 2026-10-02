"""
pdf_extract_tables.py — Extract tables from a PDF to JSON or CSV (reusable, CLI-driven)

Replaces: extract_tables.py

Usage:
    python pdf_extract_tables.py -i input/file.pdf -o extracted/tables.json --format json
    python pdf_extract_tables.py -i input/file.pdf -o extracted/tables.csv --format csv
"""

import argparse
import csv
import json
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


def extract_tables_from_doc(doc, page_indices: list[int]) -> list[dict]:
    """Return a list of table dicts (page, table_index, bbox, rows)."""
    results = []
    for i in page_indices:
        page = doc[i]
        page_num = i + 1
        try:
            tabs = page.find_tables()
            for t_idx, table in enumerate(tabs.tables):
                data = table.extract()
                results.append({
                    "page": page_num,
                    "table_index": t_idx + 1,
                    "bbox": list(table.bbox),
                    "rows": len(data),
                    "cols": len(data[0]) if data else 0,
                    "data": data,
                })
                print(f"  ✅ Page {page_num}, Table {t_idx+1}: {len(data)} rows × {len(data[0]) if data else 0} cols")
        except Exception as exc:
            print(f"  ❌ Page {page_num}: table extraction failed — {exc}")
    return results


# ---------------------------------------------------------------------------
# Serializers
# ---------------------------------------------------------------------------

def save_json(tables: list[dict], out_path: Path):
    """Save extracted tables as JSON."""
    out_path.write_text(json.dumps(tables, ensure_ascii=False, indent=2), encoding="utf-8")


def save_csv(tables: list[dict], out_path: Path):
    """Save extracted tables as CSV (all tables appended sequentially)."""
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        for t in tables:
            writer.writerow([
                f"--- Page {t['page']}, Table {t['table_index']} "
                f"({t['rows']} rows x {t['cols']} cols) ---"
            ])
            for row in t.get("data", []):
                writer.writerow([cell if cell is not None else "" for cell in row])
            writer.writerow([])  # blank separator


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Extract tables from a PDF to JSON or CSV."
    )
    parser.add_argument("-i", "--input", required=True, help="Path to PDF file")
    parser.add_argument(
        "-o", "--output", default="extracted/tables.json",
        help="Output file path (default: extracted/tables.json)"
    )
    parser.add_argument(
        "--pages", default="all",
        help="Page range, e.g. '1-9' or 'all' (default: all)"
    )
    parser.add_argument(
        "--format", choices=["json", "csv"], default="json",
        help="Output format: json | csv (default: json)"
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
    print(f"📊 Extracting tables from '{pdf_path.name}' ({len(page_indices)} page(s))")

    tables = extract_tables_from_doc(doc, page_indices)
    doc.close()

    if not tables:
        print("  ℹ️  No tables found.")
    else:
        try:
            if args.format == "json":
                save_json(tables, out_path)
            else:
                save_csv(tables, out_path)
            print(f"\n✅ Saved {len(tables)} table(s) → {out_path}")
        except Exception as exc:
            print(f"[ERROR] Could not write output: {exc}")
            sys.exit(1)


if __name__ == "__main__":
    main()
