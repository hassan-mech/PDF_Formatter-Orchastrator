import fitz
import json
from pathlib import Path

doc_path = Path("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
doc = fitz.open(str(doc_path))
print(f"Total pages: {len(doc)}")

for page_idx in range(len(doc)):
    page = doc[page_idx]
    rect = page.rect
    text = page.get_text("text").strip()
    images = page.get_images()
    tables = []
    try:
        tabs = page.find_tables()
        if tabs:
            tables = [t.bbox for t in tabs.tables]
    except Exception as e:
        pass
    print(f"\n--- PAGE {page_idx + 1} ({rect.width} x {rect.height} pt) ---")
    print(f"Images count: {len(images)}, Tables detected: {len(tables)}, Text len: {len(text)}")
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    print(f"Sample lines (first 5): {lines[:5]}")
    print(f"Sample lines (last 3): {lines[-3:]}")
