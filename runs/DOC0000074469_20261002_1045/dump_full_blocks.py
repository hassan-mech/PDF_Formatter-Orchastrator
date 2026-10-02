import fitz
import json

doc = fitz.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")

full_doc_data = []

for pno, page in enumerate(doc, start=1):
    page_data = {"page": pno, "blocks": []}
    blocks = page.get_text("blocks")
    for b in sorted(blocks, key=lambda x: (x[1], x[0])):
        x0, y0, x1, y1, text, block_no, block_type = b
        cleaned = text.strip()
        if cleaned:
            page_data["blocks"].append({
                "bbox": [round(x0, 2), round(y0, 2), round(x1, 2), round(y1, 2)],
                "text": cleaned
            })
    full_doc_data.append(page_data)

out_path = "runs/DOC0000074469_20261002_1045/extract/full_text_blocks.json"
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(full_doc_data, f, ensure_ascii=False, indent=2)

print(f"Dumped {len(full_doc_data)} pages to {out_path}")
