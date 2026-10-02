import json

with open("runs/DOC0000074469_20261002_1045/extract/full_text_blocks.json", "r", encoding="utf-8") as f:
    pages = json.load(f)

for p in pages[2:]:
    print(f"\n==================== PAGE {p['page']} ====================")
    for b in p["blocks"]:
        bbox = b["bbox"]
        txt = b["text"].replace("\n", " ")
        print(f"[{bbox[0]}, {bbox[1]}, {bbox[2]}, {bbox[3]}]: {txt[:100]}")
