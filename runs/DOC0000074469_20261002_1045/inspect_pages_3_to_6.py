import json

with open("runs/DOC0000074469_20261002_1045/extract/full_text_blocks.json", "r", encoding="utf-8") as f:
    pages = json.load(f)

for p in pages[2:6]:  # Pages 3, 4, 5, 6
    print(f"\n================ PAGE {p['page']} ================")
    for b in p["blocks"]:
        print(f"[{b['bbox']}]:\n{b['text']}\n")
