import json

with open("runs/DOC0000074469_20261002_1045/extract/full_text_blocks.json", "r", encoding="utf-8") as f:
    pages = json.load(f)

p3 = pages[2]
print("Page 3 blocks:")
for b in p3["blocks"]:
    print(f"[{b['bbox']}]:\n{b['text']}\n")
