import json

with open("runs/DOC0000074469_20261002_1045/extract/full_text_blocks.json", "r", encoding="utf-8") as f:
    pages = json.load(f)

p7 = pages[6]
print("Page 7 blocks:")
for b in p7["blocks"]:
    print(f"[{b['bbox']}]: {b['text']}\n")
