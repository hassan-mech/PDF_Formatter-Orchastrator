import json

with open("runs/DOC0000074469_20261002_1045/extract/article_data.json", "r", encoding="utf-8") as f:
    data = json.load(f)

for p in data["pages"]:
    print(f"--- Page {p['page_no']} (width={p['width']}, height={p['height']}) ---")
    print(f"  Blocks count: {len(p['blocks'])}, Images count: {len(p['images'])}")
    for i, b in enumerate(p["blocks"][:5]):
        txt = b.get("text", "").replace("\n", " ")[:80]
        bbox = [round(x, 1) for x in b.get("bbox", [])]
        print(f"  B{i} bbox={bbox}: {txt}")
    if len(p["blocks"]) > 5:
        print(f"  ... and {len(p['blocks']) - 5} more blocks")
    for j, img in enumerate(p["images"]):
        bbox = [round(x, 1) for x in img.get("bbox", [])]
        print(f"  Img{j} bbox={bbox} dim={img.get('width')}x{img.get('height')}")
