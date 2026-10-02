import fitz
import json

doc = fitz.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")

for pno, page in enumerate(doc, start=1):
    print(f"\n================ PAGE {pno} ================")
    rect = page.rect
    print(f"Page rect: {rect}")
    
    # Drawings/rectangles
    drawings = page.get_drawings()
    rect_drawings = [d for d in drawings if d.get("rect")]
    print(f"Drawings count: {len(drawings)}, rects: {len(rect_drawings)}")
    for d in rect_drawings[:5]:
        r = [round(x, 1) for x in d["rect"]]
        fill = d.get("fill")
        color = d.get("color")
        print(f"  Rect: {r}, fill={fill}, stroke={color}")

    # Text blocks
    blocks = page.get_text("blocks")
    print(f"Blocks count: {len(blocks)}")
    for b in sorted(blocks, key=lambda x: (x[1], x[0])):
        x0, y0, x1, y1, text, block_no, block_type = b
        first_line = text.strip().replace("\n", " ")[:60]
        if first_line:
            print(f"  [{round(x0,1)}, {round(y0,1)}, {round(x1,1)}, {round(y1,1)}] : {first_line}")
