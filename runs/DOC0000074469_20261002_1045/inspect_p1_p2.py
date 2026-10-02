import fitz

doc = fitz.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")

for pno in [1, 2]:
    page = doc[pno - 1]
    print(f"\n================ PAGE {pno} ================")
    blocks = page.get_text("blocks")
    for b in sorted(blocks, key=lambda x: (x[1], x[0])):
        x0, y0, x1, y1, text, block_no, block_type = b
        print(f"[{round(x0,1)}, {round(y0,1)}, {round(x1,1)}, {round(y1,1)}] :\n{text}")
