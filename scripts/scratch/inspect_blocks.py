import pymupdf

doc = pymupdf.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
for i in range(2):
    page = doc[i]
    print(f"\n================ PAGE {i+1} BLOCKS ================")
    blocks = page.get_text("blocks")
    for b in blocks:
        # b: (x0, y0, x1, y1, text, block_no, block_type)
        txt = b[4].strip()
        print(f"Block {b[5]} [{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}]:")
        print(f"   {repr(txt[:100])}")
