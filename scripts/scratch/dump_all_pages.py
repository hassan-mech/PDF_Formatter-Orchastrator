import pymupdf

doc = pymupdf.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
print("Total pages:", len(doc))

for pno in range(len(doc)):
    page = doc[pno]
    print(f"\n==================== PAGE {pno + 1} ====================")
    blocks = page.get_text("blocks")
    # sort blocks vertically
    blocks.sort(key=lambda b: (round(b[1], 1), round(b[0], 1)))
    for b in blocks:
        txt = b[4].strip().replace('\n', ' ')
        if txt:
            print(f"[{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}] ({len(txt)} chars): {txt[:90]}")
    images = page.get_images()
    if images:
        print(f"  --> {len(images)} images on page {pno + 1}")
