import pymupdf

orig = pymupdf.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
conv = pymupdf.open("temp/latest_docx.pdf")

for pno in [3, 4]: # Pages 4 and 5 (0-indexed 3, 4)
    p_orig = orig[pno]
    p_conv = conv[pno]
    print(f"\n================ PAGE {pno + 1} ================")
    print("--- ORIG IMAGES ---")
    for img in p_orig.get_images():
        for r in p_orig.get_image_rects(img[0]):
            print(f"  xref {img[0]}: [{r.x0:.1f}, {r.y0:.1f}, {r.x1:.1f}, {r.y1:.1f}] (w={r.width:.1f}, h={r.height:.1f})")
    print("--- CONV IMAGES ---")
    for img in p_conv.get_images():
        for r in p_conv.get_image_rects(img[0]):
            print(f"  xref {img[0]}: [{r.x0:.1f}, {r.y0:.1f}, {r.x1:.1f}, {r.y1:.1f}] (w={r.width:.1f}, h={r.height:.1f})")
    print("--- ORIG FIRST 3 BLOCKS ---")
    for b in sorted(p_orig.get_text("blocks"), key=lambda x: x[1])[:3]:
        print(f"  [{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}]: {b[4].strip()[:50]}")
    print("--- CONV FIRST 3 BLOCKS ---")
    for b in sorted(p_conv.get_text("blocks"), key=lambda x: x[1])[:3]:
        print(f"  [{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}]: {b[4].strip()[:50]}")
