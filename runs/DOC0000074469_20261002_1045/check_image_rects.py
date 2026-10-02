import fitz

doc = fitz.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")

for pno, page in enumerate(doc, start=1):
    img_infos = page.get_images()
    print(f"\n--- Page {pno} ({len(img_infos)} images) ---")
    for img_info in img_infos:
        xref = img_info[0]
        # Find where this image is placed on the page
        rects = page.get_image_rects(xref)
        for r in rects:
            print(f"  xref {xref}: rect=[{round(r.x0,1)}, {round(r.y0,1)}, {round(r.x1,1)}, {round(r.y1,1)}], w={round(r.width,1)}, h={round(r.height,1)}")
