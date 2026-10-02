import pymupdf

doc = pymupdf.open('finished/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf')
print(f"Total pages: {doc.page_count}")

for pno in range(doc.page_count):
    page = doc[pno]
    print(f"\n================ PAGE {pno + 1} ================")
    print(f"Dimensions: width={page.rect.width}, height={page.rect.height}")
    
    # Drawings
    drawings = page.get_drawings()
    print(f"Drawings count: {len(drawings)}")
    for d in drawings:
        if d.get('fill') or (d.get('color') and d.get('width', 0) > 0.5):
            print(f"  Drawing: rect={[round(x, 1) for x in d['rect']]}, fill={d.get('fill')}, color={d.get('color')}, w={d.get('width')}")
            
    # Images
    images = page.get_images()
    print(f"Images count: {len(images)}")
    for img in images:
        rects = page.get_image_rects(img[0])
        print(f"  Image xref={img[0]}: rects={[ [round(x, 1) for x in r] for r in rects ]}")

    # Column inspection: find text x0 distributions
    blocks = page.get_text('blocks')
    x_coords = [round(b[0], 1) for b in blocks if len(b[4].strip()) > 10]
    print(f"Text block x0 values: {sorted(list(set(x_coords)))}")
