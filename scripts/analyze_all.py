import fitz

doc = fitz.open('263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf')

print(f"Total pages: {len(doc)}")
for pno in range(len(doc)):
    page = doc[pno]
    text = page.get_text()
    imgs = page.get_images()
    print(f"Page {pno+1}: text_len={len(text)}, images={len(imgs)}")
    for img_idx, img in enumerate(imgs):
        xref = img[0]
        base_img = doc.extract_image(xref)
        w, h, ext = base_img["width"], base_img["height"], base_img["ext"]
        print(f"   img {img_idx+1}: xref={xref} {w}x{h} {ext}")
