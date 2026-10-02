import fitz

doc = fitz.open('263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf')
print(f"Total pages: {len(doc)}")

for pno in range(len(doc)):
    page = doc[pno]
    text = page.get_text().strip()
    images = page.get_images()
    print(f"--- Page {pno+1} --- (text len: {len(text)}, images: {len(images)})")
    if text:
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        print("  First 3 lines:", lines[:3])
        print("  Last 2 lines:", lines[-2:])
    for img_info in images:
        xref = img_info[0]
        base_img = doc.extract_image(xref)
        print(f"  Img xref={xref}: {base_img['width']}x{base_img['height']} {base_img['ext']}")
