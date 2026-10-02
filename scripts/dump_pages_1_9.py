import fitz

doc = fitz.open('263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf')

print("="*60)
for pno in range(9):
    page = doc[pno]
    print(f"--- PAGE {pno+1} ---")
    blocks = page.get_text("blocks")
    for b in blocks:
        text = b[4].strip()
        if text:
            # print single line representation
            lines = [l.strip() for l in text.split("\n") if l.strip()]
            print(f"[{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}] -> {' | '.join(lines[:4])}")
            if len(lines) > 4:
                print(f"    ... +{len(lines)-4} more lines")
    imgs = page.get_images()
    if imgs:
        print(f"  Images on page {pno+1}: {len(imgs)}")
        for img in imgs:
            info = doc.extract_image(img[0])
            print(f"    xref={img[0]} dim={info['width']}x{info['height']}")
