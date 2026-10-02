import fitz
from PIL import Image
import io, os

doc = fitz.open('263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf')
os.makedirs('extracted_stitched', exist_ok=True)

for pno in range(len(doc)):
    page = doc[pno]
    img_infos = page.get_image_info(xrefs=True)
    if not img_infos:
        continue
    img_infos.sort(key=lambda x: x['bbox'][1])
    
    # Check if they are slices
    images = []
    total_h = 0
    w = 0
    for info in img_infos:
        xref = info['xref']
        base = doc.extract_image(xref)
        im = Image.open(io.BytesIO(base['image']))
        images.append(im)
        w = max(w, im.width)
        total_h += im.height
    
    stitched = Image.new('RGB', (w, total_h), 'white')
    cur_y = 0
    for im in images:
        stitched.paste(im, (0, cur_y))
        cur_y += im.height
    
    out_path = f"extracted_stitched/page_{pno+1:02d}.png"
    stitched.save(out_path)
    print(f"Page {pno+1:02d}: stitched {len(images)} slices -> {w}x{total_h} -> {out_path}")
