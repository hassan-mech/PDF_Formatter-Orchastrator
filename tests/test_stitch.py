import fitz
from PIL import Image
import io

doc = fitz.open('263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf')

page10 = doc[9]
img_infos = page10.get_image_info(xrefs=True)
# Sort by y0
img_infos.sort(key=lambda x: x['bbox'][1])

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

print(f"Total stitched size: {w}x{total_h}")
stitched = Image.new('RGB', (w, total_h), 'white')
cur_y = 0
for im in images:
    stitched.paste(im, (0, cur_y))
    cur_y += im.height

stitched.save("test_stitched_p10.png")
print("Saved test_stitched_p10.png successfully!")
