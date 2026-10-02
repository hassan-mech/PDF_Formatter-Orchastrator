import pymupdf
from PIL import Image, ImageChops
import io

orig = pymupdf.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
conv = pymupdf.open("temp/latest_docx.pdf")

dpi = 150
mat = pymupdf.Matrix(dpi / 72, dpi / 72)

for i in range(7):
    pix_a = orig[i].get_pixmap(matrix=mat, alpha=False)
    pix_b = conv[i].get_pixmap(matrix=mat, alpha=False)
    img_a = Image.open(io.BytesIO(pix_a.tobytes("png"))).convert("RGB")
    img_b = Image.open(io.BytesIO(pix_b.tobytes("png"))).convert("RGB")
    
    w = min(img_a.width, img_b.width)
    h = min(img_a.height, img_b.height)
    diff = ImageChops.difference(img_a.crop((0,0,w,h)), img_b.crop((0,0,w,h)))
    hist = diff.convert("L").histogram()
    bad_pixels = sum(hist[11:])
    total = w * h
    sim = 1.0 - (bad_pixels / total)
    print(f"Page {i+1} visual similarity: {sim:.2%}")
