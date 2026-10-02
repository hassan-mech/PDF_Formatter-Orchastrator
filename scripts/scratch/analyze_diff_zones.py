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
    diff = ImageChops.difference(img_a.crop((0,0,w,h)), img_b.crop((0,0,w,h))).convert("L")
    
    # Divide page vertically into 3 zones: Top (0-33%), Mid (33-66%), Bot (66-100%)
    z_h = h // 3
    top_diff = sum(sum(1 for p in diff.crop((0, 0, w, z_h)).getdata() if p > 10) for _ in [1])
    mid_diff = sum(sum(1 for p in diff.crop((0, z_h, w, 2*z_h)).getdata() if p > 10) for _ in [1])
    bot_diff = sum(sum(1 for p in diff.crop((0, 2*z_h, w, h)).getdata() if p > 10) for _ in [1])
    total_diff = top_diff + mid_diff + bot_diff
    print(f"Page {i+1}: total diff pixels={total_diff:,} -> Top: {top_diff/total_diff:.1%}, Mid: {mid_diff/total_diff:.1%}, Bot: {bot_diff/total_diff:.1%}")
