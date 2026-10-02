import fitz
from PIL import Image, ImageChops
from io import BytesIO

def render_pdf_page(doc, idx, dpi=120):
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = doc[idx].get_pixmap(matrix=mat, alpha=False)
    return Image.open(BytesIO(pix.tobytes("png"))).convert("RGB")

def pixel_similarity(img_a, img_b, tol=10):
    w = min(img_a.width, img_b.width)
    h = min(img_a.height, img_b.height)
    a = img_a.resize((w, h), Image.LANCZOS)
    b = img_b.resize((w, h), Image.LANCZOS)
    diff = ImageChops.difference(a, b)
    if tol > 0:
        diff = diff.point(lambda p: 255 if p > tol else 0)
    hist = diff.histogram()
    total = w * h * 3
    nonzero = sum(hist[1:256]) + sum(hist[257:512]) + sum(hist[513:768])
    return max(0.0, 1.0 - nonzero / total)

orig_doc = fitz.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
conv_doc = fitz.open("runs/DOC0000074469_20261002_1045/build/verify_test.pdf")

for i in range(7):
    a = render_pdf_page(orig_doc, i)
    b = render_pdf_page(conv_doc, i)
    sim10 = pixel_similarity(a, b, tol=10)
    sim25 = pixel_similarity(a, b, tol=25)
    print(f"Page {i+1}: sim(tol=10)={sim10:.1%}, sim(tol=25)={sim25:.1%}")
