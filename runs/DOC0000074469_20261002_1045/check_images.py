import fitz
import os

pdf_path = "input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf"
doc = fitz.open(pdf_path)
out_dir = "runs/DOC0000074469_20261002_1045/extract/images"
os.makedirs(out_dir, exist_ok=True)

print(f"Total pages: {len(doc)}")
for pno, page in enumerate(doc, start=1):
    img_list = page.get_images(full=True)
    print(f"Page {pno}: {len(img_list)} images")
    for idx, img_info in enumerate(img_list):
        xref = img_info[0]
        base_img = doc.extract_image(xref)
        img_bytes = base_img["image"]
        img_ext = base_img["ext"]
        w = base_img["width"]
        h = base_img["height"]
        fname = f"p{pno}_img{idx}_{xref}.{img_ext}"
        fpath = os.path.join(out_dir, fname)
        with open(fpath, "wb") as f:
            f.write(img_bytes)
        print(f"  Saved {fname}: {w}x{h} ({img_ext})")
