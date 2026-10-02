import os
from PIL import Image

folder = "runs/DOC0000074469_20261002_1045/extract"
for f in ["p1_banner.png", "p1_eng_title.png", "p1_eng_abstract.png", "p2_eng_abstract.png"]:
    p = os.path.join(folder, f)
    if os.path.exists(p):
        img = Image.open(p)
        print(f"{f}: size={img.size}, mode={img.mode}")
