import os
from PIL import Image

for p in range(1, 42):
    fname = f"rendered_pages/page_{p:02d}.png"
    if os.path.exists(fname):
        im = Image.open(fname)
        print(f"Page {p:02d}: size={im.size}")
