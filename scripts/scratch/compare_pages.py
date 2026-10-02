import pymupdf
from pathlib import Path

orig = pymupdf.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
conv = pymupdf.open("temp/test_export.pdf")

print(f"Orig pages: {len(orig)}, Conv pages: {len(conv)}")

for i in range(min(len(orig), len(conv))):
    p_orig = orig[i]
    p_conv = conv[i]
    print(f"\n--- PAGE {i+1} ---")
    print(f"Orig rect: {p_orig.rect}, Conv rect: {p_conv.rect}")
    t_orig = p_orig.get_text("text").strip().split("\n")
    t_conv = p_conv.get_text("text").strip().split("\n")
    print(f"Orig first line: {t_orig[0] if t_orig else 'EMPTY'}")
    print(f"Conv first line: {t_conv[0] if t_conv else 'EMPTY'}")
    print(f"Orig last line: {t_orig[-1] if t_orig else 'EMPTY'}")
    print(f"Conv last line: {t_conv[-1] if t_conv else 'EMPTY'}")
