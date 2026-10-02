import pymupdf

doc = pymupdf.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
print("=== PAGE 1 ABSTRACT ===")
p1 = doc[0]
for b in p1.get_text("blocks"):
    if b[1] > 500 and b[0] < 400:
        print(f"[{b[1]:.1f} - {b[3]:.1f}] {b[4]}")

print("=== PAGE 2 TOP ===")
p2 = doc[1]
for b in p2.get_text("blocks"):
    if b[1] < 200:
        print(f"[{b[1]:.1f} - {b[3]:.1f}] {b[4]}")
