import fitz

doc = fitz.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")

print("--- Page 1 text around crops ---")
p1 = doc[0]
for b in p1.get_text("blocks"):
    if b[1] < 50 or (180 <= b[1] <= 250) or (510 <= b[1] <= 710):
        print(f"[{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}]: {b[4].strip()}")

print("\n--- Page 2 text around crops ---")
p2 = doc[1]
for b in p2.get_text("blocks"):
    if b[1] < 200:
        print(f"[{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}]: {b[4].strip()}")
