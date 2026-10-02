import fitz

orig = fitz.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
conv = fitz.open("runs/DOC0000074469_20261002_1045/build/verify_test.pdf")

for pno in [3, 4, 5]: # Page 4, 5, 6 (0-indexed 3, 4, 5)
    print(f"\n--- PAGE {pno+1} ---")
    print("ORIGINAL text blocks:")
    for b in orig[pno].get_text("blocks")[:4]:
        print(f"  [{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}]: {b[4].strip()[:50]}")
    print("CONVERTED text blocks:")
    for b in conv[pno].get_text("blocks")[:4]:
        print(f"  [{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}]: {b[4].strip()[:50]}")
