import fitz

doc = fitz.open("runs/DOC0000074469_20261002_1045/build/verify_test.pdf")

for i, page in enumerate(doc):
    print(f"\n================ TEST PAGE {i+1} ================")
    blocks = page.get_text("blocks")
    for b in sorted(blocks, key=lambda x: (x[1], x[0])):
        first_line = b[4].strip().replace("\n", " ")[:70]
        if first_line:
            print(f"  [{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}]: {first_line}")
