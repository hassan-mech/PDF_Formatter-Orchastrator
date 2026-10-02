import pymupdf

doc_pdf = pymupdf.open("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
print("=== PDF VECTOR DRAWINGS COMPLETE ANALYSIS ===")
for pno in range(len(doc_pdf)):
    page = doc_pdf[pno]
    drawings = page.get_drawings()
    print(f"\n--- Page {pno+1} ({page.rect.width:.1f} x {page.rect.height:.1f} pt) ---")
    print(f"Total drawings: {len(drawings)}")
    for idx, d in enumerate(drawings):
        rect = d.get("rect")
        color = d.get("color")
        fill = d.get("fill")
        items = d.get("items", [])
        # Check if line or rect
        types = [item[0] for item in items]
        print(f"  Drawing {idx}: rect=({rect.x0:.1f}, {rect.y0:.1f}, {rect.x1:.1f}, {rect.y1:.1f}), color={color}, fill={fill}, types={types}")
doc_pdf.close()
