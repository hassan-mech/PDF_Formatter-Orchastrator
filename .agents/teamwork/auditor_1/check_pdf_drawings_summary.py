import pymupdf

pdf_path = r"finished/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf"
doc = pymupdf.open(pdf_path)

for pno in range(doc.page_count):
    page = doc[pno]
    drawings = page.get_drawings()
    print(f"Page {pno+1}: {len(drawings)} drawings")
    for idx, d in enumerate(drawings[:5]):
        print(f"  P{pno+1} D{idx}: rect={d.get('rect')}, color={d.get('color')}, fill={d.get('fill')}")
