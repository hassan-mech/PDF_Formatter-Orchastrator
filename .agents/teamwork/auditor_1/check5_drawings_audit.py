import pymupdf
import docx
import xml.etree.ElementTree as ET

pdf_path = r"finished/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf"
docx_path = r"output/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.docx"

print("=== CHECK 5: VECTOR DRAWINGS & ANTI-HALLUCINATION AUDIT ===")

doc_pdf = pymupdf.open(pdf_path)
print(f"Total PDF pages: {doc_pdf.page_count}")

for i in range(doc_pdf.page_count):
    page = doc_pdf[i]
    drawings = page.get_drawings()
    print(f"\n--- PDF Page {i+1} Vector Drawings ({len(drawings)} drawings) ---")
    for d_idx, d in enumerate(drawings):
        rect = d.get("rect")
        color = d.get("color")
        fill = d.get("fill")
        items = d.get("items", [])
        print(f"  Drawing {d_idx}: rect={rect}, color={color}, fill={fill}, items_count={len(items)}")

# Now inspect DOCX for paragraph borders (w:pBdr) and table borders
print("\n--- DOCX Paragraph Borders (w:pBdr) ---")
doc = docx.Document(docx_path)
ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

border_paras = []
for p_idx, p in enumerate(doc.paragraphs):
    pBdr = p._p.find(".//w:pBdr", ns)
    if pBdr is not None:
        borders = [c.tag.split("}")[-1] + ": " + str(c.attrib) for c in pBdr]
        border_paras.append((p_idx, p.text[:40], borders))

print(f"Paragraphs with borders: {len(border_paras)}")
for p_idx, text, borders in border_paras:
    print(f"  P{p_idx} (text='{text}'): {borders}")
