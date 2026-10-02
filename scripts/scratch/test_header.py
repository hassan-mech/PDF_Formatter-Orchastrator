import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_BREAK
from docx.shared import Inches, Pt, RGBColor
from scripts.safe_word_export import convert_docx_to_pdf
import pymupdf

doc = docx.Document()
section = doc.sections[0]
section.top_margin = Inches(0.58)
section.left_margin = Inches(0.98)
section.right_margin = Inches(1.34)
section.header_distance = Inches(0.40)

header = section.header
p_h = header.paragraphs[0]
p_h.text = "Left Running Header\tRight Header"
# Add tab stop at right margin
tab_stops = p_h.paragraph_format.tab_stops
tab_stops.add_tab_stop(Inches(5.95), WD_TAB_ALIGNMENT.RIGHT)

p_b1 = doc.add_paragraph("Body paragraph 1")

out_path = Path("scripts/scratch/test_header.docx")
doc.save(str(out_path))

pdf_path = Path("scripts/scratch/test_header.pdf")
convert_docx_to_pdf(out_path.resolve(), pdf_path.resolve(), timeout_sec=20)

pdf = pymupdf.open(str(pdf_path))
print("Header blocks:")
for b in pdf[0].get_text("blocks"):
    print(f"  [{b[0]:.1f}, {b[1]:.1f}, {b[2]:.1f}, {b[3]:.1f}]: {b[4].strip()}")
