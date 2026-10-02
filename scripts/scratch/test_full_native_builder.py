"""
scripts/scratch/test_full_native_builder.py
Refactoring the Journal Article Builder to use Word's native 2-column feature (<w:cols>)
and continuous section breaks instead of layout tables!
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))

import io
import pymupdf
from PIL import Image
import docx
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx.enum.text import WD_BREAK, WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

from scripts.safe_word_export import convert_docx_to_pdf

COLOR_NAVY = RGBColor(0x00, 0x35, 0x84)
COLOR_CRIMSON = RGBColor(0xDB, 0x1D, 0x43)
COLOR_CHARCOAL = RGBColor(0x23, 0x1F, 0x20)
COLOR_MUTED = RGBColor(0x55, 0x55, 0x55)
COLOR_BLUE_LINK = RGBColor(0x00, 0x55, 0xA5)
HEX_PINK_FILL = "FCECEF"
HEX_CRIMSON = "DB1D43"

def set_columns(section, num_cols: int = 1, col_widths_pt: list[float] = None, space_pt: float = 19.9):
    sectPr = section._sectPr
    for old_cols in sectPr.findall(qn('w:cols')):
        sectPr.remove(old_cols)
    w_ns = nsdecls('w')
    space_dxa = int(space_pt * 20)
    
    if col_widths_pt and len(col_widths_pt) > 1:
        col_elements = ""
        for i, w_pt in enumerate(col_widths_pt):
            w_dxa = int(w_pt * 20)
            if i < len(col_widths_pt) - 1:
                col_elements += f'<w:col {w_ns} w:w="{w_dxa}" w:space="{space_dxa}"/>\n'
            else:
                col_elements += f'<w:col {w_ns} w:w="{w_dxa}"/>\n'
        cols_xml = f'<w:cols {w_ns} w:num="{len(col_widths_pt)}" w:equalWidth="0" w:space="{space_dxa}">\n{col_elements}</w:cols>'
    elif num_cols > 1:
        cols_xml = f'<w:cols {w_ns} w:num="{num_cols}" w:space="{space_dxa}"/>'
    else:
        cols_xml = f'<w:cols {w_ns} w:num="1"/>'
    sectPr.append(parse_xml(cols_xml))

print("Helper functions ready.")
