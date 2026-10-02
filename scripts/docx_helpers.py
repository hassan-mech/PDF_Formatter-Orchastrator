"""
docx_helpers.py — Shared helper functions for Word document (.docx) generation.

Provides low-level OOXML utilities for cell formatting, paragraph formatting,
run merging, and section/page size configuration.

Usage:
    from docx_helpers import set_cell_borders, set_cell_margins, merge_runs,
                             set_paragraph_format, set_page_size
"""

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn


# ---------------------------------------------------------------------------
# Cell utilities
# ---------------------------------------------------------------------------

def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    """
    Set the borders of a table cell using OOXML.

    Each border arg is either None (no border) or a dict with optional keys:
      val   : border style (default 'single')
      sz    : border width in eighths of a point (default '4')
      color : hex color code (default 'auto')
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    borders = {'top': top, 'bottom': bottom, 'left': left, 'right': right}
    for b_name, b_val in borders.items():
        if b_val:
            tag = (
                f'<w:{b_name} {nsdecls("w")} '
                f'w:val="{b_val.get("val", "single")}" '
                f'w:sz="{b_val.get("sz", "4")}" '
                f'w:space="0" '
                f'w:color="{b_val.get("color", "auto")}"/>'
            )
            tcBorders.append(parse_xml(tag))
        else:
            tag = f'<w:{b_name} {nsdecls("w")} w:val="none"/>'
            tcBorders.append(parse_xml(tag))
    tcPr.append(tcBorders)


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """
    Set the inner margins of a table cell (values in twips / DXA units).

    Default values approximate typical Word cell margins.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


# ---------------------------------------------------------------------------
# Run utilities
# ---------------------------------------------------------------------------

def _runs_have_same_format(r1, r2) -> bool:
    """Return True if two runs share identical rPr XML (or both lack it)."""
    rpr1 = r1._r.find(qn("w:rPr"))
    rpr2 = r2._r.find(qn("w:rPr"))
    if rpr1 is None and rpr2 is None:
        return True
    if rpr1 is None or rpr2 is None:
        return False
    from lxml import etree
    return etree.tostring(rpr1) == etree.tostring(rpr2)


def merge_runs(paragraph):
    """
    Merge adjacent runs that have identical run-properties (rPr) within a paragraph.

    This reduces fragmentation caused by copy-paste or programmatic construction,
    which improves file size and rendering fidelity.
    Returns the number of merges performed.
    """
    runs = paragraph.runs
    if len(runs) < 2:
        return 0
    merges = 0
    i = 0
    while i < len(runs) - 1:
        r_cur = runs[i]
        r_next = runs[i + 1]
        if _runs_have_same_format(r_cur, r_next):
            # Append next run's text to current run and remove next
            r_cur.text = (r_cur.text or "") + (r_next.text or "")
            r_next._r.getparent().remove(r_next._r)
            # Refresh runs list after mutation
            runs = paragraph.runs
            merges += 1
        else:
            i += 1
    return merges


# ---------------------------------------------------------------------------
# Paragraph format utility
# ---------------------------------------------------------------------------

_ALIGN_MAP = {
    "left":    WD_ALIGN_PARAGRAPH.LEFT,
    "right":   WD_ALIGN_PARAGRAPH.RIGHT,
    "center":  WD_ALIGN_PARAGRAPH.CENTER,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def set_paragraph_format(
    p,
    align=None,
    space_before=None,
    space_after=None,
    line_spacing=None,
    left_indent=None,
    first_line_indent=None,
):
    """
    Apply common paragraph formatting properties.

    Parameters
    ----------
    p               : docx Paragraph object
    align           : 'left'|'right'|'center'|'justify' or None
    space_before    : space before paragraph in points (float) or None
    space_after     : space after paragraph in points (float) or None
    line_spacing    : line spacing multiplier (e.g. 1.15) or None
    left_indent     : left indent in inches (float) or None
    first_line_indent: first-line indent in inches (float) or None
    """
    fmt = p.paragraph_format
    if align is not None:
        p.alignment = _ALIGN_MAP.get(align.lower(), WD_ALIGN_PARAGRAPH.LEFT)
    if space_before is not None:
        fmt.space_before = Pt(space_before)
    if space_after is not None:
        fmt.space_after = Pt(space_after)
    if line_spacing is not None:
        fmt.line_spacing = line_spacing
    if left_indent is not None:
        fmt.left_indent = Inches(left_indent)
    if first_line_indent is not None:
        fmt.first_line_indent = Inches(first_line_indent)


# ---------------------------------------------------------------------------
# Section / page size utility
# ---------------------------------------------------------------------------

def set_page_size(section, width_inches: float, height_inches: float, margins=(1, 1, 1, 1)):
    """
    Set the page size and margins of a document section.

    Parameters
    ----------
    section       : docx Section object (e.g. doc.sections[0])
    width_inches  : page width in inches
    height_inches : page height in inches
    margins       : tuple (top, right, bottom, left) in inches (default: 1 inch all)
    """
    section.page_width  = Inches(width_inches)
    section.page_height = Inches(height_inches)
    top, right, bottom, left = margins
    section.top_margin    = Inches(top)
    section.right_margin  = Inches(right)
    section.bottom_margin = Inches(bottom)
    section.left_margin   = Inches(left)
