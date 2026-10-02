"""
scripts/stages/builders/docx_builder.py — Master DocxBuilder Implementation

Assembles complete Word (.docx) documents from BuildContext:
- Pure config-driven page setup (margins, orientations, size)
- Exact page dimension matching from PDF inspection
- Typography and tag styling per Appendix A ([hw: ...], [stamp:], [signature])
- Compact paragraph line-spacing and padding to prevent overflow
- Professional tables with customized header fills, borders, and padding
- Faithful page preservation and 1:1 pagination
Registers with builder_registry under name 'docx_builder'.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

import docx
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Mm, Pt, RGBColor

SCRIPTS_DIR = Path(__file__).resolve().parent.parent.parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from core.context import BuildContext
from core.registry import builder_registry


def _hex_to_rgb(hex_str: str) -> RGBColor:
    cleaned = hex_str.lstrip("#")
    if len(cleaned) == 6:
        return RGBColor(int(cleaned[0:2], 16), int(cleaned[2:4], 16), int(cleaned[4:6], 16))
    return RGBColor(0, 0, 0)


def _set_cell_background(cell, fill_hex: str):
    fill_hex = fill_hex.lstrip("#")
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def _set_cell_margins(cell, top: int = 30, bottom: int = 30, left: int = 50, right: int = 50):
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


def _set_cell_borders(cell, top="single", bottom="single", left="single", right="single", sz="4", color="D3D3D3"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="{top}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="{bottom}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="{left}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="{right}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)


@builder_registry.register("docx_builder")
class MasterDocxBuilder:
    """Config-driven Builder assembling high-fidelity Word documents."""

    def build(self, ctx: BuildContext) -> Path:
        doc = docx.Document()
        cfg = ctx.config
        font_cfg = cfg.get("fonts", {})
        page_cfg = cfg.get("page", {})
        table_cfg = cfg.get("tables", {})
        tag_cfg = cfg.get("tags", {})

        default_font = font_cfg.get("default", "Aptos")
        default_size = font_cfg.get("size_pt", 9.5)
        hw_font = font_cfg.get("handwriting", "Segoe Script")

        # Set default Normal style
        style = doc.styles["Normal"]
        style.font.name = default_font
        style.font.size = Pt(default_size)
        style.font.color.rgb = RGBColor(0x20, 0x20, 0x20)
        style.paragraph_format.line_spacing = 1.05
        style.paragraph_format.space_before = Pt(0)
        style.paragraph_format.space_after = Pt(1.5)

        # Page setup
        margins = page_cfg.get("margins_mm", {"top": 12, "bottom": 12, "left": 15, "right": 15})
        landscape_ranges = page_cfg.get("landscape_ranges", [])

        inspect = ctx.inspect_result
        extract = ctx.extract_result
        total_pages = inspect.page_count or len(extract.text_by_page) or 1
        page_elements = extract.metadata.get("page_elements", {})

        for p_num in range(1, total_pages + 1):
            if p_num == 1:
                section = doc.sections[0]
            else:
                doc.add_page_break()
                section = doc.sections[-1]

            # Set exact page dimensions from inspection if available
            p_dim = inspect.page_dimensions[p_num - 1] if p_num <= len(inspect.page_dimensions) else None
            if p_dim:
                section.page_width = Pt(p_dim["width"])
                section.page_height = Pt(p_dim["height"])
                if p_dim.get("orientation") == "landscape":
                    section.orientation = WD_ORIENT.LANDSCAPE
                else:
                    section.orientation = WD_ORIENT.PORTRAIT
            else:
                section.page_width = Mm(210)
                section.page_height = Mm(297)

            section.top_margin = Mm(margins.get("top", 12))
            section.bottom_margin = Mm(margins.get("bottom", 12))
            section.left_margin = Mm(margins.get("left", 15))
            section.right_margin = Mm(margins.get("right", 15))

            # Header note
            if p_num == 1 and tag_cfg.get("handwriting_header_note", False):
                has_lots_hw = any("[hw:" in extract.text_by_page.get(p, "") for p in range(1, total_pages + 1))
                if has_lots_hw:
                    hp = section.header.paragraphs[0]
                    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    hrun = hp.add_run("[handwritten text is indicated in italics]")
                    hrun.font.size = Pt(8.0)
                    hrun.font.italic = True
                    hrun.font.color.rgb = RGBColor(0x70, 0x70, 0x70)

            # Content for page
            is_image_only = p_num in inspect.image_only_pages
            if is_image_only:
                img_path_str = extract.rendered_pages.get(p_num)
                if img_path_str and Path(img_path_str).exists():
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run()
                    usable_w = section.page_width - section.left_margin - section.right_margin
                    run.add_picture(img_path_str, width=min(usable_w, Inches(6.5)))
                else:
                    doc.add_paragraph("[blank page in source]")
                continue

            elements = page_elements.get(p_num, [])
            if elements:
                for el in elements:
                    el_type = el.get("type")
                    if el_type == "table":
                        self._render_table(doc, el.get("data", []), default_font, default_size, table_cfg)
                    elif el_type == "paragraph":
                        spans = el.get("spans")
                        align_str = el.get("alignment", "LEFT")
                        p = doc.add_paragraph()
                        if align_str == "CENTER":
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        elif align_str == "RIGHT":
                            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        else:
                            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(1.5)
                        p.paragraph_format.line_spacing = 1.05

                        if spans:
                            self._render_paragraph_spans(p, spans, default_font, default_size, hw_font)
                        else:
                            text = el.get("text", "")
                            self._render_line_with_tags(p, text, default_font, default_size, hw_font)
            else:
                # Fallback to plain text
                raw_text = extract.text_by_page.get(p_num, "").strip()
                if raw_text:
                    for line in raw_text.splitlines():
                        line_str = line.strip()
                        if line_str:
                            p = doc.add_paragraph()
                            self._render_line_with_tags(p, line_str, default_font, default_size, hw_font)
                else:
                    doc.add_paragraph("[blank page in source]")

        out_path = ctx.output_docx
        doc.save(str(out_path))
        return out_path

    def _render_table(self, doc, raw_data: List[List[Any]], default_font: str, default_size: float, table_cfg: dict):
        if not raw_data:
            return
        rows = len(raw_data)
        cols = max(len(r) for r in raw_data) if rows > 0 else 0
        if rows == 0 or cols == 0:
            return

        tbl = doc.add_table(rows=rows, cols=cols)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = table_cfg.get("autofit", True)

        header_fill = table_cfg.get("header_fill", "#0070C0")
        header_color = _hex_to_rgb(table_cfg.get("header_font_color", "#FFFFFF"))
        is_header_bold = table_cfg.get("header_bold", True)

        for r_idx, row in enumerate(raw_data):
            is_head = (r_idx == 0)
            for c_idx, cell_value in enumerate(row):
                if c_idx >= cols:
                    break
                cell = tbl.cell(r_idx, c_idx)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                _set_cell_margins(cell, top=30, bottom=30, left=50, right=50)
                _set_cell_borders(cell)

                if is_head:
                    _set_cell_background(cell, header_fill)

                cp = cell.paragraphs[0]
                cp.paragraph_format.space_before = Pt(0)
                cp.paragraph_format.space_after = Pt(0)
                cp.paragraph_format.line_spacing = 1.0
                c_text = str(cell_value or "").strip()
                c_run = cp.add_run(c_text)
                c_run.font.name = default_font
                c_run.font.size = Pt(max(7.5, default_size - 1.5 if is_head else default_size - 2))
                if is_head:
                    c_run.bold = is_header_bold
                    c_run.font.color.rgb = header_color

    def _render_paragraph_spans(self, paragraph, spans: List[Dict[str, Any]], default_font: str, default_size: float, hw_font: str):
        for sp in spans:
            stext = sp.get("text", "")
            if not stext:
                continue
            run = paragraph.add_run(stext)
            run.font.name = sp.get("font") or default_font
            run.font.size = Pt(sp.get("size") or default_size)
            if sp.get("bold"):
                run.bold = True
            if sp.get("italic"):
                run.italic = True
            color_int = sp.get("color", 0)
            if color_int and color_int != 0:
                r = (color_int >> 16) & 0xFF
                g = (color_int >> 8) & 0xFF
                b = color_int & 0xFF
                run.font.color.rgb = RGBColor(r, g, b)

    def _render_line_with_tags(self, paragraph, line_str: str, font_name: str, font_size: float, hw_font: str):
        pattern = r"(\[hw:\s*.*?\]|\[stamp:\s*.*?\]|\[logo:\s*.*?\]|\[icon\]|\[signature\]|\[initials\]|\[seal:\])"
        parts = re.split(pattern, line_str)

        for part in parts:
            if not part:
                continue
            if part.startswith("[hw:") and part.endswith("]"):
                content = part[4:-1].strip()
                run = paragraph.add_run(f"[hw: {content}]")
                run.italic = True
                run.font.name = hw_font
                run.font.size = Pt(font_size)
            elif part.startswith("[stamp:") and part.endswith("]"):
                run = paragraph.add_run(part)
                run.bold = True
                run.font.color.rgb = RGBColor(0x8B, 0x00, 0x00)
            elif part.startswith("[logo:") and part.endswith("]"):
                logo_payload = part[6:-1].strip()
                if not logo_payload:
                    run = paragraph.add_run("[icon]")
                else:
                    run = paragraph.add_run(f"[logo: {logo_payload}]")
                run.bold = True
                run.font.color.rgb = RGBColor(0x00, 0x35, 0x84)
            elif part == "[icon]":
                run = paragraph.add_run("[icon]")
                run.italic = True
                run.font.color.rgb = RGBColor(0x50, 0x50, 0x50)
            elif part in ("[signature]", "[initials]", "[seal:]"):
                run = paragraph.add_run(part)
                run.italic = True
                run.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
            else:
                run = paragraph.add_run(part)
                run.font.name = font_name
                run.font.size = Pt(font_size)

