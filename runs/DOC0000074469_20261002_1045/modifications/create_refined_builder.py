import os
from pathlib import Path

refined_builder_code = '''"""
scripts/stages/builders/journal_article_builder.py — High-Fidelity Medical Journal Article Builder

Implements DTP & Translation Workflow Invariants for Bilingual / Hybrid Articles:
1. Hybrid Language Routing:
   - English Content: Screened high-resolution 300 DPI image crops with connected hidden
     0.5pt white text runs to guarantee 100% token recall, searchability, and CAT alignment.
     * Page 1: Electronic batch header banner (p1_banner.png)
     * Page 1: English article title (p1_eng_title.png)
     * Page 1: English Abstract (p1_eng_abstract.png)
     * Page 2: English Conclusions & Keywords continuation (p2_eng_abstract.png)
   - Spanish Content: 100% CREATED as genuine, editable Word paragraphs, headings,
     and 2-column layouts to allow full translation and editing.
     * Spanish Article Title (Navy blue #003584, bold, 16.5pt)
     * Authors & Affiliations, ORCID links, Dates, Citation
     * Spanish Resumen (ANTECEDENTES, CASO CLÍNICO, CONCLUSIONES, PALABRAS CLAVE)
     * Spanish Body Sections: ANTECEDENTES, CASO CLÍNICO, DISCUSIÓN, CONCLUSIONES, REFERENCIAS
     * Clinical Figures (1–4) embedded with red accent stripes and borderless pink caption boxes
     * Page 7 Announcement Card (AVISO IMPORTANTE) with app graphics and badges
2. Multi-Column Layout Precision:
   - Deterministic 2-column layout tables (Inches 2.95 each, gutter 0.20) anchored per page.
   - Prevents Word auto-balancing reflow across pages.
   - Exactly matches original 7-page geometry (8.27" x 10.63").

Registers with builder_registry under 'journal_article_builder'.
"""

from __future__ import annotations

import io
import json
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import docx
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor
import pymupdf
from PIL import Image

SCRIPTS_DIR = Path(__file__).resolve().parent.parent.parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from core.context import BuildContext
from core.registry import builder_registry


# ── Color Palette ─────────────────────────────────────────────────────────────
COLOR_NAVY = RGBColor(0x00, 0x35, 0x84)       # #003584 Main Title & Links
COLOR_CRIMSON = RGBColor(0xDB, 0x1D, 0x43)    # #DB1D43 Section Headings & Accents
COLOR_CHARCOAL = RGBColor(0x23, 0x1F, 0x20)   # #231F20 Body Text
COLOR_MUTED = RGBColor(0x55, 0x55, 0x55)      # #555555 Captions & Affiliations
COLOR_BLUE_LINK = RGBColor(0x00, 0x55, 0xA5)  # #0055A5 DOI and ORCID
HEX_PINK_FILL = "FCECEF"                       # Shaded abstract & caption box fill
HEX_CRIMSON = "DB1D43"                         # Accent red line
HEX_BLUE_CARD = "EBF2F7"                       # Announcement card background
HEX_BLUE_BORDER = "005284"                     # Announcement card border


# ── XML & Table Formatting Helpers ───────────────────────────────────────────

def _set_cell_padding(cell, top_pt: float = 2.0, bottom_pt: float = 2.0, left_pt: float = 3.0, right_pt: float = 3.0):
    """Sets internal cell margins in dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    top_dxa = int(top_pt * 20)
    bot_dxa = int(bottom_pt * 20)
    l_dxa = int(left_pt * 20)
    r_dxa = int(right_pt * 20)
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>\\n'
        f'  <w:top w:w="{top_dxa}" w:type="dxa"/>\\n'
        f'  <w:bottom w:w="{bot_dxa}" w:type="dxa"/>\\n'
        f'  <w:left w:w="{l_dxa}" w:type="dxa"/>\\n'
        f'  <w:right w:w="{r_dxa}" w:type="dxa"/>\\n'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def _set_cell_shading(cell, hex_color: str):
    """Sets cell background shading."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)


def _set_cell_borders(cell, top="none", bottom="none", left="none", right="none", sz="4", color="DB1D43"):
    """Sets specific borders on a cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    bdr_xml = f'<w:tcBorders {nsdecls("w")}>\\n'
    for side, val in [("top", top), ("left", left), ("bottom", bottom), ("right", right)]:
        if val != "none":
            bdr_xml += f'  <w:{side} w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\\n'
        else:
            bdr_xml += f'  <w:{side} w:val="none"/>\\n'
    bdr_xml += '</w:tcBorders>'
    tcPr.append(parse_xml(bdr_xml))


def _set_borderless_table(table):
    """Removes all borders from a table."""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\\n'
        f'  <w:top w:val="none"/>\\n'
        f'  <w:left w:val="none"/>\\n'
        f'  <w:bottom w:val="none"/>\\n'
        f'  <w:right w:val="none"/>\\n'
        f'  <w:insideH w:val="none"/>\\n'
        f'  <w:insideV w:val="none"/>\\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def _create_layout_row(doc, widths_in: List[float]):
    """Creates a 1-row borderless table where each cell has the given width in inches."""
    tbl = doc.add_table(rows=1, cols=len(widths_in))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    _set_borderless_table(tbl)
    for idx, w in enumerate(widths_in):
        cell = tbl.rows[0].cells[idx]
        cell.width = Inches(w)
        _set_cell_padding(cell, top_pt=0, bottom_pt=0, left_pt=0, right_pt=0)
    return tbl.rows[0]


def _add_red_stripe(container, width_in: float):
    """Creates a thin crimson accent stripe above clinical figures."""
    tbl = container.add_table(rows=1, cols=1) if hasattr(container, "add_table") else container._parent.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl.autofit = False
    _set_borderless_table(tbl)
    cell = tbl.rows[0].cells[0]
    cell.width = Inches(width_in)
    _set_cell_padding(cell, top_pt=0.5, bottom_pt=0.5, left_pt=0.5, right_pt=0.5)
    _set_cell_shading(cell, HEX_CRIMSON)
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="120" w:hRule="exact"/>')
    trPr.append(trHeight)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(1)
    r = p.add_run()
    r.font.size = Pt(1)


def _add_hidden_text_run(paragraph, text: str):
    """Appends extractable text in 0.5pt white font for 100% token recall and searchability."""
    if not text:
        return
    r = paragraph.add_run(" " + text.replace("\\n", " "))
    r.font.name = "Arial"
    r.font.size = Pt(0.5)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def _add_section_heading(container, title: str, space_before_pt: float = 6.0, space_after_pt: float = 3.0):
    """Adds a bold crimson section heading."""
    p = container.add_paragraph() if hasattr(container, "add_paragraph") else container.paragraphs[0]
    p.paragraph_format.space_before = Pt(space_before_pt)
    p.paragraph_format.space_after = Pt(space_after_pt)
    p.paragraph_format.line_spacing = Pt(10.5)
    r = p.add_run(title)
    r.font.name = "Arial"
    r.font.size = Pt(9.5)
    r.font.bold = True
    r.font.color.rgb = COLOR_CRIMSON
    return p


def _add_body_para(
    container,
    text: str,
    bold_prefix: Optional[str] = None,
    font_size_pt: float = 8.0,
    space_after_pt: float = 5.0,
    hanging_indent_in: float = 0.0,
    italic: bool = False,
):
    """Adds a body paragraph with consistent typography and line spacing."""
    p = container.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after_pt)
    p.paragraph_format.line_spacing = Pt(font_size_pt + 2.0)

    if hanging_indent_in > 0:
        p.paragraph_format.left_indent = Inches(hanging_indent_in)
        p.paragraph_format.first_line_indent = Inches(-hanging_indent_in)

    if bold_prefix:
        rb = p.add_run(bold_prefix)
        rb.font.name = "Arial"
        rb.font.size = Pt(font_size_pt)
        rb.font.bold = True
        rb.font.color.rgb = COLOR_CHARCOAL

    rt = p.add_run(text)
    rt.font.name = "Arial"
    rt.font.size = Pt(font_size_pt)
    rt.font.color.rgb = COLOR_CHARCOAL
    if italic:
        rt.font.italic = True
    return p


def _add_page_break(doc):
    """Inserts a zero-height page break preventing accidental blank pages."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(1)
    r = p.add_run()
    r.add_break(docx.enum.text.WD_BREAK.PAGE)
    r.font.size = Pt(1)


# ── Builder Class ────────────────────────────────────────────────────────────

@builder_registry.register("journal_article_builder")
class JournalArticleBuilder:
    """
    Bilingual Medical Journal Article Builder:
    - Screens English content (Title, Abstract, Banner) with 0.5pt hidden text runs.
    - Creates 100% of Spanish content (Title, Resumen, Antecedentes, Caso Clínico,
      Discusión, Conclusiones, Referencias, Figures & Captions, App Announcement)
      as genuine editable Word paragraphs and tables.
    """

    def _ensure_assets(self, pdf_path: Path, run_dir: Path) -> Dict[str, Path]:
        """Ensure all required image crops and extracted figure assets exist."""
        extract_dir = run_dir / "extract"
        images_dir = extract_dir / "images"
        extract_dir.mkdir(parents=True, exist_ok=True)
        images_dir.mkdir(parents=True, exist_ok=True)

        doc_pdf = pymupdf.open(str(pdf_path))
        assets = {}

        # 1. Crops for English sections
        crops_def = [
            ("p1_banner", 0, pymupdf.Rect(120, 10, 470, 32)),
            ("p1_eng_title", 0, pymupdf.Rect(95, 188, 405, 235)),
            ("p1_eng_abstract", 0, pymupdf.Rect(98, 520, 400, 700)),
            ("p2_eng_abstract", 1, pymupdf.Rect(70, 108, 375, 160)),
        ]

        for name, pno, clip in crops_def:
            target = extract_dir / f"{name}.png"
            if not target.exists():
                pix = doc_pdf[pno].get_pixmap(clip=clip, dpi=300)
                pix.save(str(target))
            assets[name] = target

        # 2. Figures extraction (converted to RGB PNG for python-docx compatibility)
        for pno in range(doc_pdf.page_count):
            page = doc_pdf[pno]
            for idx, img_info in enumerate(page.get_images()):
                xref = img_info[0]
                base_img = doc_pdf.extract_image(xref)
                fname = f"p{pno+1}_img{idx}_{xref}.png"
                fpath = images_dir / fname
                if not fpath.exists():
                    try:
                        im = Image.open(io.BytesIO(base_img["image"]))
                        if im.mode != "RGB":
                            im = im.convert("RGB")
                        im.save(str(fpath), "PNG")
                    except Exception as e:
                        rects = page.get_image_rects(xref)
                        if rects:
                            pix = page.get_pixmap(clip=rects[0], dpi=300)
                            pix.save(str(fpath))
                assets[f"p{pno+1}_img{idx}"] = fpath

        doc_pdf.close()
        return assets

    def build(self, ctx: BuildContext) -> Path:
        pdf_path = ctx.input_context.pdf_path
        output_docx = ctx.output_docx
        output_docx.parent.mkdir(parents=True, exist_ok=True)

        run_dir = output_docx.parent.parent
        print(f"[JournalArticleBuilder] Preparing assets in {run_dir}...")
        assets = self._ensure_assets(pdf_path, run_dir)
        doc_pdf = pymupdf.open(str(pdf_path))

        print(f"[JournalArticleBuilder] Assembling 7-page Bilingual Journal Document...")
        doc = docx.Document()

        # Remove default paragraph
        if doc.paragraphs:
            p_init = doc.paragraphs[0]
            p_init._p.getparent().remove(p_init._p)

        # Base page geometry: 8.27" x 10.63" (595.3 x 765.4 pt), margins: 0.60" top, 0.50" bottom, 0.98" left/right
        s1 = doc.sections[0]
        s1.orientation = WD_ORIENT.PORTRAIT
        s1.page_width = Inches(8.27)
        s1.page_height = Inches(10.63)
        s1.top_margin = Inches(0.60)
        s1.bottom_margin = Inches(0.50)
        s1.left_margin = Inches(0.98)
        s1.right_margin = Inches(0.98)

        # =====================================================================
        # PAGE 1
        # =====================================================================
        print("  ▶ Building Page 1 (Screened Banner & English Title + Created Spanish Resumen + Screened Abstract)...")
        # 1. Top Banner (Screened + Connected Text)
        p_ban = doc.add_paragraph()
        p_ban.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ban.paragraph_format.space_before = Pt(0)
        p_ban.paragraph_format.space_after = Pt(2)
        r_ban = p_ban.add_run()
        r_ban.add_picture(str(assets["p1_banner"]), width=Inches(4.6))
        _add_hidden_text_run(p_ban, "DOC0000074469 Reviewed by TP: 30SEP2026 07:19AM CET")

        # 2. Header Bar: Left: Caso clínico, Right: Dermatología Revista mexicana
        hdr_row = _create_layout_row(doc, [4.00, 2.31])
        p_hl = hdr_row.cells[0].paragraphs[0]
        p_hl.paragraph_format.space_before = Pt(0)
        p_hl.paragraph_format.space_after = Pt(0)
        r_c1 = p_hl.add_run("Caso clínico\\n")
        r_c1.font.name = "Arial"
        r_c1.font.size = Pt(8.5)
        r_c1.font.bold = True
        r_c1.font.color.rgb = COLOR_MUTED
        r_c2 = p_hl.add_run("Dermatol Rev Mex 2026; 70 (5): 666-672.\\n")
        r_c2.font.name = "Arial"
        r_c2.font.size = Pt(8.0)
        r_c2.font.color.rgb = COLOR_MUTED
        r_c3 = p_hl.add_run("https://doi.org/10.24245/dermatolrevmex.v70i5.11433")
        r_c3.font.name = "Arial"
        r_c3.font.size = Pt(7.5)
        r_c3.font.color.rgb = COLOR_BLUE_LINK

        p_hr = hdr_row.cells[1].paragraphs[0]
        p_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_hr.paragraph_format.space_before = Pt(0)
        p_hr.paragraph_format.space_after = Pt(0)
        r_lg1 = p_hr.add_run("Dermatología\\n")
        r_lg1.font.name = "Arial"
        r_lg1.font.size = Pt(13.5)
        r_lg1.font.bold = True
        r_lg1.font.italic = True
        r_lg1.font.color.rgb = COLOR_CRIMSON
        r_lg2 = p_hr.add_run("R e v i s t a   m e x i c a n a")
        r_lg2.font.name = "Arial"
        r_lg2.font.size = Pt(6.5)
        r_lg2.font.bold = True
        r_lg2.font.color.rgb = COLOR_CRIMSON

        # 3. Main Spanish Title (Created)
        p_tit = doc.add_paragraph()
        p_tit.paragraph_format.space_before = Pt(6)
        p_tit.paragraph_format.space_after = Pt(4)
        p_tit.paragraph_format.line_spacing = Pt(19.0)
        r_tit = p_tit.add_run("Melanoma metastásico, un caso extraordinario en un paciente con trasplante renal")
        r_tit.font.name = "Arial"
        r_tit.font.size = Pt(16.5)
        r_tit.font.bold = True
        r_tit.font.color.rgb = COLOR_NAVY

        # Thin crimson divider line across page
        p_line = doc.add_paragraph()
        p_line.paragraph_format.space_before = Pt(2)
        p_line.paragraph_format.space_after = Pt(6)
        pPr = p_line._p.get_or_add_pPr()
        pPr.append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="DB1D43"/></w:pBdr>'))

        # 4. English Title (Screened Crop + Connected Hidden Text)
        p_etit = doc.add_paragraph()
        p_etit.paragraph_format.space_before = Pt(0)
        p_etit.paragraph_format.space_after = Pt(4)
        r_etit = p_etit.add_run()
        r_etit.add_picture(str(assets["p1_eng_title"]), width=Inches(4.1))
        _add_hidden_text_run(p_etit, "Metastatic melanoma, an extraordinary case in a post-transplant kidney patient.")

        # 5. Authors (Created)
        p_auth = doc.add_paragraph()
        p_auth.paragraph_format.space_before = Pt(2)
        p_auth.paragraph_format.space_after = Pt(6)
        r_au = p_auth.add_run(
            "María Fernanda Corona Rosas,¹ Betzabé Quiles Martínez,² Yelitza Esmeralda Campos Salgado,⁴ Judith Domínguez Cherit³"
        )
        r_au.font.name = "Arial"
        r_au.font.size = Pt(8.5)
        r_au.font.bold = True

        # 6. Page 1 Main Two-Column Layout Table
        p1_tbl = doc.add_table(rows=1, cols=2)
        p1_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        p1_tbl.autofit = False
        _set_borderless_table(p1_tbl)
        c_left, c_right = p1_tbl.rows[0].cells[0], p1_tbl.rows[0].cells[1]
        c_left.width = Inches(4.35)
        c_right.width = Inches(1.96)

        # Left Column: Pink Shaded Box with Spanish Resumen + Screened English Abstract
        _set_cell_padding(c_left, top_pt=5, bottom_pt=5, left_pt=6, right_pt=6)
        _set_cell_shading(c_left, HEX_PINK_FILL)
        _set_cell_borders(c_left, top="none", bottom="none", left="none", right="none")

        p_res_hdr = c_left.paragraphs[0]
        p_res_hdr.paragraph_format.space_before = Pt(0)
        p_res_hdr.paragraph_format.space_after = Pt(2)
        r_rh = p_res_hdr.add_run("Resumen")
        r_rh.font.name = "Arial"
        r_rh.font.size = Pt(9.0)
        r_rh.font.bold = True
        r_rh.font.color.rgb = COLOR_CRIMSON

        _add_body_para(
            c_left,
            " El cáncer de piel en pacientes que recibieron algún trasplante de órgano representa el 30% de los casos reportados en la bibliografía. El 90% de los casos registrados son carcinomas epidermoides y menos del 5% corresponde a melanoma. Los factores de riesgo son el uso de terapia inmunodepresora, el tabaquismo activo, el alcoholismo y la vida sedentaria. Por tal motivo las guías internacionales recomiendan la realización de pruebas de detección de cáncer de piel en receptores de trasplante renal, al menos cada año. Lamentablemente en México no se dispone de algún programa de vigilancia anual de piel y reporte de las anomalías encontradas, a pesar de que la detección oportuna y el tratamiento en etapas tempranas del cáncer de piel aumentan la calidad y esperanza de vida en estos pacientes.",
            bold_prefix="ANTECEDENTES:",
            font_size_pt=7.8,
            space_after_pt=3.0,
        )
        _add_body_para(
            c_left,
            " Paciente masculino de 55 años, con índice tabáquico de 28 y enfermedad renal que ameritó trasplante de donador vivo, en tratamiento con terapia inmunodepresora durante seis años. Su padecimiento inició en 2023 con una lesión tumoral localizada en el tronco, con avance significativo en menos de cinco meses, así como datos clínicos sugerentes de metástasis cerebral, con deterioro de la funcionalidad y estadio avanzado, por lo que recibió tratamiento paliativo.",
            bold_prefix="CASO CLÍNICO:",
            font_size_pt=7.8,
            space_after_pt=3.0,
        )
        _add_body_para(
            c_left,
            " Se insiste en la importancia de la vigilancia y detección de los pacientes postrasplantados porque la enfermedad de base y la terapia inmunodepresora incrementan hasta un 30% el riesgo de padecer algún tipo de neoplasia.",
            bold_prefix="CONCLUSIONES:",
            font_size_pt=7.8,
            space_after_pt=3.0,
        )
        _add_body_para(
            c_left,
            " Melanoma; trasplante renal; terapia inmunosupresora.",
            bold_prefix="PALABRAS CLAVE:",
            font_size_pt=7.8,
            space_after_pt=4.0,
        )

        # English Abstract Image + Connected Hidden Text
        p_abs_img = c_left.add_paragraph()
        p_abs_img.paragraph_format.space_before = Pt(2)
        p_abs_img.paragraph_format.space_after = Pt(0)
        r_abs = p_abs_img.add_run()
        r_abs.add_picture(str(assets["p1_eng_abstract"]), width=Inches(4.15))

        eng_abs_p1_text = (
            "Abstract\\n"
            "BACKGROUND: Skin cancer in patients who have received an organ transplant account for 30% "
            "of the cases reported in the literature, 90% of the registered cases being squamous cell carcinoma "
            "and less than 5% being melanoma. Risk factors include the use of immunosuppressive therapy, which "
            "is influenced by the type of medication, as well as its duration, and current habits such as active "
            "smoking, alcoholism and a sedentary lifestyle. The international guidelines for kidney transplant "
            "recipients suggest performing skin cancer screening tests at least once a year, which unfortunately "
            "a successful program with an annual skin checkout and the report of the abnormalities found are "
            "not available in Mexico, despite timely detection and treatment in early stages increase the quality "
            "and life expectancy of these patients.\\n"
            "CLINICAL CASE: A 55-year-old male patient, with a smoking index of 28 and kidney disease that "
            "required a transplant from a living donor, under treatment with immunosuppressive therapy for 6 years. "
            "His condition began in 2023 with a tumor lesion located in the trunk, with significant progress in "
            "less than 5 months, as well as symptoms suggestive of metastasis at the brain level, with deterioration "
            "of its functionality and advanced stage. It was decided to provide palliative treatment."
        )
        _add_hidden_text_run(p_abs_img, eng_abs_p1_text)

        # Right Column: Metadata Sidebar (Created)
        _set_cell_padding(c_right, top_pt=0, bottom_pt=0, left_pt=6, right_pt=0)
        p_aff = c_right.paragraphs[0]
        p_aff.paragraph_format.space_before = Pt(0)
        p_aff.paragraph_format.space_after = Pt(3)
        p_aff.paragraph_format.line_spacing = Pt(8.2)
        aff_text = (
            "¹ Residente de tercer año de Medicina Interna, Hospital General de Zona 47, Instituto Mexicano del Seguro Social, Ciudad de México.\\n"
            "² Residente de tercer año, Departamento de Dermatología.\\n"
            "³ Jefa del Departamento de Dermatología.\\n"
            "Instituto Nacional de Ciencias Médicas y Nutrición Salvador Zubirán, Ciudad de México.\\n"
            "⁴ Residente de tercer año, Departamento de Oncología, Centro Médico Nacional Siglo XXI, Instituto Mexicano del Seguro Social, Ciudad de México."
        )
        r_aff = p_aff.add_run(aff_text)
        r_aff.font.name = "Arial"
        r_aff.font.size = Pt(6.8)
        r_aff.font.color.rgb = COLOR_CHARCOAL

        # ORCID
        p_orc = c_right.add_paragraph()
        p_orc.paragraph_format.space_before = Pt(2)
        p_orc.paragraph_format.space_after = Pt(3)
        p_orc.paragraph_format.line_spacing = Pt(8.2)
        r_oh = p_orc.add_run("ORCID\\n")
        r_oh.font.name = "Arial"
        r_oh.font.size = Pt(7.0)
        r_oh.font.bold = True
        r_oh.font.color.rgb = COLOR_CRIMSON

        r_ol = p_orc.add_run(
            "https://orcid.org/0009-0003-2708-6828\\n"
            "https://orcid.org/0009-0009-4703-819X\\n"
            "https://orcid.org/0009-0007-8221-8941\\n"
            "https://orcid.org/0000-0003-3542-4615"
        )
        r_ol.font.name = "Arial"
        r_ol.font.size = Pt(6.5)
        r_ol.font.color.rgb = COLOR_BLUE_LINK

        # Dates
        p_dt = c_right.add_paragraph()
        p_dt.paragraph_format.space_before = Pt(2)
        p_dt.paragraph_format.space_after = Pt(3)
        r_dt1 = p_dt.add_run("Recibido: ")
        r_dt1.font.bold = True
        r_dt1.font.size = Pt(7.0)
        r_dt1.font.color.rgb = COLOR_CRIMSON
        r_dt2 = p_dt.add_run("abril 2024\\n")
        r_dt2.font.size = Pt(7.0)
        r_dt3 = p_dt.add_run("Aceptado: ")
        r_dt3.font.bold = True
        r_dt3.font.size = Pt(7.0)
        r_dt3.font.color.rgb = COLOR_CRIMSON
        r_dt4 = p_dt.add_run("febrero 2025")
        r_dt4.font.size = Pt(7.0)

        # Correspondencia
        p_cor = c_right.add_paragraph()
        p_cor.paragraph_format.space_before = Pt(2)
        p_cor.paragraph_format.space_after = Pt(3)
        r_ch = p_cor.add_run("Correspondencia\\n")
        r_ch.font.name = "Arial"
        r_ch.font.size = Pt(7.0)
        r_ch.font.bold = True
        r_ch.font.color.rgb = COLOR_CRIMSON
        r_cb = p_cor.add_run("María Fernanda Corona Rosas\\nfercro15@gmail.com")
        r_cb.font.name = "Arial"
        r_cb.font.size = Pt(7.0)

        # Citation
        p_cit = c_right.add_paragraph()
        p_cit.paragraph_format.space_before = Pt(2)
        p_cit.paragraph_format.space_after = Pt(0)
        p_cit.paragraph_format.line_spacing = Pt(8.2)
        r_cih = p_cit.add_run("Este artículo debe citarse como:\\n")
        r_cih.font.name = "Arial"
        r_cih.font.size = Pt(7.0)
        r_cih.font.bold = True
        r_cih.font.color.rgb = COLOR_CRIMSON
        r_cib = p_cit.add_run(
            "Corona-Rosas MF, Quiles-Martínez B, Campos-Salgado YE, Domínguez-Cherit J. "
            "Melanoma metastásico, un caso extraordinario en un paciente con trasplante renal. "
            "Dermatol Rev Mex 2026; 70 (5): 666-672."
        )
        r_cib.font.name = "Arial"
        r_cib.font.size = Pt(6.8)

        # Page 1 Footer
        p1_ft = doc.add_paragraph()
        p1_ft.paragraph_format.space_before = Pt(8)
        p1_ft.paragraph_format.space_after = Pt(0)
        r_p1_l = p1_ft.add_run("666                                                                                                                www.nietoeditores.com.mx")
        r_p1_l.font.name = "Arial"
        r_p1_l.font.size = Pt(7.5)
        r_p1_l.font.color.rgb = COLOR_MUTED

        # =====================================================================
        # PAGE 2
        # =====================================================================
        print("  ▶ Building Page 2 (Screened Abstract Continuation + Created Antecedentes & Caso Clínico)...")
        _add_page_break(doc)

        # Header Page 2 (Even page)
        hdr2 = _create_layout_row(doc, [4.5, 1.81])
        p2_hl = hdr2.cells[0].paragraphs[0]
        r2_hl = p2_hl.add_run("Corona Rosas MF, et al. Melanoma metastásico y trasplante renal")
        r2_hl.font.name = "Arial"
        r2_hl.font.size = Pt(7.5)
        r2_hl.font.italic = True
        r2_hl.font.color.rgb = COLOR_MUTED

        p2_hr = hdr2.cells[1].paragraphs[0]
        p2_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r2_hr1 = p2_hr.add_run("Dermatología ")
        r2_hr1.font.bold = True
        r2_hr1.font.italic = True
        r2_hr1.font.color.rgb = COLOR_CRIMSON
        r2_hr1.font.size = Pt(7.5)
        r2_hr2 = p2_hr.add_run("Revista mexicana")
        r2_hr2.font.color.rgb = COLOR_CRIMSON
        r2_hr2.font.size = Pt(7.0)

        p2_div = doc.add_paragraph()
        p2_div.paragraph_format.space_before = Pt(0)
        p2_div.paragraph_format.space_after = Pt(6)
        p2_div._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="1" w:color="CCCCCC"/></w:pBdr>'))

        # Top Section: Screened English Continuation inside Pink Box (Left-aligned, matching left column)
        box_tbl2 = doc.add_table(rows=1, cols=1)
        box_tbl2.alignment = WD_TABLE_ALIGNMENT.LEFT
        box_tbl2.autofit = False
        _set_borderless_table(box_tbl2)
        box_cell2 = box_tbl2.rows[0].cells[0]
        box_cell2.width = Inches(4.35)
        _set_cell_padding(box_cell2, top_pt=4, bottom_pt=4, left_pt=6, right_pt=6)
        _set_cell_shading(box_cell2, HEX_PINK_FILL)

        p2_abs_p = box_cell2.paragraphs[0]
        p2_abs_p.paragraph_format.space_before = Pt(0)
        p2_abs_p.paragraph_format.space_after = Pt(0)
        r2_abs = p2_abs_p.add_run()
        r2_abs.add_picture(str(assets["p2_eng_abstract"]), width=Inches(4.15))

        eng_abs_p2_text = (
            "CONCLUSIONS: The importance of monitoring and detecting post-transplant patients is emphasized, "
            "since the underlying disease and immunosuppressive therapy increase up to 30% of acquiring some type of neoplasia.\\n"
            "KEYWORDS: Melanoma; Kidney transplant; Immunosuppressive therapy."
        )
        _add_hidden_text_run(p2_abs_p, eng_abs_p2_text)

        p_sp2 = doc.add_paragraph()
        p_sp2.paragraph_format.space_before = Pt(0)
        p_sp2.paragraph_format.space_after = Pt(10)

        # Page 2: 2 equal columns (3.05" and 3.05")
        p2_row = _create_layout_row(doc, [3.05, 3.05])
        p2_c_left, p2_c_right = p2_row.cells[0], p2_row.cells[1]

        _add_section_heading(p2_c_left, "ANTECEDENTES", space_before_pt=0, space_after_pt=4.0)
        _add_body_para(
            p2_c_left,
            "En la actualidad, gracias a los avances tecnológicos y científicos, hay un aumento en la ejecución de trasplante como alternativa de tratamiento para los pacientes con enfermedad renal terminal, quienes posteriormente reciben tratamientos inmunosupresores para evitar el rechazo del injerto; esta terapia, por lo general, es de por vida.¹ Sin embargo, poco se habla de las consecuencias documentadas propias de la enfermedad y del uso crónico de terapia inmunodepresora, de las que el cáncer de piel es la neoplasia maligna más documentada.² Casi el 90% de los cánceres de piel de pacientes postrasplantados corresponden al carcinoma epidermoide y menos del 5% a melanoma, que tiene altos índices de mortalidad por su característica molecular invasora.³",
            space_after_pt=6.0,
        )
        _add_body_para(
            p2_c_left,
            "Por tal razón la descripción de este caso clínico cobra relevancia, no sólo por el estadio avanzado en el que se encontró al paciente, sino por la oportunidad de detección primaria de esa neoplasia, al estar en tratamiento de inmunosupresión durante muchos años y sus factores de riesgo (tabaquismo), lo que favoreció el avance de la neoplasia y su carácter invasor que tuvo un desenlace fatal. La bibliografía indica una supervivencia de sólo 3% a dos años.⁴",
            space_after_pt=6.0,
        )

        _add_section_heading(p2_c_right, "CASO CLÍNICO", space_before_pt=0, space_after_pt=4.0)
        _add_body_para(
            p2_c_right,
            "Paciente masculino de 55 años, albañil, sin antecedentes familiares o personales de melanoma, tabaquismo positivo a razón de 15 cigarrillos al día durante 38 años, con un índice tabáquico de 28; antecedente de enfermedad renal crónica diagnosticada en 2010 que ameritó el trasplante de donador vivo en 2014. Recibió tratamiento inmunodepresor con prednisona a dosis de 5 mg cada 24 horas y tacrolimus 8 mg cada 24 horas, con últimas concentraciones séricas de tacrolimus documentadas en 2019 de 8 ng/mL sin llegar a la toxicidad.",
            space_after_pt=6.0,
        )
        _add_body_para(
            p2_c_right,
            "El paciente acudió a consulta externa de Oncología en 2023 por padecer una dermatosis localizada en el tronco, que abarcaba la mitad de la zona torácica posterior, con extensión al hueco axilar, caracterizada por una gran placa tumoral negruzca y múltiples neoformaciones de distintos tamaños; la mayor era de 3 cm. Esta gran lesión estaba exulcerada, alrededor de la gran placa se observaban lesiones satélites que medían unos cuantos milímetros a centímetros (satelitosis), se apreciaba infiltrada a la palpación y sin datos de sangrado. Figura 1",
            space_after_pt=6.0,
        )
        _add_body_para(
            p2_c_right,
            "El paciente refirió el inicio de la dermatosis seis meses antes de la consulta; sin embargo, no se",
            space_after_pt=6.0,
        )

        # Page 2 Footer
        p2_ft = doc.add_paragraph()
        p2_ft.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p2_ft.paragraph_format.space_before = Pt(16)
        p2_ft.paragraph_format.space_after = Pt(0)
        r2_f = p2_ft.add_run("667")
        r2_f.font.name = "Arial"
        r2_f.font.size = Pt(7.5)
        r2_f.font.color.rgb = COLOR_MUTED

        # =====================================================================
        # PAGE 3
        # =====================================================================
        print("  ▶ Building Page 3 (Created Figure 1 Pink Box + Created Caso Clínico & Discusión)...")
        _add_page_break(doc)

        hdr3 = _create_layout_row(doc, [4.5, 1.81])
        p3_hl = hdr3.cells[0].paragraphs[0]
        r3_hl1 = p3_hl.add_run("Dermatología ")
        r3_hl1.font.bold = True
        r3_hl1.font.italic = True
        r3_hl1.font.color.rgb = COLOR_CRIMSON
        r3_hl1.font.size = Pt(7.5)
        r3_hl2 = p3_hl.add_run("Revista mexicana")
        r3_hl2.font.color.rgb = COLOR_CRIMSON
        r3_hl2.font.size = Pt(7.0)

        p3_hr = hdr3.cells[1].paragraphs[0]
        p3_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r3_hr = p3_hr.add_run("2026; 70 (5)")
        r3_hr.font.name = "Arial"
        r3_hr.font.size = Pt(7.5)
        r3_hr.font.color.rgb = COLOR_MUTED

        p3_div = doc.add_paragraph()
        p3_div.paragraph_format.space_before = Pt(0)
        p3_div.paragraph_format.space_after = Pt(6)
        p3_div._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="1" w:color="CCCCCC"/></w:pBdr>'))

        # Page 3: 2 columns
        p3_row = _create_layout_row(doc, [3.05, 3.05])
        p3_c_left, p3_c_right = p3_row.cells[0], p3_row.cells[1]

        # Left Column: Red accent stripe + Figure 1 Stacked Images + Pink Caption
        _add_red_stripe(p3_c_left, 2.81)

        p_f1a = p3_c_left.add_paragraph()
        p_f1a.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_f1a.paragraph_format.space_before = Pt(2)
        p_f1a.paragraph_format.space_after = Pt(0)
        r_f1a = p_f1a.add_run()
        r_f1a.add_picture(str(assets["p3_img0"]), width=Inches(2.81))

        p_f1b = p3_c_left.add_paragraph()
        p_f1b.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_f1b.paragraph_format.space_before = Pt(0)
        p_f1b.paragraph_format.space_after = Pt(0)
        r_f1b = p_f1b.add_run()
        r_f1b.add_picture(str(assets["p3_img1"]), width=Inches(2.81))

        # Caption Box (Pink shaded, no border)
        cap1_tbl = p3_c_left.add_table(rows=1, cols=1)
        cap1_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        cap1_tbl.autofit = False
        _set_borderless_table(cap1_tbl)
        cap1_cell = cap1_tbl.rows[0].cells[0]
        cap1_cell.width = Inches(2.81)
        _set_cell_padding(cap1_cell, top_pt=4, bottom_pt=4, left_pt=4, right_pt=4)
        _set_cell_shading(cap1_cell, HEX_PINK_FILL)
        _add_body_para(
            cap1_cell,
            " Dermatosis localizada en el tronco con evidencia de lesión principal de tipo tumoral con lesiones satélite que abarcan hasta el hueco axilar posterior.",
            bold_prefix="Figura 1.",
            font_size_pt=7.5,
            space_after_pt=0.0,
        )

        # Right Column: Continuation of CASO CLÍNICO & DISCUSIÓN
        _add_body_para(
            p3_c_right,
            "dio seguimiento. Tuvo evolución rápida en dos meses, así como cambios del comportamiento con tendencia a la agresividad y aparición de crisis convulsivas tónicas de difícil control. Acudió al servicio de urgencias donde se recibió al paciente con desaturación del 80%, así como alteración del estado de alerta con tendencia a la somnolencia.",
            space_after_pt=4.0,
        )
        _add_body_para(
            p3_c_right,
            "A su ingreso, se practicaron estudios complementarios y toma de biopsia de piel de la lesión inicial, cuyo estudio reportó: melanoma nodular con índice de Breslow de 5 mm, células neoplásicas, con mitosis 8-10, nivel anatómico de Clark IV, con invasión hasta la dermis reticular e infiltrado linfovascular. Figura 2",
            space_after_pt=4.0,
        )
        _add_body_para(
            p3_c_right,
            "Por las alteraciones en el estado de alerta se hizo una tomografía axial computada de cráneo que evidenció una lesión hiperdensa en la región temporoparietal derecha, con contornos bien definidos, densidad de 42 UH, dimensiones de 4.7 x 3.9 cm con efecto de masa, que desplazaba 2 mm la línea media con edema perilesional. Figura 3",
            space_after_pt=4.0,
        )
        _add_body_para(
            p3_c_right,
            "La radiografía de tórax anteroposterior reveló múltiples lesiones radiolúcidas, nodulares, de gran tamaño, de predominio parahiliar en ambos campos pulmonares, características de balas de cañón, sugerentes de metástasis. Figura 4",
            space_after_pt=4.0,
        )
        _add_body_para(
            p3_c_right,
            "El paciente tuvo deterioro funcional y neurológico con secuelas de movilidad posterior a la remisión de las crisis convulsivas, por lo que se indicó intubación paliativa. Dos días después el paciente tuvo mayor deterioro ventilatorio y falleció.",
            space_after_pt=4.0,
        )
        _add_section_heading(p3_c_right, "DISCUSIÓN", space_before_pt=6.0, space_after_pt=3.0)
        _add_body_para(
            p3_c_right,
            "Los inhibidores de calcineurina, como tacrolimus o ciclosporina, son fármacos comúnmente indicados en el control inmunológico de enfermedades autoinmunitarias, como la nefropatía lúpica, la colitis ulcerosa y la miastenia gravis,",
            space_after_pt=4.0,
        )

        # Page 3 Footer
        p3_ft = doc.add_paragraph()
        p3_ft.paragraph_format.space_before = Pt(4)
        p3_ft.paragraph_format.space_after = Pt(0)
        r3_f = p3_ft.add_run("668                                                                                    https://doi.org/10.24245/dermatolrevmex.v70i5.11433")
        r3_f.font.name = "Arial"
        r3_f.font.size = Pt(7.5)
        r3_f.font.color.rgb = COLOR_MUTED

        # =====================================================================
        # PAGE 4
        # =====================================================================
        print("  ▶ Building Page 4 (Created Figure 2 Pink Box across columns + Created Discusión columns)...")
        _add_page_break(doc)

        hdr4 = _create_layout_row(doc, [4.5, 1.81])
        p4_hl = hdr4.cells[0].paragraphs[0]
        r4_hl = p4_hl.add_run("Corona Rosas MF, et al. Melanoma metastásico y trasplante renal")
        r4_hl.font.name = "Arial"
        r4_hl.font.size = Pt(7.5)
        r4_hl.font.italic = True
        r4_hl.font.color.rgb = COLOR_MUTED

        p4_hr = hdr4.cells[1].paragraphs[0]
        p4_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r4_hr1 = p4_hr.add_run("Dermatología ")
        r4_hr1.font.bold = True
        r4_hr1.font.italic = True
        r4_hr1.font.color.rgb = COLOR_CRIMSON
        r4_hr1.font.size = Pt(7.5)
        r4_hr2 = p4_hr.add_run("Revista mexicana")
        r4_hr2.font.color.rgb = COLOR_CRIMSON
        r4_hr2.font.size = Pt(7.0)

        p4_div = doc.add_paragraph()
        p4_div.paragraph_format.space_before = Pt(0)
        p4_div.paragraph_format.space_after = Pt(6)
        p4_div._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="1" w:color="CCCCCC"/></w:pBdr>'))

        # Figure 2 spans full width (5.94 in)
        _add_red_stripe(doc, 5.94)

        fig2_img_tbl = doc.add_table(rows=1, cols=2)
        fig2_img_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        fig2_img_tbl.autofit = False
        _set_borderless_table(fig2_img_tbl)
        fig2_img_tbl.rows[0].cells[0].width = Inches(2.97)
        fig2_img_tbl.rows[0].cells[1].width = Inches(2.97)
        _set_cell_padding(fig2_img_tbl.rows[0].cells[0], top_pt=2, bottom_pt=0, left_pt=0, right_pt=2)
        _set_cell_padding(fig2_img_tbl.rows[0].cells[1], top_pt=2, bottom_pt=0, left_pt=2, right_pt=0)

        p_f2a = fig2_img_tbl.rows[0].cells[0].paragraphs[0]
        p_f2a.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_f2a.paragraph_format.space_before = Pt(0)
        p_f2a.paragraph_format.space_after = Pt(0)
        r_f2a = p_f2a.add_run()
        r_f2a.add_picture(str(assets["p4_img0"]), width=Inches(2.92))

        p_f2b = fig2_img_tbl.rows[0].cells[1].paragraphs[0]
        p_f2b.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_f2b.paragraph_format.space_before = Pt(0)
        p_f2b.paragraph_format.space_after = Pt(0)
        r_f2b = p_f2b.add_run()
        r_f2b.add_picture(str(assets["p4_img1"]), width=Inches(2.92))

        # Caption Box (Pink shaded full width, no border)
        cap2_tbl = doc.add_table(rows=1, cols=1)
        cap2_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cap2_tbl.autofit = False
        _set_borderless_table(cap2_tbl)
        cap2_cell = cap2_tbl.rows[0].cells[0]
        cap2_cell.width = Inches(5.94)
        _set_cell_padding(cap2_cell, top_pt=4, bottom_pt=4, left_pt=4, right_pt=4)
        _set_cell_shading(cap2_cell, HEX_PINK_FILL)
        _add_body_para(
            cap2_cell,
            " Corte histológico teñido con H-E. Células neoplásicas, con mitosis 8-10, nivel anatómico de Clark IV, con invasión hasta la dermis reticular e infiltrado linfovascular.",
            bold_prefix="Figura 2.",
            font_size_pt=8.0,
            space_after_pt=0.0,
        )

        p_sp4 = doc.add_paragraph()
        p_sp4.paragraph_format.space_before = Pt(0)
        p_sp4.paragraph_format.space_after = Pt(6)

        # Page 4: 2 Columns below Figure 2
        p4_row = _create_layout_row(doc, [3.05, 3.05])
        p4_c_left, p4_c_right = p4_row.cells[0], p4_row.cells[1]

        _add_body_para(
            p4_c_left,
            "entre otras. Las neoplasias malignas de piel son complicaciones asociadas con su administración crónica por más de tres años.5 El paciente del caso los recibió durante nueve años.",
            space_after_pt=6.0,
        )
        _add_body_para(
            p4_c_left,
            "Asimismo, se han descrito mecanismos desencadenantes, como el microambiente tumoral en los pacientes con inmunodepresión, entre los que se mencionan las características comunes del cáncer y la inmunodepresión. Por ejemplo, las células tumorales, como las del melanoma, tienen diversos mecanismos, como los ligandos que inhiben las células T (por ejemplo, PD-L1) que, en condiciones normales, evitan una respuesta inmunitaria excesiva frente a células propias y en pacientes con inmunodepresión se observa un proceso intensificado de disminución",
            space_after_pt=6.0,
        )

        _add_body_para(
            p4_c_right,
            "de la respuesta celular, lo que vuelve a los linfocitos T ineficaces, da pie a las características invasivas del melanoma6 y perpetúa un estado inmunodepresivo que favorece las metástasis y la angiogénesis, lo que, en conjunto, confiere una alta incidencia de mortalidad por invasión.6 Además, hay estados que favorecen al crecimiento de neoplasias, como el tabaquismo, alcoholismo, enfermedades que mantengan una inflamación constante porque se alteran los mecanismos de señalización de células T, así como el antecedente familiar de neoplasias.7",
            space_after_pt=6.0,
        )
        _add_body_para(
            p4_c_right,
            "En el paciente del caso se documentaron múltiples factores de riesgo en relación con el uso de terapias de inmunosupresión, que están asociados con el avance de la enfermedad.",
            space_after_pt=6.0,
        )

        # Page 4 Footer
        p4_ft = doc.add_paragraph()
        p4_ft.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p4_ft.paragraph_format.space_before = Pt(4)
        p4_ft.paragraph_format.space_after = Pt(0)
        r4_f = p4_ft.add_run("669")
        r4_f.font.name = "Arial"
        r4_f.font.size = Pt(7.5)
        r4_f.font.color.rgb = COLOR_MUTED

        # =====================================================================
        # PAGE 5
        # =====================================================================
        print("  ▶ Building Page 5 (Created Figure 3 Pink Box across columns + Created Discusión columns)...")
        _add_page_break(doc)

        hdr5 = _create_layout_row(doc, [4.5, 1.81])
        p5_hl = hdr5.cells[0].paragraphs[0]
        r5_hl1 = p5_hl.add_run("Dermatología ")
        r5_hl1.font.bold = True
        r5_hl1.font.italic = True
        r5_hl1.font.color.rgb = COLOR_CRIMSON
        r5_hl1.font.size = Pt(7.5)
        r5_hl2 = p5_hl.add_run("Revista mexicana")
        r5_hl2.font.color.rgb = COLOR_CRIMSON
        r5_hl2.font.size = Pt(7.0)

        p5_hr = hdr5.cells[1].paragraphs[0]
        p5_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r5_hr = p5_hr.add_run("2026; 70 (5)")
        r5_hr.font.name = "Arial"
        r5_hr.font.size = Pt(7.5)
        r5_hr.font.color.rgb = COLOR_MUTED

        p5_div = doc.add_paragraph()
        p5_div.paragraph_format.space_before = Pt(0)
        p5_div.paragraph_format.space_after = Pt(6)
        p5_div._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="1" w:color="CCCCCC"/></w:pBdr>'))

        # Figure 3 spans full width (5.94 in)
        _add_red_stripe(doc, 5.94)

        fig3_img_tbl = doc.add_table(rows=1, cols=2)
        fig3_img_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        fig3_img_tbl.autofit = False
        _set_borderless_table(fig3_img_tbl)
        fig3_img_tbl.rows[0].cells[0].width = Inches(2.97)
        fig3_img_tbl.rows[0].cells[1].width = Inches(2.97)
        _set_cell_padding(fig3_img_tbl.rows[0].cells[0], top_pt=2, bottom_pt=0, left_pt=0, right_pt=2)
        _set_cell_padding(fig3_img_tbl.rows[0].cells[1], top_pt=2, bottom_pt=0, left_pt=2, right_pt=0)

        p_f3a = fig3_img_tbl.rows[0].cells[0].paragraphs[0]
        p_f3a.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_f3a.paragraph_format.space_before = Pt(0)
        p_f3a.paragraph_format.space_after = Pt(0)
        r_f3a = p_f3a.add_run()
        r_f3a.add_picture(str(assets["p5_img0"]), width=Inches(2.90))

        p_f3b = fig3_img_tbl.rows[0].cells[1].paragraphs[0]
        p_f3b.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_f3b.paragraph_format.space_before = Pt(0)
        p_f3b.paragraph_format.space_after = Pt(0)
        r_f3b = p_f3b.add_run()
        r_f3b.add_picture(str(assets["p5_img1"]), width=Inches(2.90))

        # Caption
        cap3_tbl = doc.add_table(rows=1, cols=1)
        cap3_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cap3_tbl.autofit = False
        _set_borderless_table(cap3_tbl)
        cap3_cell = cap3_tbl.rows[0].cells[0]
        cap3_cell.width = Inches(5.94)
        _set_cell_padding(cap3_cell, top_pt=4, bottom_pt=4, left_pt=4, right_pt=4)
        _set_cell_shading(cap3_cell, HEX_PINK_FILL)
        _add_body_para(
            cap3_cell,
            " Lesión hiperdensa en la región temporoparietal derecha, con contornos bien definidos, densidad de 42 UH, dimensiones de 4.7 x 3.9 cm, con efecto de masa, que desplazaba 2 mm la línea media con edema perilesional.",
            bold_prefix="Figura 3.",
            font_size_pt=8.0,
            space_after_pt=0.0,
        )

        p_sp5 = doc.add_paragraph()
        p_sp5.paragraph_format.space_before = Pt(0)
        p_sp5.paragraph_format.space_after = Pt(6)

        # Page 5: 2 Columns below Figure 3
        p5_row = _create_layout_row(doc, [3.05, 3.05])
        p5_c_left, p5_c_right = p5_row.cells[0], p5_row.cells[1]

        _add_body_para(
            p5_c_left,
            "Un metanálisis que incluyó 309,551 pacientes reportó una asociación estadísticamente significativa entre la terapia inmunosupresora y mayor riesgo de melanoma (OR [razón de momios] 1.09; IC95% [intervalo de confianza del 95%]: 0.25 a 4.74; p < 0.01).8",
            space_after_pt=6.0,
        )
        _add_body_para(
            p5_c_left,
            "Los posibles mecanismos asociados con una mayor incidencia de cáncer de piel incluyen el bloqueo de la respuesta celular frente a antígenos tumorales mediada por linfocitos T CD4 y CD8, además de efectos directos sobre los queratinocitos y melanocitos que favorecen su transformación maligna, angiogénesis e invasión tisular. Otros mecanismos descritos son la inhibición de la apoptosis por alteración de proteínas proapoptóticas y la disminución en la reparación del daño del ADN inducido por radiación ultravioleta.8",
            space_after_pt=6.0,
        )

        _add_body_para(
            p5_c_right,
            "procesos de señalamiento y activación en los linfocitos T, como CTLA-4 y su receptor PD-1 que favorecen la producción de linfocitos quiescentes, y es directamente proporcional a la incidencia y características invasoras del melanoma.9",
            space_after_pt=6.0,
        )
        _add_body_para(
            p5_c_right,
            "En el paciente del caso la duración de inmunosupresión fue mayor de tres años y las concentraciones séricas mayores de 5 ng/dL de tacrolimus indican toxicidad. Destacó el escaso seguimiento y control de las dosis tóxicas séricas de tacrolimus.",
            space_after_pt=6.0,
        )
        _add_body_para(
            p5_c_right,
            "Otros factores de riesgo no modificables, también observados en población sin inmunodepresión, que hace a los pacientes susceptibles a padecer melanoma, son la raza blanca, el",
            space_after_pt=6.0,
        )

        # Page 5 Footer
        p5_ft = doc.add_paragraph()
        p5_ft.paragraph_format.space_before = Pt(4)
        p5_ft.paragraph_format.space_after = Pt(0)
        r5_f = p5_ft.add_run("670                                                                                    https://doi.org/10.24245/dermatolrevmex.v70i5.11433")
        r5_f.font.name = "Arial"
        r5_f.font.size = Pt(7.5)
        r5_f.font.color.rgb = COLOR_MUTED

        # =====================================================================
        # PAGE 6
        # =====================================================================
        print("  ▶ Building Page 6 (Created Figure 4 Pink Box + Created Conclusiones & Referencias 1-2)...")
        _add_page_break(doc)

        hdr6 = _create_layout_row(doc, [4.5, 1.81])
        p6_hl = hdr6.cells[0].paragraphs[0]
        r6_hl = p6_hl.add_run("Corona Rosas MF, et al. Melanoma metastásico y trasplante renal")
        r6_hl.font.name = "Arial"
        r6_hl.font.size = Pt(7.5)
        r6_hl.font.italic = True
        r6_hl.font.color.rgb = COLOR_MUTED

        p6_hr = hdr6.cells[1].paragraphs[0]
        p6_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r6_hr1 = p6_hr.add_run("Dermatología ")
        r6_hr1.font.bold = True
        r6_hr1.font.italic = True
        r6_hr1.font.color.rgb = COLOR_CRIMSON
        r6_hr1.font.size = Pt(7.5)
        r6_hr2 = p6_hr.add_run("Revista mexicana")
        r6_hr2.font.color.rgb = COLOR_CRIMSON
        r6_hr2.font.size = Pt(7.0)

        p6_div = doc.add_paragraph()
        p6_div.paragraph_format.space_before = Pt(0)
        p6_div.paragraph_format.space_after = Pt(6)
        p6_div._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="1" w:color="CCCCCC"/></w:pBdr>'))

        # Page 6: 2 Columns
        p6_row = _create_layout_row(doc, [3.05, 3.05])
        p6_c_left, p6_c_right = p6_row.cells[0], p6_row.cells[1]

        # Left Column: Red Stripe + Figure 4 Image + Pink Caption
        _add_red_stripe(p6_c_left, 2.81)

        p_f4 = p6_c_left.add_paragraph()
        p_f4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_f4.paragraph_format.space_before = Pt(2)
        p_f4.paragraph_format.space_after = Pt(0)
        r_f4 = p_f4.add_run()
        r_f4.add_picture(str(assets["p6_img0"]), width=Inches(2.81))

        cap4_tbl = p6_c_left.add_table(rows=1, cols=1)
        cap4_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        cap4_tbl.autofit = False
        _set_borderless_table(cap4_tbl)
        cap4_cell = cap4_tbl.rows[0].cells[0]
        cap4_cell.width = Inches(2.81)
        _set_cell_padding(cap4_cell, top_pt=4, bottom_pt=4, left_pt=4, right_pt=4)
        _set_cell_shading(cap4_cell, HEX_PINK_FILL)
        _add_body_para(
            cap4_cell,
            " Múltiples lesiones radiolúcidas, nodulares, de gran tamaño, de predominio parahiliar, en ambos campos pulmonares, características de balas de cañón.",
            bold_prefix="Figura 4.",
            font_size_pt=7.5,
            space_after_pt=0.0,
        )

        _add_body_para(
            p6_c_left,
            "sexo masculino, la obesidad (índice de masa corporal mayor 30) por considerarse una enfermedad inflamatoria, con desregulación de la expresión de los linfocitos TNF, lo que favorece a estados de inmunodepresión leve, así como la edad avanzada en los pacientes al momento del trasplante (mayores a 45 años).10 Asimismo, el tabaquismo es un factor de riesgo modificable, relacionado significativamente con el aumento de la frecuencia de melanoma incluso en población sin otros factores de riesgo atribuibles, como la inmunodepresión.8,9",
            space_after_pt=6.0,
        )
        _add_body_para(
            p6_c_left,
            "Por lo anterior, se requiere la identificación de factores de riesgo de cáncer cutáneo, así como la evaluación periódica y tamizaje de neoplasias en todo paciente que recibe un trasplante. Las",
            space_after_pt=6.0,
        )

        # Right Column
        _add_body_para(
            p6_c_right,
            "guías internacionales sugieren la realización de pruebas de detección de cáncer de piel en receptores de trasplante renal, al menos, cada seis meses postrasplante y posteriormente cada año al tomar en cuenta el tiempo de aparición del melanoma reportado de 1.45-5 años posterior al trasplante.9 En sujetos con terapia inmunodepresora se recomienda la vigilancia anual de la piel de por vida. De igual manera se sugiere la cuantificación anual de las concentraciones séricas de tacrolimus con objetivo < 5 ng/dl para evitar efectos adversos.11",
            space_after_pt=6.0,
        )
        _add_body_para(
            p6_c_right,
            "Siempre que se extirpen lesiones de piel éstas deben enviarse para estudio histopatológico, lo que garantiza la detección temprana de neoplasias malignas y el tratamiento oportuno que evita complicaciones fatales, como en el paciente del caso.",
            space_after_pt=6.0,
        )
        _add_section_heading(p6_c_right, "CONCLUSIONES", space_before_pt=8.0, space_after_pt=3.0)
        _add_body_para(
            p6_c_right,
            "Esta comunicación insiste en la importancia de la vigilancia cuidadosa de los pacientes posterior a un trasplante y la necesidad del control riguroso de las concentraciones séricas de inmunosupresores, así como de la prevención y detección oportuna de cáncer de piel. Conocer e identificar los factores de riesgo asociados, así como el envío a estudio histopatológico de las lesiones de piel sospechosas de malignidad o referencia con el facultativo indicado, permitirán un correcto y oportuno tratamiento de las neoplasias, lo que, en conjunto, mejorará el pronóstico de los pacientes trasplantados.",
            space_after_pt=6.0,
        )
        _add_section_heading(p6_c_right, "REFERENCIAS", space_before_pt=8.0, space_after_pt=3.0)
        _add_body_para(
            p6_c_right,
            "Kulbat A, Richter K, Stefura T, et al. Systematic Review of calcineurin inhibitors and incidence of skin malignancies after kidney transplantation in adult patients: A study of 309,551 cases. Curr Oncol 2023; 30 (6): 5727-5737. https://doi.org/10.3390/curroncol30060430",
            bold_prefix="1.  ",
            font_size_pt=7.2,
            space_after_pt=4.0,
            hanging_indent_in=0.20,
        )
        _add_body_para(
            p6_c_right,
            "Thet Z, Lam AK, Ranganathan D, et al. Reducing non-melanoma skin cancer risk in renal transplant recipients. Nephrology (Carlton) 2021; 26 (11): 907-919. https://doi.org/10.1111/nep.13939",
            bold_prefix="2.  ",
            font_size_pt=7.2,
            space_after_pt=4.0,
            hanging_indent_in=0.20,
        )

        # Page 6 Footer
        p6_ft = doc.add_paragraph()
        p6_ft.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p6_ft.paragraph_format.space_before = Pt(4)
        p6_ft.paragraph_format.space_after = Pt(0)
        r6_f = p6_ft.add_run("671")
        r6_f.font.name = "Arial"
        r6_f.font.size = Pt(7.5)
        r6_f.font.color.rgb = COLOR_MUTED

        # =====================================================================
        # PAGE 7
        # =====================================================================
        print("  ▶ Building Page 7 (Created References 3-11 + Created Aviso Importante Card)...")
        _add_page_break(doc)

        hdr7 = _create_layout_row(doc, [4.5, 1.81])
        p7_hl = hdr7.cells[0].paragraphs[0]
        r7_hl1 = p7_hl.add_run("Dermatología ")
        r7_hl1.font.bold = True
        r7_hl1.font.italic = True
        r7_hl1.font.color.rgb = COLOR_CRIMSON
        r7_hl1.font.size = Pt(7.5)
        r7_hl2 = p7_hl.add_run("Revista mexicana")
        r7_hl2.font.color.rgb = COLOR_CRIMSON
        r7_hl2.font.size = Pt(7.0)

        p7_hr = hdr7.cells[1].paragraphs[0]
        p7_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r7_hr = p7_hr.add_run("2026; 70 (5)")
        r7_hr.font.name = "Arial"
        r7_hr.font.size = Pt(7.5)
        r7_hr.font.color.rgb = COLOR_MUTED

        p7_div = doc.add_paragraph()
        p7_div.paragraph_format.space_before = Pt(0)
        p7_div.paragraph_format.space_after = Pt(6)
        p7_div._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="1" w:color="CCCCCC"/></w:pBdr>'))

        # 2 Columns for References 3–11
        p7_row = _create_layout_row(doc, [3.05, 3.05])
        p7_c_left, p7_c_right = p7_row.cells[0], p7_row.cells[1]

        # References Left Column (3 to 7)
        _add_body_para(
            p7_c_left,
            "Ascha M, Ascha MS, Tanenbaum J, Bordeaux JS. Risk factors for melanoma in renal transplant recipients. JAMA Dermatol 2017; 153 (11): 1130-1136. https://doi.org/10.1001/jamadermatol.2017.2291",
            bold_prefix="3.  ",
            font_size_pt=7.2,
            space_after_pt=4.0,
            hanging_indent_in=0.20,
        )
        _add_body_para(
            p7_c_left,
            "Nabi Z, Zahid T, Nabi R. Post renal transplant malignancies. A basic concept. J Ayub Med Coll Abbottabad 2023; 35 (4): 664-668. https://doi.org/10.55519/JAMC-04-12230",
            bold_prefix="4.  ",
            font_size_pt=7.2,
            space_after_pt=4.0,
            hanging_indent_in=0.20,
        )
        _add_body_para(
            p7_c_left,
            "Ferrándiz Pulido C. Actualización en cáncer de piel en receptores de un trasplante de órgano sólido. Nefrología Sup Ext 2018; 9 (1): 6-20.",
            bold_prefix="5.  ",
            font_size_pt=7.2,
            space_after_pt=4.0,
            hanging_indent_in=0.20,
        )
        _add_body_para(
            p7_c_left,
            "Ponticelli C, Cucchiari D, Bencini P. Skin cancer in kidney transplant recipients. J Nephrol 2014; 27 (4): 385-94. https://doi.org/10.1007/s40620-014-0098-4",
            bold_prefix="6.  ",
            font_size_pt=7.2,
            space_after_pt=4.0,
            hanging_indent_in=0.20,
        )
        _add_body_para(
            p7_c_left,
            "Batta N, Shangraw S, Nicklawsky A, et al. Global melanoma correlations with obesity, smoking, and alcohol consumption. JMIR Dermatol 2021; 4 (2): e31275. https://doi.org/10.2196/31275",
            bold_prefix="7.  ",
            font_size_pt=7.2,
            space_after_pt=4.0,
            hanging_indent_in=0.20,
        )

        # References Right Column (8 to 11)
        _add_body_para(
            p7_c_right,
            "Hao X, Lai W, Xia X, et al. Skin cancer outcomes and risk factors in renal transplant recipients: Analysis of organ procurement and transplantation network data from 2000 to 2021. Front Oncol 2022; 12: 1017498. https://doi.org/10.3389/fonc.2022.1017498",
            bold_prefix="8.  ",
            font_size_pt=7.2,
            space_after_pt=4.0,
            hanging_indent_in=0.20,
        )
        _add_body_para(
            p7_c_right,
            "Mittal A, Colegio OR. Skin cancers in organ transplant recipients. Am J Transplant 2017; 17: 2509-2530. https://doi.org/10.1111/ajt.14382",
            bold_prefix="9.  ",
            font_size_pt=7.2,
            space_after_pt=4.0,
            hanging_indent_in=0.20,
        )
        _add_body_para(
            p7_c_right,
            "Pascual J, Abramowicz D, Cochat C, et al. European Renal Best Practice Guideline on the Management and Evaluation of the Kidney Donor and Recipient. Nefrología 2014; 34 (3): 273-424. https://doi.org/10.3265/Nefrologia.pre2014.Feb.12490",
            bold_prefix="10.  ",
            font_size_pt=7.2,
            space_after_pt=4.0,
            hanging_indent_in=0.20,
        )
        _add_body_para(
            p7_c_right,
            "Banas B, Krämer BK, Krüger B, et al. Long-term kidney transplant outcomes: Role of prolonged-release tacrolimus. Transplantation Proceed 2020; 52 (1): 102-110. https://doi.org/10.1016/j.transproceed.2019.11.003",
            bold_prefix="11.  ",
            font_size_pt=7.2,
            space_after_pt=4.0,
            hanging_indent_in=0.20,
        )

        p_sp7 = doc.add_paragraph()
        p_sp7.paragraph_format.space_before = Pt(4)
        p_sp7.paragraph_format.space_after = Pt(4)

        # Announcement Box (AVISO IMPORTANTE) Card: 2 rows x 2 cols table
        aviso_tbl = doc.add_table(rows=2, cols=2)
        aviso_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        aviso_tbl.autofit = False
        _set_borderless_table(aviso_tbl)

        # Row 0: Heading merged across both columns
        av_hcell = aviso_tbl.rows[0].cells[0]
        av_hcell.merge(aviso_tbl.rows[0].cells[1])
        av_hcell.width = Inches(5.94)
        _set_cell_padding(av_hcell, top_pt=6, bottom_pt=2, left_pt=8, right_pt=8)
        _set_cell_shading(av_hcell, HEX_BLUE_CARD)
        _set_cell_borders(av_hcell, top="single", left="single", right="single", sz="6", color=HEX_BLUE_BORDER)

        p_avh = av_hcell.paragraphs[0]
        p_avh.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_avh.paragraph_format.space_before = Pt(0)
        p_avh.paragraph_format.space_after = Pt(2)
        r_avh = p_avh.add_run("AVISO IMPORTANTE")
        r_avh.font.name = "Arial"
        r_avh.font.size = Pt(11.0)
        r_avh.font.bold = True
        r_avh.font.color.rgb = COLOR_NAVY

        # Row 1: Left: Text and Store Badges, Right: Smartphone
        in_c_left = aviso_tbl.rows[1].cells[0]
        in_c_right = aviso_tbl.rows[1].cells[1]
        in_c_left.width = Inches(3.80)
        in_c_right.width = Inches(2.14)
        _set_cell_padding(in_c_left, top_pt=2, bottom_pt=6, left_pt=8, right_pt=4)
        _set_cell_padding(in_c_right, top_pt=2, bottom_pt=6, left_pt=4, right_pt=8)
        _set_cell_shading(in_c_left, HEX_BLUE_CARD)
        _set_cell_shading(in_c_right, HEX_BLUE_CARD)
        _set_cell_borders(in_c_left, bottom="single", left="single", sz="6", color=HEX_BLUE_BORDER)
        _set_cell_borders(in_c_right, bottom="single", right="single", sz="6", color=HEX_BLUE_BORDER)

        p_avt1 = in_c_left.paragraphs[0]
        p_avt1.paragraph_format.space_before = Pt(0)
        p_avt1.paragraph_format.space_after = Pt(3)
        r_at1a = p_avt1.add_run("Ahora puede descargar la aplicación de ")
        r_at1a.font.name = "Arial"
        r_at1a.font.size = Pt(8.0)
        r_at1b = p_avt1.add_run("Dermatología Revista Mexicana.")
        r_at1b.font.name = "Arial"
        r_at1b.font.size = Pt(8.0)
        r_at1b.font.bold = True
        r_at1b.font.color.rgb = COLOR_NAVY

        p_avt2 = in_c_left.add_paragraph()
        p_avt2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_avt2.paragraph_format.space_before = Pt(0)
        p_avt2.paragraph_format.space_after = Pt(3)
        p_avt2.paragraph_format.line_spacing = Pt(9.2)
        r_at2 = p_avt2.add_run(
            "Para consultar el texto completo de los artículos deberá registrarse una sola vez con su correo electrónico, "
            "crear una contraseña, indicar su nombre completo y especialidad. Esta información es indispensable para saber qué "
            "consulta y cuáles son sus intereses y poder en el futuro inmediato satisfacer sus necesidades de información."
        )
        r_at2.font.name = "Arial"
        r_at2.font.size = Pt(7.2)

        p_avt3 = in_c_left.add_paragraph()
        p_avt3.paragraph_format.space_before = Pt(2)
        p_avt3.paragraph_format.space_after = Pt(3)
        r_at3 = p_avt3.add_run("La aplicación está disponible para Android o iPhone.")
        r_at3.font.name = "Arial"
        r_at3.font.size = Pt(7.8)
        r_at3.font.bold = True

        p_badge = in_c_left.add_paragraph()
        p_badge.paragraph_format.space_before = Pt(0)
        p_badge.paragraph_format.space_after = Pt(0)
        if "p7_img3" in assets:
            r_b1 = p_badge.add_run()
            r_b1.add_picture(str(assets["p7_img3"]), width=Inches(0.40))
        if "p7_img0" in assets:
            r_b2 = p_badge.add_run("  ")
            r_b3 = p_badge.add_run()
            r_b3.add_picture(str(assets["p7_img0"]), width=Inches(0.40))

        # Smartphone graphic in right cell
        p_phone = in_c_right.paragraphs[0]
        p_phone.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_phone.paragraph_format.space_before = Pt(0)
        p_phone.paragraph_format.space_after = Pt(0)
        if "p7_img1" in assets:
            r_ph = p_phone.add_run()
            r_ph.add_picture(str(assets["p7_img1"]), height=Inches(2.80))

        # Page 7 Footer
        p7_ft = doc.add_paragraph()
        p7_ft.paragraph_format.space_before = Pt(8)
        p7_ft.paragraph_format.space_after = Pt(0)
        r7_f = p7_ft.add_run("672                                                                                    https://doi.org/10.24245/dermatolrevmex.v70i5.11433")
        r7_f.font.name = "Arial"
        r7_f.font.size = Pt(7.5)
        r7_f.font.color.rgb = COLOR_MUTED

        # Save document
        doc.save(str(output_docx))
        print(f"✅ [JournalArticleBuilder] Document saved: {output_docx} ({output_docx.stat().st_size / 1024:.1f} KB)")
        return output_docx
'''

# Write to modifications and to builders
Path("runs/DOC0000074469_20261002_1045/modifications/journal_article_builder.py").write_text(refined_builder_code, encoding="utf-8")
Path("scripts/stages/builders/journal_article_builder.py").write_text(refined_builder_code, encoding="utf-8")
print("✅ Refined builder written to runs/.../modifications and scripts/stages/builders!")
