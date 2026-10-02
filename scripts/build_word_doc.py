import os
import sys
from pathlib import Path
from typing import Optional
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def resolve_image(img_name: str, input_pdf: Optional[str] = None) -> str:
    """Find stitched image on disk, or extract & stitch dynamically if missing."""
    base_name = os.path.basename(img_name)
    workspace = Path(__file__).resolve().parent.parent

    candidates = [
        workspace / "extracted" / "extracted_stitched" / base_name,
        workspace / "extracted_stitched" / base_name,
        Path("extracted") / "extracted_stitched" / base_name,
        Path("extracted_stitched") / base_name,
        Path(img_name),
    ]
    for c in candidates:
        if c.exists():
            return str(c.resolve())

    # Dynamic extraction fallback if input_pdf is provided
    search_pdf = input_pdf
    if not search_pdf:
        for default_p in [
            workspace / "input" / "263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf",
            workspace / "finished" / "263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf",
        ]:
            if default_p.exists():
                search_pdf = str(default_p)
                break

    if search_pdf and os.path.exists(search_pdf):
        try:
            import fitz
            from PIL import Image
            import io
            import re
            m = re.search(r"page_(\d+)", base_name)
            if m:
                pno = int(m.group(1)) - 1
                doc = fitz.open(search_pdf)
                if 0 <= pno < len(doc):
                    page = doc[pno]
                    img_infos = page.get_image_info(xrefs=True)
                    if img_infos:
                        img_infos.sort(key=lambda x: x['bbox'][1])
                        images = []
                        w = 0
                        total_h = 0
                        for info in img_infos:
                            xref = info['xref']
                            base = doc.extract_image(xref)
                            im = Image.open(io.BytesIO(base['image']))
                            images.append(im)
                            w = max(w, im.width)
                            total_h += im.height
                        stitched = Image.new('RGB', (w, total_h), 'white')
                        cur_y = 0
                        for im in images:
                            stitched.paste(im, (0, cur_y))
                            cur_y += im.height
                        out_dir = workspace / "extracted" / "extracted_stitched"
                        out_dir.mkdir(parents=True, exist_ok=True)
                        out_path = out_dir / base_name
                        stitched.save(str(out_path))
                        doc.close()
                        return str(out_path.resolve())
                    doc.close()
        except Exception as exc:
            print(f"  ⚠️ Dynamic extraction failed for {base_name}: {exc}")

    raise FileNotFoundError(f"Could not find or extract image: {img_name}")


def create_document(input_pdf: Optional[str] = None, output_path: Optional[str] = None) -> str:
    workspace = Path(__file__).resolve().parent.parent

    if output_path is None:
        out_file = workspace / "output" / "263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.docx"
    else:
        out_file = Path(output_path).resolve()

    out_file.parent.mkdir(parents=True, exist_ok=True)

    doc = docx.Document()
    
    # Page setup: Letter Landscape, 0.5 in margins
    for section in doc.sections:
        section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11.0)
        section.page_height = Inches(8.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.header_distance = Inches(0.3)
        section.footer_distance = Inches(0.3)

    BLUE = RGBColor(0, 112, 192)  # #0070C0
    BLACK = RGBColor(0, 0, 0)
    FONT_NAME = "Aptos"
    BORDER_COLOR = "0070C0"

    def set_font(run, size=11, bold=False, color=BLUE):
        run.font.name = FONT_NAME
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        rPr = run._r.get_or_add_rPr()
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{FONT_NAME}" w:hAnsi="{FONT_NAME}" w:cs="{FONT_NAME}"/>')
        rPr.append(rFonts)

    def add_p(text="", size=11, bold=False, color=BLUE, space_before=0, space_after=0, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.0
        if indent:
            p.paragraph_format.left_indent = Inches(indent)
        if text:
            run = p.add_run(text)
            set_font(run, size=size, bold=bold, color=color)
        return p

    def set_table_borders(table, color=BORDER_COLOR, sz="4"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(f'''
            <w:tblBorders {nsdecls("w")}>
                <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
                <w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
            </w:tblBorders>
        ''')
        tblPr.append(borders)

    def set_table_margins(table, top=0, bottom=0, left=30, right=30):
        tblPr = table._tbl.tblPr
        cell_mar = parse_xml(f'''
            <w:tblCellMar {nsdecls("w")}>
                <w:top w:w="{top}" w:type="dxa"/>
                <w:bottom w:w="{bottom}" w:type="dxa"/>
                <w:left w:w="{left}" w:type="dxa"/>
                <w:right w:w="{right}" w:type="dxa"/>
            </w:tblCellMar>
        ''')
        tblPr.append(cell_mar)

    def format_cell(cell, text, width_in, size=10, bold=False, color=BLUE, align=WD_ALIGN_PARAGRAPH.LEFT, v_align=WD_ALIGN_VERTICAL.CENTER):
        cell.width = Inches(width_in)
        cell.vertical_alignment = v_align
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(text)
        set_font(run, size=size, bold=bold, color=color)

    def format_row(row, height_pt=None, cant_split=True):
        trPr = row._tr.get_or_add_trPr()
        if cant_split:
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        if height_pt:
            val = int(height_pt * 20)
            trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{val}" w:hRule="atLeast"/>'))

    # ==========================================
    # PAGE 1
    # ==========================================
    p1 = add_p("1. Please provide relevant medical history, disease stage/status, prior anticancer therapies, concomitant medications, and baseline laboratory findings.",
               size=11.5, bold=True, color=BLACK, space_before=0, space_after=1.5, indent=0.25)
    
    add_p("1.1 medical history", size=11.5, bold=True, color=BLACK, space_before=1.5, space_after=1.5)

    # Table 1.1 (15 rows, 8 cols)
    t1_widths = [0.82, 2.05, 0.95, 0.95, 0.95, 2.11, 0.70, 1.23] # Sum = ~9.76 inches
    t1_data = [
        ["Subject", "Medical Condition or Event", "Category", "Start Date", "Start Date", "Status at Screening", "End Date", "NCI CTCAE Grade"],
        ["263116", "mitral regurgitation", "Other (General)", "3/12/2026", "12-Mar-26", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "tricuspid regurgitation", "Other (General)", "3/12/2026", "12-Mar-26", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "aortic regurgitation", "Other (General)", "3/12/2026", "12-Mar-26", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "struma nodosa", "Other (General)", "3/13/2026", "13-Mar-26", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "Leukocytosis", "Other (General)", "3/13/2026", "13-Mar-26", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "The neutrophil count is elevated", "Other (General)", "3/13/2026", "13-Mar-26", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "Anemia", "Other (General)", "3/13/2026", "13-Mar-26", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "Hypoalbuminemia", "Other (General)", "3/13/2026", "13-Mar-26", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "Hyperglycemia", "Other (General)", "3/13/2026", "13-Mar-26", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "Subclinical hyperthyroidism", "Other (General)", "3/13/2026", "13-Mar-26", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "sour regurgitation", "Other (General)", "2022-08-UN", "UN-Aug-2022", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "abdominal pain", "Cancer related", "2025-12-UN", "UN-Dec-2025", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "ventosity", "Cancer related", "2025-12-UN", "UN-Dec-2025", "Ongoing at Screening", "", "Grade 1"],
        ["263116", "Implantation of venous access port", "Other (General)", "2/9/2026", "9-Feb-26", "Resolved Before Screening", "9-Feb-26", ""]
    ]
    t1 = doc.add_table(rows=len(t1_data), cols=len(t1_widths))
    t1.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_table_borders(t1, color=BORDER_COLOR)
    set_table_margins(t1, top=0, bottom=0, left=25, right=25)
    for r_idx, row in enumerate(t1.rows):
        format_row(row, height_pt=12)
        is_header = (r_idx == 0)
        for c_idx, cell in enumerate(row.cells):
            format_cell(cell, t1_data[r_idx][c_idx], t1_widths[c_idx], size=9.5, bold=is_header, color=BLUE)

    add_p("1.2 disease stage/status", size=11.5, bold=True, color=BLACK, space_before=2, space_after=1)
    add_p("cT4bN2M0 IIIC  before surgery on 8 Aug 2026， ypT0N0M0", size=11.5, bold=False, color=BLUE, space_before=0, space_after=2)

    add_p("1.3 prior anticancer therapies", size=11.5, bold=True, color=BLACK, space_before=1.5, space_after=1.5)
    
    # Table 1.3 (2 rows, 7 cols) - exactly 6.58 inches in original PDF (0.94 in per col)
    t3_widths = [0.94, 0.94, 0.94, 0.94, 0.94, 0.94, 0.94]
    t3_data = [
        ["P1-C1D1", "P1-C2D1", "P1-C3D1", "P1-C4D1", "Surgery", "P2-C1D1", "P2-C2D1"],
        ["3/14/2026", "4/3/2026", "4/24/2026", "5/15/2026", "6/8/2026", "7/8/2026", "8/19/2026"]
    ]
    t3 = doc.add_table(rows=2, cols=7)
    t3.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_table_borders(t3, color=BORDER_COLOR)
    set_table_margins(t3, top=0, bottom=0, left=20, right=20)
    for r_idx, row in enumerate(t3.rows):
        format_row(row, height_pt=12)
        is_h = (r_idx == 0)
        for c_idx, cell in enumerate(row.cells):
            format_cell(cell, t3_data[r_idx][c_idx], t3_widths[c_idx], size=9.5, bold=is_h, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER)

    add_p("1.4 concomitant medications", size=11.5, bold=True, color=BLACK, space_before=2, space_after=1.5)

    # Table 1.4 Header row (at bottom of Page 1)
    t4_widths = [0.62, 2.00, 0.81, 0.81, 0.50, 0.81, 0.75, 0.60, 0.81, 0.42, 0.75, 0.79] # Sum = ~9.67 in
    t4_hdr = [
        ["Subject", "Medication or Therapy", "Medication/Therapy Taken Prior Screening", "Start Date", "Ongoing", "End Date", "Indication", "Specify Indication", "Dose per Administration", "Dose Unit", "Frequency", "Route of Administration"]
    ]
    t4_p1 = doc.add_table(rows=1, cols=12)
    set_table_borders(t4_p1, color=BORDER_COLOR)
    set_table_margins(t4_p1, top=0, bottom=0, left=20, right=20)
    format_row(t4_p1.rows[0], height_pt=35)
    for c_idx, cell in enumerate(t4_p1.rows[0].cells):
        format_cell(cell, t4_hdr[0][c_idx], t4_widths[c_idx], size=9, bold=False, color=BLUE, align=WD_ALIGN_PARAGRAPH.LEFT)

    doc.add_page_break()

    # ==========================================
    # PAGE 2
    # ==========================================
    t4_p2_data = [
        ["263116", "Esomeprazole Magnesium Enteric-coated Tablets", "No", "15-Mar-26", "No", "21-Mar-26", "Medical History", "", "20", "mg", "QD (every day)", "Oral"],
        ["263116", "Granisetron Hydrochloride Capsules", "No", "15-Mar-26", "No", "17-Mar-26", "Adverse Event", "", "1", "mg", "BID (twice a day)", "Oral"],
        ["263116", "Esomeprazole Magnesium Enteric-coated Tablets", "No", "24-Apr-26", "No", "4-Aug-26", "Medical History", "", "20", "mg", "PRN (as necessary)", "Oral"],
        ["263116", "Potassium Chloride Injection", "No", "6-Jun-26", "No", "8-Jun-26", "Adverse Event", "", "10", "mL", "BID (twice a day)", "Intravenous"],
        ["263116", "Human Albumin", "No", "9-Jun-26", "No", "12-Jun-26", "Adverse Event", "", "100", "mL", "QD (every day)", "Intravenous"],
        ["263116", "Montmorillonite Powder", "No", "7-Jul-26", "No", "UN-Jul-2026", "Other", "Antidiarrheal (related to bowel cancer resection surgery)", "3", "g", "TID (three times daily)", "Oral"],
        ["263116", "Thiamazole Tablets", "No", "8-Jul-26", "Yes", "", "Adverse Event", "", "10", "mg", "QD (every day)", "Oral"],
        ["263116", "Potassium Chloride Sustained-release Tablets", "No", "8-Jul-26", "No", "19-Jul-26", "Adverse Event", "", "1", "g", "BID (twice a day)", "Oral"],
        ["263116", "Compound Glycyrrhizin Capsules", "No", "16-Jul-26", "No", "31-Jul-26", "Adverse Event", "", "2", "Capsule", "TID (three times daily)", "Oral"],
        ["263116", "Domperidone Tablets", "No", "31-Jul-26", "No", "1-Aug-26", "Adverse Event", "", "1", "Caplet", "TID (three times daily)", "Oral"],
        ["263116", "Potassium Chloride Sustained-release Tablets", "No", "19-Aug-26", "Yes", "", "Adverse Event", "", "1", "g", "BID (twice a day)", "Oral"]
    ]
    t4_p2 = doc.add_table(rows=len(t4_p2_data), cols=12)
    set_table_borders(t4_p2, color=BORDER_COLOR)
    set_table_margins(t4_p2, top=6, bottom=6, left=20, right=20)
    for r_idx, row in enumerate(t4_p2.rows):
        format_row(row, height_pt=24)
        for c_idx, cell in enumerate(row.cells):
            format_cell(cell, t4_p2_data[r_idx][c_idx], t4_widths[c_idx], size=9.5, bold=False, color=BLUE)

    doc.add_page_break()

    # ==========================================
    # PAGE 3
    # ==========================================
    t4_p3_data = [
        ["263116", "Imipenem and Cilastatin Sodium for lnjection", "No", "27-Sep-26", "Yes", "", "Adverse Event", "", "1", "g", "Q8H (every 8 hours)", "Intravenous"],
        ["263116", "Norepinephrine Bitartrate Injection", "No", "20-Sep-26", "No", "28-Sep-26", "Adverse Event", "", "6", "mg", "Continuous", "Intravenous"],
        ["263116", "Hydroocortisone Injection", "No", "27-Sep-26", "No", "30-Sep-26", "Adverse Event", "", "30", "mL", "Continuous", "Intravenous"],
        ["263116", "Sodium Bicarbonate Injection", "No", "27-Sep-26", "Yes", "", "Adverse Event", "", "50", "mL", "Continuous", "Intravenous"],
        ["263116", "Compound Sodium Chloride Injection", "No", "27-Sep-26", "Yes", "", "Adverse Event", "", "500", "mL", "Continuous", "Intravenous"],
        ["263116", "Human Insulin Injection", "No", "27-Sep-26", "Yes", "", "Adverse Event", "", "8", "IU", "PRN (as necessary)", "Subcutaneous"],
        ["263116", "Omeprazole Sodium for Injection", "No", "27-Sep-26", "Yes", "", "Adverse Event", "", "40", "mg", "Continuous", "Intravenous"],
        ["263116", "Amiodarone Hydrochloride Injection", "No", "27-Sep-26", "No", "27-Sep-26", "Adverse Event", "", "0.15", "g", "Once", "Intravenous"],
        ["263116", "Vitamin C Injection", "No", "28-Sep-26", "Yes", "", "Adverse Event", "", "2", "g", "Continuous", "Intravenous"],
        ["263116", "Vitamin B6 Injection", "No", "28-Sep-26", "Yes", "", "Adverse Event", "", "100", "mg", "Continuous", "Intravenous"],
        ["263116", "Methylprednisolone Sodium Succinate for Injection", "No", "30-Sep-26", "Yes", "", "Adverse Event", "", "100", "mg", "Continuous", "Intravenous"],
        ["263116", "Human Immunoglobulin???pH4???for Intravenous Injection", "No", "30-Sep-26", "Yes", "", "Adverse Event", "", "2.5", "g", "Continuous", "Intravenous"]
    ]
    t4_p3 = doc.add_table(rows=len(t4_p3_data), cols=12)
    set_table_borders(t4_p3, color=BORDER_COLOR)
    set_table_margins(t4_p3, top=4, bottom=4, left=20, right=20)
    for r_idx, row in enumerate(t4_p3.rows):
        format_row(row, height_pt=18)
        for c_idx, cell in enumerate(row.cells):
            format_cell(cell, t4_p3_data[r_idx][c_idx], t4_widths[c_idx], size=9.5, bold=False, color=BLUE)

    # 1.5 baseline laboratory findings-refer to the separate attachment
    p_15 = doc.add_paragraph()
    p_15.paragraph_format.space_before = Pt(6)
    p_15.paragraph_format.space_after = Pt(3)
    r1 = p_15.add_run("1.5 baseline laboratory findings-")
    set_font(r1, size=11.5, bold=True, color=BLACK)
    r2 = p_15.add_run("refer to the separate attachment")
    set_font(r2, size=11.5, bold=False, color=BLUE)

    add_p("2. Please provide details of the endocrine evaluation, including cortisol, ACTH, other pituitary hormone levels, and pituitary MRI results (if performed). Please specify whether hormone samples were collected prior to corticosteroid administration.",
          size=11.5, bold=True, color=BLACK, space_before=3, space_after=2, indent=0.25)
    
    add_p("2.1 Site confirmed: Hormone samples were collected prior to corticosteroid administration.",
          size=11.5, bold=False, color=BLUE, space_before=1.5, space_after=1.5, indent=0.50)
    
    add_p("2.2 Test results", size=11.5, bold=False, color=BLUE, space_before=1.5, space_after=2, indent=0.50)

    # Table 2.2 Header + Row 1
    t22_widths = [1.93, 1.93, 1.91, 1.95] # Sum = 7.72 in
    t22_p3 = doc.add_table(rows=2, cols=4)
    set_table_borders(t22_p3, color=BORDER_COLOR)
    set_table_margins(t22_p3, top=4, bottom=4, left=25, right=25)
    format_row(t22_p3.rows[0], height_pt=14)
    format_row(t22_p3.rows[1], height_pt=14)
    for c_idx, text in enumerate(["Testing", "Result", "Unit", "Normal range"]):
        format_cell(t22_p3.rows[0].cells[c_idx], text, t22_widths[c_idx], size=10, bold=False, color=BLUE)
    for c_idx, text in enumerate(["cortisol", "12.64", "ug/dl", "7-9am:4.2-24.85"]):
        format_cell(t22_p3.rows[1].cells[c_idx], text, t22_widths[c_idx], size=10, bold=False, color=BLUE)

    doc.add_page_break()

    # ==========================================
    # PAGE 4
    # ==========================================
    t22_p4_data = [
        ["", "", "", "3-5pm: 2.9-17.3"],
        ["ACTH", "<1.00", "pg/ml", "7-10am 7.2-63.3"],
        ["TT3", "0.61", "nmol/L", "0.54-2.96"],
        ["TT4", "2.44↓", "ug/dl", "4.87-11.72"],
        ["TSH", "14.9941↑", "uIU/ml", "0.35-4.94"],
        ["FT3", "1.51↓", "pg/ml", "1.58-3.91"],
        ["FT4", "0.50↓", "ng/dl", "0.7-1.48"]
    ]
    t22_p4 = doc.add_table(rows=len(t22_p4_data), cols=4)
    set_table_borders(t22_p4, color=BORDER_COLOR)
    set_table_margins(t22_p4, top=3, bottom=3, left=25, right=25)
    for r_idx, row in enumerate(t22_p4.rows):
        format_row(row, height_pt=13)
        for c_idx, cell in enumerate(row.cells):
            format_cell(cell, t22_p4_data[r_idx][c_idx], t22_widths[c_idx], size=10, bold=False, color=BLUE)

    # Image for Page 4
    img4_path = resolve_image("page_04.png", input_pdf=input_pdf)
    p_img4 = doc.add_paragraph()
    p_img4.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_img4.paragraph_format.left_indent = Inches(0.5)
    p_img4.paragraph_format.space_before = Pt(2)
    p_img4.paragraph_format.space_after = Pt(0)
    run_img4 = p_img4.add_run()
    run_img4.add_picture(img4_path, width=Inches(8.0))

    doc.add_page_break()

    # ==========================================
    # PAGE 5
    # ==========================================
    img5_path = resolve_image("page_05.png", input_pdf=input_pdf)
    p_img5 = doc.add_paragraph()
    p_img5.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_img5.paragraph_format.left_indent = Inches(0.5)
    p_img5.paragraph_format.space_before = Pt(0)
    p_img5.paragraph_format.space_after = Pt(2)
    run_img5 = p_img5.add_run()
    run_img5.add_picture(img5_path, width=Inches(8.4))

    add_p("3. Was immune-mediated hypophysitis and/or secondary adrenal insufficiency diagnosed? If yes, please provide details.",
          size=11.5, bold=True, color=BLACK, space_before=3, space_after=2)
    add_p("Endocrinology department assessment: The patient was diagnosed with a tumor and underwent immunotherapy. Upon admission, ACTH, cortisol, and thyroid function were found to be decreased. Considering medical history, secondary adrenal insufficiency and secondary hypothyroidism were considered.",
          size=11.5, bold=False, color=BLUE, space_before=1.5, space_after=0, indent=0.25)

    doc.add_page_break()

    # ==========================================
    # PAGE 6
    # ==========================================
    add_p("4. Please provide the results of all microbiological investigations, including blood cultures, urine cultures, and other relevant cultures. Were any pathogens identified and was a source of infection confirmed?",
          size=11.5, bold=True, color=BLACK, space_before=0, space_after=2)
    
    add_p("1.  Urine culture, normal", size=11.5, bold=False, color=BLUE, space_before=1, space_after=1, indent=0.45)
    add_p("2.  Blood culture, results still pending", size=11.5, bold=False, color=BLUE, space_before=1, space_after=1, indent=0.45)
    add_p("3.  sputum Culture ：The source of infection is Pseudomonas aeruginosa.", size=11.5, bold=False, color=BLUE, space_before=1, space_after=2, indent=0.45)

    img6_path = resolve_image("page_06.png", input_pdf=input_pdf)
    p_img6 = doc.add_paragraph()
    p_img6.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_img6.paragraph_format.left_indent = Inches(0.93)
    p_img6.paragraph_format.space_before = Pt(2)
    p_img6.paragraph_format.space_after = Pt(2)
    run_img6 = p_img6.add_run()
    run_img6.add_picture(img6_path, width=Inches(5.97))

    add_p("5. Please provide details of hydrocortisone treatment, including dose, start date, and response to treatment.",
          size=11.5, bold=True, color=BLACK, space_before=2, space_after=1.5)
    add_p("Hydrocortisone injection, 100mg, was continuously pumped in with a micro pump. Starting from September 27, 2026, the patient's blood pressure gradually stabilized after hormone treatment and the antihypertensive drugs were discontinued.",
          size=11.5, bold=False, color=BLUE, space_before=1, space_after=2, indent=0.25)

    add_p("6. Please provide abdominal imaging and gastrointestinal evaluation findings. Was an intra-abdominal source of infection identified (e.g., obstruction, perforation, abscess, colitis, ischemia, or mesenteric panniculitis)?",
          size=11.5, bold=True, color=BLACK, space_before=2, space_after=1.5)
    add_p("6.1 Septic shock is considered to have an intestinal source of infection, but the specific source of infection is unknown",
          size=11.5, bold=False, color=BLUE, space_before=1, space_after=0, indent=0.25)

    doc.add_page_break()

    # ==========================================
    # PAGE 7
    # ==========================================
    add_p("6.2 Gastrointestinal surgery consultation: Physical examination shows that the entire abdomen is flat and soft, with tenderness and no expression of pain. In combination with CT, there is currently no special treatment in the field of gastrointestinal surgery.",
          size=11.5, bold=False, color=BLUE, space_before=0, space_after=2, indent=0.25)
    add_p("6.3  CT scan report(CT enhancement [lower abdomen, upper abdomen, pelvic cavity], CT plain scan [cranio, thoracic])",
          size=11.5, bold=False, color=BLUE, space_before=1.5, space_after=2, indent=0.25)

    p7_items = [
        "1. Multiple ischemic lesions in bilateral basal ganglion areas and radiating crowns; Brain atrophy;",
        "2. Bronchitis, emphysema signs;",
        "3. Thickening of the interlobular septa in some lungs, pseudo-interstitial pulmonary edema:",
        "4. The solid nodule in the anterior inner basal segment of the left lower lobe of the lung is not obvious:",
        "5.The medial segment of the right middle lobe and the lower lingual segment of the left upper lobe have some inflammatory cords as before;",
        "6.Inflammation of the lower lobes of both lungs; recommended follow-up after treatment; Small mucus plugs in both main bronchi;",
        "7. Aortic and coronary artery sclerosis are the same as before, with minor effusion in both pleural cavities;",
        "8.Swelling and thickening of the middle and lower esophageal wall. Please combine with clinical and related examinations:",
        "9. Postoperative defects of the right hemicolon are as before; the intestinal wall of the anastomotic mouth shows no thickening, and the fat space around the anastomosis is slightly reduced compared to before.",
        "Slightly thickened peritoneum near the peritoneum improves compared to before; A few small lymph nodes in the mesenteric region are the same as before; follow-up and follow-up examination are recommended:",
        "10. Mild fatty liver is no longer obvious;",
        "11. Chronic cholecystitis with cholestasis:",
        "12. Left kidney cyst same as before;",
        "13. Prostate calcification lesions are the same as before;",
        "14. Slight pelvic effusion."
    ]
    for item in p7_items:
        add_p(item, size=11.5, bold=False, color=BLUE, space_before=1, space_after=1, indent=0.25)

    doc.add_page_break()

    # ==========================================
    # PAGE 8
    # ==========================================
    img8_path = resolve_image("page_08.png", input_pdf=input_pdf)
    p_img8 = doc.add_paragraph()
    p_img8.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_img8.paragraph_format.left_indent = Inches(0.5)
    p_img8.paragraph_format.space_before = Pt(0)
    p_img8.paragraph_format.space_after = Pt(2)
    run_img8 = p_img8.add_run()
    run_img8.add_picture(img8_path, width=Inches(4.24))

    add_p("7. Please provide relevant laboratory results at baseline and at the time of the event, including hematology, inflammatory markers, chemistry, blood gas analysis, lactate, and coagulation parameters.",
          size=11.5, bold=True, color=BLACK, space_before=3, space_after=1.5)
    add_p("laboratory results at baseline: refer to the separate attachment.",
          size=11.5, bold=False, color=BLUE, space_before=1, space_after=0, indent=0.25)

    doc.add_page_break()

    # ==========================================
    # PAGE 9
    # ==========================================
    add_p("laboratory results at the time of the event: at the end of the text.",
          size=11.5, bold=False, color=BLUE, space_before=0, space_after=2, indent=0.50)
    
    add_p("8. Please provide details supporting the investigator's assessment that no alternative etiology adequately explains the event.",
          size=11.5, bold=True, color=BLACK, space_before=2, space_after=1.5, indent=0.25)
    
    add_p("After the patient was admitted to the hospital, a multidisciplinary consultation was organized, including the Department of Gastrointestinal Surgery, the Department of Intestinal and Hepatobiliary Pancreatic Oncology, the Department of Neurology, the antibacterial drug consultation expert group, the Department of Endocrinology, and the Department of Critical Care Medicine. The multidisciplinary consensus was: The patient is a typical case of severe sepsis secondary to immunotherapy. The infection focus is considered to be of intestinal origin and related to immunotherapy. There is also immune-related multiple organ function impairment, which overlaps with severe infection.",
          size=11.5, bold=False, color=BLUE, space_before=1.5, space_after=2, indent=0.50)
    
    add_p("On September 30, 2026, a hospital-wide consultation was held again. The multi-disciplinary consultation of the entire hospital reached a consensus: The patient has typical immune-related multiple organ function damage that occurred after immunotherapy, combined with sepsis. The infection focus is considered to be of intestinal origin and related to immunotherapy. Anti-infection treatment for sepsis is currently continuing. Immune-damaged organs: encephalitis, pituitary inflammation, blood system, heart system, intestinal system, liver.",
          size=11.5, bold=False, color=BLUE, space_before=1.5, space_after=2, indent=0.50)
    
    add_p("Treatment suggestion: Administer intravenous human immunoglobulin and methylprednisolone (100mg intravenously every day) for symptomatic treatment. Plasma exchange therapy is planned to be performed on October 1, 2026.",
          size=11.5, bold=False, color=BLUE, space_before=1.5, space_after=2, indent=0.50)
    
    add_p('Subsequently, the name of SAE "hypophysitis" will be revised to "Immune-related multiple organ dysfunction (hypophysitis, encephalitis, colitis)", and the grade will be corrected to CTCA grade 4 (life-threatening). The follow-up of SA "immune-related multiple organ dysfunction" will be updated and reported.',
          size=11.5, bold=False, color=BLUE, space_before=1.5, space_after=2, indent=0.50)
    
    add_p("9. Please clarify whether the multidisciplinary team considered the event to be due to sepsis/septic shock, an immune-mediated endocrine event, or a combination of both.",
          size=11.5, bold=True, color=BLACK, space_before=3, space_after=1.5, indent=0.25)
    
    add_p("The hospital's multiple disciplines have reached a consensus: The patient is a typical case of severe sepsis secondary to immunotherapy, with immune-related multi-organ functional impairment, which overlaps with severe infection.",
          size=11.5, bold=False, color=BLUE, space_before=1.5, space_after=2, indent=0.50)
    
    add_p("laboratory results at the time of the event as below for reference:",
          size=11.5, bold=False, color=BLUE, space_before=1.5, space_after=0, indent=0.50)

    doc.add_page_break()

    # ==========================================
    # PAGES 10 to 41 (Lab reports)
    # ==========================================
    for p in range(10, 42):
        img_name = f"page_{p:02d}.png"
        img_path = resolve_image(img_name, input_pdf=input_pdf)
        p_img = doc.add_paragraph()
        if p >= 38:
            p_img.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_img.paragraph_format.left_indent = Inches(0)
            img_width = Inches(8.45)
        else:
            p_img.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_img.paragraph_format.left_indent = Inches(0.5)
            img_width = Inches(8.35)

        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after = Pt(0)
        p_img.paragraph_format.line_spacing = 1.0
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=img_width)
        if p < 41:
            doc.add_page_break()

    doc.save(str(out_file))
    print(f"Document successfully created and saved to: {out_file}")
    return str(out_file)

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Build Word doc matching Chinese Medical PDF")
    parser.add_argument("-i", "--input", default=None, help="Input PDF path")
    parser.add_argument("-o", "--output", default=None, help="Output DOCX path")
    args = parser.parse_args()
    create_document(input_pdf=args.input, output_path=args.output)
