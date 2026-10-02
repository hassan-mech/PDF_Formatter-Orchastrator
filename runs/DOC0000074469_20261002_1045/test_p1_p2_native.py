import docx
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_BREAK
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
from pathlib import Path
import pymupdf

pdf_path = Path("finished/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
doc_pdf = pymupdf.open(str(pdf_path))

crop_dir = Path("runs/DOC0000074469_20261002_1045/extract")
crop_dir.mkdir(parents=True, exist_ok=True)
p1_abs_path = crop_dir / "p1_eng_abstract_clean.png"
if not p1_abs_path.exists():
    pix = doc_pdf[0].get_pixmap(clip=pymupdf.Rect(98, 520, 398, 696), dpi=300)
    pix.save(str(p1_abs_path))

p2_abs_path = crop_dir / "p2_eng_abstract_clean.png"
pix2 = doc_pdf[1].get_pixmap(clip=pymupdf.Rect(70, 113, 369, 158), dpi=300)
pix2.save(str(p2_abs_path))

doc = docx.Document()

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1
# ══════════════════════════════════════════════════════════════════════════════
sec1 = doc.sections[0]
sec1.page_width = Inches(8.27)
sec1.page_height = Inches(10.63)
sec1.top_margin = Inches(0.40)
sec1.bottom_margin = Inches(0.40)
sec1.left_margin = Inches(1.38)
sec1.right_margin = Inches(0.96)

# 1. Top Stamp
p_stamp = doc.add_paragraph()
p_stamp.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_stamp.paragraph_format.space_before = Pt(0)
p_stamp.paragraph_format.space_after = Pt(8)
r_st = p_stamp.add_run("DOC0000074469   Reviewed by TP: 30SEP2026 07:19AM CET")
r_st.font.name = "Arial"
r_st.font.size = Pt(9.0)
r_st.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

# 2. Header Line with Tab Stop
p_hdr1 = doc.add_paragraph()
p_hdr1.paragraph_format.space_before = Pt(0)
p_hdr1.paragraph_format.space_after = Pt(0)
p_hdr1.paragraph_format.tab_stops.add_tab_stop(Inches(5.93), WD_TAB_ALIGNMENT.RIGHT)

r_h1 = p_hdr1.add_run("CASO CLÍNICO")
r_h1.font.name = "Arial"
r_h1.font.size = Pt(8.5)
r_h1.font.bold = True
r_h1.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

p_hdr1.add_run("\t")
r_logo = p_hdr1.add_run("Dermatología")
r_logo.font.name = "Arial"
r_logo.font.size = Pt(15.5)
r_logo.font.italic = True
r_logo.font.bold = True
r_logo.font.underline = True
r_logo.font.color.rgb = RGBColor(0xDC, 0x1D, 0x44)

p_hdr2 = doc.add_paragraph()
p_hdr2.paragraph_format.space_before = Pt(0)
p_hdr2.paragraph_format.space_after = Pt(0)
p_hdr2.paragraph_format.tab_stops.add_tab_stop(Inches(5.93), WD_TAB_ALIGNMENT.RIGHT)

r_h2 = p_hdr2.add_run("Dermatol Rev Mex 2026; 70 (5): 666-672.")
r_h2.font.name = "Arial"
r_h2.font.size = Pt(8.0)
r_h2.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

p_hdr2.add_run("\t")
r_h2_tab = p_hdr2.add_run("R e v i s t a   m e x i c a n a")
r_h2_tab.font.name = "Arial"
r_h2_tab.font.size = Pt(6.5)
r_h2_tab.font.bold = True
r_h2_tab.font.color.rgb = RGBColor(0xDC, 0x1D, 0x44)

p_hdr3 = doc.add_paragraph()
p_hdr3.paragraph_format.space_before = Pt(0)
p_hdr3.paragraph_format.space_after = Pt(14)
r_h3 = p_hdr3.add_run("https://doi.org/10.24245/dermatolrevmex.v70i5.11433")
r_h3.font.name = "Arial"
r_h3.font.size = Pt(7.5)
r_h3.font.color.rgb = RGBColor(0x00, 0x55, 0xA5)

# 3. Main Spanish Title (Exact 3 lines)
p_tit = doc.add_paragraph()
p_tit.paragraph_format.space_before = Pt(8)
p_tit.paragraph_format.space_after = Pt(4)
p_tit.paragraph_format.line_spacing = Pt(19.0)
r_tit = p_tit.add_run("Melanoma metastásico, un caso\nextraordinario en un paciente con\ntrasplante renal")
r_tit.font.name = "Arial"
r_tit.font.size = Pt(16.5)
r_tit.font.bold = True
r_tit.font.color.rgb = RGBColor(0x00, 0x35, 0x84)

# 4. Crimson divider line spanning only 4.13 inches
p_line = doc.add_paragraph()
p_line.paragraph_format.space_before = Pt(2)
p_line.paragraph_format.space_after = Pt(6)
p_line.paragraph_format.right_indent = Inches(1.80)
pPr = p_line._p.get_or_add_pPr()
pPr.append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="1" w:color="DB1D43"/></w:pBdr>'))

# 5. English Title (Genuine Styled Text, 2 lines)
p_etit = doc.add_paragraph()
p_etit.paragraph_format.space_before = Pt(4)
p_etit.paragraph_format.space_after = Pt(10)
p_etit.paragraph_format.line_spacing = Pt(16.0)
r_etit = p_etit.add_run("Metastatic melanoma, an extraordinary case\nin a post-transplant kidney patient.")
r_etit.font.name = "Georgia"
r_etit.font.size = Pt(13.5)
r_etit.font.italic = True
r_etit.font.bold = True
r_etit.font.color.rgb = RGBColor(0x6D, 0x6E, 0x71)

# 6. Authors
p_auth = doc.add_paragraph()
p_auth.paragraph_format.space_before = Pt(2)
p_auth.paragraph_format.space_after = Pt(12)
r_au = p_auth.add_run(
    "María Fernanda Corona Rosas,¹ Betzabé Quiles Martínez,² Yelitza Esmeralda Campos Salgado,⁴ Judith Domínguez Cherit³"
)
r_au.font.name = "Arial"
r_au.font.size = Pt(8.5)
r_au.font.bold = True
r_au.font.color.rgb = RGBColor(0x23, 0x1F, 0x20)

# ── Section 2: Continuous 2 Unequal Columns ──
sec2 = doc.add_section(WD_SECTION_START.CONTINUOUS)
sec2.top_margin = Inches(0.40)
sec2.bottom_margin = Inches(0.40)
sec2.left_margin = Inches(1.38)
sec2.right_margin = Inches(0.96)
sectPr2 = sec2._sectPr
cols_xml = f'''<w:cols {nsdecls("w")} w:num="2" w:equalWidth="0">
  <w:col w:w="5980" w:space="360"/>
  <w:col w:w="2220"/>
</w:cols>'''
sectPr2.append(parse_xml(cols_xml))

def add_shaded_p(text, bold_prefix="", font_size=7.8, space_after=3.0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = Pt(10.5)
    p.paragraph_format.left_indent = Pt(4)
    p.paragraph_format.right_indent = Pt(4)
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FCECEF"/>'))
    if bold_prefix:
        r_b = p.add_run(bold_prefix)
        r_b.font.name = "Arial"
        r_b.font.size = Pt(font_size)
        r_b.font.bold = True
        r_b.font.color.rgb = RGBColor(0xDB, 0x1D, 0x43)
    r_t = p.add_run(text)
    r_t.font.name = "Arial"
    r_t.font.size = Pt(font_size)
    r_t.font.color.rgb = RGBColor(0x23, 0x1F, 0x20)
    return p

# Left Column Content (Pink Box)
p_rh = doc.add_paragraph()
p_rh.paragraph_format.space_before = Pt(4)
p_rh.paragraph_format.space_after = Pt(2)
p_rh.paragraph_format.left_indent = Pt(4)
p_rh._p.get_or_add_pPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FCECEF"/>'))
r_rh = p_rh.add_run("Resumen")
r_rh.font.name = "Arial"
r_rh.font.size = Pt(9.0)
r_rh.font.bold = True
r_rh.font.color.rgb = RGBColor(0xDB, 0x1D, 0x43)

add_shaded_p(" El cáncer de piel en pacientes que recibieron algún trasplante de órgano representa el 30% de los casos reportados en la bibliografía. El 90% de los casos registrados son carcinomas epidermoides y menos del 5% corresponde a melanoma. Los factores de riesgo son el uso de terapia inmunodepresora, el tabaquismo activo, el alcoholismo y la vida sedentaria. Por tal motivo las guías internacionales recomiendan la realización de pruebas de detección de cáncer de piel en receptores de trasplante renal, al menos cada año. Lamentablemente en México no se dispone de algún programa de vigilancia anual de piel y reporte de las anomalías encontradas, a pesar de que la detección oportuna y el tratamiento en etapas tempranas del cáncer de piel aumentan la calidad y esperanza de vida en estos pacientes.", bold_prefix="ANTECEDENTES:", space_after=2.5)

add_shaded_p(" Paciente masculino de 55 años, con índice tabáquico de 28 y enfermedad renal que ameritó trasplante de donador vivo, en tratamiento con terapia inmunodepresora durante seis años. Su padecimiento inició en 2023 con una lesión tumoral localizada en el tronco, con avance significativo en menos de cinco meses, así como datos clínicos sugerentes de metástasis cerebral, con deterioro de la funcionalidad y estadio avanzado, por lo que recibió tratamiento paliativo.", bold_prefix="CASO CLÍNICO:", space_after=2.5)

add_shaded_p(" Se insiste en la importancia de la vigilancia y detección de los pacientes postrasplantados porque la enfermedad de base y la terapia inmunodepresora incrementan hasta un 30% el riesgo de padecer algún tipo de neoplasia.", bold_prefix="CONCLUSIONES:", space_after=2.5)

add_shaded_p(" Melanoma; trasplante renal; terapia inmunosupresora.", bold_prefix="PALABRAS CLAVE:", space_after=4.0)

# English Abstract Screened Image in Left Column
p_abs = doc.add_paragraph()
p_abs.paragraph_format.space_before = Pt(0)
p_abs.paragraph_format.space_after = Pt(4)
p_abs.paragraph_format.left_indent = Pt(4)
p_abs._p.get_or_add_pPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FCECEF"/>'))
r_abs = p_abs.add_run()
r_abs.add_picture(str(p1_abs_path), width=Inches(4.08))

# Connected Hidden Text
p_hid = doc.add_paragraph()
p_hid.paragraph_format.space_before = Pt(0)
p_hid.paragraph_format.space_after = Pt(0)
p_hid.paragraph_format.line_spacing = Pt(1)
r_h = p_hid.add_run(
    "Abstract BACKGROUND: Skin cancer in patients who have received an organ transplant account for 30% "
    "of the cases reported in the literature, 90% of the registered cases being squamous cell carcinoma "
    "and less than 5% being melanoma. Risk factors include the use of immunosuppressive therapy, which "
    "is influenced by the type of medication, as well as its duration, and current habits such as active "
    "smoking, alcoholism and a sedentary lifestyle. The international guidelines for kidney transplant "
    "recipients suggest performing skin cancer screening tests at least once a year, which unfortunately "
    "a successful program with an annual skin checkout and the report of the abnormalities found are "
    "not available in Mexico, despite timely detection and treatment in early stages increase the quality "
    "and life expectancy of these patients. CLINICAL CASE: A 55-year-old male patient, with a smoking "
    "index of 28 and kidney disease that required a transplant from a living donor, under treatment with "
    "immunosuppressive therapy for 6 years. His condition began in 2023 with a tumor lesion located in "
    "the trunk, with significant progress in less than 5 months, as well as symptoms suggestive of "
    "metastasis at the brain level, with deterioration of its functionality and advanced stage. It was "
    "decided to provide palliative treatment."
)
r_h.font.size = Pt(0.5)
r_h.font.hidden = True

# Column Break to go to Right Column
r_break = p_hid.add_run()
r_break.add_break(WD_BREAK.COLUMN)

def add_affil_p(text, sup_prefix=""):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = Pt(8.5)
    if sup_prefix:
        r_s = p.add_run(sup_prefix + " ")
        r_s.font.name = "Arial"
        r_s.font.size = Pt(6.5)
        r_s.font.bold = True
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(6.8)
    r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

add_affil_p("Residente de tercer año de Medicina Interna, Hospital General de Zona 47, Instituto Mexicano del Seguro Social, Ciudad de México.", sup_prefix="¹")
add_affil_p("Residente de tercer año, Departamento de Dermatología.", sup_prefix="²")
add_affil_p("Jefa del Departamento de Dermatología. Instituto Nacional de Ciencias Médicas y Nutrición Salvador Zubirán, Ciudad de México.", sup_prefix="³")
add_affil_p("Residente de tercer año, Departamento de Oncología, Centro Médico Nacional Siglo XXI, Instituto Mexicano del Seguro Social, Ciudad de México.", sup_prefix="⁴")

p_meta = doc.add_paragraph()
p_meta.paragraph_format.space_before = Pt(6)
p_meta.paragraph_format.space_after = Pt(1)
r_m1 = p_meta.add_run("ORCID\n")
r_m1.font.name = "Arial"
r_m1.font.size = Pt(7.0)
r_m1.font.bold = True
r_m1.font.color.rgb = RGBColor(0xDB, 0x1D, 0x43)

r_m2 = p_meta.add_run(
    "https://orcid.org/0009-0003-2708-6828\n"
    "https://orcid.org/0009-0009-4703-819X\n"
    "https://orcid.org/0009-0007-8221-8941\n"
    "https://orcid.org/0000-0003-3542-4615"
)
r_m2.font.name = "Arial"
r_m2.font.size = Pt(6.5)
r_m2.font.color.rgb = RGBColor(0x00, 0x55, 0xA5)

p_dates = doc.add_paragraph()
p_dates.paragraph_format.space_before = Pt(6)
p_dates.paragraph_format.space_after = Pt(6)
r_d1 = p_dates.add_run("Recibido: ")
r_d1.font.bold = True
r_d1.font.color.rgb = RGBColor(0xDB, 0x1D, 0x43)
r_d1.font.size = Pt(7.0)
r_d2 = p_dates.add_run("abril 2024\n")
r_d2.font.size = Pt(7.0)

r_d3 = p_dates.add_run("Aceptado: ")
r_d3.font.bold = True
r_d3.font.color.rgb = RGBColor(0xDB, 0x1D, 0x43)
r_d3.font.size = Pt(7.0)
r_d4 = p_dates.add_run("febrero 2025")
r_d4.font.size = Pt(7.0)

p_cor = doc.add_paragraph()
p_cor.paragraph_format.space_before = Pt(4)
p_cor.paragraph_format.space_after = Pt(6)
r_c1 = p_cor.add_run("Correspondencia\n")
r_c1.font.bold = True
r_c1.font.color.rgb = RGBColor(0xDB, 0x1D, 0x43)
r_c1.font.size = Pt(7.0)
r_c2 = p_cor.add_run("María Fernanda Corona Rosas\nfercro15@gmail.com")
r_c2.font.size = Pt(7.0)

p_cite = doc.add_paragraph()
p_cite.paragraph_format.space_before = Pt(4)
p_cite.paragraph_format.space_after = Pt(0)
r_ci1 = p_cite.add_run("Este artículo debe citarse como:\n")
r_ci1.font.bold = True
r_ci1.font.color.rgb = RGBColor(0xDB, 0x1D, 0x43)
r_ci1.font.size = Pt(7.0)
r_ci2 = p_cite.add_run(
    "Corona-Rosas MF, Quiles-Martínez B, Campos-Salgado YE, Domínguez-Cherit J. "
    "Melanoma metastásico, un caso extraordinario en un paciente con trasplante renal. "
    "Dermatol Rev Mex 2026; 70 (5): 666-672."
)
r_ci2.font.size = Pt(6.8)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2 (Even Page)
# ══════════════════════════════════════════════════════════════════════════════
sec_p2_top = doc.add_section(WD_SECTION_START.NEW_PAGE)
sec_p2_top.page_width = Inches(8.27)
sec_p2_top.page_height = Inches(10.63)
sec_p2_top.top_margin = Inches(0.40)
sec_p2_top.bottom_margin = Inches(0.40)
sec_p2_top.left_margin = Inches(0.98)
sec_p2_top.right_margin = Inches(1.34)

# Running Header Page 2
p2_hdr = doc.add_paragraph()
p2_hdr.paragraph_format.space_before = Pt(0)
p2_hdr.paragraph_format.space_after = Pt(0)
p2_hdr.paragraph_format.tab_stops.add_tab_stop(Inches(5.95), WD_TAB_ALIGNMENT.RIGHT)

r2_h1 = p2_hdr.add_run("Corona Rosas MF, et al. Melanoma metastásico y trasplante renal")
r2_h1.font.name = "Arial"
r2_h1.font.size = Pt(7.5)
r2_h1.font.italic = True
r2_h1.font.color.rgb = RGBColor(0x77, 0x77, 0x77)

p2_hdr.add_run("\t")
r2_h2 = p2_hdr.add_run("Dermatología ")
r2_h2.font.name = "Arial"
r2_h2.font.size = Pt(7.5)
r2_h2.font.bold = True
r2_h2.font.italic = True
r2_h2.font.color.rgb = RGBColor(0xDC, 0x1D, 0x44)

r2_h3 = p2_hdr.add_run("Revista mexicana")
r2_h3.font.name = "Arial"
r2_h3.font.size = Pt(7.0)
r2_h3.font.color.rgb = RGBColor(0xDC, 0x1D, 0x44)

# Divider line
p2_div = doc.add_paragraph()
p2_div.paragraph_format.space_before = Pt(2)
p2_div.paragraph_format.space_after = Pt(4)
p2_div._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:space="1" w:color="CCCCCC"/></w:pBdr>'))

# Screened English Abstract Continuation (Pink box width 4.13 in)
p2_abs = doc.add_paragraph()
p2_abs.paragraph_format.space_before = Pt(2)
p2_abs.paragraph_format.space_after = Pt(0)
p2_abs.paragraph_format.left_indent = Pt(4)
p2_abs.paragraph_format.right_indent = Inches(1.82)
p2_abs._p.get_or_add_pPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="FCECEF"/>'))
r2_ab = p2_abs.add_run()
r2_ab.add_picture(str(p2_abs_path), width=Inches(4.13))

# Hidden text
p2_hid = doc.add_paragraph()
p2_hid.paragraph_format.space_before = Pt(0)
p2_hid.paragraph_format.space_after = Pt(65)
p2_hid.paragraph_format.line_spacing = Pt(1)
r2_hd = p2_hid.add_run(
    "CONCLUSIONS: The importance of monitoring and detecting post-transplant patients is emphasized, "
    "since the underlying disease and immunosuppressive therapy increase up to 30% of acquiring some type of neoplasia. "
    "KEYWORDS: Melanoma; Kidney transplant; Immunosuppressive therapy."
)
r2_hd.font.size = Pt(0.5)
r2_hd.font.hidden = True

# Continuous Break to 2 Equal Columns for Body Text
sec_p2_cols = doc.add_section(WD_SECTION_START.CONTINUOUS)
sec_p2_cols.top_margin = Inches(0.40)
sec_p2_cols.bottom_margin = Inches(0.40)
sec_p2_cols.left_margin = Inches(0.98)
sec_p2_cols.right_margin = Inches(1.34)
sectPr_p2 = sec_p2_cols._sectPr
cols_p2_xml = f'<w:cols {nsdecls("w")} w:num="2" w:space="346" w:equalWidth="1"/>'
sectPr_p2.append(parse_xml(cols_p2_xml))

# Left Column
p_h_ant = doc.add_paragraph()
p_h_ant.paragraph_format.space_before = Pt(0)
p_h_ant.paragraph_format.space_after = Pt(3)
r_ha = p_h_ant.add_run("ANTECEDENTES")
r_ha.font.name = "Arial"
r_ha.font.size = Pt(10.5)
r_ha.font.bold = True
r_ha.font.color.rgb = RGBColor(0xDB, 0x1D, 0x43)

def add_body_p(text, space_after=4.0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = Pt(10.2)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(8.0)
    r.font.color.rgb = RGBColor(0x23, 0x1F, 0x20)
    return p

add_body_p("En la actualidad, gracias a los avances tecnológicos y científicos, hay un aumento en la ejecución de trasplante como alternativa de tratamiento para los pacientes con enfermedad renal terminal, quienes posteriormente reciben tratamientos inmunosupresores para evitar el rechazo del injerto; esta terapia, por lo general, es de por vida.¹ Sin embargo, poco se habla de las consecuencias documentadas propias de la enfermedad y del uso crónico de terapia inmunodepresora, de las que el cáncer de piel es la neoplasia maligna más documentada.² Casi el 90% de los cánceres de piel de pacientes postrasplantados corresponden al carcinoma epidermoide y menos del 5% a melanoma, que tiene altos índices de mortalidad por su característica molecular invasora.³")

add_body_p("Por tal razón la descripción de este caso clínico cobra relevancia, no sólo por el estadio avanzado en el que se encontró al paciente, sino por la oportunidad de detección primaria de esa neoplasia, al estar en tratamiento de inmunosupresión durante muchos años y sus factores de riesgo (tabaquismo), lo que favoreció el avance de la neoplasia y su carácter invasor que tuvo un desenlace fatal. La bibliografía indica una supervivencia de sólo 3% a dos años.⁴")

# Column Break to Right Column!
p_cbrk = doc.add_paragraph()
p_cbrk.paragraph_format.space_before = Pt(0)
p_cbrk.paragraph_format.space_after = Pt(0)
p_cbrk.paragraph_format.line_spacing = Pt(1)
r_cbrk = p_cbrk.add_run()
r_cbrk.add_break(WD_BREAK.COLUMN)

# Right Column
p_h_caso = doc.add_paragraph()
p_h_caso.paragraph_format.space_before = Pt(0)
p_h_caso.paragraph_format.space_after = Pt(3)
r_hc = p_h_caso.add_run("CASO CLÍNICO")
r_hc.font.name = "Arial"
r_hc.font.size = Pt(10.5)
r_hc.font.bold = True
r_hc.font.color.rgb = RGBColor(0xDB, 0x1D, 0x43)

add_body_p("Paciente masculino de 55 años, albañil, sin antecedentes familiares o personales de melanoma, tabaquismo positivo a razón de 15 cigarrillos al día durante 38 años, con un índice tabáquico de 28; antecedente de enfermedad renal crónica diagnosticada en 2010 que ameritó el trasplante de donador vivo en 2014. Recibió tratamiento inmunodepresor con prednisona a dosis de 5 mg cada 24 horas y tacrolimus 8 mg cada 24 horas, con últimas concentraciones séricas de tacrolimus documentadas en 2019 de 8 ng/mL sin llegar a la toxicidad.")

add_body_p("El paciente acudió a consulta externa de Oncología en 2023 por padecer una dermatosis localizada en el tronco, que abarcaba la mitad de la zona torácica posterior, con extensión al hueco axilar, caracterizada por una gran placa tumoral negruzca y múltiples neoformaciones de distintos tamaños; la mayor era de 3 cm. Esta gran lesión estaba exulcerada, alrededor de la gran placa se observaban lesiones satélites que medían unos cuantos milímetros a centímetros (satelitosis), se apreciaba infiltrada a la palpación y sin datos de sangrado. Figura 1")

add_body_p("El paciente refirió el inicio de la dermatosis seis meses antes de la consulta; sin embargo, no se")

out_docx = Path("runs/DOC0000074469_20261002_1045/test_p1_p2_native.docx")
doc.save(str(out_docx))
print(f"Saved {out_docx}")
