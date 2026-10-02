import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def build_full_approved_judgment(output_path=None):
    doc = docx.Document()
    
    # Base styling
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)
    normal_style.paragraph_format.space_before = Pt(0)
    
    # -------------------------------------------------------------
    # SECTION 1: Pages 1 & 2
    # -------------------------------------------------------------
    s1 = doc.sections[0]
    s1.page_width = Inches(8.27)
    s1.page_height = Inches(11.69)
    s1.top_margin = Inches(0.8)
    s1.bottom_margin = Inches(0.8)
    s1.left_margin = Inches(1.0)
    s1.right_margin = Inches(1.0)
    
    # --- PAGE 1: Notary Certificate ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("SAVILLE & CO")
    r.bold = True
    r.font.size = Pt(18)
    p.paragraph_format.space_after = Pt(2)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("—   SCRIVENER NOTARIES   —")
    r.font.size = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Saville Notaries LLP   11 Old Jewry   London EC2R 8DU\n")
    r.font.size = Pt(8.5)
    r2 = p.add_run("Tel: +44 (0)20 7776 9800   www.savillenotaries.com   mail@savillenotaries.com")
    r2.font.size = Pt(8.5)
    p.paragraph_format.space_after = Pt(4)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Sophie Milburn   Nicholas Thompson   Robert Kerss   Andrew MacNab   Christopher Higgins*\n")
    r.font.size = Pt(8)
    r2 = p.add_run("Eleonora Ceolin*   Kyriaki Manika   Saffiyah Mengrani*   Olga Kulikovskaya*   Roman Egorov*   Yoana Georgieva*   Marta Maldonado Prados*")
    r2.font.size = Pt(8)
    p.paragraph_format.space_after = Pt(40)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.2
    p.paragraph_format.space_after = Pt(16)
    r = p.add_run("TO ALL TO WHOM THESE PRESENTS SHALL COME, I MARTA MALDONADO PRADOS of the City of London NOTARY PUBLIC by royal authority duly admitted and sworn DO HEREBY CERTIFY that the document hereunto annexed is a true printout of the digital transcription of the ")
    r_b1 = p.add_run("approved judgment")
    r_b1.bold = True
    r2 = p.add_run(" relating to INVEST INTERNATIONAL CAPITAL B.V. and LEWA TRADING INDUSTRY AND CONTRACTING CO. LTD., in respect of claim number ")
    r_b2 = p.add_run("CL-2026-000221")
    r_b2.bold = True
    r3 = p.add_run(", in THE HIGH COURT OF JUSTICE, BUSINESS AND PROPERTY COURTS OF ENGLAND AND WALES, COMMERCIAL COURT (KBD), issued electronically by MARTEN WALSH CHERER LIMITED, a private limited company duly organised and existing under the laws of England and Wales, registered with the Registrar of Companies for England and Wales under number 2669638 and with registered address at 27 Old Gloucester Street, London WC1N 3AX, England.")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.2
    p.paragraph_format.space_after = Pt(32)
    p.add_run("IN FAITH AND TESTIMONY WHEREOF I the said notary have subscribed my name and set and affixed my seal of office at London aforesaid this sixteenth day of September two thousand and twenty six.")
    
    # Seal & Signature Table
    table_sig = doc.add_table(rows=1, cols=2)
    table_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_sig.autofit = False
    table_sig.columns[0].width = Inches(3.2)
    table_sig.columns[1].width = Inches(3.0)
    
    c_seal = table_sig.cell(0, 0)
    c_sig = table_sig.cell(0, 1)
    
    p_seal = c_seal.paragraphs[0]
    p_seal.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p_seal.add_run("[seal: MARTA MALDONADO PRADOS * NOTARY PUBLIC *]")
    r.bold = True
    
    p_sig = c_sig.paragraphs[0]
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p_sig.add_run("[hw: Marta M.P.]\n")
    r.italic = True
    r_sig = p_sig.add_run("[signature]")
    r_sig.bold = True
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(45)
    p.paragraph_format.space_after = Pt(0)
    
    # Footer table with top border
    table_foot = doc.add_table(rows=1, cols=2)
    table_foot.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_foot.autofit = False
    table_foot.columns[0].width = Inches(4.5)
    table_foot.columns[1].width = Inches(1.8)
    
    for c in table_foot.rows[0].cells:
        tcPr = c._tc.get_or_add_tcPr()
        tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>')
        tcPr.append(tcBorders)
    
    p = table_foot.cell(0, 0).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run("Saville & Co. Scrivener Notaries is the trading name of Saville Notaries LLP, a limited liability partnership registered in England and Wales with registered number OC420687 and with registered office at 11 Old Jewry, London EC2R 8DU. Regulated through the Faculty Office of the Archbishop of Canterbury\n*General Notary")
    r.font.size = Pt(7)
    
    p = table_foot.cell(0, 1).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("[logo: SCRIVENER NOTARIES] | Member Notariat of the International Union of Notaries (UINL)")
    r.font.size = Pt(7)
    
    # --- PAGE 2: Apostille ---
    doc.add_page_break()
    
    table_apo = doc.add_table(rows=11, cols=2)
    table_apo.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_apo.autofit = False
    table_apo.columns[0].width = Inches(2.6)
    table_apo.columns[1].width = Inches(3.6)
    
    tblPr = table_apo._tbl.tblPr
    tblBorders = parse_xml(f'''<w:tblBorders {nsdecls("w")}>
        <w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>
        <w:left w:val="single" w:sz="12" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>
        <w:right w:val="single" w:sz="12" w:space="0" w:color="000000"/>
        <w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        <w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>
    </w:tblBorders>''')
    tblPr.append(tblBorders)
    
    def format_cell(cell, bold_label, sub_label=None, value=None, align=WD_ALIGN_PARAGRAPH.LEFT):
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.line_spacing = 1.05
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        if bold_label:
            r = p.add_run(bold_label)
            r.bold = True
            r.font.size = Pt(9.5)
        if sub_label:
            if bold_label:
                p.add_run("\n")
            r = p.add_run(sub_label)
            r.font.size = Pt(8.5)
        if value:
            if bold_label or sub_label:
                p.add_run("\n")
            r = p.add_run(value)
            r.font.size = Pt(10)

    # R0
    r0_c0 = table_apo.cell(0, 0)
    r0_c1 = table_apo.cell(0, 1)
    r0_c0.merge(r0_c1)
    format_cell(r0_c0, "APOSTILLE", "(Convention de La Haye du 5 octobre 1961)", align=WD_ALIGN_PARAGRAPH.CENTER)
    r0_c0.paragraphs[0].runs[0].font.size = Pt(12)
    
    # R1
    format_cell(table_apo.cell(1, 0), "1. Country:", "Pays / Pais:")
    format_cell(table_apo.cell(1, 1), "", "", "United Kingdom of Great Britain and Northern Ireland")
    
    # R2
    r2_c0 = table_apo.cell(2, 0)
    r2_c1 = table_apo.cell(2, 1)
    r2_c0.merge(r2_c1)
    format_cell(r2_c0, "This public document", "Le présent acte public / El presente documento público", align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # R3
    format_cell(table_apo.cell(3, 0), "2. Has been signed by", "a été signé par\nha sido firmado por")
    format_cell(table_apo.cell(3, 1), "", "", "Marta Maldonado Prados")
    
    # R4
    format_cell(table_apo.cell(4, 0), "3. Acting in the capacity of", "agissant en qualité de\nquien actúa en calidad de")
    format_cell(table_apo.cell(4, 1), "", "", "Notary Public")
    
    # R5
    format_cell(table_apo.cell(5, 0), "4. Bears the seal / stamp of", "est revêtu du sceau / timbre de\ny está revestido del sello / timbre de")
    format_cell(table_apo.cell(5, 1), "", "", "The Said Notary Public")
    
    # R6
    r6_c0 = table_apo.cell(6, 0)
    r6_c1 = table_apo.cell(6, 1)
    r6_c0.merge(r6_c1)
    format_cell(r6_c0, "Certified", "Attesté / Certificado", align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # R7
    format_cell(table_apo.cell(7, 0), "5. at  London", "á / en")
    format_cell(table_apo.cell(7, 1), "6. the  17 September 2026", "le / el día")
    
    # R8
    format_cell(table_apo.cell(8, 0), "7. by", "par / por")
    format_cell(table_apo.cell(8, 1), "", "", "His Majesty's Principal Secretary of State for Foreign, Commonwealth and Development Affairs")
    
    # R9
    format_cell(table_apo.cell(9, 0), "8. Number", "sous no / bajo el numero")
    format_cell(table_apo.cell(9, 1), "", "", "APO-0C4M-NL6S-CMSG-ER3P")
    
    # R10
    c_seal_apo = table_apo.cell(10, 0)
    p = c_seal_apo.paragraphs[0]
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("9. Seal / stamp\n")
    r.bold = True
    r.font.size = Pt(9.5)
    r = p.add_run("Sceau / timbre\nSello / timbre\n\n")
    r.font.size = Pt(8.5)
    r_st = p.add_run("[stamp: FOREIGN, COMMONWEALTH & DEVELOPMENT OFFICE LONDON]")
    r_st.bold = True
    r_st.font.size = Pt(8.5)
    
    c_sig_apo = table_apo.cell(10, 1)
    p = c_sig_apo.paragraphs[0]
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("10. Signature  ")
    r.bold = True
    r.font.size = Pt(9.5)
    r_ingram = p.add_run("L. Ingram\n")
    r_ingram.font.size = Pt(9.5)
    r = p.add_run("Signature\nFirma\n\n")
    r.font.size = Pt(8.5)
    r_sig = p.add_run("[signature]")
    r_sig.bold = True
    r_sig.font.size = Pt(9.5)
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.05
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run("This Apostille is not to be used in the UK and only confirms the authenticity of the signature, seal or stamp on the attached UK public document. It does not confirm the authenticity of the underlying document. Apostilles attached to documents that have been photocopied and certified in the UK confirm the signature of the UK official who conducted the certification only. It does not authenticate either the signature on the original document or the contents of the original document in any way.")
    r.font.size = Pt(8)
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.05
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("If this document is to be used in a country not party to the Hague Convention of the 5th of October 1961, it should be presented to the consular section of the mission representing that country.")
    r.font.size = Pt(8)
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("To verify this apostille go to www.verifyapostille.service.gov.uk")
    r.bold = True
    r.font.size = Pt(8.5)
    
    # -------------------------------------------------------------
    # SECTION 2: Page 3 (Judgment Cover / Title Page)
    # -------------------------------------------------------------
    s2 = doc.add_section(docx.enum.section.WD_SECTION.NEW_PAGE)
    s2.header.is_linked_to_previous = False
    s2.top_margin = Inches(0.8)
    s2.bottom_margin = Inches(0.8)
    s2.left_margin = Inches(1.0)
    s2.right_margin = Inches(1.0)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run("[emblem: ROYAL COURTS OF JUSTICE]")
    r.bold = True
    r.font.size = Pt(10)
    
    table_cit = doc.add_table(rows=1, cols=2)
    table_cit.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_cit.autofit = False
    table_cit.columns[0].width = Inches(3.8)
    table_cit.columns[1].width = Inches(2.4)
    
    p = table_cit.cell(0, 0).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run("Neutral Citation Number: [2026] EWHC 1634 (Comm)")
    r.underline = True
    r.font.size = Pt(9.5)
    
    p = table_cit.cell(0, 1).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("Claim No: CL-2026-000221")
    r.underline = True
    r.font.size = Pt(9.5)
    
    table_court = doc.add_table(rows=1, cols=2)
    table_court.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_court.autofit = False
    table_court.columns[0].width = Inches(4.0)
    table_court.columns[1].width = Inches(2.2)
    
    p = table_court.cell(0, 0).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run("IN THE HIGH COURT OF JUSTICE\nBUSINESS AND PROPERTY COURTS OF ENGLAND AND WALES\nCOMMERCIAL COURT (KBD)")
    r.bold = True
    r.underline = True
    r.font.size = Pt(9.5)
    
    p = table_court.cell(0, 1).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run("The Rolls Building\n7 Rolls Buildings\nFetter Lane\nLondon EC4A 1NL\n\nDate: Friday, 26th June 2026")
    r.underline = True
    r.font.size = Pt(9.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("Before:")
    r.bold = True
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("PAUL STANLEY KC\n(Sitting as a Judge of the High Court)")
    r.bold = True
    r.underline = True
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("(Remotely via MS Teams)")
    r.bold = True
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    p.add_run("- - - - - - - - - - - - - - - - - - - - -")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("Between:")
    r.bold = True
    
    table_parties = doc.add_table(rows=1, cols=2)
    table_parties.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_parties.autofit = False
    table_parties.columns[0].width = Inches(4.5)
    table_parties.columns[1].width = Inches(1.7)
    
    p = table_parties.cell(0, 0).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.2
    r = p.add_run("INVEST INTERNATIONAL CAPITAL B.V.\n- and -\nLEWA TRADING INDUSTRY AND CONTRACTING CO. LTD")
    r.bold = True
    
    p = table_parties.cell(0, 1).paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.line_spacing = 1.2
    r1 = p.add_run("Claimant\n\n\n")
    r1.bold = True
    r1.underline = True
    r2 = p.add_run("Defendant")
    r2.bold = True
    r2.underline = True
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    p.add_run("- - - - - - - - - - - - - - - - - - - - -")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    p.add_run("- - - - - - - - - - - - - - - - - - - - -")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("MR. SIMON OAKES ")
    r.bold = True
    r_mid = p.add_run("(instructed by ")
    r_firm = p.add_run("Norton Rose Fulbright LLP")
    r_firm.bold = True
    r_end = p.add_run(" ) for the ")
    r_cl = p.add_run("Claimant")
    r_cl.bold = True
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("THE DEFENDANTS ")
    r.bold = True
    p.add_run("did not appear and were not represented")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    p.add_run("- - - - - - - - - - - - - - - - - - - - -")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run("Approved Judgment")
    r.bold = True
    r.font.size = Pt(14)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run("If this Transcript is to be reported or published, there is a requirement to ensure that no reporting restriction will be breached. This is particularly important in relation to any case involving a sexual offence, where the victim is guaranteed lifetime anonymity (Sexual Offences (Amendment) Act 1992), or where an order has been made in relation to a young person.")
    r.italic = True
    r.font.size = Pt(8.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("This Transcript is Crown Copyright. It may not be reproduced in whole or in part other than in accordance with relevant licence or with the express consent of the Authority. All rights are reserved.")
    r.italic = True
    r.font.size = Pt(8.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run("Transcript of the Stenograph Notes of Marten Walsh Cherer Ltd.,\n2nd Floor, Quality House, 6-9 Quality Court, Chancery Lane, London WC2A 1HP.\nTelephone No: 020 7067 2900. DX 410 LDE\nEmail: info@martenwalshcherer.com\nWeb: www.martenwalshcherer.com")
    r.font.size = Pt(8.5)
    
    # -------------------------------------------------------------
    # SECTION 3: Pages 4 to 13 (Approved Judgment Text)
    # -------------------------------------------------------------
    s3 = doc.add_section(docx.enum.section.WD_SECTION.NEW_PAGE)
    s3.header.is_linked_to_previous = False
    s3.top_margin = Inches(1.1)
    s3.bottom_margin = Inches(0.8)
    s3.left_margin = Inches(1.0)
    s3.right_margin = Inches(1.0)
    s3.header_distance = Inches(0.5)
    
    # Running Header
    hdr = s3.header
    table_hdr = hdr.add_table(rows=1, cols=2, width=Inches(6.27))
    table_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_hdr.autofit = False
    table_hdr.columns[0].width = Inches(3.8)
    table_hdr.columns[1].width = Inches(2.4)
    
    p_hl = table_hdr.cell(0, 0).paragraphs[0]
    p_hl.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_hl.paragraph_format.line_spacing = 1.0
    p_hl.paragraph_format.space_after = Pt(0)
    r1 = p_hl.add_run("Paul Stanley KC (Sitting as a Judge of the High Court)\n")
    r1.underline = True
    r1.font.size = Pt(9)
    r2 = p_hl.add_run("Approved Judgment")
    r2.underline = True
    r2.font.size = Pt(9)
    
    p_hr = table_hdr.cell(0, 1).paragraphs[0]
    p_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_hr.paragraph_format.line_spacing = 1.0
    p_hr.paragraph_format.space_after = Pt(0)
    r1 = p_hr.add_run("Invest Trading v Lew Trading\n")
    r1.font.size = Pt(9)
    r2 = p_hr.add_run("26.06.26")
    r2.font.size = Pt(9)
    
    def add_sec_heading(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(title)
        r.underline = True
        return p

    def add_num_para(num_str, runs, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.4)
        p.paragraph_format.first_line_indent = Inches(-0.4)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        
        r_num = p.add_run(f"{num_str}\t")
        for text, italic, bold, underline in runs:
            r = p.add_run(text)
            r.italic = italic
            r.bold = bold
            r.underline = underline
        return p

    def add_quote_p(runs, space_after=4, left_indent=0.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(left_indent)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        p.alignment = align
        for text, italic, bold, underline in runs:
            r = p.add_run(text)
            r.italic = italic
            r.bold = bold
            r.underline = underline
        return p

    # --- PAGE 4 ---
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("PAUL STANLEY KC :")
    r.bold = True
    
    add_num_para("1.", [
        ("The matter listed before me this morning is a summary judgment application. The claimant (which I shall call \"Invest\") seeks summary judgment under a guarantee given by the defendant (which I shall call \"Lewa Trading\"). Lewa Trading has not attended the application. It is not represented. It has not said anything to indicate its position formally or informally. I must decide whether to proceed with the hearing in the absence of the defendant. I am satisfied, as I told Mr. Oakes at the outset of the hearing, that I should do so.", False, False, False)
    ])
    
    add_num_para("2.", [
        ("I take into account in reaching that decision the factors identified in Butcher J in ", False, False, False),
        ("European Union v Syria", True, False, False),
        (" [2023] EWHC 1116, and by Lionel Persey KC in ", False, False, False),
        ("African Export-Import Bank v National Government of the Republic of South Sudan", True, False, False),
        (" [2025] EWHC 1079 (Comm). I am satisfied that the steps taken to bring this application to Lewa Trading's attention are likely to have done so. The application and the information about its listing have been notified to it both through the agent that it appointed to receive court process and through various e-mail addresses, and by post to its offices in Saudi Arabia, in sufficient time to make it likely that it will have been aware of the hearing. It was notified of this application on 27th May 2026 and of the date for the hearing on 2nd June 2026. Its lack of any response seems more likely to reflect a deliberate decision not to engage with proceedings than ignorance.", False, False, False)
    ])
    
    add_num_para("3.", [
        ("There is a strong public interest in defendants being notified of hearings in good time so they can participate in them, but there is equally a strong public interest in the orderly determination of claimants' rights. The purpose of notifying defendants of hearings is to enable them to participate if they chose to do so. They are not compelled to do so, but their voluntary decision not to cannot stop the wheels of justice in their tracks; and is no less important where the amounts at stake are, as they are here, substantial. I am satisfied, therefore, there are compelling reasons for the hearing to proceed and no injustice to Lewa Trading if it does. Adjournment would serve no useful purpose.", False, False, False)
    ])
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run("(For continuation of proceedings: please see separate transcript)")
    r.bold = True
    r.italic = True
    
    add_num_para("4.", [
        ("The application before me is for summary judgment. At the opening of the hearing today, I explained my reasons for proceeding with the hearing in the absence of the defendant, which I am satisfied has notice of it.", False, False, False)
    ])
    
    add_num_para("5.", [
        ("The case arises out of a Facility Agreement dated 2nd December 2022. The parties to that agreement were Naylor Nutrition UK Limited as borrower, Naylor Nutrition Group Limited as Guarantor, and the claimant (which I call \"Invest\"), which is a Dutch company, as lender.", False, False, False)
    ])
    
    add_num_para("6.", [
        ("In the evidence before me Naylor Nutrition Group Limited has been called \"Naylor Guarantor\" and I shall follow that practice too.", False, False, False)
    ])
    
    add_num_para("7.", [
        ("The agreement was a term Facility Agreement to provide export finance credit in the sum of €34.6 million.", False, False, False)
    ])
    
    # --- PAGE 5 ---
    doc.add_page_break()
    
    add_sec_heading("The facility")
    
    add_num_para("8.", [
        ("For present purposes the key provisions of the Facility Agreement are Clauses 23.1 and 23.2. Clause 23.1 provided that it was an Event of Default under the agreement if:", False, False, False)
    ])
    
    add_quote_p([
        ("\"An Obligor does not pay on the due date any amount payable pursuant to a Finance Document at the place and in the currency in which it is expressed to be payable unless its failure to pay is caused by administrative or technical error and payment is made within three Business Days of its due date.\"", False, False, False)
    ], space_after=6)
    
    add_num_para("9.", [
        ("Clause 23.22 then provided a right to accelerate if an Event of Default occurred:", False, False, False)
    ])
    
    add_quote_p([
        ("\"On and at any time after the occurrence of an Event of Default the Lender may by notice to the Borrower:", False, False, False)
    ], space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_quote_p([
        ("...", False, False, False)
    ], space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_quote_p([
        ("(b) declare that all or part of the Loans, together with accrued interest, and all other amounts accrued or outstanding under the Finance Documents be immediately due and payable, whereupon they shall become immediately due and payable.\"", False, False, False)
    ], space_after=6)
    
    add_sec_heading("The Guarantee")
    
    add_num_para("10.", [
        ("The Guarantee which is the subject of the claim is given by Lewa Trading by deed dated 5th December 2022. Facility Agreement means a defined Facility Agreement as follows:", False, False, False)
    ])
    
    add_quote_p([
        ("\"Facility Agreement means the facility agreement dated 2 December 2022 between the Borrower, Naylor Nutrition Group Limited and the Lender, as it may from time to time be amended, restated, novated or replaced (however fundamentally, including by an increase of any size in the amount of the facilities made available under it, the alteration of the nature, purpose or period of those facilities or the change of its parties).\"", False, False, False)
    ], space_after=6)
    
    add_num_para("11.", [
        ("The material operative Clauses of the Guarantee were expressed in Clause 2 as follows:", False, False, False)
    ])
    
    add_quote_p([
        ("\"2.1 The Lewa Guarantor irrevocably and unconditionally:", False, False, False)
    ], space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_quote_p([
        ("(a) guarantees to the Lender the punctual payment and discharge of all Obligations from time to time incurred by the Borrower under or in connection with the Finance Documents;", False, False, False)
    ], space_after=4)
    add_quote_p([
        ("(b) undertakes with the Lender that, whenever the Borrower does not pay or discharge any of those Obligations when they become due for payment or discharge, it will immediately on demand do so itself, as if it were the principal obligor; and", False, False, False)
    ], space_after=4)
    
    # --- PAGE 6 ---
    doc.add_page_break()
    
    add_quote_p([
        ("(c) agrees with the Lender that if, for any reason, any amount claimed by the Lender under this clause 2 is not recoverable on the basis of a guarantee, it will be liable as a principal debtor and primary obligor to indemnify the Lender against any cost, loss or liability it incurs as a result of the Borrower not paying any amount expressed to be payable by it under any Finance Document on the date when it is expressed to be due; the amount payable by the Lewa Guarantor under this indemnity will not exceed the amount it would have had to pay under this clause 2 if the amount claimed had been recoverable on the basis of a guarantee.\"", False, False, False)
    ], space_after=6)
    
    add_num_para("12.", [
        ("Clause 3.1 provided:", False, False, False)
    ])
    
    add_quote_p([
        ("\"This guarantee is a continuing guarantee and will extend to the ultimate balance of sums payable by the Borrower under the Finance Documents, regardless of any intermediate payment or discharge in whole or in part.\"", False, False, False)
    ], space_after=6)
    
    add_num_para("13.", [
        ("Clause 8.1 provided:", False, False, False)
    ])
    
    add_quote_p([
        ("\"The Lewa Guarantor will, on demand, pay all legal and other costs and expenses incurred by the Lender in connection with this Deed. This includes any costs and expenses relating to the enforcement of this Deed.\"", False, False, False)
    ], space_after=6)
    
    add_num_para("14.", [
        ("Clause 8.2 provided:", False, False, False)
    ])
    
    add_quote_p([
        ("\"Neither the Lender nor any Officer of the Lender will be in any way liable or responsible to the Lewa Guarantor for any loss or liability of any kind arising from any act or omission by it of any kind in relation to this Deed, except to the extent caused by its own negligence or wilful misconduct.\"", False, False, False)
    ], space_after=6)
    
    add_num_para("15.", [
        ("Clause 9.1 provided:", False, False, False)
    ])
    
    add_quote_p([
        ("\"All payments by the Lewa Guarantor under this Deed will be made in full, without any set-off or other deduction \"", False, False, False)
    ], space_after=6)
    
    add_num_para("16.", [
        ("Clause 9.4 provided:", False, False, False)
    ])
    
    add_quote_p([
        ("\"If the Lewa Guarantor fails to make a payment to a person under this Deed, it will pay interest to that person on the amount concerned at the Default Rate from the date it should have made the payment until the date of payment (after, as well as before, judgment).\"", False, False, False)
    ], space_after=6)
    
    add_num_para("17.", [
        ("Clause 11.4 provided:", False, False, False)
    ])
    
    add_quote_p([
        ("\"Any notice to the Lewa Guarantor may alternatively be sent to its registered office or to any of its places of business or to any of its directors or its company secretary; and it will be deemed", False, False, False)
    ], space_after=4)
    
    # --- PAGE 7 ---
    doc.add_page_break()
    
    add_quote_p([
        ("to have been received when delivered to any such places or persons.\"", False, False, False)
    ], space_after=6)
    
    add_num_para("18.", [
        ("Clause 12.1 provided that the Guarantee was governed by English law.", False, False, False)
    ])
    
    add_num_para("19.", [
        ("By Clause 12.2, the parties agreed to English jurisdiction which was, as against Lewa Trading, exclusive. Among the provisions of that clause were the following:", False, False, False)
    ])
    
    add_quote_p([
        ("\" ...", False, False, False)
    ], space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_quote_p([
        ("(ii) The parties agree that the courts of England are the most appropriate and convenient courts to settle Disputes and, accordingly, that they will not argue to the contrary.", False, False, False)
    ], space_after=4)
    add_quote_p([
        ("...", False, False, False)
    ], space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_quote_p([
        ("(iv) The Lewa Guarantor irrevocably appoints Naylor Nutrition Group Limited (attention: Simon Naylor) a company incorporated under the laws of England and Wales ... at its registered office from time to time to receive on its behalf process issued out of the English courts in connection with this Deed.\"", False, False, False)
    ], space_after=4)
    add_quote_p([
        ("(v) Failure by the process agent to notify the Lewa Guarantor of the process will not invalidate the proceedings concerned.", False, False, False)
    ], space_after=4)
    add_quote_p([
        ("(vi) If this appointment is terminated for any reason, the Lewa Guarantor will appoint a replacement agent and will ensure that the new agent notifies the Lender of its acceptance of appointment.\"", False, False, False)
    ], space_after=6)
    
    add_sec_heading("The commercial position")
    
    add_num_para("20.", [
        ("I turn to the underlying commercial position based on the evidence of Ms. Charlotte Winter. Ms. Winter served two witness statements in support of the application. The first deals with the substance and the second was a statement whose purpose was largely to update the court on the procedural position, service of documents and notice about the hearing.", False, False, False)
    ])
    
    add_num_para("21.", [
        ("I have read both statements and the very helpful and thorough skeleton prepared by Mr. Oakes. In what follows I rely principally on the first statement.", False, False, False)
    ])
    
    add_num_para("22.", [
        ("The amounts drawn down totalled €34,384,074.50. This was just less than the maximum amount of the facility. Invest applied as part of this application to amend its particulars of claim to make this clear. It does not seek summary judgment for the full amount of the facility. Although it may not be strictly necessary, I grant permission to amend, and proceed on the basis of the amount set out in the evidence.", False, False, False)
    ])
    
    add_num_para("23.", [
        ("As Ms. Winter describes, there were defaults in making payment under the Facility Agreement. Notices of Events of Default were given in June 2025 for", False, False, False)
    ])
    
    # --- PAGE 8 ---
    doc.add_page_break()
    
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("the first time. Those were not cured and further Events of Default occurred and were notified in December 2025. A Notice of Demand was sent in January 2026. Invest accepts that in one minor respect the Notice of Demand was not accurate because a small part of the interest claimed was, although due, not yet payable. No payment whatsoever was made.")
    
    add_num_para("24.", [
        ("On 5th February 2026, Invest served Notice to Accelerate and demanded payment of €36,425,230.12. On the same day it served Notice of Demand on Lewa Trading under the Guarantee. Out of an abundance of caution the notices were each served by a variety of means. I am satisfied, based on Ms. Winter's evidence, that they were validly served.", False, False, False)
    ])
    
    add_num_para("25.", [
        ("There was subsequent correspondence, apparently or possibly without prejudice. I have not been asked to read that. But I accept it must show that Lewa Trading received the notice. No payment has been made.", False, False, False)
    ])
    
    add_sec_heading("These proceedings")
    
    add_num_para("26.", [
        ("A Claim Form was issued on 9th April 2026. It was preceded by a letter before action to which no response was received. The Claim Form was served by First Class Post on Naylor Guarantor at its registered address. Copies for information were provided to various e-mail addresses. Acknowledgment of service was due by 28th April 2026 and a defence by no later than 12th May 2026. Service was not acknowledged and no defence has been served or filed.", False, False, False)
    ])
    
    add_num_para("27.", [
        ("This application was then issued. It and all documents relating to it have been served on Naylor Guarantor at its registered address and also provided by various e-mails and by post to Lewa Trading in Saudi Arabia.", False, False, False)
    ])
    
    add_sec_heading("Procedural Points")
    
    add_num_para("28.", [
        ("It is convenient to deal first with certain procedural points.", False, False, False)
    ])
    
    add_num_para("29.", [
        ("First, it is clear that the court has jurisdiction. Quite apart from the fact that the claim has permissibly been served in the jurisdiction on an agent, as I shall explain below, the Guarantee contains an English Jurisdiction Clause. Mr. Oakes, who appears for Invest, very properly drew my attention to the possibility that a ", False, False, False),
        ("forum non conveniens", True, False, False),
        (" argument might technically be open to Lewa Trading. But such an argument would be completely doomed to fail, given the exclusive jurisdiction agreement, the express renunciation in the guarantee of any such argument, the English governing law, the close connection between the Guarantee to the Facility Agreement. For all these reasons it is clear beyond any shadow of argument that England is not merely an appropriate jurisdiction but the only appropriate jurisdiction for this claim.", False, False, False)
    ])
    
    add_num_para("30.", [
        ("Secondly, I am satisfied that the Claim Form was validly served. The Guarantee appointed Naylor Guarantor as agent for service of process. There is no indication that authority has ever purportedly been revoked. Service required under the Guarantor was service at Naylor Guarantor's registered office from time to time. That happened. Service accordingly took place in accordance with CPR 6.11.", False, False, False)
    ])
    
    # --- PAGE 9 ---
    doc.add_page_break()
    
    add_num_para("31.", [
        ("Third, however, I consider that despite the contractual permission to serve other procedural documents in the same way, the court's permission was also required. That is because CPR 6.11 applies to a Claim Form: it does not apply to other documents such as Application Notices. Other documents must ", False, False, False),
        ("prima facie", True, False, False),
        (" be served either on the party or at an address for service given the acknowledgement for service, or after an order for service by alternative means has been made.", False, False, False)
    ])
    
    add_num_para("32.", [
        ("However, in ", False, False, False),
        ("DVB Bank SE v Vega Marine Ltd", True, False, False),
        (" [2020] EWHC 1494 (Comm) at paragraphs 48-52, Henshaw J held that the court can authorise service by alternative means (which is plainly right) and that it can do so retrospectively. The question is simply whether there is a good reason for that to happen, there being no question in this case of an \"exceptional\" reason being required because Saudi Arabia is not party to the Hague Service Convention.", False, False, False)
    ])
    
    add_num_para("33.", [
        ("In my view, where retrospective validation is sought, the good reason needs also to be good enough to justify retrospective validation rather than prospective grant. I am sure that there is a good reason here. The method of service adopted was contractually agreed. Invest did not merely follow it to the letter, but it followed it to the spirit too, and went beyond it by providing copies of relevant documents to other addresses which were likely to ensure that they came to the attention of Lewa Trading. It would be disproportionate and serve no useful purpose whatever to insist on doing again all that has been done.", False, False, False)
    ])
    
    add_num_para("34.", [
        ("I should perhaps make it clear, however, for future reference, that though I have authorised such service retrospectively, the better course would always be for an application to be made, ", False, False, False),
        ("ex parte", True, False, False),
        (" and in writing (I would expect) in advance, in any case where alternative service is required. If nothing else, that obviates any risk that the hearing of the relevant application may need to be adjourned.", False, False, False)
    ])
    
    add_num_para("35.", [
        ("Summary judgment before Acknowledgment of Service", False, False, True)
    ], space_after=8)
    
    add_num_para("36.", [
        ("CPR Part 24 provides that in most cases an application for summary judgment cannot be made before the defendant has acknowledged service. As Bryan J explained, however, in ", False, False, False),
        ("European Union v Syria", True, False, False),
        (" [2018] EWHC 1712 (Comm) there are circumstances in which it is possible and appropriate for the court to permit such an application. The main purpose of the general rule is to ensure the defendant has had an opportunity to participate in the proceedings and to make any jurisdiction objections. Where it is clear that the defendant has had that opportunity and is choosing not to take it, the court is satisfied that it has jurisdiction and that the defendant has notice of the proceedings, and the application and there is a good reason for summary judgment rather than default judgment, the court may permit such an application to be made. It may grant that permission retrospectively at the time it hears that application.", False, False, False)
    ])
    
    add_num_para("37.", [
        ("A common reason for seeking summary judgment rather than default judgment is that some jurisdictions are more likely to enforce a judgment if they know that the English court has considered the merits. ", False, False, False),
        ("European Union v Syria", True, False, False),
        (" recognises that as a legitimate potential reason; and it has been, in my experience, frequently applied in many other cases.", False, False, False)
    ])
    
    # --- PAGE 10 ---
    doc.add_page_break()
    
    add_num_para("38.", [
        ("In ", False, False, False),
        ("DVB", True, False, False),
        (", at 58, Henshaw J said this:", False, False, False)
    ])
    
    add_quote_p([
        ("\"Bryan J summarised the principles relevant to the exercise of the court's discretion under CPR 24.4(1) in ", False, False, False),
        ("European Union v Syria", True, False, False),
        (":", False, False, False)
    ], space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_quote_p([
        ("'(1) The purpose of the rule are to ensure that no application for summary judgment is made before a defendant has had an opportunity to participate in the proceedings and to protect a defendant who wishes to challenge the Court's jurisdiction from having to engage on the merits pending such application.", False, False, False)
    ], space_after=4)
    add_quote_p([
        ("'(2) Generally, permission should be granted only where the Court is satisfied that the claim has been validly served and that the Court has jurisdiction to hear it. Once those conditions are met there is generally no reason why the Court should prevent a claimant with a legitimate claim from seeking summary judgment.", False, False, False)
    ], space_after=4)
    add_quote_p([
        ("'(3) The fact that a summary judgment may be more readily enforced in other jurisdictions than a default judgment is a proper reason for seeking permission under CPR 24.4(1).' (§ 61)", False, False, False)
    ], space_after=4)
    add_quote_p([
        ("I would add, in relation to (3), that it would in my view be sufficient that the claimant has a reasonable belief that a summary judgment may be more readily enforced than a default judgment. There is no justification for the court subjecting any such belief to minute examination, when the permission the claimant is seeking is in reality no more than the opportunity to obtain a reasoned judgment on the merits of its claim.\"", False, False, False)
    ], space_after=6)
    
    add_num_para("39.", [
        ("For reasons that I have already given, I am satisfied on points (1) and (2). The evidence in this case in relation to point (3), however, is somewhat thin. All Ms. Winter's evidence said was as follows:", False, False, False)
    ])
    
    add_quote_p([
        ("\"However, since the Lewa Guarantor has its registered office and (to the best of my knowledge) its assets in the Kingdom of Saudi Arabia, Invest seeks a merits-based summary judgment since it believes that this will better facilitate the enforcement of judgment in Saudi Arabia. The Court will be well aware of the difficulties of enforcing default judgments overseas.\"", False, False, False)
    ], space_after=6)
    
    add_num_para("40.", [
        ("I regard this as borderline. Although a minute examination is not required, more than a merely subjective view is needed and it is, not, in my view, acceptable simply to refer in general terms to possible difficulties of enforcement abroad. Enough should be said to make it clear that the view is a reasonable one although it should not be subjected to minute examination and one should bear in mind the fact that, as Bryan J said, once the conditions (1) and (2) have been satisfied:", False, False, False)
    ])
    
    # --- PAGE 11 ---
    doc.add_page_break()
    
    add_quote_p([
        ("\"... there is generally no reason why the Court should prevent a claimant with a legitimate claim from seeking summary judgment.\"", False, False, False)
    ], space_after=6)
    
    add_num_para("41.", [
        ("In this case, however, Mr. Oakes told me on instructions that the view expressed was not merely a subjective one on Invest's part, but was informed by legal advice about the position in Saudi Arabia. The costs schedule, as it happens, confirms the involvement of Saudi Arabian lawyers. Although it would certainly have been better if that had been made clear, and better still if the reasons for that advice had been given, I am prepared to accept that what Ms. Winter says is sufficient to meet the relatively low bar that Henshaw J described.", False, False, False)
    ])
    
    add_sec_heading("The substance")
    
    add_num_para("42.", [
        ("I turn then to the substance. I must be satisfied before I grant summary judgment that there does not appear to be any defence which has a real prospect of success and that there is no other compelling reason for trial. In short, I am so satisfied, for the reasons given in Ms. Winter's evidence.", False, False, False)
    ])
    
    add_num_para("43.", [
        ("In some ways, of course, the absence of any express defence is a reason to tread carefully. It would be wrong to treat the failure to serve a defence as very positively indicative of the absence of any realistically arguable defence, for that would be to turn summary judgment into default judgment. However, a claim such as this is a common sort of claim, where the likely defences can be intelligently assessed. There are no grounds in this case to doubt the validity of the core documents. There are no grounds to doubt that the loan was advanced, that defaults occurred and that it was properly accelerated. In those circumstances, there are no grounds to doubt that Lewa Trading is liable under the Guarantee. It has not merely failed to file the defence, but it has not put forward any contrary contention in open correspondence. Ms. Winter's evidence thoughtfully and fully sets out the position and this is not merely a matter of assertion, but of carefully explained reasons which satisfy me that there appears to be no real prospect of a defence succeeding. It appears on the evidence to be an open and shut case.", False, False, False)
    ])
    
    add_num_para("44.", [
        ("In reaching that conclusion, I have specifically considered three points. The first is the admitted error in the calculation of the demand, sent in January 2026, to which Ms. Winter refers. I am satisfied that this can give no real prospect of a defence. The error was minuscule. There was no doubt there were very substantial sums due and there was therefore no doubt that there was, when the acceleration notice was given, an Event of Default which validated that notice. At that point, all sums became immediately payable.", False, False, False)
    ])
    
    add_num_para("45.", [
        ("The second matter I considered is the possibility, based as I understand it on things that might have been said in without prejudice communications, that Lewa Trading would argue that Invest was at fault in guiding it about the underlying transaction. In my view, even imagining that case at its highest, it would be a hopeless argument. Quite apart from the extreme improbability of Lewa Trading establishing any duty on the part of Invest to do anything of that sort—a prospect I would regard as fanciful, not real—the most that this would", False, False, False)
    ])
    
    # --- PAGE 12 ---
    doc.add_page_break()
    
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("give rise to would be a cross-claim which would not constitute a defence because the Guarantee includes a provision precluding set off.")
    
    add_num_para("46.", [
        ("The third matter that I have considered, at Mr. Oakes' suggestion, is the possibility that Lewa Trading might argue that the acceleration provisions or the default interest rate (which is EURIBOR, plus the loan margin of 1.95%, plus 2%) are penal. They very plainly are not, and there is no reasonable prospect that Lewa Trading would establish that they are. Acceleration Clauses of this sort, which provide for an acceleration of payment of sums accrued due are not penal. They are simply an ordinary and necessary part of the bargain specifying the relevant date for repayment: see in this respect, ", False, False, False),
        ("Chitty On Contracts", True, False, False),
        (" at paragraph 30-253. I have no doubt that what the editors there say is correct.", False, False, False)
    ])
    
    add_num_para("47.", [
        ("As to default interest, there are sound commercial reasons based on the increased risk attendant on the loan which is not being kept current which justify default interest rates. There may of course come a point at which a rate is so high that the proper conclusion is that it has become a penalty. But the rate in question here, a margin of 2% above that payable if the loan were current, is modest and well within the range that could not conceivably be regarded as engaging a rule against penalties.", False, False, False)
    ])
    
    add_num_para("48.", [
        ("Given the legitimate interest served by the default interest rate, a relatively modest rate in this case, falling far short of anything that could be regarded as extravagant, there is nothing remotely resembling a real prospect that it would be regarded as penal. The purpose of a default interest here is plainly not penal but to reflect the legitimate commercial interests of a lender in circumstances where the risk has changed: see, in that respect, ", False, False, False),
        ("Houssein & Ors v London Credit", True, False, False),
        (" [2024] EWCA Civ 721, paragraphs 49 and 50.", False, False, False)
    ])
    
    add_sec_heading("Conclusion")
    
    add_num_para("49.", [
        ("Accordingly, I will grant summary judgment in the terms set out in the draft order which is supported by the evidence and argument I have seen. Two questions then arise. As to post-judgment interest, the question is whether I should award interest at the contractually-agreed rate. It is clear I have power under the Administration of Justice Act 1970 section 44A to do so, in circumstances where the judgment is not in sterling. I am clear that it is appropriate that I should exercise that power given the compensatory purpose of interest: see ", False, False, False),
        ("Novoship (UK) Ltd v Mikhaylyuk", True, False, False),
        (" [2014] EWCA Civ 908 at 136.", False, False, False)
    ])
    
    add_num_para("50.", [
        ("Moreover, like Blair J in ", False, False, False),
        ("Law Debenture Trust Corp v Ukraine", True, False, False),
        (" [2017] EWHC 1902 (Comm) at 13, I consider that where the parties have specified, as they have here, a rate of interest which is to apply post-judgment, then unless that right is penal rather than compensatory, there are powerful reasons to uphold their bargain.", False, False, False)
    ])
    
    add_num_para("51.", [
        ("This rate, as I have made clear, is not penal. It is compensatory. It is also lower than the judgment rate. I shall therefore apply it. I shall, however, crystallise that rate at 6.49% in order to prevent any doubt about the precise sum due at any time. I can be confident that that leaves Lewa Trading in a better position than it would have been if the Judgment Act rate had applied.", False, False, False)
    ])
    
    # --- PAGE 13 ---
    doc.add_page_break()
    
    add_num_para("52.", [
        ("As to costs, in my judgment the terms of the guarantee entail that costs be awarded on an indemnity basis. That is the effect of the term that all legal costs should be paid. I do not accept that the reference to all legal costs, however, means that Invest is entitled to recover unreasonable or unnecessary costs. It means that the limits are broad, but it does not introduce limitless liability without regard to reasonableness. I would not be able to accept—certainly not for the purposes of giving summary judgment—that unreasonable costs are recoverable.", False, False, False)
    ])
    
    add_num_para("53.", [
        ("I have considered the costs set out in the costs schedule which I have been provided with. In terms of the time that has been spent, they appear to me to be reasonable. That is not, however, the position as far as the amount is concerned. I accept that this is heavy litigation in the sense that there is a large sum at stake, but it is not complex litigation and, indeed, it is Invest's submission, which I have accepted, that it is legally straightforward. The procedural complexity is not profound either and there is nothing that begins to take this outside the ordinary range of heavy commercial litigation which Guideline Rates foresee.", False, False, False)
    ])
    
    add_num_para("54.", [
        ("I see no justification for departure from the guideline rate much less the near doubling of some of them that the rates claim involved. Far from their being a clear and compelling justification for departing from guideline rates there seems to me to be no good reason at all. That involves no criticism of the decision to instruct Norton Rose, but they are a firm whose rates should in any event, so far as reasonableness is concerned when assessing costs, fall within the London 1 Band. I shall therefore award costs on an indemnity basis, but limited to the costs which result from applying the hours claimed at London 1 rates.", False, False, False)
    ])
    
    add_num_para("55.", [
        ("I will ask the claimant to provide me with a revised figure prepared on that basis which I make it clear may also include the costs of producing that revision.", False, False, False)
    ])
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(14)
    r = p.add_run("(For continuation of proceedings: please see separate transcript)")
    r.bold = True
    r.italic = True
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    p.add_run("- - - - - - - - - - - - -")
    
    if output_path is None:
        target_filename = "Approved Judgment_Notarised+Legalised_SN 193344 (1)-Non-Parsable-en-US#SRECFMT_DXBRBK#.docx"
    else:
        target_filename = str(output_path)
    
    import os
    os.makedirs(os.path.dirname(os.path.abspath(target_filename)), exist_ok=True)
    doc.save(target_filename)
    print(f"Successfully generated and saved: {target_filename}")
    return target_filename

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Generate Approved Judgment document")
    parser.add_argument("-i", "--input", default=None, help="Input PDF path")
    parser.add_argument("-o", "--output", default=None, help="Output DOCX path")
    args = parser.parse_args()
    build_full_approved_judgment(output_path=args.output)
