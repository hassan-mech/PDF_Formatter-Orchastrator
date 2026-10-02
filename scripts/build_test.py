import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_document():
    doc = docx.Document()
    
    # Configure default style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)
    normal_style.paragraph_format.space_before = Pt(0)
    
    # Page setup for Section 1 (Pages 1 & 2)
    s1 = doc.sections[0]
    s1.page_width = Inches(8.27)
    s1.page_height = Inches(11.69)
    s1.top_margin = Inches(0.8)
    s1.bottom_margin = Inches(0.8)
    s1.left_margin = Inches(1.0)
    s1.right_margin = Inches(1.0)
    
    # ---------------------------------------------------------
    # PAGE 1: Notary Certificate
    # ---------------------------------------------------------
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
    p.paragraph_format.space_after = Pt(36)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.2
    p.paragraph_format.space_after = Pt(14)
    r = p.add_run("TO ALL TO WHOM THESE PRESENTS SHALL COME, I MARTA MALDONADO PRADOS of the City of London NOTARY PUBLIC by royal authority duly admitted and sworn DO HEREBY CERTIFY that the document hereunto annexed is a true printout of the digital transcription of the ")
    r_bold1 = p.add_run("approved judgment")
    r_bold1.bold = True
    r2 = p.add_run(" relating to INVEST INTERNATIONAL CAPITAL B.V. and LEWA TRADING INDUSTRY AND CONTRACTING CO. LTD., in respect of claim number ")
    r_bold2 = p.add_run("CL-2026-000221")
    r_bold2.bold = True
    r3 = p.add_run(", in THE HIGH COURT OF JUSTICE, BUSINESS AND PROPERTY COURTS OF ENGLAND AND WALES, COMMERCIAL COURT (KBD), issued electronically by MARTEN WALSH CHERER LIMITED, a private limited company duly organised and existing under the laws of England and Wales, registered with the Registrar of Companies for England and Wales under number 2669638 and with registered address at 27 Old Gloucester Street, London WC1N 3AX, England.")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.2
    p.paragraph_format.space_after = Pt(28)
    p.add_run("IN FAITH AND TESTIMONY WHEREOF I the said notary have subscribed my name and set and affixed my seal of office at London aforesaid this sixteenth day of September two thousand and twenty six.")
    
    # Table for Seal & Signature
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
    
    # Spacing before footer
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(40)
    p.paragraph_format.space_after = Pt(0)
    
    # Footer table with top border
    table_foot = doc.add_table(rows=1, cols=2)
    table_foot.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_foot.autofit = False
    table_foot.columns[0].width = Inches(4.5)
    table_foot.columns[1].width = Inches(1.8)
    
    # set border on table
    for c in table_foot.rows[0].cells:
        tcPr = c._tc.get_or_add_tcPr()
        tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>')
        tcPr.append(tcBorders)
    
    c_f1 = table_foot.cell(0, 0)
    p = c_f1.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run("Saville & Co. Scrivener Notaries is the trading name of Saville Notaries LLP, a limited liability partnership registered in England and Wales with registered number OC420687 and with registered office at 11 Old Jewry, London EC2R 8DU. Regulated through the Faculty Office of the Archbishop of Canterbury\n*General Notary")
    r.font.size = Pt(7)
    
    c_f2 = table_foot.cell(0, 1)
    p = c_f2.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("[logo: SCRIVENER NOTARIES] | Member Notariat of the International Union of Notaries (UINL)")
    r.font.size = Pt(7)
    
    # ---------------------------------------------------------
    # PAGE 2: Apostille
    # ---------------------------------------------------------
    doc.add_page_break()
    
    # Apostille Table
    table_apo = doc.add_table(rows=11, cols=2)
    table_apo.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_apo.autofit = False
    table_apo.columns[0].width = Inches(2.6)
    table_apo.columns[1].width = Inches(3.6)
    
    # Set borders for Apostille table
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
    
    def format_cell(cell, bold_label, sub_label=None, value=None, align=WD_ALIGN_PARAGRAPH.LEFT, bold_val=False):
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.line_spacing = 1.05
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
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
            if bold_val:
                r.bold = True

    # Row 0: Merged title
    r0_c0 = table_apo.cell(0, 0)
    r0_c1 = table_apo.cell(0, 1)
    r0_c0.merge(r0_c1)
    format_cell(r0_c0, "APOSTILLE", "(Convention de La Haye du 5 octobre 1961)", align=WD_ALIGN_PARAGRAPH.CENTER)
    r0_c0.paragraphs[0].runs[0].font.size = Pt(12)
    
    # Row 1: Country
    format_cell(table_apo.cell(1, 0), "1. Country:", "Pays / Pais:")
    format_cell(table_apo.cell(1, 1), "", "", "United Kingdom of Great Britain and Northern Ireland")
    
    # Row 2: Merged "This public document"
    r2_c0 = table_apo.cell(2, 0)
    r2_c1 = table_apo.cell(2, 1)
    r2_c0.merge(r2_c1)
    format_cell(r2_c0, "This public document", "Le présent acte public / El presente documento público", align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Row 3: Has been signed by
    format_cell(table_apo.cell(3, 0), "2. Has been signed by", "a été signé par\nha sido firmado por")
    format_cell(table_apo.cell(3, 1), "", "", "Marta Maldonado Prados")
    
    # Row 4: Acting in capacity of
    format_cell(table_apo.cell(4, 0), "3. Acting in the capacity of", "agissant en qualité de\nquien actúa en calidad de")
    format_cell(table_apo.cell(4, 1), "", "", "Notary Public")
    
    # Row 5: Bears the seal / stamp of
    format_cell(table_apo.cell(5, 0), "4. Bears the seal / stamp of", "est revêtu du sceau / timbre de\ny está revestido del sello / timbre de")
    format_cell(table_apo.cell(5, 1), "", "", "The Said Notary Public")
    
    # Row 6: Merged "Certified"
    r6_c0 = table_apo.cell(6, 0)
    r6_c1 = table_apo.cell(6, 1)
    r6_c0.merge(r6_c1)
    format_cell(r6_c0, "Certified", "Attesté / Certificado", align=WD_ALIGN_PARAGRAPH.CENTER)
    
    # Row 7: at / the
    format_cell(table_apo.cell(7, 0), "5. at  London", "á / en")
    format_cell(table_apo.cell(7, 1), "6. the  17 September 2026", "le / el día")
    
    # Row 8: by
    format_cell(table_apo.cell(8, 0), "7. by", "par / por")
    format_cell(table_apo.cell(8, 1), "", "", "His Majesty's Principal Secretary of State for Foreign, Commonwealth and Development Affairs")
    
    # Row 9: Number
    format_cell(table_apo.cell(9, 0), "8. Number", "sous no / bajo el numero")
    format_cell(table_apo.cell(9, 1), "", "", "APO-0C4M-NL6S-CMSG-ER3P")
    
    # Row 10: Seal / Signature
    c_seal_apo = table_apo.cell(10, 0)
    p = c_seal_apo.paragraphs[0]
    p.paragraph_format.line_spacing = 1.05
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
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
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
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
    
    # Disclaimers below table
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
    
    # ---------------------------------------------------------
    # SECTION 2: Page 3 (Judgment Cover / Title Page)
    # ---------------------------------------------------------
    s2 = doc.add_section(docx.enum.section.WD_SECTION.NEW_PAGE)
    s2.header.is_linked_to_previous = False
    s2.top_margin = Inches(0.8)
    s2.bottom_margin = Inches(0.8)
    s2.left_margin = Inches(1.0)
    s2.right_margin = Inches(1.0)
    
    # Top emblem
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run("[emblem: ROYAL COURTS OF JUSTICE]")
    r.bold = True
    r.font.size = Pt(10)
    
    # Citation & Claim No table
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
    
    # Court & Address table
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
    
    # Before
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
    
    # Parties table
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
    
    # ---------------------------------------------------------
    # SECTION 3: Pages 4 to 13 (Approved Judgment Text)
    # ---------------------------------------------------------
    s3 = doc.add_section(docx.enum.section.WD_SECTION.NEW_PAGE)
    s3.header.is_linked_to_previous = False
    s3.top_margin = Inches(0.8)
    s3.bottom_margin = Inches(0.8)
    s3.left_margin = Inches(1.0)
    s3.right_margin = Inches(1.0)
    
    # Configure running header for Section 3
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
    
    print("Document structure through Section 3 created successfully")
    return doc

print("create_document defined")
