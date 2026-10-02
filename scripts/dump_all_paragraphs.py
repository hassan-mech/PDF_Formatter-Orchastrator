import docx

doc = docx.Document('Approved Judgment_Notarised+Legalised_SN 193344 (1)-Non-Parsable-en-US#SRECFMT_DXBRBK#.docx')

for i, p in enumerate(doc.paragraphs):
    print(f"{i}: {repr(p.text)}")
