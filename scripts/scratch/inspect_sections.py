import docx

doc = docx.Document('output/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.docx')
print('Number of sections:', len(doc.sections))
for i, s in enumerate(doc.sections):
    sectPr = s._sectPr
    cols = sectPr.xpath('./w:cols')
    col_info = [c.attrib for c in cols] if cols else 'None'
    type_info = sectPr.xpath('./w:type')
    type_val = [t.attrib for t in type_info] if type_info else 'NEW_PAGE (default)'
    footer_p = [p.text for p in s.footer.paragraphs if p.text]
    print(f'Section {i+1}: type={type_val}, cols={col_info}, margins=(L={s.left_margin.inches:.2f}\", R={s.right_margin.inches:.2f}\", T={s.top_margin.inches:.2f}\", B={s.bottom_margin.inches:.2f}\"), footer={footer_p}')
