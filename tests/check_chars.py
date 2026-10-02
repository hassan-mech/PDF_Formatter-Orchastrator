import fitz

doc = fitz.open('263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf')

for pno in [2, 6, 8]:
    text = doc[pno].get_text()
    for line in text.split("\n"):
        if any(x in line for x in ['pH4', 'usion', 'eƯ', 'insu', 'flammation', 'conﬁrmed', 'identiﬁed']):
            print(f"P{pno+1}: {repr(line)}")
