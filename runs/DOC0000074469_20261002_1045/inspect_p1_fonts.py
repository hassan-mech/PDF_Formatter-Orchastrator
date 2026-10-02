import pymupdf

doc = pymupdf.open('finished/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf')
p1 = doc[0]
for b in p1.get_text('dict')['blocks']:
    if 'lines' in b:
        for l in b['lines']:
            for s in l['spans']:
                text = s['text'].strip()
                if any(k in text for k in ['Dermatolog', 'Revista', 'Melanoma', 'Metastatic', 'CASO', 'Reviewed', 'Corona', 'ANTECEDENTES']):
                    print(f"{text[:35]:35} | font: {s['font']:25} | size: {s['size']:5.2f} | color: {hex(s['color'])} | bbox: {[round(x, 1) for x in s['bbox']]}")
