import fitz

doc = fitz.open('263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.pdf')
page1 = doc[0]
for b in page1.get_text('dict')['blocks']:
    if 'lines' in b:
        for l in b['lines']:
            for s in l['spans']:
                text = s['text'].strip()
                if any(w in text for w in ['Subject', 'Medical', 'P1-C1D1', 'Please', '1.1', '1.2']):
                    color = s['color']
                    r = (color >> 16) & 255
                    g = (color >> 8) & 255
                    b_val = color & 255
                    flags = s['flags']
                    print(f"{text[:30]}: font={s['font']}, size={s['size']:.1f}, color=({r},{g},{b_val}), flags={flags}")
