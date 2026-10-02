import docx

doc = docx.Document('output/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.docx')
print('=== IMAGES IN DOCUMENT ===')
for p_idx, p in enumerate(doc.paragraphs):
    # check for blip / drawing
    drawings = p._p.xpath('.//a:blip')
    if drawings:
        r_text = p.text.strip()
        hidden_runs = [r.text for r in p.runs if r.font.hidden]
        print(f'Para {p_idx}: blips={len(drawings)}, text_len={len(r_text)}, hidden_runs={len(hidden_runs)}')
        if hidden_runs:
            print(f'   Hidden text preview: {hidden_runs[0][:80]}...')
