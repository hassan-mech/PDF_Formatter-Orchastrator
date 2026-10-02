import docx

doc = docx.Document('output/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.docx')
print('=== HIDDEN RUNS IN DOCUMENT ===')
for p_idx, p in enumerate(doc.paragraphs):
    hidden_runs = [r for r in p.runs if r.font.hidden]
    if hidden_runs:
        print(f'Para {p_idx}: {len(hidden_runs)} hidden run(s), size={hidden_runs[0].font.size.pt}pt, color={hidden_runs[0].font.color.rgb}, text_preview={hidden_runs[0].text[:80]}...')
