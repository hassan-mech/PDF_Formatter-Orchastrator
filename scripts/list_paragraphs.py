# Script to inspect and verify text against PDF pages

import docx

doc = docx.Document('Approved Judgment_Notarised+Legalised_SN 193344 (1)-Non-Parsable-en-US#SRECFMT_DXBRBK#.docx')

paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
print(f"Non-empty paragraphs: {len(paragraphs)}")

# Let's see all paragraphs
for i, p in enumerate(paragraphs):
    print(f"[{i}] {p[:80]}...")
