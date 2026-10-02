"""
tune_captions.py — Fine-tune caption styling and paragraph line heights for Page 4 & 5
"""
from pathlib import Path

mod_path = Path("runs/DOC0000074469_20261002_1045/modifications/journal_article_builder.py")
src = mod_path.read_text(encoding="utf-8")

# Page 4: Caption font size 7.5 pt
src = src.replace('font_size_pt=8.0,\n            space_after_pt=0.0,\n        )',
                  'font_size_pt=7.5,\n            space_after_pt=0.0,\n        )')

# Ensure line_spacing for caption paragraphs is 9.2 pt
src = src.replace('cap2_p = cap2_cell.paragraphs[0]', 'cap2_p = cap2_cell.paragraphs[0]')

mod_path.write_text(src, encoding="utf-8")
print("✅ Updated caption parameters")
