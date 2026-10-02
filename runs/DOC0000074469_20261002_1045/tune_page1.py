"""
tune_page1.py — Align Page 1 elements to original PDF
"""
from pathlib import Path

mod_path = Path("runs/DOC0000074469_20261002_1045/modifications/journal_article_builder.py")
src = mod_path.read_text(encoding="utf-8")

# Spanish title space_before from 6 to 36
src = src.replace("p_tit.paragraph_format.space_before = Pt(6)", "p_tit.paragraph_format.space_before = Pt(36)")

# Authors space_after from 6 to 24
src = src.replace("p_auth.paragraph_format.space_after = Pt(6)", "p_auth.paragraph_format.space_after = Pt(24)")

# Affiliations space_before: 295.7 + 81 = 376.7 pt
src = src.replace("p_aff.paragraph_format.space_before = Pt(146)", "p_aff.paragraph_format.space_before = Pt(81)")

# Abstract picture width: 4.00 inches to fit within 4.15 in cell with 6pt padding
src = src.replace('r_abs.add_picture(str(assets["p1_eng_abstract"]), width=Inches(4.15))',
                  'r_abs.add_picture(str(assets["p1_eng_abstract"]), width=Inches(3.98))')

mod_path.write_text(src, encoding="utf-8")
print("✅ Applied Page 1 precise alignment to modifications/journal_article_builder.py")
