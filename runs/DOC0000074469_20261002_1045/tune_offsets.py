"""
tune_offsets.py — Exact alignment of Page 3 and Page 6 figures
"""
from pathlib import Path

mod_path = Path("runs/DOC0000074469_20261002_1045/modifications/journal_article_builder.py")
src = mod_path.read_text(encoding="utf-8")

# Page 3: Figure 1 from 112.5 pt to 129.2 pt (+16.7 pt)
src = src.replace("p3_fig_sp.paragraph_format.space_before = Pt(0)", "p3_fig_sp.paragraph_format.space_before = Pt(16.7)")

# Page 6: Figure 4 from 112.5 pt to 129.6 pt (+17.1 pt)
src = src.replace("p6_fig_sp.paragraph_format.space_before = Pt(0)", "p6_fig_sp.paragraph_format.space_before = Pt(17.1)")

mod_path.write_text(src, encoding="utf-8")
print("✅ Applied exact vertical offsets to Page 3 and Page 6")
