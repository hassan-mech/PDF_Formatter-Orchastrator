"""
apply_analytical_fixes.py — Applies precise analytical coordinates to journal_article_builder in modifications/
"""
from pathlib import Path

src = Path('scripts/stages/builders/journal_article_builder.py').read_text(encoding='utf-8')

# 1. Add _create_2col_row right after _create_layout_row
c2_func = '''
def _create_2col_row(doc, col1_w: float = 2.85, gutter_w: float = 0.24, col2_w: float = 2.85):
    """Creates a deterministic 2-column table with a gutter column, total width 5.94 inches.
    Aligned LEFT so that Column 1 starts exactly at the left margin.
    Returns (left_cell, right_cell).
    """
    tbl = doc.add_table(rows=1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl.autofit = False
    _set_borderless_table(tbl)
    tbl.rows[0].cells[0].width = Inches(col1_w)
    tbl.rows[0].cells[1].width = Inches(gutter_w)
    tbl.rows[0].cells[2].width = Inches(col2_w)
    _set_cell_padding(tbl.rows[0].cells[0], top_pt=0, bottom_pt=0, left_pt=0, right_pt=0)
    _set_cell_padding(tbl.rows[0].cells[1], top_pt=0, bottom_pt=0, left_pt=0, right_pt=0)
    _set_cell_padding(tbl.rows[0].cells[2], top_pt=0, bottom_pt=0, left_pt=0, right_pt=0)
    return tbl.rows[0].cells[0], tbl.rows[0].cells[2]
'''
idx = src.find('def _add_red_stripe')
assert idx > 0, "Could not find _add_red_stripe in source"
src = src[:idx] + c2_func + '\n\n' + src[idx:]

# 2. Update _create_layout_row alignment to LEFT
src = src.replace('def _create_layout_row(doc, widths_in: List[float]):\n    """Creates a 1-row borderless table where each cell has the given width in inches."""\n    tbl = doc.add_table(rows=1, cols=len(widths_in))\n    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER',
                  'def _create_layout_row(doc, widths_in: List[float]):\n    """Creates a 1-row borderless table where each cell has the given width in inches."""\n    tbl = doc.add_table(rows=1, cols=len(widths_in))\n    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT')

# 3. Page 1 2-column layout: 3 cells (4.15, 0.25, 1.54)
old_p1_tbl = '''        p1_tbl = doc.add_table(rows=1, cols=2)
        p1_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        p1_tbl.autofit = False
        _set_borderless_table(p1_tbl)
        c_left, c_right = p1_tbl.rows[0].cells[0], p1_tbl.rows[0].cells[1]
        c_left.width = Inches(4.15)
        c_right.width = Inches(1.78)'''
new_p1_tbl = '''        p1_tbl = doc.add_table(rows=1, cols=3)
        p1_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        p1_tbl.autofit = False
        _set_borderless_table(p1_tbl)
        c_left = p1_tbl.rows[0].cells[0]
        c_gutter = p1_tbl.rows[0].cells[1]
        c_right = p1_tbl.rows[0].cells[2]
        c_left.width = Inches(4.15)
        c_gutter.width = Inches(0.25)
        c_right.width = Inches(1.54)
        _set_cell_padding(c_gutter, top_pt=0, bottom_pt=0, left_pt=0, right_pt=0)'''
assert old_p1_tbl in src, "Could not find old_p1_tbl"
src = src.replace(old_p1_tbl, new_p1_tbl)

# 4. Page 1 right col space_before for affiliations (starts at y=376.8 pt in orig)
old_p1_aff = '''        p_aff = c_right.paragraphs[0]
        p_aff.paragraph_format.space_before = Pt(0)'''
new_p1_aff = '''        p_aff = c_right.paragraphs[0]
        p_aff.paragraph_format.space_before = Pt(146)'''
assert old_p1_aff in src, "Could not find old_p1_aff"
src = src.replace(old_p1_aff, new_p1_aff)

# 5. Page 2 section setup and spacing (starts at y=294.3 pt in orig)
old_p2_sec = '''        print("  ▶ Building Page 2 (Screened Abstract Continuation + Created Antecedentes & Caso Clínico)...")
        _add_page_break(doc)'''
new_p2_sec = '''        print("  ▶ Building Page 2 (Screened Abstract Continuation + Created Antecedentes & Caso Clínico)...")
        _setup_page_section(doc, is_odd=False, top_margin_in=0.55)'''
assert old_p2_sec in src, "Could not find old_p2_sec"
src = src.replace(old_p2_sec, new_p2_sec)

old_p2_sp = '''        p_sp2 = doc.add_paragraph()
        p_sp2.paragraph_format.space_before = Pt(0)
        p_sp2.paragraph_format.space_after = Pt(10)

        # Page 2: 2 equal columns (3.05" and 3.05")
        p2_row = _create_layout_row(doc, [3.05, 3.05])
        p2_c_left, p2_c_right = p2_row.cells[0], p2_row.cells[1]'''
new_p2_sp = '''        p_sp2 = doc.add_paragraph()
        p_sp2.paragraph_format.space_before = Pt(0)
        p_sp2.paragraph_format.space_after = Pt(114)

        # Page 2: 2 equal columns (2.85" and 2.85" with 0.24" gutter)
        p2_c_left, p2_c_right = _create_2col_row(doc, 2.85, 0.24, 2.85)'''
assert old_p2_sp in src, "Could not find old_p2_sp"
src = src.replace(old_p2_sp, new_p2_sp)

# 6. Page 3 section setup, offset, and 2-column layout (Right col starts at y=113.4, Fig 1 starts at y=129.2)
old_p3 = '''        print("  ▶ Building Page 3 (Created Figure 1 Pink Box + Created Caso Clínico & Discusión)...")
        _add_page_break(doc)

        hdr3 = _create_layout_row(doc, [4.5, 1.81])'''
new_p3 = '''        print("  ▶ Building Page 3 (Created Figure 1 Pink Box + Created Caso Clínico & Discusión)...")
        _setup_page_section(doc, is_odd=True, top_margin_in=0.58)

        hdr3 = _create_layout_row(doc, [3.80, 2.14])'''
assert old_p3 in src, "Could not find old_p3"
src = src.replace(old_p3, new_p3)

# Table starts at y=113.4 pt
assert 'p3_div.paragraph_format.space_after = Pt(23)' in src
src = src.replace('p3_div.paragraph_format.space_after = Pt(23)', 'p3_div.paragraph_format.space_after = Pt(8)')

old_p3_row = '''        # Page 3: 2 columns
        p3_row = _create_layout_row(doc, [2.85, 2.85])
        p3_c_left, p3_c_right = p3_row.cells[0], p3_row.cells[1]
        _set_cell_padding(p3_c_left, top_pt=0, bottom_pt=0, left_pt=0, right_pt=2)

        # Left Column: Red accent stripe + Figure 1 Stacked Images + Pink Caption
        _add_red_stripe(p3_c_left, 2.81)'''
new_p3_row = '''        # Page 3: 2 columns with gutter
        p3_c_left, p3_c_right = _create_2col_row(doc, 2.85, 0.24, 2.85)

        # In Left Column: space before red stripe to hit Fig 1 y=129.2 pt exactly
        p3_fig_sp = p3_c_left.paragraphs[0]
        p3_fig_sp.paragraph_format.space_before = Pt(15)
        p3_fig_sp.paragraph_format.space_after = Pt(0)
        p3_fig_sp.paragraph_format.line_spacing = Pt(1)

        # Left Column: Red accent stripe + Figure 1 Stacked Images + Pink Caption
        _add_red_stripe(p3_c_left, 2.81)'''
assert old_p3_row in src, "Could not find old_p3_row"
src = src.replace(old_p3_row, new_p3_row)

# 7. Page 4 offset and 2-column layout (Fig 2 starts at y=130.0 pt)
assert 'p4_div.paragraph_format.space_after = Pt(36)' in src
src = src.replace('p4_div.paragraph_format.space_after = Pt(36)', 'p4_div.paragraph_format.space_after = Pt(43)')

old_p4_row = '''        # Page 4: 2 Columns below Figure 2
        p4_row = _create_layout_row(doc, [2.85, 2.85])
        p4_c_left, p4_c_right = p4_row.cells[0], p4_row.cells[1]'''
new_p4_row = '''        # Page 4: 2 Columns below Figure 2 with gutter
        p4_c_left, p4_c_right = _create_2col_row(doc, 2.85, 0.24, 2.85)'''
assert old_p4_row in src, "Could not find old_p4_row"
src = src.replace(old_p4_row, new_p4_row)

# 8. Page 5 offset and 2-column layout (Fig 3 starts at y=128.8 pt)
assert 'p5_div.paragraph_format.space_after = Pt(34)' in src
src = src.replace('p5_div.paragraph_format.space_after = Pt(34)', 'p5_div.paragraph_format.space_after = Pt(42)')

old_p5_row = '''        # Page 5: 2 Columns below Figure 3
        p5_row = _create_layout_row(doc, [2.85, 2.85])
        p5_c_left, p5_c_right = p5_row.cells[0], p5_row.cells[1]'''
new_p5_row = '''        # Page 5: 2 Columns below Figure 3 with gutter
        p5_c_left, p5_c_right = _create_2col_row(doc, 2.85, 0.24, 2.85)'''
assert old_p5_row in src, "Could not find old_p5_row"
src = src.replace(old_p5_row, new_p5_row)

# 9. Page 6 offset and 2-column layout (Right col starts at y=113.4, Fig 4 starts at y=129.6)
assert 'p6_div.paragraph_format.space_after = Pt(1)' in src
src = src.replace('p6_div.paragraph_format.space_after = Pt(1)', 'p6_div.paragraph_format.space_after = Pt(8)')

old_p6_row = '''        # Page 6: 2 Columns
        p6_row = _create_layout_row(doc, [2.85, 2.85])
        p6_c_left, p6_c_right = p6_row.cells[0], p6_row.cells[1]

        # Left Column: Red Stripe + Figure 4 Image + Pink Caption
        _add_red_stripe(p6_c_left, 2.81)'''
new_p6_row = '''        # Page 6: 2 Columns with gutter
        p6_c_left, p6_c_right = _create_2col_row(doc, 2.85, 0.24, 2.85)

        # In Left Column: space before red stripe to hit Fig 4 y=129.6 pt exactly
        p6_fig_sp = p6_c_left.paragraphs[0]
        p6_fig_sp.paragraph_format.space_before = Pt(15)
        p6_fig_sp.paragraph_format.space_after = Pt(0)
        p6_fig_sp.paragraph_format.line_spacing = Pt(1)

        # Left Column: Red Stripe + Figure 4 Image + Pink Caption
        _add_red_stripe(p6_c_left, 2.81)'''
assert old_p6_row in src, "Could not find old_p6_row"
src = src.replace(old_p6_row, new_p6_row)

# 10. Page 7 offset, 2-column layout, and card spacing (References start at y=111.7 pt)
assert 'p7_div.paragraph_format.space_after = Pt(6)' in src
src = src.replace('p7_div.paragraph_format.space_after = Pt(6)', 'p7_div.paragraph_format.space_after = Pt(8)')

old_p7_row = '''        # 2 Columns for References 3–11
        p7_row = _create_layout_row(doc, [2.85, 2.85])
        p7_c_left, p7_c_right = p7_row.cells[0], p7_row.cells[1]'''
new_p7_row = '''        # 2 Columns for References 3–11 with gutter
        p7_c_left, p7_c_right = _create_2col_row(doc, 2.85, 0.24, 2.85)'''
assert old_p7_row in src, "Could not find old_p7_row"
src = src.replace(old_p7_row, new_p7_row)

assert 'p_sp7.paragraph_format.space_before = Pt(4)' in src
src = src.replace('p_sp7.paragraph_format.space_before = Pt(4)', 'p_sp7.paragraph_format.space_before = Pt(78)')

# Page 7 Aviso card sub-layout: Badges next to Android/iPhone text
old_p7_badges = '''        p_avt3 = in_c_left.add_paragraph()
        p_avt3.paragraph_format.space_before = Pt(2)
        p_avt3.paragraph_format.space_after = Pt(3)
        r_at3 = p_avt3.add_run("La aplicación está disponible para Android o iPhone.")
        r_at3.font.name = "Arial"
        r_at3.font.size = Pt(7.8)
        r_at3.font.bold = True

        p_badge = in_c_left.add_paragraph()
        p_badge.paragraph_format.space_before = Pt(0)
        p_badge.paragraph_format.space_after = Pt(0)
        if "p7_img3" in assets:
            r_b1 = p_badge.add_run()
            r_b1.add_picture(str(assets["p7_img3"]), width=Inches(0.40))
        if "p7_img0" in assets:
            r_b2 = p_badge.add_run("  ")
            r_b3 = p_badge.add_run()
            r_b3.add_picture(str(assets["p7_img0"]), width=Inches(0.40))'''

new_p7_badges = '''        # Bottom sub-row: Left: text (2.10 in), Right: Badges (1.50 in)
        sub_tbl = in_c_left.add_table(rows=1, cols=2)
        sub_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        sub_tbl.autofit = False
        _set_borderless_table(sub_tbl)
        sub_tbl.rows[0].cells[0].width = Inches(2.10)
        sub_tbl.rows[0].cells[1].width = Inches(1.50)
        _set_cell_padding(sub_tbl.rows[0].cells[0], top_pt=2, bottom_pt=0, left_pt=0, right_pt=2)
        _set_cell_padding(sub_tbl.rows[0].cells[1], top_pt=2, bottom_pt=0, left_pt=2, right_pt=0)

        p_avt3 = sub_tbl.rows[0].cells[0].paragraphs[0]
        p_avt3.paragraph_format.space_before = Pt(0)
        p_avt3.paragraph_format.space_after = Pt(0)
        r_at3 = p_avt3.add_run("La aplicación está disponible\\npara Android o iPhone.")
        r_at3.font.name = "Arial"
        r_at3.font.size = Pt(7.8)
        r_at3.font.bold = True

        p_badge = sub_tbl.rows[0].cells[1].paragraphs[0]
        p_badge.paragraph_format.space_before = Pt(0)
        p_badge.paragraph_format.space_after = Pt(0)
        if "p7_img3" in assets:
            r_b1 = p_badge.add_run()
            r_b1.add_picture(str(assets["p7_img3"]), width=Inches(0.45))
        if "p7_img0" in assets:
            r_b2 = p_badge.add_run(" ")
            r_b3 = p_badge.add_run()
            r_b3.add_picture(str(assets["p7_img0"]), width=Inches(0.45))'''
assert old_p7_badges in src, "Could not find old_p7_badges"
src = src.replace(old_p7_badges, new_p7_badges)

# Also align tables to LEFT
src = src.replace('aviso_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER', 'aviso_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT')
src = src.replace('fig2_img_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER', 'fig2_img_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT')
src = src.replace('cap2_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER', 'cap2_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT')
src = src.replace('fig3_img_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER', 'fig3_img_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT')
src = src.replace('cap3_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER', 'cap3_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT')

# Write to modifications
out_mod = Path('runs/DOC0000074469_20261002_1045/modifications/journal_article_builder.py')
out_mod.write_text(src, encoding='utf-8')
print('✅ Successfully wrote refined builder to:', out_mod)
