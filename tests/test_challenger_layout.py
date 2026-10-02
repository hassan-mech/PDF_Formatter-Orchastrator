"""
Adversarial Layout and Structural Test Harness
Target: output/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.docx
Original: input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf
"""

import sys
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsdecls
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_BREAK
import pymupdf

DOCX_PATH = "output/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#" + ".docx"
PDF_PATH = "input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#" + ".pdf"

def run_tests():
    print(f"Loading DOCX: {DOCX_PATH}")
    doc = docx.Document(DOCX_PATH)
    pdf = pymupdf.open(PDF_PATH)
    
    results = {}
    failures = []

    # =========================================================================
    # TEST 1: Table Check
    # =========================================================================
    print("\n--- TEST 1: Table Check ---")
    doc_tables = doc.tables
    print(f"doc.tables count: {len(doc_tables)}")
    
    # Check all <w:tbl> elements in entire document body
    all_tbls = doc._element.xpath(".//w:tbl")
    print(f"Total <w:tbl> elements in document body: {len(all_tbls)}")
    
    # Check headers and footers for any hidden tables
    header_footer_tbls = 0
    for i, s in enumerate(doc.sections):
        for part in [s.header, s.footer]:
            tbls = part._element.xpath(".//w:tbl")
            header_footer_tbls += len(tbls)
    print(f"Total <w:tbl> elements in headers/footers: {header_footer_tbls}")

    # Check table count assertion
    t1_pass = True
    if len(doc_tables) != 1 or len(all_tbls) != 1 or header_footer_tbls != 0:
        failures.append(f"T1 FAIL: Expected exactly 1 table, found doc.tables={len(doc_tables)}, body_tbls={len(all_tbls)}, hf_tbls={header_footer_tbls}")
        t1_pass = False
    else:
        tbl = doc_tables[0]
        tbl_text = "".join(cell.text for row in tbl.rows for cell in row.cells)
        print(f"Table 1 text preview: {tbl_text[:80]!r}")
        if "AVISO IMPORTANTE" not in tbl_text:
            failures.append("T1 FAIL: Table 1 does not contain 'AVISO IMPORTANTE'")
            t1_pass = False
        else:
            print("Table 1 confirmed to be Announcement Box Card on Page 7.")
            
    # Check whether any layout tables are used for multi-column body text
    # A layout table typically has invisible borders or multiple columns containing paragraphs of article body text
    # Here all_tbls == 1 and it's the announcement card.
    results["T1_table_check"] = {
        "pass": t1_pass,
        "table_count": len(doc_tables),
        "total_w_tbl": len(all_tbls),
        "header_footer_tbl": header_footer_tbls,
        "is_announcement_card": ("AVISO IMPORTANTE" in tbl_text) if len(doc_tables) == 1 else False
    }

    # =========================================================================
    # TEST 2: Section & Column Check
    # =========================================================================
    print("\n--- TEST 2: Section & Column Check ---")
    sections = doc.sections
    print(f"Total sections: {len(sections)}")
    
    t2_pass = True
    if len(sections) != 15:
        failures.append(f"T2 FAIL: Expected 15 sections, found {len(sections)}")
        t2_pass = False

    section_details = []
    multi_col_sections = 0
    single_col_sections = 0

    # Native column breaks check in paragraphs
    col_break_count = 0
    for p_idx, p in enumerate(doc.paragraphs):
        c_brs = p._element.xpath('.//w:br[@w:type="column"]')
        if c_brs:
            col_break_count += len(c_brs)
            print(f"  Para {p_idx} has {len(c_brs)} column break(s)")
    print(f"Total native column breaks found in document: {col_break_count}")
    
    for i, s in enumerate(sections):
        sectPr = s._sectPr
        type_elem = sectPr.xpath("./w:type")
        start_type = type_elem[0].get(qn("w:val")) if type_elem else "nextPage"
        
        cols_elem = sectPr.xpath("./w:cols")
        cols_num = cols_elem[0].get(qn("w:num")) if cols_elem else "1"
        cols_space = cols_elem[0].get(qn("w:space")) if cols_elem else None
        cols_equal = cols_elem[0].get(qn("w:equalWidth")) if cols_elem else None
        
        # Check sub-columns <w:col> if not equalWidth
        NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "v": "urn:schemas-microsoft-com:vml"}
        sub_cols = cols_elem[0].xpath("./w:col", namespaces=NS) if cols_elem else []
        sub_col_widths = [c.get(qn("w:w")) for c in sub_cols]

        is_multi = (cols_num == "2")
        if is_multi:
            multi_col_sections += 1
        else:
            single_col_sections += 1

        detail = {
            "section_index": i + 1,
            "start_type": start_type,
            "num_cols": cols_num,
            "space": cols_space,
            "equal_width": cols_equal,
            "sub_cols": sub_col_widths
        }
        section_details.append(detail)
        print(f"  Sec {i+1:2d}: start={start_type:10s} cols={cols_num} space={str(cols_space):5s} equal={str(cols_equal):5s} sub_cols={sub_col_widths}")

    print(f"Single-column sections: {single_col_sections}, Multi-column (num=2) sections: {multi_col_sections}")

    # Verify that multi-column sections exist and have num=2
    if multi_col_sections < 7:
        failures.append(f"T2 FAIL: Expected at least 7 multi-column sections, found {multi_col_sections}")
        t2_pass = False
        
    if col_break_count < 6:
        failures.append(f"T2 FAIL: Expected at least 6 column breaks across 7 pages, found {col_break_count}")
        t2_pass = False

    results["T2_section_column_check"] = {
        "pass": t2_pass,
        "total_sections": len(sections),
        "multi_col_sections": multi_col_sections,
        "single_col_sections": single_col_sections,
        "native_column_breaks": col_break_count,
        "details": section_details
    }

    # =========================================================================
    # TEST 3: Mirror Margin Check
    # =========================================================================
    print("\n--- TEST 3: Mirror Margin Check ---")
    # Expected page-to-section mapping:
    # Page 1 (Odd): Sec 1 (Banner), Sec 2 (Body)
    # Page 2 (Even): Sec 3 (Banner), Sec 4 (Body)
    # Page 3 (Odd): Sec 5 (Banner), Sec 6 (Body)
    # Page 4 (Even): Sec 7 (Banner), Sec 8 (Body)
    # Page 5 (Odd): Sec 9 (Banner), Sec 10 (Body)
    # Page 6 (Even): Sec 11 (Banner), Sec 12 (Body)
    # Page 7 (Odd): Sec 13 (Banner), Sec 14 (Body), Sec 15 (Card)
    
    odd_page_sec_indices = [1, 2, 5, 6, 9, 10, 13, 14, 15]  # 1-indexed
    even_page_sec_indices = [3, 4, 7, 8, 11, 12]             # 1-indexed
    
    t3_pass = True
    margin_audit = []
    
    for i, s in enumerate(sections):
        sec_num = i + 1
        l_in = s.left_margin.inches
        r_in = s.right_margin.inches
        t_in = s.top_margin.inches
        b_in = s.bottom_margin.inches
        page_w = s.page_width.inches
        grid_w = page_w - (l_in + r_in)
        
        # Check odd vs even page margins
        if sec_num in odd_page_sec_indices:
            page_type = "Odd"
            exp_l, exp_r = 1.38, 0.96
        elif sec_num in even_page_sec_indices:
            page_type = "Even"
            exp_l, exp_r = 0.98, 1.34
        else:
            page_type = "Unknown"
            exp_l, exp_r = None, None

        l_diff = abs(l_in - exp_l) if exp_l else 0
        r_diff = abs(r_in - exp_r) if exp_r else 0
        grid_diff = abs(grid_w - 5.94)

        audit_entry = {
            "sec": sec_num,
            "page_type": page_type,
            "left": round(l_in, 3),
            "right": round(r_in, 3),
            "top": round(t_in, 3),
            "bottom": round(b_in, 3),
            "grid_w": round(grid_w, 3),
            "grid_diff": round(grid_diff, 3)
        }
        margin_audit.append(audit_entry)
        
        print(f"  Sec {sec_num:2d} ({page_type}): L={l_in:.2f}\" (exp {exp_l:.2f}\"), R={r_in:.2f}\" (exp {exp_r:.2f}\"), Grid={grid_w:.2f}\" (exp 5.94\" +/- 0.02\")")
        
        if l_diff > 0.02:
            failures.append(f"T3 FAIL: Sec {sec_num} Left margin {l_in:.2f} != expected {exp_l:.2f}")
            t3_pass = False
        if r_diff > 0.02:
            failures.append(f"T3 FAIL: Sec {sec_num} Right margin {r_in:.2f} != expected {exp_r:.2f}")
            t3_pass = False
        if grid_diff > 0.02:
            failures.append(f"T3 FAIL: Sec {sec_num} Grid width {grid_w:.2f} outside 5.94 +/- 0.02")
            t3_pass = False

    results["T3_mirror_margin_check"] = {
        "pass": t3_pass,
        "margin_audit": margin_audit
    }

    # =========================================================================
    # TEST 4: Dual-Section Footer Check
    # =========================================================================
    print("\n--- TEST 4: Dual-Section Footer Check ---")
    # Verify that footers across all 15 sections are configured on both NEW_PAGE and CONTINUOUS sections.
    # Pages 1 to 7 should have page numbers 666 to 672.
    expected_page_numbers = {
        1: 666, 2: 666,
        3: 667, 4: 667,
        5: 668, 6: 668,
        7: 669, 8: 669,
        9: 670, 10: 670,
        11: 671, 12: 671,
        13: 672, 14: 672, 15: 672
    }
    
    t4_pass = True
    footer_audit = []
    
    for i, s in enumerate(sections):
        sec_num = i + 1
        footer = s.footer
        is_linked = footer.is_linked_to_previous
        footer_paras = [p.text.strip() for p in footer.paragraphs if p.text.strip()]
        footer_text = " | ".join(footer_paras)
        exp_pnum = str(expected_page_numbers[sec_num])
        
        has_exp_pnum = exp_pnum in footer_text
        print(f"  Sec {sec_num:2d}: linked={str(is_linked):5s} footer_text={footer_text!r} (expects pnum {exp_pnum}: {has_exp_pnum})")
        
        if not has_exp_pnum:
            failures.append(f"T4 FAIL: Sec {sec_num} footer missing expected page number {exp_pnum}. Found: {footer_text!r}")
            t4_pass = False
            
        if is_linked:
            # We want footers explicitly configured per section (or if linked, only within the same page)
            # The dual-section footer invariant requires _set_section_footer() on both sections with is_linked_to_previous = False
            failures.append(f"T4 FAIL: Sec {sec_num} footer is_linked_to_previous is True (should be False for independent section mapping)")
            t4_pass = False
            
        footer_audit.append({
            "sec": sec_num,
            "is_linked": is_linked,
            "footer_text": footer_text,
            "has_expected_page_number": has_exp_pnum
        })

    results["T4_dual_section_footer_check"] = {
        "pass": t4_pass,
        "footer_audit": footer_audit
    }

    # =========================================================================
    # TEST 5: Vector Drawing Anti-Hallucination Check
    # =========================================================================
    print("\n--- TEST 5: Vector Drawing Anti-Hallucination Check ---")
    # Check PDF drawings across all pages
    pdf_drawings_summary = []
    for pno in range(len(pdf)):
        page = pdf[pno]
        drawings = page.get_drawings()
        print(f"  PDF Page {pno+1}: {len(drawings)} drawing item(s)")
        # Look specifically for header horizontal lines in PDF (y < 70 pt)
        header_lines_in_pdf = []
        for d in drawings:
            rect = d.get("rect")
            # If drawing is near top header area
            if rect and rect.y1 < 75:
                header_lines_in_pdf.append(d)
        print(f"    Header area drawings in PDF: {len(header_lines_in_pdf)}")
        pdf_drawings_summary.append({
            "page": pno + 1,
            "total_drawings": len(drawings),
            "header_drawings": len(header_lines_in_pdf)
        })

    # In DOCX: Check for hallucinated horizontal lines under running headers
    t5_pass = True
    docx_header_lines = []
    for i, s in enumerate(sections):
        sec_num = i + 1
        header = s.header
        # Check for paragraph borders (e.g. bottom border)
        for p_idx, p in enumerate(header.paragraphs):
            pBdr = p._element.xpath("./w:pPr/w:pBdr")
            if pBdr:
                bottom_bdr = pBdr[0].xpath("./w:bottom")
                if bottom_bdr:
                    bdr_val = bottom_bdr[0].get(qn("w:val"))
                    if bdr_val != "none":
                        docx_header_lines.append(f"Sec {sec_num} header p{p_idx} has bottom border: {bdr_val}")
            # Check for shapes or drawings in header
            drawings_in_p = p._element.xpath(".//w:drawing | .//w:pict")
            if drawings_in_p:
                docx_header_lines.append(f"Sec {sec_num} header p{p_idx} has drawing/shape elements: {len(drawings_in_p)}")

    print(f"DOCX header divider lines found: {len(docx_header_lines)}")
    if docx_header_lines:
        for hl in docx_header_lines:
            print(f"  WARNING/FAIL: {hl}")
            failures.append(f"T5 FAIL: Hallucinated header divider line found: {hl}")
        t5_pass = False
    else:
        print("  Confirmed: Zero hallucinated horizontal lines exist under running headers in DOCX.")

    # Also check body paragraph 0 of each page to see if an artificial divider line was placed right under header
    # Note: On Page 1, there is a legitimate crimson title divider rule below the title (PDF Page 1 Drawing 0).
    # But running headers on Pages 2-7 must NOT have divider lines below them.
    for i, s in enumerate(sections):
        # Header text
        header_text = "".join(p.text for p in s.header.paragraphs).strip()
        print(f"  Sec {i+1} header text: {header_text!r}")

    results["T5_vector_anti_hallucination_check"] = {
        "pass": t5_pass,
        "docx_header_lines": docx_header_lines,
        "pdf_drawings": pdf_drawings_summary
    }

    # =========================================================================
    # SUMMARY & VERDICT
    # =========================================================================
    print("\n================== SUMMARY ==================")
    all_passed = t1_pass and t2_pass and t3_pass and t4_pass and t5_pass
    verdict = "APPROVE" if all_passed else "FAIL"
    print(f"VERDICT: {verdict}")
    if failures:
        print(f"Failures ({len(failures)}):")
        for f in failures:
            print(f"  - {f}")
    else:
        print("All 5 challenge axes passed empirically with zero defects!")
        
    return verdict, results, failures

if __name__ == "__main__":
    verdict, results, failures = run_tests()
    sys.exit(0 if verdict == "APPROVE" else 1)
