# Original User Request

## 2026-10-02T11:40:38Z

Convert "DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf" to Word (.docx) with pixel-accurate layout and typography preservation using the collaborative multi-agent team, enforcing native Word multi-column sections, zero layout tables for text columns, and selective screening of English abstract continuation crops.

Working directory: d:\Hassan\Amouna\Hasssan&Eman
Integrity mode: development

## Requirements

### R1. Generic Pipeline Reuse (Zero Custom Scripts)
- Route the document through the existing configurable pipeline (`scripts/convert_to_word.py` / `configs/medical_journal_portrait.yaml`) and registered stages in `scripts/stages/`.
- Standalone one-off builder scripts are prohibited (strict negative points). Work within the established architectural components.
- Run in an isolated workspace directory (`runs/DOC0000074469_<timestamp>/`) containing dedicated subdirectories: `extract/`, `ocr/`, `build/`, `qa/`, and `modifications/`.

### R2. Pure Native Multi-Column Layout Architecture
- Multi-column sections must use Word's native `<w:cols>` configuration (`WD_SECTION_START.CONTINUOUS`) and native column breaks (`WD_BREAK.COLUMN`).
- Zero layout tables (`<w:tbl>`) for body text columns.
- Maintain authentic alternating facing-page mirror margins (Odd: 1.38" left, 0.96" right; Even: 0.98" left, 1.34" right) across a constant 5.94" printable column grid.
- Apply the dual-section footer mapping architecture: call `_set_section_footer()` on BOTH the `NEW_PAGE` header section and the `CONTINUOUS` body column sections for every page.

### R3. Language Routing & Selective Screening
- Screen ONLY the English abstract continuation crops (`p1_eng_abstract_clean.png` and `p2_eng_abstract_clean.png`).
- Attach 0.5pt hidden text runs exclusively to the screened images to ensure 100% text token recall without visual duplication.
- All Spanish body text, headings, metadata, and review stamp must be 100% genuine styled editable Word text.

### R4. Span-Level Typography & Graphical Accents
- Auto-extract bold, italic, font size, and color from PDF text spans (`page.get_text("dict")`).
- Validate vector drawings using PyMuPDF `page.get_drawings()` before emitting any lines; do not inject hallucinated decorative divider rules.

## Acceptance Criteria

### A1. Structural & Layout Integrity
- [ ] No layout tables (`<w:tbl>`) used for body text columns.
- [ ] Alternating mirror margins preserved across all facing pages.
- [ ] Running headers and footers present and synchronized on both 1-column and 2-column sections.

### A2. Content & Typography Fidelity
- [ ] Only English abstract continuation crops are screened; all Spanish text and review stamp are editable text.
- [ ] 0.5pt hidden text attached to screened crops for full token recall.
- [ ] Span-level font styling (bold, italic, size, color) matches source PDF.

### A3. 5-Axis QA Verification Gate (tol = 20)
- [ ] Visual similarity score ≥ 0.85 (with tolerance parameter `tol=20`).
- [ ] Text token recall score ≥ 0.85.
- [ ] Layout geometry score ≥ 0.85.
- [ ] Table structure score ≥ 0.85.
- [ ] Placeholder / tag lexicon compliance score ≥ 0.85.
- [ ] Overall QA status reports `PASS`.
- [ ] Final DOCX promoted to `output/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.docx`.
