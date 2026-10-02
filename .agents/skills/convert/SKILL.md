---
name: convert
description: Convert a PDF from input/ to DOCX with pixel-faithful fidelity
---

# Convert PDF to DOCX

## Task
$ARGUMENTS

## Parse
- First token = filename (assume it lives in `input/`).
- Everything after the first line = special instructions.
- If no instructions, use: "none — use configs/default.yaml as-is."
- Echo back the parsed (FILE, INSTRUCTIONS) tuple before doing anything.

## Pre-flight
1. Read `.agents/rules/` fully (SOLID contract, workflow, tags, symbols).
2. Workspace containment check: Verify that ALL code, test suites, and builders run 100% within the repository. NEVER run or write scratch scripts in `brain/scratch/` or external temp folders.
3. Read last 20 lines of `DECISIONS.md` and `CHANGELOG.md`.
4. List `configs/*.yaml` and `scripts/stages/**/*.py`.
5. Reuse check: can an existing config + existing stages handle this file?
   - YES → use it, justify in one line.
   - NO → log `GAP: <one line>` in DECISIONS.md.
     - Config-only gap → create `configs/<doc_type>.yaml`. Done.
     - Code gap → propose ONE handler behind an existing interface.
       Requires user approval, recorded in DECISIONS.md.

## Builder Selection Rules (MANDATORY — read before choosing a builder)

### Rule 1 — Screen-Capture First
When the user says ANY of these (or similar):
  - "screen", "screenshot", "screened", "put it as image", "same as the PDF",
    "don't recreate", "just embed", "pixel-perfect image"
→ ALWAYS use `screen_capture_builder` + `configs/screen_capture.yaml`.
→ Do NOT use `docx_builder` or `chinese_medical_builder`.
→ Do NOT recreate tables, fonts, or text as Word content.

### Rule 2 — Hybrid Document Routing (Mixed Scanned + Digital / Clinical Reports)
When the file has administrative sections (questionnaires/tables) alongside foreign clinical scans (e.g. GSK questionnaires with Chinese hospital reports):
→ ALWAYS use `chinese_medical_builder` + `configs/medical_chinese_landscape.yaml`.
- **Pure Administrative Pages (Pages 1, 2, 3, 9):** Full-bleed 300 DPI screen capture + 0.5pt hidden white text run for 100% token recall.
- **Hybrid Pages with Embedded Medical Scans (Pages 4, 5, 6, 8):**
  - Use a visually hidden 1-column layout container table (`<w:tblBorders>` all `none`, padding = 0) spanning 11.0" width.
  - Embed screened English sections into container rows at 11.0" width (300 DPI).
  - Reconstruct foreign clinical reports into editable Word tables/text inside container rows with matching left indent (`<w:tblInd>`):
    - Page 4: Screened top English table + Created Cortisol lab table.
    - Page 5: Created Thyroid lab table + Screened bottom Question 3.
    - Page 6: Screened top Question 4 + Created Sputum PCR table + Screened bottom Questions 5/6.
    - Page 8: Created Chinese CT Scan Report (`梅州市人民医院 CT检查报告书`) + Screened bottom Question 7.
- **Pure Foreign Clinical Reports (Pages 10–41):** Real editable `docx.Table` with strict 5-column / 8-column separation (Item Name, Result Value, Flag, Unit, Reference Interval, Method). Result values (e.g. `20.76`) MUST be in their own column.
- **Header-Value Alignment:** Center numeric results and status text directly beneath `结果` / `提示` / `单位` column headers. Never right-align values against the far border.
- **Font Scale Hierarchy:** Match natural visual flow: 12–14 pt bold for titles, 8.5–9.5 pt for table cells and body text, 7.5–8.0 pt for footers.

### Rule 3 — Standard Digital PDF
All other PDFs with extractable digital text and standard layout:
→ Use `docx_builder` + `configs/default.yaml`.

### Rule 4 — Bilingual Medical Journal Portrait Protocol
When converting bilingual medical journal articles (e.g. *Dermatología Revista mexicana*):
→ Use `journal_article_builder` + `configs/medical_journal_portrait.yaml`.
- **English Administrative / Abstract Sections:** Screened (300 DPI high-res crop images) + connected 0.5pt hidden white text runs attached directly to the image paragraph for 100% token recall.
- **Foreign / Spanish Clinical Content:** 100% created as genuine editable Word text, native 2-column `<w:cols>` sections with `WD_BREAK.COLUMN` (ZERO layout tables), and figure containers.
- **Facing-Page Margins:** Alternating mirror margins (Odd: 1.38" left, 0.96" right; Even: 0.98" left, 1.34" right) across a constant 5.94" printable column grid.
- **Figures:** Embedded across columns with pink shaded caption containers, positioned via analytical bounding-box calculation ($\Delta y = y_{\text{target}} - y_{\text{current}}$).

## Screen-Capture Quality Rules (non-negotiable when using screen_capture_builder)
- DPI MUST be 300 minimum. Use 300 in `configs/screen_capture.yaml`.
- Both `width=` AND `height=` MUST be passed to `add_picture()`.
  Passing only width causes white space at the bottom — this is a bug.
- All section margins MUST be 0 (top, bottom, left, right, header, footer, gutter).
- Each PDF page MUST become exactly ONE DOCX section.
- Page size in DOCX MUST match the PDF page dimensions in points, converted to twips.
- Do NOT add any header, footer, watermark, or extra paragraphs.
- Delete all default paragraphs from the blank Document before adding pages.
- The output DOCX opened in Word must look identical to the PDF — no white borders,
  no clipping, no rotation, no scaling artifacts.

## Workflow (per governance rules, multi-agent isolation, 2-Tier Fast-Path)
// turbo
0. **Initialize Teamwork & Sub-Agent Orchestration:**
   - Prompt user to run `/teamwork-preview` (or toggle Plan Mode) to activate collaborative multi-agent team mode.
   - Coordinate specialized subagents (`inspector`, `extractor`, `ocr-engine`, `docx-builder`, `qa-verifier`) concurrently to maximize speed and eliminate serial bottlenecks.
1. **Pre-flight, Multi-File Isolation & Reuse Check:**
   - Scan `runs/` for active or recently finished runs from other conversations.
   - In multi-file operations, EVERY file MUST have its own isolated directory: `runs/<stem>_<timestamp>/` containing subdirectories: `extract/`, `ocr/`, `build/`, `qa/`, and `modifications/`.
   - **Zero Custom Code Penalty:** Re-use standard pipeline stages (`scripts/stages/`). Writing one-off builder code per file is strictly penalized with negative points. If a feature is needed, generically extend the core stage.
2. **Span-Level Style & Typography Extraction:**
   - Extract text spans preserving `bold` (`flags & 16`), `italic` (`flags & 2`), `font`, `size`, and `color` via PyMuPDF `dict` extraction.
   - Do NOT flatten styled spans to plain unstyled strings or manually guess styles.
3. **Execution Timers & Safety:**
   - Wrap external tasks with strict execution timers (max 15s for Word COM, max 60s for OCR).
   - Auto-terminate hung tasks and kill lingering `WINWORD.EXE` processes immediately upon timeout.
4. **Build Candidate via Analytical Coordinate Derivation (≤ 2.0s):**
   - Extract exact element bounding boxes (`Rect(x0, y0, x1, y1)`) via PyMuPDF in ONE pass.
   - Calculate all section margins, column widths, and vertical paragraph offsets analytically ($\Delta y = y_{\text{target\_PDF}} - y_{\text{current\_DOCX}}$).
   - Generate candidate DOCX into `runs/<operation_id>/build/<stem>.docx`.
5. **Tier 1 Fast-Path In-Memory Verification (≤ 0.5s):**
   - Check section count, page geometry, table widths, and text token recall directly in Python using `pymupdf` and `python-docx`.
   - If Tier 1 fails: fix parameters analytically in code without launching Word.
6. **Tier 2 Final Word COM QA Gate (≤ 15.0s, run ONCE):**
   - Run `tests/qa_agent.py` outputting report to `runs/<operation_id>/qa/report.json`.
   - All 5 axes (Visual, Text, Layout, Tables, Placeholders) must achieve ≥ 0.85.
7. **Consensus Aggregation & Code Promotion (/learn):**
   - On full PASS (all 5 axes ≥ 0.85):
     a. Cross-check modifications in `runs/<operation_id>/modifications/` against other recent runs for contradictions.
     b. Promote verified, conflict-free logic into shared core builders, configs, and `.agents/rules/`.
     c. Promote output DOCX to `output/<stem>.docx` and source PDF to `finished/`.
     d. Append entries to `CHANGELOG.md` and `DECISIONS.md`. Mark run manifest as `promoted`.

## Fidelity rules (non-negotiable)
- Page count, size, orientation (per page), margins, fonts, sizes, line
  spacing, alignment must match the PDF — read from PDF, not guessed.
- Style fidelity: bold and italic runs within the same paragraph must be rendered identically to the source PDF.
- Tables = real `docx.Table` with preserved widths/fills/borders.
- Scans/stamps/logos/signatures/handwriting → Appendix A tags, not images.
  EXCEPTION: in screen_capture mode, tags are NOT emitted (images speak for themselves).
- NEVER emit `[logo:]` without readable text. Use `[icon]` or embedded graphics for graphical badges/icons.
- `[hw: ...]` payload in Italic. Header note only if file is mostly handwritten.
- Symbols from Appendix B only. No ASCII substitutions.


## Output
Per governance rules § Response Format: reuse verdict, SOLID check,
files touched, QA 5 axes with scores, promotion, gap log.
