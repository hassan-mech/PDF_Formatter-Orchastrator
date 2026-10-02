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
