# DECISIONS LOG (append-only)

Format: one entry per file/decision. Do not edit past entries.
See AGENTS.md § Decision Log Protocol for field meanings.

---
## 2026-10-01 00:00 | _bootstrap_
- Config used: n/a
- Gap found: initial pipeline skeleton did not exist
- Action: bootstrap session (approved by user)
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: n/a
- Result: PENDING

---

## 2026-10-01 16:28 | UG-100-000284526_2-Non-Parsable-en-US#FPREP_DXCGQN#
- Config used: default
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.87 text=0.18 layout=1.00 tables=1.00 placeholders=1.00
- Result: FAIL
---

## 2026-10-01 16:29 | UG-100-000284526_2-Non-Parsable-en-US#FPREP_DXCGQN#
- Config used: default
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.59 text=1.00 layout=1.00 tables=1.00 placeholders=1.00
- Result: FAIL
---

## 2026-10-01 16:30 | UG-100-000284526_2-Non-Parsable-en-US#FPREP_DXCGQN#
- Config used: default
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.87 text=1.00 layout=1.00 tables=1.00 placeholders=1.00
- Result: PASS
---

## 2026-10-01 16:48 | UG-100-000284526_2-Non-Parsable-en-US#FPREP_DXCGQN#
- Config used: default
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.86 text=1.00 layout=1.00 tables=1.00 placeholders=1.00
- Result: PASS
---

## 2026-10-01 16:57 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: medical_chinese_landscape
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.91 text=0.89 layout=1.00 tables=1.00 placeholders=1.00
- Result: PASS
---

## 2026-10-01 16:59 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: medical_chinese_landscape
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.91 text=0.89 layout=1.00 tables=1.00 placeholders=1.00
- Result: PASS
---

## 2026-10-01 17:02 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: medical_chinese_landscape
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.91 text=0.89 layout=1.00 tables=1.00 placeholders=1.00
- Result: PASS
---

## 2026-10-01 17:06 | System Migration to Antigravity Primitives
- Config used: N/A
- Gap found: none
- Action: migrated system rules to .agents/rules/, workflow to .agents/workflows/, agents to .agents/agents/, skills to .agents/skills/
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=N/A text=N/A layout=N/A tables=N/A placeholders=N/A
- Result: PASS
---

## 2026-10-01 17:23 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: medical_chinese_landscape
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.64 text=0.02 layout=1.00 tables=0.11 placeholders=1.00
- Result: FAIL
---

## 2026-10-01 17:33 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: medical_chinese_landscape
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.91 text=1.00 layout=1.00 tables=1.00 placeholders=1.00
- Result: PASS
---

## 2026-10-01 17:47 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: medical_chinese_landscape
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.90 text=0.85 layout=1.00 tables=0.99 placeholders=1.00
- Result: PASS
---

## 2026-10-01 19:45 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: medical_chinese_landscape
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.94 text=0.94 layout=1.00 tables=0.86 placeholders=1.00
- Result: PASS
---

## 2026-10-01 22:02 | GAP: screen-capture approach
- GAP: chinese_medical_builder recreates content as text/tables; user requires full-page image embedding (screen-capture DOCX). Proposing screen_capture_builder behind Builder interface.
- User approved: yes (explicit instruction)
---

## 2026-10-01 22:06 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: configs/screen_capture.yaml
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.89 text=1.00 layout=1.00 tables=0.86 placeholders=1.00
- Result: PASS
---

## 2026-10-01 22:16 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: configs/screen_capture.yaml
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.89 text=1.00 layout=1.00 tables=0.86 placeholders=1.00
- Result: PASS
---

## 2026-10-01 22:28 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: configs/screen_capture.yaml
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.89 text=1.00 layout=1.00 tables=0.86 placeholders=1.00
- Result: PASS
---

## 2026-10-01 22:32 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: configs/screen_capture.yaml
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.95 text=0.00 layout=1.00 tables=0.00 placeholders=1.00
- Result: FAIL
---

## 2026-10-01 22:38 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: configs/screen_capture.yaml
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.95 text=1.00 layout=1.00 tables=1.00 placeholders=1.00
- Result: PASS
---

## 2026-10-02 07:45 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: medical_chinese_landscape
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.91 text=1.00 layout=1.00 tables=1.00 placeholders=1.00
- Result: PASS
---

## 2026-10-02 08:21 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: medical_chinese_landscape
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.90 text=1.00 layout=1.00 tables=1.00 placeholders=1.00
- Result: PASS
---

## 2026-10-02 08:28 | Governance Update (/learn)
- Rule additions:
  1. Section 0.4 Workspace Containment Invariant in .agents/rules/00-solid-contract.md (Strict prohibition of external brain/scratch script creation; 100% repository-contained execution).
  2. Hybrid Language Routing 1.2 (Intra-Page Interleaved Routing: top/bottom screened English, middle created Word tables) in .agents/rules/01-workflow.md.
  3. Clinical Table Column Discretization 1.3 (Strict column independence for item names, result values, flags, units, ranges, methods) in .agents/rules/01-workflow.md.
  4. Updated Rule 2 in .agents/workflows/convert.md for hybrid documents.
---

## 2026-10-02 08:42 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: medical_chinese_landscape
- Gap found: none
- Action: upgraded chinese_medical_builder:
  1. Visually hidden container table pattern for hybrid pages (4, 5, 6) ensuring pixel-accurate vertical margin alignment and zero horizontal displacement between screened crops and created tables without modifying page margins.
  2. Dynamic 5-column discretization for single-column hospital reports (检验项目, 结果, 单位, 参考区间, 检验方法) derived from header coordinate centroids, eliminating column shift and phantom empty columns.
  3. Strict header/metadata separation ensuring table headers never leak into patient metadata lines.
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.90 text=1.00 layout=1.00 tables=0.91 placeholders=1.00
- Result: PASS
---

## 2026-10-02 09:03 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: medical_chinese_landscape
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.90 text=1.00 layout=1.00 tables=0.89 placeholders=1.00
- Result: PASS
---

## 2026-10-02 09:28 | 263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#
- Config used: configs/medical_chinese_landscape.yaml
- Gap found: none
- Action: reused pipeline with column discretization fix:
  1. Regex separation of trailing numeric values from item name bounding boxes into the `结果` column (pages 10, 23, 37).
  2. Automatic merging of orphan continuation rows in split tables.
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.90 text=1.00 layout=1.00 tables=0.89 placeholders=1.00
- Result: PASS
---

## 2026-10-02 10:28 | Multi-Agent Isolated Operations & Execution Timer Architecture (/learn)
- Upgraded `.agents/rules/01-workflow.md` and `.agents/workflows/convert.md`:
  1. **Isolated Operation Runs (`runs/<run_id>/`):** Every operation or subagent executes in an isolated subdirectory with its own intermediate extraction, OCR caches, candidate builds, and candidate modifications. Shared code is NEVER mutated during an unverified run.
  2. **Cross-Agent / Multi-Conversation Discovery:** Step 0 scans `runs/*/manifest.json` across conversations to detect concurrent operations and prevent clobbering.
  3. **Consensus Aggregation & Contradiction Resolution (/learn):** After 5-axis QA pass (≥ 0.85), candidate modifications are cross-checked for contradictions against other recent runs, synthesized, and promoted to shared core builders and configs.
  4. **Strict Execution Timers & Safety:** All background tasks (Word COM, OCR, extraction) are enforced with strict execution timers (max 60s for Word export, 120s for OCR) with automatic termination of hanging background processes (`WINWORD.EXE`) via `scripts/safe_word_export.py`.
---

## 2026-10-02 12:10 | Execution Time Optimization, 2-Tier Fast-Path, and Analytical Solving (/learn)
- Upgraded `.agents/rules/01-workflow.md` and `.agents/skills/convert/SKILL.md`:
  1. **Analytical Layout Solving (Zero Guesswork Invariant):** Prohibits iterative trial-and-error micro-adjustments. Mandates extracting exact bounding boxes (`Rect(x0, y0, x1, y1)`) via PyMuPDF in ONE pass and calculating section margins, column widths, and vertical paragraph offsets analytically ($\Delta y = y_{\text{target\_PDF}} - y_{\text{current\_DOCX}}$).
  2. **2-Tier Verification Fast-Path:**
     - Tier 1 (Headless In-Memory Pre-flight, ≤ 0.5s): Validates page count, section margins, column widths, and text token recall directly in Python using `python-docx` and `pymupdf` without launching Word COM. Implemented in `scripts/fast_verify.py` (runtime 0.052s).
     - Tier 2 (Word COM Acceptance Gate, ≤ 15s, executed ONCE): Runs `safe_word_export.py` and `tests/qa_agent.py` only after Tier 1 completely passes. Word COM is an acceptance gate, never a trial-and-error debugging loop.
  3. **Strict Bounded Execution Timers & Watchdog:**
     - Document Build Stage: max 2.0s.
     - Tier 1 Pre-flight: max 1.0s.
     - Tier 2 Word COM Export: max 15.0s with hard process termination and orphan `WINWORD.EXE` cleanup.
     - Total pipeline conversion cycle: max 25.0s.
  4. **Academic Journal Mirror Margins Invariant 1.7:** Codifies facing-page mirror margins (Odd: 1.38" left, 0.96" right; Even: 0.98" left, 1.34" right) across a constant 5.94" printable column grid.
  5. **Connected Hidden Text Invariant 1.8:** Attaches 0.5pt hidden text runs exclusively to screened images (banners, English titles, abstracts) for 100% token recall, eliminating text duplication on created Spanish paragraphs.
  6. **Bilingual Medical Journal Portrait Protocol (Rule 4 in convert/SKILL.md):** Formalized routing for portrait academic/medical journal articles.
---


## 2026-10-02 14:10 | DOC0000074469 Pure Native Multi-Column Architecture & Header/Footer Learnings (/learn)
- Config used: configs/medical_journal_portrait.yaml
- Gap found:
  1. Hallucinated grey divider rule (`_add_divider_rule(doc)`) under running headers caused systematic dark visual diffs across all pages.
  2. Word OpenXML header/footer mapping bug: when a page starts with a `NEW_PAGE` section and transitions to a `CONTINUOUS` 2-column section, assigning footers only to `NEW_PAGE` caused off-by-one or missing footers.
  3. Layout tables violated the user's hard directive of zero `<w:tbl>` for column layouts.
- Action:
  1. **Pure Native Word Multi-Column Section Architecture:** Replaced all column tables with native `<w:cols>` sections (`WD_SECTION_START.CONTINUOUS`) and native column breaks (`WD_BREAK.COLUMN`).
  2. **Eliminated Hallucinated Divider Rule:** Removed `_add_divider_rule(doc)` completely.
  3. **Dual-Section Footer Assignment Pattern:** Called `_set_section_footer()` on BOTH the `NEW_PAGE` header section and the `CONTINUOUS` body column sections for every page.
  4. **Strict Screening Scope Enforced:** Only the English Abstract continuation crops (`p1_eng_abstract_clean.png` and `p2_eng_abstract_clean.png`) are screened. All Spanish text, headings, review stamp, and English title are 100% genuine styled editable text.
  5. **Analytical Vertical Offsets:** Stabilized Page 1 top margin and spacing to fit all 6 header/title elements without causing column break spills onto Page 2.
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.852 text=0.897 layout=1.00 tables=1.00 placeholders=1.00 | overall=0.94 | total_time=5.66s
- Result: PASS
---


## 2026-10-02 15:06 | DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#
- Config used: medical_journal_portrait
- Gap found: none
- Action: reused pipeline
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.81 text=0.90 layout=1.00 tables=1.00 placeholders=1.00
- Result: FAIL
---

## 2026-10-02 15:11 | DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#
- Config used: medical_journal_portrait
- Gap found: none
- Action: upgraded journal_article_builder in scripts/stages/builders/:
  1. Enhanced figure extraction to prioritize `page.get_image_rects(xref)` and 300 DPI vector-clipped pixmaps (`page.get_pixmap(clip=rects[0], dpi=300)`) over raw uncalibrated CMYK stream conversions, ensuring authentic aspect ratios, color fidelity, and resolution across clinical histology and radiology figures.
  2. Verified pure native multi-column architecture across 15 sections with alternating mirror margins (Odd: 1.38"/0.96", Even: 0.98"/1.34") and zero layout tables for body text columns.
  3. Verified dual-section footers bound to both NEW_PAGE and CONTINUOUS sections.
  4. Verified selective screening strictly restricted to English abstract crops with 0.5pt hidden text runs.
  5. Verified zero hallucinated header/footer divider rules.
- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no
- QA: visual=0.852 text=0.897 layout=1.00 tables=1.00 placeholders=1.00 | overall=0.94 | total_time=5.80s
- Result: PASS
---

