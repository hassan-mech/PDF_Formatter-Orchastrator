---
description: Standard conversion workflow, QA gate requirements, anti-patterns, and decision protocols
always_apply: true
---

# STANDARD WORKFLOW & QA GATE

## QA Gate Absolute
- No `input/` → `finished/` without `tests/reports/<stem>_qa.json`
  with `overall_pass == true` and every axis ≥ 0.85.
- Visual < 0.85 → halt, inspect `renders/diffs/diff_p###.png`, fix config, re-run.
- Styling errors = config bugs, not one-off fixes.

## Tags and Symbols Contractual
Every tag matches Appendix A verbatim (`.agents/skills/tag-emitter/SKILL.md`). Every symbol matches Appendix B.
Deviation fails QA axis Placeholders.

---

# DTP & TRANSLATION WORKFLOW INVARIANTS

### 1. Hybrid Language Routing (Screened vs. Created)

#### 1.1 Full-Page Routing
- **Administrative / Non-Translatable Source Content (e.g. English Questionnaires, Logos, Banners):**
  Must be **screened** (embedded as high-resolution 300 DPI page snapshots) to preserve 100% corporate branding, icons, and fixed layouts.
- **Translatable Foreign Content (e.g. Chinese / Greek Hospital Lab Reports & Medical Records):**
  Must be **created** as genuine, editable Word tables and text (`docx.Table`), enabling translators to edit and translate terms.
- NEVER invert this by screening translatable foreign pages and typing administrative questionnaires.

#### 1.2 Intra-Page Hybrid Routing (Interleaved Content / Mixed Pages)
When an administrative page contains an embedded foreign lab report or clinical scan (e.g. an English questionnaire with an embedded Chinese lab report scan in the middle):
- **Top Administrative Section:** Crop from the source PDF at 300 DPI (`clip=Rect(...)`) and embed as an exact image snapshot.
- **Middle Foreign Section:** Reconstruct into an editable, genuine Word table (`docx.Table`) matching the hospital report format.
- **Bottom Administrative Section:** Crop from the source PDF at 300 DPI and embed directly below the reconstructed table.
- **Hidden Text Requirement:** Embed a 0.5pt hidden white text run of the full page's extractable text so that the document retains 100% token recall and full searchability for screen readers and QA checks.

#### 1.3 Clinical Table Column Discretization
- **Strict Column Independence:** Clinical report tables must discretize each field into its own column:
  - `序号` (Index / Number)
  - `检验项目` (Test Item Name)
  - `结果` (Result Value) — **NEVER** concatenate numeric results (e.g. `20.76`, `12.64`) into the item name cell.
  - `提示` (Flag: `↓`, `↑`)
  - `单位` (Unit)
  - `参考区间` (Reference Interval)
  - `检验方法` (Methodology)
- **Cell Continuation Merging:** If OCR splits a cell across lines (e.g. `ml/ (min-` on line 1 and `1.73m^2)` on line 2), merge the continuation string back into the existing cell; **NEVER** create an empty orphaned table row.

#### 1.4 Table Header-Value Horizontal Alignment
- **Coordinated Column Alignment:** In all clinical and data tables, cells within a column must share coordinated horizontal alignment with their header:
  - If the header `结果` / `提示` / `单位` is centered, all data cells below it MUST be centered (`WD_ALIGN_PARAGRAPH.CENTER`).
  - If the header is left-aligned, data cells MUST be left-aligned.
  - **NEVER** right-align numeric values against the right cell border when the header is left-aligned or centered, as this displaces the values away from the Chinese header character.
  - Flag symbols (`↓`, `↑`, `★`) and status text (`未检出`, `检出`) must sit directly aligned with the column title.

#### 1.5 Clinical Document Font Scale Hierarchy
Reconstructed foreign clinical documents (lab reports, CT/MRI scan sheets) must match the natural typographic scale and visual flow of the original document:
- **Hospital Title & Department Header:** 12.0 pt – 14.0 pt bold (`SimSun` or appropriate document font).
- **Patient & Sample Metadata:** 8.5 pt – 9.5 pt regular.
- **Table Headers & Data Cells:** 8.5 pt – 9.5 pt (minimum 8.0 pt for high-density 8-column tables). NEVER shrink table text down to 6.0–7.0 pt.
- **Report Footers, Timestamps & Notes:** 7.5 pt – 8.0 pt.

#### 1.6 Embedded Clinical Scan Coverage (Universal Medical Report Extraction)
- Any page in an administrative questionnaire that embeds a foreign medical scan—including **CT reports, MRI reports, ultrasound, pathology, cytology, or lab test sheets**—must NOT be treated as a pure image screen.
- The foreign clinical scan section must be **reconstructed into genuine Word text and tables** to allow complete translation and searchability.
- The administrative questionnaire sections above and/or below the scan must be cropped and embedded via the layout container table.

#### 1.7 Academic Journal Facing-Page Mirror Margins (Odd vs. Even Layouts)
For book spreads and academic/medical journal articles (e.g. *Dermatología Revista mexicana*):
- **Odd Pages (1, 3, 5, 7):** Right-hand page of spread.
  - Inside binding margin is on the LEFT: `left_margin = Inches(1.38)` (~99.2 pt).
  - Outside margin is on the RIGHT: `right_margin = Inches(0.96)` (~68.5 pt).
- **Even Pages (2, 4, 6):** Left-hand page of spread.
  - Outside margin is on the LEFT: `left_margin = Inches(0.98)` (~70.6 pt).
  - Inside binding margin is on the RIGHT: `right_margin = Inches(1.34)` (~96.4 pt).
- **Grid Invariant:** Both Odd and Even pages share identical printable column widths ($2 \times 2.85''$ with $0.24''$ gutter = $5.94''$ grid).
- **Implementation:** Configure section geometry using `WD_SECTION_START.NEW_PAGE` per page rather than arbitrary soft page breaks.

#### 1.8 Connected Hidden Text Invariant for CAT / QA Alignment
- When screening administrative banners, logos, or English titles/abstracts in a bilingual foreign document:
  - Attach 0.5pt white hidden text runs (`font.hidden = True`, `font.size = Pt(0.5)`) strictly to the screened container paragraph.
  - **NEVER** duplicate the entire page's text into hidden runs when foreign content is already created as editable Word text. Duplicate text causes word-count bloat and line-wrap artifacts.

#### 1.9 Pure Native Multi-Column Architecture (Zero Layout Tables for Text Columns)
- Multi-column text flows (academic papers, journals, articles) MUST use Word's native multi-column section architecture (`<w:cols>`) configured via `WD_SECTION_START.CONTINUOUS` and native column breaks (`WD_BREAK.COLUMN`).
- Using invisible or borderless layout tables (`<w:tbl>`) for multi-column body text is strictly prohibited. Tables are reserved strictly for tabular data cards and enclosed forms.

#### 1.10 Dual-Section Word Header & Footer Mapping Architecture
- In Word OpenXML, when a page starts with a `NEW_PAGE` section (for 1-column headers or banners) and transitions into a `CONTINUOUS` 2-column body section, Word assigns headers and footers based on the section spanning the page boundary.
- To eliminate missing or off-by-one footers, ALWAYS explicitly call `_set_section_footer()` on BOTH the `NEW_PAGE` top section AND the `CONTINUOUS` body column sections for that page.

#### 1.11 Verification Against Hallucinated Graphical Accents
- Never assume or inject decorative lines (such as running header divider rules, separator borders, or frame outlines) without confirming their presence in the source PDF vector stream.
- Hallucinated rules produce severe visual diff penalties across every page. Validate vector drawings via PyMuPDF `page.get_drawings()` before adding lines.

#### 1.12 Zero Custom Code Invariant & Reuse Penalty (Strict Negative Points)
- **STRICT PENALTY:** Writing standalone, one-off custom builder scripts per document (e.g. hardcoding text strings and layout logic into separate python files) is strictly penalized with negative evaluation points.
- **MANDATORY REUSE:** Work with existing core pipeline components (`scripts/stages/extractors/`, `scripts/stages/builders/`, `scripts/stages/inspectors/`).
- **GENERIC ENHANCEMENT ONLY:** If an existing stage lacks a required capability (such as span-level bold detection or multi-column section handling), enhance the generic component in `scripts/stages/` so that all current and future documents benefit.

#### 1.13 Span-Level Typography & Style Extraction Invariant
- Text styling (bold, italic, font family, font size, font color) must NEVER be guessed, hardcoded, or flattened to plain unstyled text.
- The extractor must inspect span dictionaries (`page.get_text("dict")`):
  - **Bold:** `flags & 16 != 0` or `"bold"` in `span["font"].lower()` -> `run.bold = True`.
  - **Italic:** `flags & 2 != 0` or `"italic"` / `"oblique"` in `span["font"].lower()` -> `run.italic = True`.
  - **Font Size:** `span["size"]` -> `run.font.size = Pt(size)`.
  - **Font Color:** `span["color"]` -> convert sRGB integer to `RGBColor(r, g, b)`.
- Mixed-style sentences (e.g. regular text containing bold titles like `"Ahora puede descargar la aplicación de **Dermatología Revista Mexicana.**"`) must be emitted as distinct, styled runs within the same paragraph.

#### 1.14 Strict Appendix A Logo & Icon Governance
- `[logo: xxxxxxxxxxx]` MUST have a non-empty readable text payload.
- **FORBIDDEN:** Emitting `[logo:]` with an empty payload or splitting logo tags across artificial table cells.
- If a graphic logo has NO readable text (e.g. app store badges, social media icons, brand emblems without text), emit `[icon]` or `[emblem:]`.

#### 1.15 Multi-File Operation Folder Isolation
- In batch or multi-file operations, EVERY file MUST have its own isolated directory: `runs/<stem>_<timestamp>/` containing dedicated subdirectories:
  - `extract/`
  - `ocr/`
  - `build/`
  - `qa/`
  - `modifications/`
- Mixing outputs or intermediate state across different files in a shared unisolated directory is strictly prohibited.

#### 1.16 Mandatory Teamwork & Collaborative Sub-Agent Delegation
- At the start of every conversion session, initialize collaborative teamwork using the dedicated specialist subagents:
  1. `inspector`: Structure, page geometry, font discovery.
  2. `extractor`: Text, span-level styles, and table extraction.
  3. `ocr-engine`: Concurrent OCR on raster scans and stamps.
  4. `docx-builder`: Config-driven assembly of `.docx`.
  5. `qa-verifier`: Independent 5-axis QA gate verification.
- Always recommend `/teamwork-preview` (or Plan Mode) to the user when starting a complex or multi-page session to leverage concurrent execution.




### 2. Form Box and Border Invariants
- Any form section enclosed in borders in the source PDF must be constructed as a **real bordered Word table** (`cell_borders: single`), with proper cell padding and leader dashes sized to fit the printable area.
- NEVER dump boxed form fields as plain unbordered paragraphs with dashes.
- Table headers must match source fill colors; NEVER apply default colored fills (e.g. `#0070C0`) to white tables.

### 3. Appendix A & B Enforcement in Medical Forms
- Medical flags: `↓` (U+2193), `↑` (U+2191), `★` (U+2605), `■` (U+25A0), `□` (U+25A1).
- Doctor & validator signatures: emit `[signature]`, stamps as `[stamp: xxxx]`, handwritten entries as `[hw: xxxx]`.

---

# STANDARD WORKFLOW & MULTI-AGENT ISOLATION

### Step 0 — Cross-Conversation & Reuse Check (FIRST)
1. **Multi-Agent Discovery:** Check `runs/` for active or recently finished runs from other conversations or subagents.
   - Read existing `runs/*/manifest.json` to detect concurrent work on the same file or shared modules.
   - Ensure operations do NOT clobber each other.
2. **Reuse Check:**
   - List `configs/*.yaml` and `scripts/stages/**/*.py`.
   - Can existing config + stages handle this file?
   - YES → select best config.
   - NO → log `GAP: <one line>` in DECISIONS.md (new YAML or interface implementation with user approval).

### Step 1 — Workspace Isolation per Operation
- Every operation/conversation MUST create its own isolated run folder: `runs/<stem>_<timestamp>/` (or `runs/<operation_id>/`).
- **Isolation Invariant:** All intermediate artifacts (`inspect.json`, `extract/`, `ocr/`, candidate DOCX, and experimental patches) remain strictly inside `runs/<operation_id>/`.
- **Zero In-Place Mutation:** NEVER mutate shared core code (`scripts/stages/builders/`, `configs/`, `.agents/rules/`) during an unverified candidate run.

### Step 2 — Timed Execution & Monitoring
- **Mandatory Execution Timers:** Every external command, background task, and Word COM export MUST be bounded by a strict execution timer (max 60s for Word export, max 120s for OCR).
- **Auto-Kill on Timeout:** If an execution exceeds its timer, terminate it immediately and auto-kill any lingering background processes (e.g. `WINWORD.EXE`) to avoid deadlocks.

### Step 3 — Build Candidate in Run Folder
- Generate the candidate document inside `runs/<operation_id>/build/<stem>.docx`.
- If candidate builder/config modifications were made, record them under `runs/<operation_id>/modifications/`.

### Step 4 — Verify (5 axes ≥ 0.85)
- Run QA agent comparing original PDF vs candidate DOCX inside `runs/<operation_id>/qa/`.
- All 5 axes (Visual, Text, Layout, Tables, Placeholders) must achieve ≥ 0.85.

### Step 5 — Consensus Aggregation & Code Promotion (/learn)
- When the operation achieves a full PASS:
  1. **Contradiction Check:** Compare modifications in `runs/<operation_id>/modifications/` against other recent runs in `runs/*/manifest.json`.
     - Detect and resolve any contradictory rules (e.g. conflicting column boundaries, incompatible margin rules).
     - Generalized rules always take precedence over brittle one-off hacks.
  2. **Code Promotion:** Apply verified, contradiction-free enhancements into:
     - Shared core builders (`scripts/stages/builders/`)
     - Shared configs (`configs/`)
     - Shared project rules (`.agents/rules/`) and workflows (`.agents/workflows/`)
  3. **Artifact Promotion:** Promote final DOCX to `output/<stem>.docx`, move source PDF to `finished/`, and record to `CHANGELOG.md` and `DECISIONS.md`.
  4. **Run Archive:** Update `runs/<operation_id>/manifest.json` to `status: promoted`.

---

# ANTI-PATTERNS (auto-reject)
- Script named after an input file
- Hardcoded margins/fonts/colors/page ranges in Python
- `if doc_type == "..."` in `pipeline.py`
- A stage importing another stage
- Business logic in a CLI file
- `Converter`/`Manager`/`Utils` god-class
- Skipping QA
- Invented tag syntax (`[sig]`, `<stamp>`, `[stamp-here]`)
- Tables as plain paragraphs
- Dropping image-only pages
- Hand-editing `output/*.docx`
- Adding a new sub-agent when config suffices

---

# SUB-AGENT CONTRACTS
- `inspector(InputContext) -> InspectResult`
- `extractor(InputContext, page_range) -> ExtractResult`
- `ocr(list[Path], lang) -> OcrResult`
- `builder(ExtractResult, Config) -> Path`
- `verifier(Path, Path, dpi) -> QaReport`
- `diff_reviewer(list[Path]) -> DiffReview`

All dataclasses, JSON-serializable, implement ABCs in `core/interfaces.py`.

---

# DECISION LOG PROTOCOL
`DECISIONS.md` is append-only:
```
## <YYYY-MM-DD HH:MM> | <stem>
- Config used: <name>
- Gap found: <one line or "none">
- Action: reused pipeline | added config | added handler (approved by user)
- SOLID check: interfaces-implemented=<list> | pipeline-edited=no
- QA: visual=<x> text=<x> layout=<x> tables=<x> placeholders=<x>
- Result: PASS | FAIL(retry n)
```

---

# SPEED RULES (STRICT EXECUTION TIME ENFORCEMENT)
1. **Parallelize stages** per page-range with `--workers N`.
2. **Skip stages** with unchanged config hash.
3. **Cache OCR** on `(image_hash, lang, engine)`.
4. **Analytical Layout Solving (Zero Guesswork Invariant):**
   - **PROHIBITED:** Iterative trial-and-error micro-adjustments of margins, spacings, or font sizes.
   - **MANDATORY:** Extract exact bounding boxes (`Rect(x0, y0, x1, y1)`) via PyMuPDF in ONE pass before writing builder code.
   - Calculate all section margins, table widths, and vertical paragraph offsets analytically:
     $$\Delta y = y_{\text{target\_PDF}} - y_{\text{current\_DOCX}}$$
   - Set all layout parameters in a single deterministic pass.
5. **2-Tier Verification Fast-Path:**
   - **Tier 1 — Headless In-Memory Pre-flight (≤ 0.5s):**
     Validate page count, section margins, column widths, and text token recall directly in Python using `python-docx` and `pymupdf` *without* launching Word or converting to PDF.
     - If Tier 1 fails (e.g. missing tokens, wrong page count, overflow paragraph): fix immediately in code.
   - **Tier 2 — Word COM Acceptance Gate (≤ 10s, executed ONCE):**
     Run `safe_word_export.py` and `tests/qa_agent.py` ONLY when Tier 1 completely passes.
     - Word COM is an acceptance gate, NEVER an exploratory trial-and-error tool.
6. **Bounded Execution Timers & Watchdog:**
   - Document Build Stage: max **2.0s**.
   - Tier 1 Pre-flight: max **1.0s**.
   - Tier 2 Word COM Export: max **15.0s** with hard process termination and orphan `WINWORD.EXE` cleanup.
   - Total pipeline conversion cycle: max **25.0s**.

---

# RESPONSE FORMAT
1. Reuse verdict
2. SOLID check
3. Plan
4. Execution
5. QA result
6. Promotion
7. Gap log

---

# FIRST ACTION EVERY SESSION
1. Read Appendix A and Appendix B (`.agents/rules/02-tags-symbols.md` / `.agents/skills/tag-emitter/SKILL.md`)
2. Read last 20 lines of `DECISIONS.md` and `CHANGELOG.md`
3. List `configs/` and `scripts/stages/`
4. Then accept the new input file
