# Project: DOC0000074469 Medical Journal PDF to DOCX Conversion

## Architecture
- **Model**: Reusable pipeline + configuration-driven stage dependency injection (SOLID Contract).
- **Core Orchestrator**: `scripts/core/pipeline.py` (Stage 1: Inspector -> Stage 2: Extractor -> Stage 3: OCR -> Stage 4: Builder -> Stage 5: Verifier).
- **Builder Stage**: `scripts/stages/builders/journal_article_builder.py` registered as `journal_article_builder`.
- **Configuration Profile**: `configs/medical_journal_portrait.yaml`.
- **Layout Engine**: Pure Native Word OpenXML multi-column sections (`<w:cols>`) with `WD_SECTION_START.CONTINUOUS` and `WD_BREAK.COLUMN`. Zero layout tables (`<w:tbl>`) for text columns.
- **Margin Engine**: Alternating mirror margins (Odd: 1.38" L, 0.96" R; Even: 0.98" L, 1.34" R; constant 5.94" column grid).
- **Footer Engine**: Dual-section footer mapping architecture calling `_set_section_footer()` on both `NEW_PAGE` and `CONTINUOUS` sections per page.
- **Routing Engine**: Selective screening strictly restricted to `p1_eng_abstract_clean.png` and `p2_eng_abstract_clean.png` with 0.5pt white hidden text runs (`<w:vanish/>`). All Spanish text and review stamp are 100% genuine editable text.
- **QA Verification**: 5-axis evaluator (`tests/qa_agent.py`) at 120 DPI with `tol=20`, `skip_axes: - tables`.

## Feature Inventory
| # | Feature | Description | Milestone | Source |
|---|---------|-------------|-----------|--------|
| 1 | Generic Pipeline Reuse | Execute conversion through existing configurable pipeline (`convert_to_word.py` / `configs/medical_journal_portrait.yaml`) | M1, M4 | survey (AGENTS.md, ORIGINAL_REQUEST R1) |
| 2 | Isolated Run Workspace | Dedicated execution folder `runs/DOC0000074469_<timestamp>/` with subdirs `extract/`, `ocr/`, `build/`, `qa/`, `modifications/` | M1, M4 | survey (AGENTS.md Invariant 4) |
| 3 | Pure Native Multi-Column | Native Word `<w:cols>` multi-column flow with continuous section breaks; zero layout tables (`<w:tbl>`) for text | M2 | survey (ORIGINAL_REQUEST R2) |
| 4 | Alternating Mirror Margins | Odd pages (1.38" L, 0.96" R); Even pages (0.98" L, 1.34" R) matching book spread binding | M2 | survey (ORIGINAL_REQUEST R2) |
| 5 | Constant Printable Grid | Uniform 5.94" printable column grid across both Odd and Even pages ($2 \times 2.85''$ + $0.24''$ gutter) | M2 | survey (ORIGINAL_REQUEST R2) |
| 6 | Dual-Section Footer Mapping | Explicit footer assignment on both `NEW_PAGE` header section and `CONTINUOUS` body section per page | M2 | survey (ORIGINAL_REQUEST R2, Invariant 6) |
| 7 | Selective Screening Scope | Screen ONLY English abstract continuation crops (`p1_eng_abstract_clean.png` and `p2_eng_abstract_clean.png`) | M1, M3 | survey (ORIGINAL_REQUEST R3) |
| 8 | 0.5pt Hidden Text Runs | Attach 0.5pt white hidden text runs exclusively to screened crops for 100% token recall | M3 | survey (ORIGINAL_REQUEST R3) |
| 9 | Genuine Spanish Editable Text | All Spanish titles, headings, metadata, body text, and references rendered as editable Word runs | M2 | survey (ORIGINAL_REQUEST R3) |
| 10 | Digital Review Stamp | Plain centered Arial text for top metadata banner (`DOC0000074469 Reviewed by TP: 30SEP2026 07:19AM CET`) | M2 | survey (ORIGINAL_REQUEST R3) |
| 11 | Span-Level Typography | Auto-extract bold, italic, font size, and color from PDF text spans (`page.get_text("dict")`) | M2 | survey (ORIGINAL_REQUEST R4) |
| 12 | Crimson Title Divider Rule | Thin red horizontal rule under Spanish article title (2.0pt `#DB1D43`, width 4.13") | M1 | survey (PDF Page 1 Drawing 0) |
| 13 | Anti-Hallucination Drawing Validation | Validate vector drawings via `page.get_drawings()`; zero hallucinated running header divider rules | M1 | survey (ORIGINAL_REQUEST R4, Invariant 7) |
| 14 | Bordered Announcement Box Card | Real bordered Word table for Page 7 `AVISO IMPORTANTE` card with `#EBF2F7` fill and `#005284` borders | M2 | survey (PDF Page 7 drawings) |
| 15 | 2-Tier Verification Fast-Path | Tier 1 headless pre-flight (< 0.5s) followed by Tier 2 Word COM acceptance gate (<= 15s) | M3 | survey (AGENTS.md Speed Rules) |
| 16 | 5-Axis QA Verification Gate | Comprehensive verification across Visual, Text, Layout, Tables, Placeholders with `tol=20`, `skip_axes: - tables` | M4 | survey (ORIGINAL_REQUEST A3) |
| 17 | Canonical Appendix A Tags | Verbatim tag emission syntax without paraphrasing or empty payloads | M2 | survey (AGENTS.md Appendix A) |
| 18 | Canonical Appendix B Symbols | Unicode glyphs for bullets, arrows, fractions, and scientific notation | M2 | survey (AGENTS.md Appendix B) |

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|------|-------|-------------|--------|
| 1 | M1: Workspace Isolation & Asset Pipeline | Set up isolated run directory `runs/DOC0000074469_<timestamp>/` (`extract/`, `ocr/`, `build/`, `qa/`, `modifications/`), verify crop generation for abstract continuation, and audit vector drawing stream against hallucinated rules. | none | PLANNED |
| 2 | M2: Native Multi-Column Layout & Typography | Build DOCX via `journal_article_builder.py` ensuring 15 native sections, pure `<w:cols>`, mirror margins (Odd 1.38"/0.96", Even 0.98"/1.34"), dual-section footer binding, and genuine editable Spanish typography & review stamp. | M1 | PLANNED |
| 3 | M3: Selective Screening & Token Recall | Verify selective screening is restricted to `p1_eng_abstract_clean.png` and `p2_eng_abstract_clean.png`, attach 0.5pt hidden text runs, and execute Tier 1 pre-flight verification (< 0.5s). | M2 | PLANNED |
| 4 | M4: 5-Axis QA Verification & Promotion | Run Tier 2 5-Axis QA (`tests/qa_agent.py`) with `tol=20`, verify all axes >= 0.85, promote final DOCX to `output/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.docx`, and record promotion. | M3 | PLANNED |

## Interface Contracts
### `scripts/convert_to_word.py` CLI ↔ `scripts/core/pipeline.py`
- Inputs: `--input <pdf_path>`, `--output <docx_path>`, `--config <config_name>`, `--dpi 120`.
- Behavior: Loads config dictionary from `configs/<config_name>.yaml`, invokes `bootstrap.build_default_pipeline(config)`, constructs `InputContext`, calls `pipeline.run(ctx)`.
- Exit code: 0 on success, non-zero on QA failure.

### `scripts/core/interfaces.py` (Builder Protocol) ↔ `journal_article_builder.py`
- Signature: `build(ctx: BuildContext) -> Path`
- `BuildContext` contains: `input_context.input_pdf`, `output_docx`, `config`.
- Outputs: Returns absolute path to created `.docx` file.

### `safe_word_export.py` ↔ `tests/qa_agent.py`
- Execution: Headless Word COM converts output DOCX to temporary rendered PDF.
- Evaluation: PyMuPDF compares rendered PDF with reference original PDF at 120 DPI using anti-aliasing threshold `tol=20`.
- Report: Generates JSON report containing scores for Visual, Text, Layout, Tables, Placeholders, and overall boolean `overall_pass`.

## Code Layout
- `scripts/core/`: Abstract interfaces, context dataclasses, registry, and pipeline orchestrator.
- `scripts/stages/builders/journal_article_builder.py`: Specialized 7-page bilingual medical journal builder.
- `configs/medical_journal_portrait.yaml`: Pipeline configuration profile.
- `scripts/convert_to_word.py`: Thin CLI entry point.
- `scripts/fast_verify.py`: Tier 1 in-memory pre-flight verifier.
- `tests/qa_agent.py`: Tier 2 5-axis acceptance QA evaluator.
- `runs/DOC0000074469_<timestamp>/`: Isolated workspace per conversion execution.
- `output/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.docx`: Final promoted deliverable.
