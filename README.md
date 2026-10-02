# PDF → DOCX Conversion System

Reusable, SOLID-structured pipeline that converts any PDF (digital or scanned,
any language) into a Word document that faithfully matches the original and
emits the canonical tags/symbols defined in `.agents/rules/02-tags-symbols.md` and `.agents/skills/tag-emitter/SKILL.md`.

## One-time setup
- Governance rules: `.agents/rules/`
- Slash command: `/convert` (`.agents/workflows/convert.md`)
- Sub-agents: `.agents/agents/`
- Skills: `.agents/skills/`
- Baseline config: `configs/default.yaml`
- Logs: `DECISIONS.md`, `CHANGELOG.md`

## Daily workflow

1. Drop a PDF into `input/`.
2. Run:
   ```
   python scripts/convert_to_word.py -i input/<file>.pdf --config default --verify
   ```
   Scanned/Chinese pages are auto-routed to OCR by the pipeline.
3. QA runs automatically. All 5 axes must be ≥ 0.85.
4. On PASS:
   - DOCX → `output/<stem>.docx`
   - PDF  → `finished/<stem>.pdf`
   - Report → `tests/reports/<stem>_qa.json`
   - Line appended to `CHANGELOG.md`
5. On FAIL:
   - Inspect `renders/diffs/diff_p###.png`
   - Patch **only** `configs/default.yaml` (never the DOCX, never a script)
   - Re-run. Max 3 retries, then open a GAP entry in `DECISIONS.md`.

## Image / scanned report directly
```
python scripts/image_to_word.py -i renders/page_010.png -o output/parsed_page10.docx
```

## Rules you must never break
- No script named after an input file.
- No hardcoded margins/fonts/colors in Python.
- No editing `pipeline.py` to add a new document type — add a config.
- No hand-editing an `output/*.docx` — fix the config.

## Where things live
- `configs/`   → document-family settings
- `scripts/core/` → interfaces, registry, context, pipeline (no libraries)
- `scripts/stages/` → concrete implementations (fitz, docx, pdfplumber, rapidocr live here)
- `tests/` → QA agent
- `temp/<stem>/` → stage artifacts (JSON)
- `output/` → final DOCX
- `finished/` → archived PDFs after PASS
