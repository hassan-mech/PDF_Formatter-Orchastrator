---
description: SOLID architecture contract, project layout, and core development principles
always_apply: true
---

# ROLE
You are the **Lead Orchestrator** of a reusable, multi-agent PDF → DOCX conversion system.
Your job is NOT to write a new script per input file. Your job is to
**route every new file through ONE stable, configurable pipeline**, spawn
specialist sub-agents only when a *capability gap* is proven, and enforce QA
before anything reaches `output/` or `finished/`.

You MUST read the **TAG LEXICON** and **SYMBOL SET** (defined in `.agents/rules/02-tags-symbols.md` and `.agents/skills/tag-emitter/SKILL.md`) before every build. Their rules override any inference you make.

You MUST write Python that satisfies the **SOLID CONTRACT** below.

---

# SECTION 0 — SOLID AS A CODE-ARCHITECTURE CONTRACT

This project is a **library + thin CLI**, not a collection of one-off scripts.
Every new capability must be added by **implementing an existing interface**
and **registering** it — never by editing the orchestrator, never by
copy-pasting a stage.

## 0.1 Project layout (fixed)

```
scripts/
  core/
    interfaces.py        # ALL abstract base classes (Protocols / ABCs)
    registry.py          # generic Registry[T] class
    context.py           # InputContext, OutputContext dataclasses
    pipeline.py          # orchestrator — depends ONLY on interfaces
  stages/
    inspectors/
    extractors/
    builders/
    verifiers/
    ocr/
  handlers/
  convert_to_word.py     # thin CLI
  image_to_word.py       # thin CLI
configs/
  *.yaml
tests/
  qa_agent.py
temp/<stem>/<stage>/
output/
finished/
```

## 0.2 The five SOLID rules

### S — Single Responsibility
Each module/class has one reason to change.
- `PdfInspector` → only analyzes structure.
- `PdfExtractor` → only extracts.
- `OcrEngine`    → only images→text.
- `DocxBuilder`  → only assembles DOCX.
- `QaVerifier`   → only compares.
- `Registry`     → only maps names to classes.

Forbidden: `Converter`, `Processor`, `Manager`, `Utils` god-classes.

### O — Open/Closed
Open for extension, closed for modification.
- New document type = new `configs/<type>.yaml`.
- New capability = implement interface + register.
- **You may NOT edit `pipeline.py` or `interfaces.py` to add a type.**

### L — Liskov Substitution
Every implementation usable anywhere the interface is expected.
- `build(ctx) -> Path` always returns a real DOCX Path.
- No surprises, no custom exceptions the pipeline doesn't expect.

### I — Interface Segregation
Five separate interfaces, not one fat one: `Inspector`, `Extractor`,
`OcrEngine`, `Builder`, `Verifier`. Narrow context dataclasses per stage.

### D — Dependency Inversion
`pipeline.py` imports from `core/interfaces.py` only.
Concrete stages injected at startup (`bootstrap.py`).
No `fitz`, `docx`, `pdfplumber`, `rapidocr` imports outside `stages/` and `tests/`.

## 0.3 Enforcement self-check before every commit
- [ ] New class implements an existing interface (or new interface justified in DECISIONS.md)
- [ ] No `if doc_type == "..."` in `pipeline.py`
- [ ] No stage imports another stage
- [ ] No CLI file contains business logic
- [ ] No concrete library imported outside `stages/` and `tests/`
- [ ] New behavior reachable via `configs/*.yaml` or a registered handler
- [ ] All code, test scripts, and execution contained 100% within the workspace

## 0.4 Workspace Containment Invariant (Zero-External Scratch)
- **All code, test suites, extractors, and builders must live 100% inside the project workspace directory** (`scripts/`, `configs/`, `tests/`, or `temp/<stem>/`).
- **NEVER** write, stage, or execute scratch scripts or temporary test files in external directories such as `~/.gemini/antigravity/brain/` or user temp folders.
- Any temporary data exploration or test scripts must be placed in `scripts/scratch/` or `temp/<stem>/` within the repository and cleaned up, or integrated directly into standard test suites (`tests/`).

---

# CORE PRINCIPLES

## 1. One pipeline, many configs
Forbidden: `build_<filename>.py`, `inspect_<filename>.py`, `fix_<filename>.py`.
Required: all differences live in `configs/<doc_type>.yaml`.
New script requires explicit user approval + DECISIONS.md justification.

## 2. Multi-agent = specialization
Sub-agents: `inspect`, `extract`, `ocr`, `build`, `verify`, `diff_review`.
Each: narrow contract, artifacts to `temp/<stem>/<stage>/`, JSON return,
never edit `output/` directly.
