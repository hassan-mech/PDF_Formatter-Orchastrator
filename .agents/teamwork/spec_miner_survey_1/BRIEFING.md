# BRIEFING — 2026-10-02T11:56:30Z

## Mission
Extract and document all precise contractual requirements, acceptance criteria, layout constraints, formatting invariants, tag lexicon rules, symbol sets, and QA pass thresholds for converting DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf to DOCX.

## 🔒 My Identity
- Archetype: Specification Miner
- Roles: Specification Mining Specialist, DTP & Translation Workflow Auditor
- Working directory: d:\Hassan\Amouna\Hasssan&Eman\.agents\teamwork\spec_miner_survey_1
- Original parent: 7f25a33e-af26-4a54-b55a-d1814880ab41
- Milestone: Milestone 1 — Specification Mining & Interface Contract Audit

## 🔒 Key Constraints
- Do NOT implement anything — read-only specification miner role.
- Prioritize authoritative sources (ORIGINAL_REQUEST.md, AGENTS.md, rules, skills, reference code, target PDF).
- Enforce SOLID contract (Section 0) and zero one-off custom builder script invariant.
- Probe pure native multi-column layout (<w:cols>, WD_SECTION_START.CONTINUOUS, WD_BREAK.COLUMN, zero <w:tbl> for text).
- Probe alternating mirror margins (Odd vs Even) and column grid.
- Probe dual-section footer mapping architecture (_set_section_footer() on both sections).
- Probe selective screening of English abstract continuation crops with 0.5pt hidden text runs.
- Probe Spanish text, headings, metadata, review stamp as genuine editable text.
- Probe span-level typography and vector drawings validation (no hallucinated rules).
- Probe 5-axis QA verification gate (tol=20, all axes >= 0.85, overall pass).

## Current Parent
- Conversation ID: 7f25a33e-af26-4a54-b55a-d1814880ab41
- Updated: 2026-10-02T11:56:30Z

## Task Summary
- **What to build**: Comprehensive specification report documenting contractual rules, geometry, layout, typography, tags, and QA thresholds for target PDF.
- **Success criteria**: Exhaustive enumeration of features, interface contracts, layout invariants, edge cases, and QA criteria in handoff.md.
- **Interface contracts**: scripts/core/interfaces.py, scripts/core/context.py, configs/medical_journal_portrait.yaml
- **Code layout**: scripts/core/, scripts/stages/builders/, configs/, tests/

## Key Decisions Made
- Used py -3.14 to inspect target PDF, vector drawings, fonts, and existing builds.
- Probed and verified pure native multi-column structure (<w:cols>), facing-page mirror margins (Odd vs Even), dual-section footer mapping, selective screening crops with 0.5pt hidden runs, span typography, drawing validation (no hallucinated rules), and 5-axis QA gate (tol=20, skip_axes: tables).
- Verified edge cases: Table mismatch when skip_axes is omitted, footer binding disconnection across continuous sections, and vertical budgeting on Page 1.

## Artifact Index
- .agents/teamwork/spec_miner_survey_1/DISPATCH.md — Task assignment
- .agents/teamwork/spec_miner_survey_1/progress.md — Liveness heartbeat
- .agents/teamwork/spec_miner_survey_1/handoff.md — Final deliverable report
- .agents/teamwork/spec_miner_survey_1/skills/convert.md — Local convert skill summary
- .agents/teamwork/spec_miner_survey_1/skills/table-builder.md — Local table-builder skill summary
- .agents/teamwork/spec_miner_survey_1/skills/tag-emitter.md — Local tag-emitter skill summary

## Loaded Skills
- **Source**: d:\Hassan\Amouna\Hasssan&Eman\.agents\skills\convert\SKILL.md
  - **Local copy**: d:\Hassan\Amouna\Hasssan&Eman\.agents\teamwork\spec_miner_survey_1\skills\convert.md
  - **Core methodology**: High-fidelity conversion of PDF to DOCX using config-driven pipeline, zero one-off scripts, hybrid screening, native multi-column layout, and 2-tier verification.
- **Source**: d:\Hassan\Amouna\Hasssan&Eman\.agents\skills\table-builder\SKILL.md
  - **Local copy**: d:\Hassan\Amouna\Hasssan&Eman\.agents\teamwork\spec_miner_survey_1\skills\table-builder.md
  - **Core methodology**: Real Word tables with explicit column widths on tblGrid and cells, XML shading fills, borders, padding in dxa, cantSplit row properties, and bidi support.
- **Source**: d:\Hassan\Amouna\Hasssan&Eman\.agents\skills\tag-emitter\SKILL.md
  - **Local copy**: d:\Hassan\Amouna\Hasssan&Eman\.agents\teamwork\spec_miner_survey_1\skills\tag-emitter.md
  - **Core methodology**: Canonical tags per Appendix A ([stamp:], [hw:], [signature], etc.) and Unicode characters per Appendix B without paraphrasing or ASCII substitutions.
