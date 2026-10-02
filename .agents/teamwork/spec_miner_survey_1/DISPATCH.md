## 2026-10-02T11:43:58Z
You are Spec Miner 1 (spec_miner_survey_1).
Your working directory is: d:\Hassan\Amouna\Hasssan&Eman\.agents\teamwork\spec_miner_survey_1

MANDATORY: You MUST read the original user request at:
d:\Hassan\Amouna\Hasssan&Eman\.agents\teamwork\ORIGINAL_REQUEST.md
before starting your analysis.

Also inspect and read:
- d:\Hassan\Amouna\Hasssan&Eman\AGENTS.md (including Section 0 SOLID, Appendix A Tag Lexicon, Appendix B Symbol Set)
- d:\Hassan\Amouna\Hasssan&Eman\GEMINI.md
- d:\Hassan\Amouna\Hasssan&Eman\.agents\rules\01-workflow.md
- d:\Hassan\Amouna\Hasssan&Eman\.agents\rules\02-tags-symbols.md
- d:\Hassan\Amouna\Hasssan&Eman\.agents\skills\convert\SKILL.md
- d:\Hassan\Amouna\Hasssan&Eman\.agents\skills\table-builder\SKILL.md
- d:\Hassan\Amouna\Hasssan&Eman\.agents\skills\tag-emitter\SKILL.md

Task:
Extract and systematically document all precise contractual requirements, acceptance criteria, layout constraints, formatting invariants, tag lexicon rules, symbol sets, and QA pass thresholds for converting "DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf" to DOCX.
Pay special attention to:
1. Pure Native Multi-Column Layout Architecture (<w:cols>, WD_SECTION_START.CONTINUOUS, WD_BREAK.COLUMN, zero <w:tbl> for text).
2. Alternating mirror margins (Odd vs Even) and column grid.
3. Dual-section footer mapping architecture (_set_section_footer() on both sections).
4. Selective screening of English abstract continuation crops (p1_eng_abstract_clean.png and p2_eng_abstract_clean.png) with 0.5pt hidden text runs.
5. Spanish text, headings, metadata, review stamp as genuine editable text.
6. Span-level typography and vector drawings validation (no hallucinated rules).
7. 5-Axis QA Verification Gate (tol=20, all axes >= 0.85, overall pass).

Deliverable:
Write a comprehensive specification report to:
d:\Hassan\Amouna\Hasssan&Eman\.agents\teamwork\spec_miner_survey_1\handoff.md
Update your progress.md in your working directory as you work.
When done, send a completion message back to caller.
