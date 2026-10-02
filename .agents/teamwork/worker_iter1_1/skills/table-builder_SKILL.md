---
name: table-builder
description: Best practices, XML schema conventions, and python-docx patterns for generating pixel-faithful tables with preserved widths, fills, borders, and alignments.
---

# Table Builder Skill Summary
1. Real Tables: Never output tabular data as plain text. Always create a genuine docx.Table.
2. Explicit Column Widths on tblGrid and tcW.
3. Preserve Cell Fills via parse_xml shading.
4. Preserve Borders explicitly.
5. No Orphan Rows: cantSplit.
6. Repeat Headers: tblHeader.
7. RTL Support: bidiVisual when needed.
