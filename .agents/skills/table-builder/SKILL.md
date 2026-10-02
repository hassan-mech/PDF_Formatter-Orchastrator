---
name: table-builder
description: Best practices, XML schema conventions, and python-docx patterns for generating pixel-faithful tables with preserved widths, fills, borders, and alignments.
---

# Table Builder Skill

## When to use
Use whenever extracting, constructing, or styling Word tables (`docx.Table`) from PDF sources or images.

---

## Core Rules

1. **Always Real Tables:** Never output tabular data as plain text, bullet points, or space/tab-separated paragraphs. Always create a genuine `docx.Table`.
2. **Explicit Column Widths:** Word auto-layout will break geometry unless column widths are explicitly defined on both `tblGrid` (`w:gridCol`) and every cell (`w:tcW`).
3. **Preserve Cell Fills:** Extract RGB background colors from PDF fills and apply via `parse_xml` (`<w:shd ... w:fill="HEX"/>`).
4. **Preserve Borders:** Custom borders (colors, weights, dashed/single) must be applied explicitly to avoid default Word thick gridlines or disappearing borders.
5. **No Orphan Rows:** Apply `<w:cantSplit/>` to rows to prevent awkward line breaks across page boundaries.
6. **Repeat Headers:** Apply `<w:tblHeader/>` to header rows on multi-page tables.
7. **RTL Support:** When Arabic text is present in the table, set table direction to Right-to-Left (`<w:bidiVisual/>`).

---

## Standard python-docx XML Snippets

### 1. Cell Background Shading (Fill)
```python
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_shading(cell, color_hex: str):
    """Sets background color of a table cell (e.g. '0070C0')."""
    clean_hex = color_hex.lstrip('#')
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{clean_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)
```

### 2. Cell Borders
```python
def set_cell_borders(cell, top="single", bottom="single", left="none", right="none", 
                     color="CCCCCC", sz="4"):
    """Applies specific borders to a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="{top}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:bottom w:val="{bottom}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:left w:val="{left}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:right w:val="{right}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)
```

### 3. Cell Padding / Margins
```python
def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Sets internal cell padding in dxa (1 pt = 20 dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>\n'
        f'  <w:top w:w="{top}" w:type="dxa"/>\n'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>\n'
        f'  <w:left w:w="{left}" w:type="dxa"/>\n'
        f'  <w:right w:w="{right}" w:type="dxa"/>\n'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)
```

### 4. Row Properties (cantSplit & tblHeader)
```python
def make_row_header(row):
    """Marks row as repeating header across pages."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

def prevent_row_split(row):
    """Prevents row from being split across page boundary."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
```

### 5. Table Right-to-Left (Bidi)
```python
def set_table_rtl(table):
    """Enables right-to-left column ordering for Arabic tables."""
    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
```
