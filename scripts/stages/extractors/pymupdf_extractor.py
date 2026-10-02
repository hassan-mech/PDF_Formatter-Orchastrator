"""
scripts/stages/extractors/pymupdf_extractor.py — PyMuPDF Extractor Implementation

Extracts text, page images, and structured layout elements (paragraphs + tables in vertical reading order)
with geometric alignment detection (LEFT, CENTER, RIGHT).
Registers with extractor_registry under name 'pymupdf'.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import fitz  # PyMuPDF

SCRIPTS_DIR = Path(__file__).resolve().parent.parent.parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from core.context import ExtractResult, InputContext
from core.registry import extractor_registry


@extractor_registry.register("pymupdf")
class PyMuPdfExtractor:
    """Extractor implementation using PyMuPDF (fitz)."""

    def extract(self, ctx: InputContext, page_range: Optional[Tuple[int, int]] = None) -> ExtractResult:
        doc = fitz.open(str(ctx.pdf_path))
        stem = ctx.pdf_path.stem
        renders_dir = Path("renders") / stem
        renders_dir.mkdir(parents=True, exist_ok=True)

        text_by_page: Dict[int, str] = {}
        rendered_pages: Dict[int, str] = {}
        tables_by_page: Dict[int, List[Dict[str, Any]]] = {}
        page_elements: Dict[int, List[Dict[str, Any]]] = {}

        start_p = page_range[0] - 1 if page_range else 0
        end_p = page_range[1] if page_range else doc.page_count
        start_p = max(0, start_p)
        end_p = min(doc.page_count, end_p)

        dpi = ctx.dpi or 200
        zoom = dpi / 72.0
        matrix = fitz.Matrix(zoom, zoom)

        for p_idx in range(start_p, end_p):
            p_num = p_idx + 1
            page = doc[p_idx]
            page_w = page.rect.width
            page_mid = page_w / 2.0

            # 1. Text extraction
            text_by_page[p_num] = page.get_text("text")

            # 2. Render page to high-res image
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            img_path = renders_dir / f"page_{p_num:03d}.png"
            pix.save(str(img_path))
            rendered_pages[p_num] = str(img_path)

            # 3. Tables & Elements in reading order
            elements: List[Dict[str, Any]] = []
            extracted_tables: List[Dict[str, Any]] = []
            tb_boxes: List[List[float]] = []

            try:
                tabs = page.find_tables()
                if tabs and tabs.tables:
                    for t in tabs.tables:
                        t_data = t.extract()
                        table_dict = {
                            "type": "table",
                            "y0": t.bbox[1],
                            "bbox": list(t.bbox),
                            "row_count": t.row_count,
                            "col_count": t.col_count,
                            "data": t_data,
                        }
                        extracted_tables.append(table_dict)
                        elements.append(table_dict)
                        tb_boxes.append(list(t.bbox))
                    tables_by_page[p_num] = extracted_tables
            except Exception:
                pass

            # 4. Paragraph blocks outside tables with geometric alignment detection & rich span styles
            page_dict = page.get_text("dict")
            for b in page_dict.get("blocks", []):
                if b.get("type") != 0:  # 0 is text block
                    continue
                bx0, by0, bx1, by1 = b["bbox"]
                # Check overlap with any table bbox
                in_table = False
                for tx0, ty0, tx1, ty1 in tb_boxes:
                    if not (bx1 < tx0 or bx0 > tx1 or by1 < ty0 or by0 > ty1):
                        in_table = True
                        break
                if in_table:
                    continue

                spans_data = []
                full_lines = []
                for line in b.get("lines", []):
                    line_spans = []
                    for sp in line.get("spans", []):
                        stext = sp.get("text", "")
                        if not stext:
                            continue
                        flags = sp.get("flags", 0)
                        font_str = sp.get("font", "")
                        f_lower = font_str.lower()
                        is_bold = bool(flags & 16) or ("bold" in f_lower) or ("black" in f_lower) or ("heavy" in f_lower)
                        is_italic = bool(flags & 2) or ("italic" in f_lower) or ("oblique" in f_lower)
                        span_info = {
                            "text": stext,
                            "bold": is_bold,
                            "italic": is_italic,
                            "size": round(sp.get("size", 10.0), 1),
                            "color": sp.get("color", 0),
                            "font": font_str,
                        }
                        line_spans.append(span_info)
                        spans_data.append(span_info)
                    if line_spans:
                        full_lines.append("".join(s["text"] for s in line_spans))

                block_text = "\n".join(full_lines).strip()
                if not block_text:
                    continue

                b_mid = (bx0 + bx1) / 2.0
                b_width = bx1 - bx0
                if abs(b_mid - page_mid) < 30 and b_width < page_w * 0.75:
                    align = "CENTER"
                elif bx0 > page_mid and bx1 > page_w - 110:
                    align = "RIGHT"
                else:
                    align = "LEFT"

                elements.append({
                    "type": "paragraph",
                    "y0": by0,
                    "bbox": [bx0, by0, bx1, by1],
                    "alignment": align,
                    "text": block_text,
                    "spans": spans_data,
                })

            # Sort elements vertically
            elements.sort(key=lambda el: el["y0"])
            page_elements[p_num] = elements

        doc.close()

        return ExtractResult(
            text_by_page=text_by_page,
            rendered_pages=rendered_pages,
            tables_by_page=tables_by_page,
            metadata={"dpi": dpi, "stem": stem, "page_elements": page_elements},
        )
