"""
scripts/stages/extractors/table_extractor.py — Specialized Table Extractor

Extracts detailed table structures, bounding boxes, cell coordinates, and text matrices.
Registers with extractor_registry under name 'table_extractor'.
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


@extractor_registry.register("table_extractor")
class TableExtractor:
    """Specialized table extractor using PyMuPDF table detection."""

    def extract(self, ctx: InputContext, page_range: Optional[Tuple[int, int]] = None) -> ExtractResult:
        doc = fitz.open(str(ctx.pdf_path))
        stem = ctx.pdf_path.stem
        tables_by_page: Dict[int, List[Dict[str, Any]]] = {}
        text_by_page: Dict[int, str] = {}
        rendered_pages: Dict[int, str] = {}

        start_p = page_range[0] - 1 if page_range else 0
        end_p = page_range[1] if page_range else doc.page_count

        for p_idx in range(start_p, end_p):
            p_num = p_idx + 1
            page = doc[p_idx]
            text_by_page[p_num] = page.get_text("text")

            try:
                tabs = page.find_tables()
                if tabs and tabs.tables:
                    page_tables = []
                    for t in tabs.tables:
                        cells = []
                        for row in t.extract():
                            cells.append([str(c or "").strip() for c in row])
                        page_tables.append({
                            "bbox": list(t.bbox),
                            "row_count": t.row_count,
                            "col_count": t.col_count,
                            "data": cells,
                        })
                    if page_tables:
                        tables_by_page[p_num] = page_tables
            except Exception:
                pass

        doc.close()

        return ExtractResult(
            text_by_page=text_by_page,
            rendered_pages=rendered_pages,
            tables_by_page=tables_by_page,
            metadata={"strategy": "table_focused", "stem": stem},
        )
