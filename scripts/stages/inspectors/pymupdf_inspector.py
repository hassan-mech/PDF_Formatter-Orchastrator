"""
scripts/stages/inspectors/pymupdf_inspector.py — PyMuPDF Inspector Implementation

Inspects document structure, dimensions, orientations, text density, and image-only pages.
Registers with inspector_registry under name 'pymupdf'.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, List

import fitz  # PyMuPDF

# Ensure scripts is on sys.path
SCRIPTS_DIR = Path(__file__).resolve().parent.parent.parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from core.context import InputContext, InspectResult
from core.registry import inspector_registry


@inspector_registry.register("pymupdf")
class PyMuPdfInspector:
    """Inspector implementation using PyMuPDF (fitz)."""

    def inspect(self, ctx: InputContext) -> InspectResult:
        doc = fitz.open(str(ctx.pdf_path))
        page_count = doc.page_count
        page_dimensions: List[Dict[str, float]] = []
        image_only_pages: List[int] = []
        has_any_tables = False
        detected_chars: set[str] = set()

        for page_idx in range(page_count):
            page = doc[page_idx]
            rect = page.rect
            text = page.get_text("text").strip()
            images = page.get_images()

            is_img_only = (len(text) == 0 and len(images) > 0)
            if is_img_only:
                image_only_pages.append(page_idx + 1)

            # Sample characters for language hint
            if text:
                detected_chars.update(text[:200])

            # Quick table check via tabs/find_tables
            try:
                tables = page.find_tables()
                if tables and len(tables.tables) > 0:
                    has_any_tables = True
            except Exception:
                pass

            page_dimensions.append({
                "page": page_idx + 1,
                "width": rect.width,
                "height": rect.height,
                "orientation": "landscape" if rect.width > rect.height else "portrait",
            })

        # Basic language detection
        languages = []
        chars_str = "".join(detected_chars)
        if any('\u0600' <= c <= '\u06FF' for c in chars_str):
            languages.append("ar")
        if any('\u4e00' <= c <= '\u9fff' for c in chars_str):
            languages.append("zh")
        if any('\u0370' <= c <= '\u03FF' for c in chars_str):
            languages.append("el")
        if any('a' <= c.lower() <= 'z' for c in chars_str):
            languages.append("en")
        if not languages:
            languages = ["en"]

        meta = {
            "title": doc.metadata.get("title", ""),
            "author": doc.metadata.get("author", ""),
            "producer": doc.metadata.get("producer", ""),
        }
        doc.close()

        return InspectResult(
            page_count=page_count,
            page_dimensions=page_dimensions,
            image_only_pages=image_only_pages,
            has_tables=has_any_tables,
            detected_languages=languages,
            metadata=meta,
        )
