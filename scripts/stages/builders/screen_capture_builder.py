"""
scripts/stages/builders/screen_capture_builder.py — Full-Page Screen-Capture Builder
Generalised for ANY PDF. Implements the Builder protocol.

Approach
--------
1. Use python-docx's built-in `add_picture()` to embed images (handles all
   relationship/part wiring correctly).
2. After building, walk every section's XML and forcibly set:
   - pgSz to the PDF page dimensions
   - pgMar to 0 on all sides
   - The Drawing's <wp:extent> and <a:ext> to full-page EMU dimensions
   so the image fills 100% of the page with zero border.

This two-pass approach avoids all relationship-management bugs.

Registers under ``screen_capture_builder`` in ``builder_registry``.
"""

from __future__ import annotations

import io
import sys
from lxml import etree
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pymupdf

import docx
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

_SCRIPTS_DIR = Path(__file__).resolve().parents[3]
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from core.context import BuildContext
from core.registry import builder_registry


# ─────────────────────────────────────────────────────────────────────────────
# Namespace constants
# ─────────────────────────────────────────────────────────────────────────────
W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A   = "http://schemas.openxmlformats.org/drawingml/2006/main"

def _w(t):  return f"{{{W}}}{t}"
def _wp(t): return f"{{{WP}}}{t}"
def _a(t):  return f"{{{A}}}{t}"
def _twips(pt: float) -> str:
    return str(max(0, int(round(pt * 20))))
def _emu(inches: float) -> int:
    return int(round(inches * 914400))


# ─────────────────────────────────────────────────────────────────────────────
# Rendering helper
# ─────────────────────────────────────────────────────────────────────────────

def _render_page(page: pymupdf.Page, dpi: int, fmt: str, quality: int) -> bytes:
    zoom = dpi / 72.0
    mat  = pymupdf.Matrix(zoom, zoom)
    pix  = page.get_pixmap(matrix=mat, alpha=False, colorspace=pymupdf.csRGB)
    if fmt == "jpeg":
        return pix.tobytes("jpeg", jpg_quality=quality)
    return pix.tobytes("png")


# ─────────────────────────────────────────────────────────────────────────────
# Post-build fixup helpers
# ─────────────────────────────────────────────────────────────────────────────

def _fix_sectPr(sectPr: etree._Element, w_pt: float, h_pt: float) -> None:
    """Force exact page size and zero margins on a sectPr element."""
    # --- Page size ---
    for old in sectPr.findall(_w("pgSz")):
        sectPr.remove(old)
    pgSz = etree.SubElement(sectPr, _w("pgSz"))
    pgSz.set(_w("w"), _twips(w_pt))
    pgSz.set(_w("h"), _twips(h_pt))

    # --- Margins: all zero ---
    for old in sectPr.findall(_w("pgMar")):
        sectPr.remove(old)
    pgMar = etree.SubElement(sectPr, _w("pgMar"))
    for attr in ("top", "bottom", "left", "right", "header", "footer", "gutter"):
        pgMar.set(_w(attr), "0")


def _fix_drawing_size(drawing_el: etree._Element, w_pt: float, h_pt: float) -> None:
    """Pin the inline drawing to exact page dimensions (cx, cy in EMU)."""
    w_emu_s = str(_emu(w_pt / 72.0))
    h_emu_s = str(_emu(h_pt / 72.0))

    # wp:extent
    for ext in drawing_el.findall(_wp("inline") + "/" + _wp("extent")):
        ext.set("cx", w_emu_s)
        ext.set("cy", h_emu_s)
    # a:ext inside xfrm
    for ext in drawing_el.findall(".//" + _a("ext")):
        ext.set("cx", w_emu_s)
        ext.set("cy", h_emu_s)
    # wp:inline distT/B/L/R
    for inline in drawing_el.findall(_wp("inline")):
        inline.set("distT", "0")
        inline.set("distB", "0")
        inline.set("distL", "0")
        inline.set("distR", "0")


def _fix_para_spacing(para_el: etree._Element) -> None:
    """Zero paragraph spacing before/after."""
    pPr = para_el.find(_w("pPr"))
    if pPr is None:
        pPr = etree.Element(_w("pPr"))
        para_el.insert(0, pPr)
    for old in pPr.findall(_w("spacing")):
        pPr.remove(old)
    spacing = etree.SubElement(pPr, _w("spacing"))
    spacing.set(_w("before"), "0")
    spacing.set(_w("after"),  "0")
    spacing.set(_w("line"),   "240")
    spacing.set(_w("lineRule"), "auto")
    for old in pPr.findall(_w("ind")):
        pPr.remove(old)
    ind = etree.SubElement(pPr, _w("ind"))
    ind.set(_w("left"),  "0")
    ind.set(_w("right"), "0")


# ─────────────────────────────────────────────────────────────────────────────
# Builder
# ─────────────────────────────────────────────────────────────────────────────

@builder_registry.register("screen_capture_builder")
class ScreenCaptureBuilder:
    """
    Universal full-page screen-capture DOCX builder.

    Pass 1 — build:  For every PDF page, render to image and call
                      `doc.add_picture()` + `doc.add_section()` (python-docx
                      handles all image relationship wiring correctly).

    Pass 2 — fix:    Walk the document XML and:
                      - Set every sectPr to exact PDF page dimensions.
                      - Zero all margins on every sectPr.
                      - Pin every Drawing to full page EMU size.
                      - Zero paragraph spacing on every image paragraph.

    Works for any PDF: any page size, orientation, mixed per-page.
    """

    def build(self, ctx: BuildContext) -> Path:
        cfg         = ctx.config
        pdf_path    = ctx.input_context.pdf_path
        output_path = ctx.output_docx
        output_path.parent.mkdir(parents=True, exist_ok=True)

        cap     = cfg.get("capture", {})
        dpi     = int(cap.get("dpi", 300))
        img_fmt = str(cap.get("image_format", "png")).lower().strip()
        quality = int(cap.get("jpeg_quality", 95))

        pdf = pymupdf.open(str(pdf_path))
        n   = pdf.page_count
        print(f"[ScreenCaptureBuilder] {n} pages | DPI={dpi} | fmt={img_fmt}")

        # ── PASS 1: Build with python-docx ────────────────────────────────
        doc = docx.Document()

        # Remove all default paragraphs (keep doc structure intact)
        body = doc.element.body
        for child in list(body):
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag in ("p", "tbl"):
                body.remove(child)

        page_dims: List[Tuple[float, float]] = []  # (w_pt, h_pt) per page

        for idx in range(n):
            page  = pdf[idx]
            rect  = page.rect
            w_pt  = rect.width
            h_pt  = rect.height
            page_dims.append((w_pt, h_pt))

            print(f"  p{idx+1:02d}/{n} {w_pt:.0f}x{h_pt:.0f}pt ", end="", flush=True)
            img_bytes = _render_page(page, dpi, img_fmt, quality)
            print(f"{len(img_bytes)//1024}KB")

            # Add new section for every page after the first
            if idx > 0:
                doc.add_section()

            # Add paragraph with the image (python-docx handles relationships)
            para = doc.add_paragraph()
            run  = para.add_run()
            # Use width only here — we'll fix the height in Pass 2
            run.add_picture(io.BytesIO(img_bytes), width=Inches(w_pt / 72.0))

        pdf.close()

        # ── PASS 2: Fix page sizes, margins, and drawing dimensions ───────
        print("[ScreenCaptureBuilder] Pass 2: fixing page sizes and margins...")
        body = doc.element.body
        paragraphs = body.findall(_w("p"))

        # Collect all sectPrs in document order:
        # - sectPrs embedded in pPr (section breaks)
        # - The final doc-level sectPr at end of body
        all_sectPrs: List[Tuple[etree._Element, int]] = []  # (sectPr, page_idx)

        for p_idx, para_el in enumerate(paragraphs):
            pPr = para_el.find(_w("pPr"))
            if pPr is not None:
                s = pPr.find(_w("sectPr"))
                if s is not None:
                    all_sectPrs.append((s, p_idx))

        # The doc-level sectPr is for the LAST page
        doc_sectPr = body.find(_w("sectPr"))
        if doc_sectPr is not None:
            all_sectPrs.append((doc_sectPr, len(paragraphs) - 1))

        # Apply fixes to each sectPr
        for sect_el, para_idx in all_sectPrs:
            # Map to page index: pPr sectPrs break from the PREVIOUS page
            # The doc-level sectPr is always the last page
            page_i = min(para_idx, n - 1)
            w_pt, h_pt = page_dims[page_i]
            _fix_sectPr(sect_el, w_pt, h_pt)

        # Fix every paragraph's spacing and drawing size
        for p_idx, para_el in enumerate(paragraphs):
            _fix_para_spacing(para_el)
            w_pt, h_pt = page_dims[min(p_idx, n - 1)]
            for drawing in para_el.findall(".//" + _w("drawing")):
                _fix_drawing_size(drawing, w_pt, h_pt)

        # Force document defaults: zero page border on docDefaults
        # This prevents Word from applying Normal.dotm defaults
        doc_sectPr2 = body.find(_w("sectPr"))
        if doc_sectPr2 is not None:
            w_pt, h_pt = page_dims[-1]
            _fix_sectPr(doc_sectPr2, w_pt, h_pt)

        doc.save(str(output_path))
        sz_mb = output_path.stat().st_size / 1_048_576
        print(f"[ScreenCaptureBuilder] Saved → {output_path}  ({sz_mb:.1f} MB)")
        return output_path
