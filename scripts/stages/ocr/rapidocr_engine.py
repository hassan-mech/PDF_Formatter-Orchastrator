"""
scripts/stages/ocr/rapidocr_engine.py — RapidOCR Engine Implementation

Performs OCR on raster images / scanned pages using RapidOCR ONNX runtime.
Supports multilingual OCR (English, Chinese, Arabic numbers/symbols).
Registers with ocr_registry under name 'rapidocr'.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, List

SCRIPTS_DIR = Path(__file__).resolve().parent.parent.parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from core.registry import ocr_registry

try:
    from rapidocr_onnxruntime import RapidOCR
    HAS_RAPIDOCR = True
except ImportError:
    HAS_RAPIDOCR = False


@ocr_registry.register("rapidocr")
class RapidOcrEngine:
    """OCR Engine implementation using RapidOCR."""

    def __init__(self):
        self._engine = RapidOCR() if HAS_RAPIDOCR else None

    def ocr(self, images: List[Path], langs: List[str]) -> Dict[str, Any]:
        results: Dict[str, Any] = {}
        if not HAS_RAPIDOCR or not self._engine:
            print("  ⚠️  RapidOCR not available; OCR stage will return empty results.")
            return results

        for img_path in images:
            p = Path(img_path)
            if not p.exists():
                continue

            try:
                ocr_result, elapse = self._engine(str(p))
                page_items = []
                if ocr_result:
                    for item in ocr_result:
                        # item format: [dt_boxes, text, score]
                        bbox, text, score = item
                        page_items.append({
                            "bbox": bbox,
                            "text": str(text).strip(),
                            "score": float(score),
                        })
                results[p.name] = {
                    "path": str(p),
                    "items": page_items,
                    "elapse": elapse,
                }
            except Exception as exc:
                print(f"  ⚠️  Error running RapidOCR on {p.name}: {exc}")
                results[p.name] = {"error": str(exc), "items": []}

        return results
