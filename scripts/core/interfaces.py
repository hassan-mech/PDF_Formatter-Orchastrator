"""
scripts/core/interfaces.py — SOLID Abstract Interface Protocols

Defines the 5 core stage interfaces:
1. Inspector  — Analyzes document structure and characteristics.
2. Extractor  — Extracts text, images, and tables per page.
3. OcrEngine  — Performs OCR on image-only pages or crops.
4. Builder    — Assembles high-fidelity Word (.docx) documents from extracted context.
5. Verifier   — Runs 5-axis QA verification against the original document.

NO concrete third-party libraries (fitz, docx, rapidocr, etc.) may be imported here.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Protocol, runtime_checkable

from .context import BuildContext, ExtractResult, InputContext, InspectResult, QaReport


@runtime_checkable
class Inspector(Protocol):
    """Protocol for document inspection and structural categorization."""

    def inspect(self, ctx: InputContext) -> InspectResult:
        """Inspect the input document and return structural metadata."""
        ...


@runtime_checkable
class Extractor(Protocol):
    """Protocol for extracting text, tables, and render images from document."""

    def extract(self, ctx: InputContext, page_range: Optional[tuple[int, int]] = None) -> ExtractResult:
        """Extract content from pages within the input document."""
        ...


@runtime_checkable
class OcrEngine(Protocol):
    """Protocol for OCR processing of rasterized pages or images."""

    def ocr(self, images: List[Path], langs: List[str]) -> Dict[str, Any]:
        """Perform OCR on a list of image paths and return structured text/layout."""
        ...


@runtime_checkable
class Builder(Protocol):
    """Protocol for assembling a target Word (.docx) document."""

    def build(self, ctx: BuildContext) -> Path:
        """Build the Word document from BuildContext and return target docx Path."""
        ...


@runtime_checkable
class Verifier(Protocol):
    """Protocol for 5-axis quality assurance verification."""

    def verify(
        self,
        original_pdf: Path,
        docx_path: Path,
        report_path: Optional[Path] = None,
        dpi: int = 150,
        config: Optional[Dict[str, Any]] = None,
    ) -> QaReport:
        """Run QA checks comparing output DOCX against original PDF."""
        ...
