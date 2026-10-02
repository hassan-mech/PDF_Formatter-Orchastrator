"""
scripts/core/context.py — Context and Data Transfer Objects (DTOs)

Dataclasses for communication between pipeline stages:
- InputContext: Context defining the input file, configuration, parameters, and paths.
- InspectResult: Findings from inspection (page count, orientation, language, image-only pages).
- ExtractResult: Extracted textual content, tables, and page render artifacts.
- BuildContext: Full payload supplied to Builder stage.
- QaReport: Structured result from Verifier stage.

All objects are JSON-serializable. NO concrete third-party libraries may be imported here.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional


@dataclass
class InputContext:
    """Input configuration and runtime context for the conversion pipeline."""
    pdf_path: Path
    output_docx: Path
    config: Dict[str, Any] = field(default_factory=dict)
    temp_dir: Path = field(default_factory=lambda: Path("temp"))
    workers: int = 4
    dpi: int = 200
    enable_ocr: bool = False
    enable_verify: bool = True

    def to_dict(self) -> Dict[str, Any]:
        return {
            "pdf_path": str(self.pdf_path),
            "output_docx": str(self.output_docx),
            "config": self.config,
            "temp_dir": str(self.temp_dir),
            "workers": self.workers,
            "dpi": self.dpi,
            "enable_ocr": self.enable_ocr,
            "enable_verify": self.enable_verify,
        }

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)


@dataclass
class InspectResult:
    """Results from inspecting the PDF structure and document traits."""
    page_count: int = 0
    page_dimensions: List[Dict[str, float]] = field(default_factory=list)
    image_only_pages: List[int] = field(default_factory=list)
    has_tables: bool = False
    detected_languages: List[str] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)


@dataclass
class ExtractResult:
    """Extracted text, tables, and render paths from the document."""
    text_by_page: Dict[int, str] = field(default_factory=dict)
    rendered_pages: Dict[int, str] = field(default_factory=dict)
    tables_by_page: Dict[int, List[Dict[str, Any]]] = field(default_factory=dict)
    ocr_results: Dict[int, Any] = field(default_factory=dict)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)


@dataclass
class BuildContext:
    """Composite context passed into Builder implementation."""
    input_context: InputContext
    inspect_result: InspectResult
    extract_result: ExtractResult
    output_docx: Path
    config: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "input_context": self.input_context.to_dict(),
            "inspect_result": self.inspect_result.to_dict(),
            "extract_result": self.extract_result.to_dict(),
            "output_docx": str(self.output_docx),
            "config": self.config,
        }

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)


@dataclass
class QaCheckResult:
    """Individual QA axis result."""
    name: str
    passed: bool
    score: float
    details: str
    warnings: List[str] = field(default_factory=list)
    fix_hints: List[str] = field(default_factory=list)


@dataclass
class QaReport:
    """Complete 5-axis quality assurance report."""
    original_pdf: str
    output_docx: str
    overall_pass: bool
    overall_score: float
    checks: List[Dict[str, Any]] = field(default_factory=list)
    verdict: str = ""
    total_time_s: float = 0.0
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)
