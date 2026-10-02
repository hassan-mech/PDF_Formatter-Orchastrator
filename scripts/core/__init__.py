"""
scripts/core package — Interfaces, context DTOs, registry, and pipeline orchestrator.
"""

from .context import BuildContext, ExtractResult, InputContext, InspectResult, QaCheckResult, QaReport
from .interfaces import Builder, Extractor, Inspector, OcrEngine, Verifier
from .pipeline import Pipeline
from .registry import Registry, builder_registry, extractor_registry, inspector_registry, ocr_registry, verifier_registry

__all__ = [
    "InputContext",
    "InspectResult",
    "ExtractResult",
    "BuildContext",
    "QaCheckResult",
    "QaReport",
    "Inspector",
    "Extractor",
    "OcrEngine",
    "Builder",
    "Verifier",
    "Registry",
    "inspector_registry",
    "extractor_registry",
    "ocr_registry",
    "builder_registry",
    "verifier_registry",
    "Pipeline",
]
