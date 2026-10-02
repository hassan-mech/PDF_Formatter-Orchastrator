"""
scripts/bootstrap.py — Composition Root

Wires abstract stage interfaces to concrete stage implementations using Registry
or direct Dependency Injection.
Satisfies Dependency Inversion Principle (DIP):
- Concrete imports occur only here at the application boundary.
- Constructs and returns a ready-to-run Pipeline instance.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any, Dict, Optional

# Ensure scripts directory is in sys.path
SCRIPTS_DIR = Path(__file__).resolve().parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

# Import core
from core.pipeline import Pipeline
from core.registry import (
    builder_registry,
    extractor_registry,
    inspector_registry,
    ocr_registry,
    verifier_registry,
)

# Import and register concrete stages
import stages.inspectors.pymupdf_inspector
import stages.extractors.pymupdf_extractor
import stages.extractors.table_extractor
import stages.ocr.rapidocr_engine
import stages.builders.docx_builder
import stages.builders.chinese_medical_builder
import stages.builders.screen_capture_builder
import stages.builders.journal_article_builder
import stages.verifiers.qa_verifier


def build_default_pipeline(config: Optional[Dict[str, Any]] = None) -> Pipeline:
    """Build and wire the default 5-stage conversion pipeline.

    Stage selection is configuration-driven via registries, defaulting to:
    - Inspector: 'pymupdf'
    - Extractor: 'pymupdf'
    - OcrEngine: 'rapidocr'
    - Builder:   'docx_builder'
    - Verifier:  'qa_verifier'
    """
    cfg = config or {}

    inspector_name = cfg.get("inspector", "pymupdf")
    extractor_name = cfg.get("extractor", "pymupdf")
    ocr_name = cfg.get("ocr", {}).get("engine", "rapidocr")
    builder_name = cfg.get("builder", "docx_builder")
    verifier_name = cfg.get("verifier", "qa_verifier")

    inspector_cls = inspector_registry.get(inspector_name)
    extractor_cls = extractor_registry.get(extractor_name)
    ocr_cls = ocr_registry.get(ocr_name)
    builder_cls = builder_registry.get(builder_name)
    verifier_cls = verifier_registry.get(verifier_name)

    return Pipeline(
        inspector=inspector_cls(),
        extractor=extractor_cls(),
        ocr_engine=ocr_cls(),
        builder=builder_cls(),
        verifier=verifier_cls(),
    )
