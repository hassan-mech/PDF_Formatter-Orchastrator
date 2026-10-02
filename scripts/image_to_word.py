"""
scripts/image_to_word.py — Thin CLI for Image / Scanned Page → DOCX Conversion

Provides a direct CLI interface to convert images or scanned document pages into DOCX.
Satisfies SOLID Single Responsibility:
- Thin CLI wrapper delegating to OcrEngine and Builder.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any, Dict

import yaml

SCRIPTS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPTS_DIR.parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from core.context import BuildContext, ExtractResult, InputContext, InspectResult
from core.registry import builder_registry, ocr_registry
import stages.builders.docx_builder
import stages.ocr.rapidocr_engine


def main():
    parser = argparse.ArgumentParser(
        description="image_to_word.py — Convert image or scan directly to DOCX",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("-i", "--input", required=True, help="Input image or PDF scan path")
    parser.add_argument("-o", "--output", default=None, help="Output DOCX path")
    parser.add_argument("--config", default="default", help="Config name (configs/<name>.yaml)")
    parser.add_argument("--dpi", type=int, default=300, help="Target DPI")

    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"❌ Input file not found: {input_path}")
        sys.exit(1)

    stem = input_path.stem
    output_docx = Path(args.output) if args.output else PROJECT_ROOT / "output" / f"{stem}.docx"

    # If input is a PDF, forward directly to convert_to_word
    if input_path.suffix.lower() == ".pdf":
        from convert_to_word import main as c_main
        sys.argv = ["convert_to_word.py", "-i", str(input_path), "-o", str(output_docx), "--config", args.config, "--ocr"]
        return c_main()

    # Load config
    cfg_file = PROJECT_ROOT / "configs" / f"{args.config}.yaml"
    cfg: Dict[str, Any] = {}
    if cfg_file.exists():
        with open(cfg_file, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}

    print(f"▶ Converting image '{input_path.name}' to DOCX '{output_docx.name}'...")

    # Run OCR on image
    ocr_cls = ocr_registry.get(cfg.get("ocr", {}).get("engine", "rapidocr"))
    ocr_engine = ocr_cls()
    ocr_res = ocr_engine.ocr([input_path], cfg.get("ocr", {}).get("langs", ["en", "zh", "ar"]))

    # Build contexts
    input_ctx = InputContext(
        pdf_path=input_path,
        output_docx=output_docx,
        config=cfg,
        dpi=args.dpi,
    )
    inspect_res = InspectResult(
        page_count=1,
        image_only_pages=[1],
    )
    extract_res = ExtractResult(
        rendered_pages={1: str(input_path)},
        ocr_results={1: ocr_res},
    )

    build_ctx = BuildContext(
        input_context=input_ctx,
        inspect_result=inspect_res,
        extract_result=extract_res,
        output_docx=output_docx,
        config=cfg,
    )

    builder_cls = builder_registry.get(cfg.get("builder", "docx_builder"))
    builder = builder_cls()
    built_path = builder.build(build_ctx)

    print(f"✅ Document successfully created: {built_path}")
    sys.exit(0)


if __name__ == "__main__":
    main()
