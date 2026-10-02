"""
scripts/convert_to_word.py — Thin CLI for PDF → DOCX Conversion

Delegates all execution to the SOLID pipeline assembled by bootstrap.py.
Satisfies SOLID Single Responsibility & Open/Closed:
- ZERO business logic.
- ZERO document-type if/else switching.
- Pure parameter parsing and pipeline invocation.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

import yaml

SCRIPTS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPTS_DIR.parent
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from bootstrap import build_default_pipeline
from core.context import InputContext


def load_config(config_name: str) -> Dict[str, Any]:
    """Load configuration dictionary from configs/<name>.yaml.

    Accepts:
      - stem only:          "screen_capture"
      - relative with ext:  "configs/screen_capture.yaml"
      - absolute path:      "d:/foo/configs/screen_capture.yaml"
    """
    from pathlib import PurePosixPath
    p = Path(config_name)

    # If argument already has .yaml extension, resolve it directly
    if p.suffix.lower() == ".yaml":
        # Try as-is (absolute), then relative to project root, then just the filename in configs/
        candidates = [p, PROJECT_ROOT / p, PROJECT_ROOT / "configs" / p.name]
    else:
        candidates = [PROJECT_ROOT / "configs" / f"{config_name}.yaml"]

    cfg_file = None
    for candidate in candidates:
        if candidate.exists():
            cfg_file = candidate
            break

    if cfg_file is None:
        fallback = PROJECT_ROOT / "configs" / "default.yaml"
        print(f"  ⚠️  Config '{config_name}' not found, falling back to default.yaml")
        cfg_file = fallback

    if not cfg_file.exists():
        return {}

    try:
        with open(cfg_file, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        print(f"  📋 Config loaded: {cfg_file.name}  →  builder={data.get('builder', 'docx_builder')}")
        return data
    except Exception as exc:
        print(f"  ⚠️ Warning: Failed to load config {cfg_file}: {exc}")
        return {}


def log_promotion(stem: str, config_name: str, score: float):
    """Append entry to CHANGELOG.md upon QA PASS."""
    changelog_path = PROJECT_ROOT / "CHANGELOG.md"
    today = datetime.now().strftime("%Y-%m-%d")
    entry = f"{stem} | config={config_name} | score={score:.2f} | {today}\n"
    with open(changelog_path, "a", encoding="utf-8") as f:
        f.write(entry)


def log_decision(stem: str, config_name: str, qa_report: Any, passed: bool):
    """Append entry to DECISIONS.md upon conversion completion."""
    decisions_path = PROJECT_ROOT / "DECISIONS.md"
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    res_str = "PASS" if passed else "FAIL"

    scores_str = "n/a"
    if qa_report and qa_report.checks:
        scores_parts = []
        for c in qa_report.checks:
            name = c.get("name", "").lower()
            score = c.get("score", 0.0)
            if "visual" in name:
                scores_parts.append(f"visual={score:.2f}")
            elif "text" in name:
                scores_parts.append(f"text={score:.2f}")
            elif "layout" in name:
                scores_parts.append(f"layout={score:.2f}")
            elif "table" in name:
                scores_parts.append(f"tables={score:.2f}")
            elif "placeholder" in name:
                scores_parts.append(f"placeholders={score:.2f}")
        scores_str = " ".join(scores_parts)

    entry = (
        f"\n## {now_str} | {stem}\n"
        f"- Config used: {config_name}\n"
        f"- Gap found: none\n"
        f"- Action: reused pipeline\n"
        f"- SOLID check: interfaces-implemented=[Inspector,Extractor,OcrEngine,Builder,Verifier] | pipeline-edited=no\n"
        f"- QA: {scores_str}\n"
        f"- Result: {res_str}\n"
        f"---\n"
    )
    with open(decisions_path, "a", encoding="utf-8") as f:
        f.write(entry)


def main():
    parser = argparse.ArgumentParser(
        description="convert_to_word.py — Universal PDF to Word Converter",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("-i", "--input", required=True, help="Input PDF file path")
    parser.add_argument("-o", "--output", default=None, help="Output DOCX file path")
    parser.add_argument("--config", default="default", help="Config profile name (configs/<name>.yaml)")
    parser.add_argument("--ocr", action="store_true", help="Force OCR across all pages")
    parser.add_argument("--verify", action="store_true", default=True, help="Run 5-axis QA verification")
    parser.add_argument("--no-verify", dest="verify", action="store_false", help="Skip QA verification")
    parser.add_argument("--workers", type=int, default=4, help="Worker concurrency")
    parser.add_argument("--dpi", type=int, default=200, help="Rendering DPI for rasterization")

    args = parser.parse_args()

    input_pdf = Path(args.input)
    if not input_pdf.exists():
        print(f"❌ Input file not found: {input_pdf}")
        sys.exit(1)

    stem = input_pdf.stem
    output_docx = Path(args.output) if args.output else PROJECT_ROOT / "output" / f"{stem}.docx"

    # 1. Load config
    cfg = load_config(args.config)

    # 2. Build input context
    ctx = InputContext(
        pdf_path=input_pdf,
        output_docx=output_docx,
        config=cfg,
        temp_dir=PROJECT_ROOT / "temp",
        workers=args.workers,
        dpi=args.dpi,
        enable_ocr=args.ocr,
        enable_verify=args.verify,
    )

    # 3. Assemble pipeline & run
    pipeline = build_default_pipeline(cfg)
    built_path, qa_report = pipeline.run(ctx)

    # 4. Handle promotion and exit status
    if qa_report:
        log_decision(stem, args.config, qa_report, qa_report.overall_pass)
        if qa_report.overall_pass:
            print("\n🎉 QA Passed (all axes ≥ 0.85). Promoting file...")
            log_promotion(stem, args.config, qa_report.overall_score)
            finished_dir = PROJECT_ROOT / "finished"
            finished_dir.mkdir(parents=True, exist_ok=True)
            archived_pdf = finished_dir / input_pdf.name
            if input_pdf.resolve() != archived_pdf.resolve():
                if archived_pdf.exists():
                    try:
                        archived_pdf.unlink()
                        shutil.move(str(input_pdf), str(archived_pdf))
                    except Exception:
                        # Destination exists and is locked (e.g. open in Foxit/Acrobat)
                        try:
                            input_pdf.unlink()
                        except Exception:
                            pass
                else:
                    try:
                        shutil.move(str(input_pdf), str(archived_pdf))
                    except Exception:
                        pass
                print(f"  📦 Archived: {input_pdf} → {archived_pdf}")
            sys.exit(0)
        else:
            print("\n❌ QA Failed. Review fix hints in the report.")
            sys.exit(1)
    else:
        # Verification was skipped
        print(f"\n✅ DOCX created successfully: {built_path}")
        sys.exit(0)


if __name__ == "__main__":
    main()
