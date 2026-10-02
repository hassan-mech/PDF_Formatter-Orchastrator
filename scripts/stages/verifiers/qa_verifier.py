"""
scripts/stages/verifiers/qa_verifier.py — QaVerifier Implementation

Executes 5-axis verification (Visual, Text, Layout, Tables, Placeholders).
Satisfies Verifier protocol and registers with verifier_registry under name 'qa_verifier'.
"""

from __future__ import annotations

import sys
from dataclasses import asdict
from pathlib import Path
from typing import Any, Dict, Optional

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

SCRIPTS_DIR = PROJECT_ROOT / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from core.context import QaReport
from core.registry import verifier_registry


@verifier_registry.register("qa_verifier")
class QaVerifier:
    """Verifier implementation utilizing tests/qa_agent.py."""

    def verify(
        self,
        original_pdf: Path,
        docx_path: Path,
        report_path: Optional[Path] = None,
        dpi: int = 150,
        config: Optional[Dict[str, Any]] = None,
    ) -> QaReport:
        from tests.qa_agent import QAAgent

        threshold = 0.15
        tol = 15
        skip_axes = []
        if config and "verification" in config:
            threshold = config["verification"].get("visual_diff_threshold", 0.15)
            tol = config["verification"].get("anti_aliasing_tolerance", config["verification"].get("tol", 15))
            dpi = config["verification"].get("dpi_for_render", dpi)
            skip_axes = config["verification"].get("skip_axes", [])

        agent = QAAgent(
            original_pdf=str(original_pdf),
            output_docx=str(docx_path),
            dpi=dpi,
            threshold=threshold,
            diff_dir="renders/diffs",
            show_fix_hints=True,
            skip_axes=skip_axes,
            tol=tol,
        )

        rep = agent.run()
        agent.print_report(rep)

        if report_path:
            agent.save_report(rep, report_path)

        # Convert to core.context.QaReport
        checks_data = []
        for c in rep.checks:
            checks_data.append({
                "name": c.name,
                "passed": c.passed,
                "score": c.score,
                "details": c.details,
                "warnings": c.warnings,
                "fix_hints": c.fix_hints,
            })

        return QaReport(
            original_pdf=rep.original_pdf,
            output_docx=rep.output_docx,
            overall_pass=rep.overall_pass,
            overall_score=rep.overall_score,
            checks=checks_data,
            verdict=rep.verdict,
            total_time_s=rep.total_time_s,
            timestamp=rep.timestamp,
        )
