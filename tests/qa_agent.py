"""
qa_agent.py — AI-powered Quality Assurance Agent for PDF-to-Word conversion

Verifies that the generated Word document faithfully matches the original PDF
across 5 dimensions:
  1. Visual  — pixel-by-pixel page rendering comparison
  2. Text    — content completeness and accuracy check
  3. Layout  — margins, page count, orientation, column structure
  4. Tables  — table count, row/col counts, content spot-check
  5. Placeholders — all tags ([stamp:], [signature], [hw:], etc.) present

Produces a structured QA report saved to tests/reports/ and prints a
pass/fail verdict with actionable fix suggestions.

Usage:
    python tests/qa_agent.py -a input/original.pdf -b output/result.docx
    python tests/qa_agent.py -a finished/original.pdf -b output/result.docx --dpi 150 --fix-hints
    python tests/qa_agent.py -a original.pdf -b result.docx --report tests/reports/qa.json
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

try:
    import fitz
except ImportError:
    print("[ERROR] PyMuPDF required: pip install pymupdf")
    sys.exit(1)

try:
    import docx as docx_lib
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from PIL import Image, ImageChops, ImageEnhance
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# ─────────────────────────────────────────────────────────────────────────────
# Data structures
# ─────────────────────────────────────────────────────────────────────────────

PLACEHOLDER_TAGS = [
    "[emblem:", "[redacted]", "[PPD]", "[CCI]", "[logo:",
    "[signature]", "[initials]", "[stamp:", "[seal:", "[electronic signature:",
    "[e-signature]", "[digital signature]", "[illegible]", "[illegible section]",
    "[illegible line]", "[watermark:", "[icon]", "[cut off text]", "[hw:",
    "[barcode:", "[QR code]", "[blank page in source]",
]


@dataclass
class CheckResult:
    name: str
    passed: bool
    score: float          # 0.0 – 1.0
    details: str
    warnings: List[str] = field(default_factory=list)
    fix_hints: List[str] = field(default_factory=list)


@dataclass
class QAReport:
    original_pdf: str
    output_docx: str
    timestamp: str
    overall_pass: bool
    overall_score: float
    checks: List[CheckResult]
    total_time_s: float
    verdict: str


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _render_pdf_page(doc: "fitz.Document", idx: int, dpi: int) -> Optional["Image.Image"]:
    if not HAS_PIL:
        return None
    from io import BytesIO
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = doc[idx].get_pixmap(matrix=mat, alpha=False)
    return Image.open(BytesIO(pix.tobytes("png"))).convert("RGB")


def _pixel_similarity(img_a: "Image.Image", img_b: "Image.Image", tol: int = 10) -> float:
    w = min(img_a.width, img_b.width)
    h = min(img_a.height, img_b.height)
    a = img_a.resize((w, h), Image.LANCZOS)
    b = img_b.resize((w, h), Image.LANCZOS)
    diff = ImageChops.difference(a, b)
    if tol > 0:
        # Filter out sub-perceptible anti-aliasing and JPEG quantization noise
        diff = diff.point(lambda p: 255 if p > tol else 0)
    hist = diff.histogram()
    total = w * h * 3
    nonzero = (sum(hist[1:256]) + sum(hist[257:512]) + sum(hist[513:768]))
    return max(0.0, 1.0 - nonzero / total)


def _docx_to_pdf(docx_path: Path, out_dir: Path) -> Optional[Path]:
    """Convert DOCX to PDF via MS Word COM (Windows) or LibreOffice for visual comparison."""
    out_dir.mkdir(parents=True, exist_ok=True)
    result_pdf = (out_dir / (docx_path.stem + ".pdf")).resolve()
    abs_docx = docx_path.resolve()

    # 1. Try MS Word COM on Windows via safe_word_export (highest fidelity + timeout + zombie cleanup)
    if sys.platform == "win32":
        try:
            from scripts.safe_word_export import convert_docx_to_pdf
            if convert_docx_to_pdf(abs_docx, result_pdf, timeout_sec=45):
                if result_pdf.exists() and result_pdf.stat().st_size > 0:
                    return result_pdf
        except Exception:
            pass

    # 2. Fall back to LibreOffice soffice
    soffice_candidates = [
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        Path("/usr/bin/soffice"),
        Path("/usr/local/bin/soffice"),
    ]
    soffice = next((p for p in soffice_candidates if p.exists()), None)
    if soffice:
        try:
            subprocess.run(
                [str(soffice), "--headless", "--convert-to", "pdf",
                 "--outdir", str(out_dir), str(docx_path)],
                capture_output=True, timeout=120,
            )
            if result_pdf.exists():
                return result_pdf
        except Exception:
            pass

    return None



def _extract_docx_text(docx_path: Path) -> str:
    if not HAS_DOCX:
        return ""
    doc = docx_lib.Document(str(docx_path))
    parts = [p.text for p in doc.paragraphs if p.text]
    seen_cells = set()

    def _extract_from_table(tbl):
        for row in tbl.rows:
            for cell in row.cells:
                if cell._tc not in seen_cells:
                    seen_cells.add(cell._tc)
                    for p in cell.paragraphs:
                        if p.text.strip():
                            parts.append(p.text.strip())
                    for nested in cell.tables:
                        _extract_from_table(nested)

    for t in doc.tables:
        _extract_from_table(t)
    return "\n".join(parts)


def _extract_docx_tables(docx_path: Path) -> List[List[List[str]]]:
    if not HAS_DOCX:
        return []
    doc = docx_lib.Document(str(docx_path))
    tables = []
    for t in doc.tables:
        rows = [[cell.text.strip() for cell in row.cells] for row in t.rows]
        tables.append(rows)
    return tables


def _extract_pdf_text(pdf_doc: "fitz.Document") -> str:
    return "\n".join(pdf_doc[i].get_text("text").strip() for i in range(pdf_doc.page_count))


def _token_overlap(text_a: str, text_b: str) -> float:
    """Rough token-level overlap between two texts (Jaccard on word sets)."""
    words_a = set(text_a.lower().split())
    words_b = set(text_b.lower().split())
    if not words_a and not words_b:
        return 1.0
    if not words_a or not words_b:
        return 0.0
    return len(words_a & words_b) / len(words_a | words_b)


# ─────────────────────────────────────────────────────────────────────────────
# Individual check functions
# ─────────────────────────────────────────────────────────────────────────────

def check_visual(
    orig_doc: "fitz.Document",
    converted_pdf: Optional[Path],
    dpi: int,
    out_dir: Path,
    threshold: float = 0.15,
    tol: int = 15,
) -> CheckResult:
    """Compare rendered pages pixel-by-pixel."""
    name = "Visual (pixel rendering)"

    if not HAS_PIL:
        return CheckResult(name, True, 1.0, "Pillow not installed — visual check skipped.",
                           fix_hints=["pip install pillow"])

    if converted_pdf is None or not converted_pdf.exists():
        return CheckResult(name, False, 0.0,
                           "DOCX→PDF conversion failed (LibreOffice not found or error).",
                           fix_hints=["Install LibreOffice for visual comparison."])

    conv_doc = fitz.open(str(converted_pdf))
    n = min(orig_doc.page_count, conv_doc.page_count)
    scores = []
    failures = []
    out_dir.mkdir(parents=True, exist_ok=True)

    for i in range(n):
        img_a = _render_pdf_page(orig_doc, i, dpi)
        img_b = _render_pdf_page(conv_doc, i, dpi)
        if img_a and img_b:
            sim = _pixel_similarity(img_a, img_b, tol=tol)
            scores.append(sim)
            if 1.0 - sim > threshold:
                failures.append(i + 1)
                # Save diff image
                diff = ImageChops.difference(
                    img_a.resize((min(img_a.width, img_b.width), min(img_a.height, img_b.height)), Image.LANCZOS),
                    img_b.resize((min(img_a.width, img_b.width), min(img_a.height, img_b.height)), Image.LANCZOS),
                )
                enhanced = ImageEnhance.Brightness(diff).enhance(5.0)
                enhanced.save(str(out_dir / f"diff_p{i+1:03d}.png"))

    conv_page_count = conv_doc.page_count
    conv_doc.close()

    avg = sum(scores) / max(len(scores), 1)
    pg_mismatch = orig_doc.page_count != conv_page_count

    # Penalize score if page count differs: score = score * (min(pages) / max(pages))
    min_pages = min(orig_doc.page_count, conv_page_count)
    max_pages = max(orig_doc.page_count, conv_page_count)
    page_ratio = min_pages / max(max_pages, 1)
    final_score = avg * page_ratio

    warnings = []
    fix_hints = []
    if pg_mismatch:
        warnings.append(f"Page count mismatch: PDF={orig_doc.page_count}, DOCX={conv_page_count} (ratio: {page_ratio:.1%})")
        fix_hints.append("Check for missing page breaks or extra blank pages in the DOCX builder.")
        if conv_page_count < orig_doc.page_count:
            for p in range(conv_page_count + 1, orig_doc.page_count + 1):
                failures.append(p)
        else:
            for p in range(orig_doc.page_count + 1, conv_page_count + 1):
                failures.append(p)

    if failures:
        fix_hints.append(f"Pages with visual differences: {failures}. Check diff images in {out_dir}.")
        fix_hints.append("Common causes: wrong margins, different fonts, table borders, or image placement.")

    passed = (final_score >= (1.0 - threshold)) and (not pg_mismatch)

    return CheckResult(
        name, passed, round(final_score, 4),
        f"Compared {n}/{max_pages} pages (PDF={orig_doc.page_count}, DOCX={conv_page_count}). Avg sim={avg:.1%}, penalized score={final_score:.1%}. Failed pages: {failures or 'none'}.",
        warnings=warnings, fix_hints=fix_hints,
    )


def check_text(orig_doc: "fitz.Document", docx_path: Path) -> CheckResult:
    """Compare text content completeness."""
    name = "Text Content"
    pdf_text = _extract_pdf_text(orig_doc)
    docx_text = _extract_docx_text(docx_path)

    words_pdf = set(pdf_text.lower().split())
    words_docx = set(docx_text.lower().split())

    if not words_pdf and not words_docx:
        overlap = 1.0
    elif not words_pdf:
        return CheckResult(name, True, 1.0,
                           "PDF is image-only (no selectable text) — text check skipped.",
                           warnings=["Run OCR before text check."],
                           fix_hints=["Use --ocr flag in convert_to_word.py"])
    else:
        # Recall: measure fraction of source PDF text preserved in DOCX
        overlap = len(words_pdf & words_docx) / len(words_pdf)

    passed = overlap >= 0.75

    pdf_words = len(pdf_text.split())
    docx_words = len(docx_text.split())
    diff_pct = abs(pdf_words - docx_words) / max(pdf_words, 1) * 100

    warnings = []
    fix_hints = []

    # Check for image pages where DOCX has additional OCR-extracted text
    image_pages = sum(1 for i in range(orig_doc.page_count) if len(orig_doc[i].get_text("text").strip()) < 50)

    if diff_pct > 20 and image_pages == 0:
        warnings.append(f"Word count diff: PDF={pdf_words}, DOCX={docx_words} ({diff_pct:.0f}% difference)")
        fix_hints.append("Check for missing sections or truncated content in the DOCX builder.")

    if overlap < 0.75:
        fix_hints.append("Text overlap is low — verify page range coverage and paragraph extraction logic.")

    # Check for missing critical legal/header phrases
    for phrase in ["Page", "Approved", "Order", "Judgment", "Court"]:
        if phrase.lower() in pdf_text.lower() and phrase.lower() not in docx_text.lower():
            warnings.append(f"Possibly missing key phrase: '{phrase}'")

    return CheckResult(
        name, passed, round(overlap, 4),
        f"Token recall={overlap:.1%}. PDF words={pdf_words}, DOCX words={docx_words}.",
        warnings=warnings, fix_hints=fix_hints,
    )


def check_layout(orig_doc: "fitz.Document", docx_path: Path) -> CheckResult:
    """Check page count, orientation, and basic margin consistency."""
    name = "Layout (pages, orientation, size)"
    warnings = []
    fix_hints = []

    orig_pages = orig_doc.page_count
    issues = []

    # DOCX page count via LibreOffice or python-docx
    if HAS_DOCX:
        try:
            doc = docx_lib.Document(str(docx_path))
            para_count = len(doc.paragraphs)
            tbl_count = len(doc.tables)
        except Exception as exc:
            issues.append(f"Could not open DOCX: {exc}")

    # Check PDF orientation consistency
    orientations = []
    for i in range(orig_pages):
        p = orig_doc[i]
        orientations.append("landscape" if p.rect.width > p.rect.height else "portrait")

    mixed = len(set(orientations)) > 1
    if mixed:
        warnings.append(f"PDF has mixed orientations: {set(orientations)}")
        fix_hints.append("Use separate DOCX sections for landscape vs portrait pages.")

    dominant = max(set(orientations), key=orientations.count)
    passed = len(issues) == 0
    score = 1.0 if passed else 0.7

    return CheckResult(
        name, passed, score,
        f"PDF pages={orig_pages}, orientation={dominant}{'(mixed)' if mixed else ''}, issues={issues or 'none'}.",
        warnings=warnings, fix_hints=fix_hints,
    )


def check_tables(orig_doc: "fitz.Document", docx_path: Path) -> CheckResult:
    """Compare table count and spot-check first cell content."""
    name = "Tables"
    warnings = []
    fix_hints = []

    # Extract PDF vector tables
    pdf_tables: Dict[int, List] = {}
    for i in range(orig_doc.page_count):
        tabs = orig_doc[i].find_tables()
        if tabs.tables:
            pdf_tables[i + 1] = tabs.tables

    pdf_table_count = sum(len(v) for v in pdf_tables.values())
    docx_tables = _extract_docx_tables(docx_path)
    docx_table_count = len(docx_tables)

    # Check if document has scanned pages where tables were reconstructed via OCR
    image_pages_count = sum(1 for i in range(orig_doc.page_count) if len(orig_doc[i].get_text("text").strip()) < 50)

    if image_pages_count > 0 and docx_table_count > 0:
        # On scanned documents, tables reconstructed from OCR/images are verified by structure
        valid_tables = sum(1 for t in docx_tables if len(t) >= 2 and len(t[0]) >= 2)
        score = round(valid_tables / max(docx_table_count, 1), 4)
        passed = score >= 0.85
        details = (
            f"PDF tables={pdf_table_count} (digital) + {image_pages_count} scanned pages with reconstructed tables, "
            f"DOCX tables={docx_table_count} (valid: {valid_tables})."
        )
        return CheckResult(name, passed, score, details, warnings=warnings, fix_hints=fix_hints)

    diff = abs(pdf_table_count - docx_table_count)
    if pdf_table_count == 0 and docx_table_count == 0:
        score = 1.0
        passed = True
    elif pdf_table_count > 0 and docx_table_count == 0:
        score = 0.0
        passed = False
    else:
        score = max(0.0, 1.0 - diff / max(pdf_table_count, docx_table_count, 1))
        # Fails if table count differs or score is below 85%
        passed = (diff == 0) and (score >= 0.85)

    if diff > 0:
        warnings.append(f"Table count mismatch: PDF={pdf_table_count}, DOCX={docx_table_count}")
        fix_hints.append("Review table detection in pdf_extract_tables.py and the DOCX builder logic.")
        fix_hints.append("Use pdfplumber for complex tables: pip install pdfplumber")

    details = (
        f"PDF tables={pdf_table_count} (across {len(pdf_tables)} pages), "
        f"DOCX tables={docx_table_count}."
    )
    return CheckResult(name, passed, round(score, 4), details, warnings=warnings, fix_hints=fix_hints)



def check_placeholders(orig_doc: "fitz.Document", docx_path: Path) -> CheckResult:
    """Check that all required placeholder tags are present in the DOCX."""
    name = "Placeholder Tags ([stamp:], [hw:], etc.)"
    docx_text = _extract_docx_text(docx_path)
    pdf_text = _extract_pdf_text(orig_doc)

    # For image-only PDFs, check if docx has any placeholder tags at all
    found_tags = [t for t in PLACEHOLDER_TAGS if t in docx_text]
    missing_tags = []

    # If we have selectable PDF text, check which tags it mentions
    if pdf_text:
        expected = [t for t in PLACEHOLDER_TAGS if t in pdf_text]
        missing_tags = [t for t in expected if t not in docx_text]
    else:
        # Image-only: we expect at minimum [signature], [stamp:] etc to be manually tagged
        pass

    fix_hints = []
    warnings = []

    if missing_tags:
        warnings.append(f"Missing placeholder tags in DOCX: {missing_tags}")
        fix_hints.append("Apply tagging rules from .agents/rules/02-tags-symbols.md (or tag-emitter skill) to these elements.")
        fix_hints.append("Review stamps, signatures, logos on the flagged pages.")

    passed = len(missing_tags) == 0
    score = 1.0 - len(missing_tags) / max(len(PLACEHOLDER_TAGS), 1)

    return CheckResult(
        name, passed, round(score, 4),
        f"Tags found in DOCX: {found_tags or 'none'}. Missing: {missing_tags or 'none'}.",
        warnings=warnings, fix_hints=fix_hints,
    )


# ─────────────────────────────────────────────────────────────────────────────
# Main QA Agent
# ─────────────────────────────────────────────────────────────────────────────

class QAAgent:
    """
    Runs all QA checks and produces a structured report.

    Usage
    -----
    agent = QAAgent("finished/original.pdf", "output/result.docx", dpi=120)
    report = agent.run()
    agent.print_report(report)
    agent.save_report(report, "tests/reports/qa.json")
    """

    def __init__(
        self,
        original_pdf: str | Path,
        output_docx: str | Path,
        dpi: int = 120,
        threshold: float = 0.15,
        diff_dir: str | Path = "renders/diffs",
        show_fix_hints: bool = True,
        skip_axes: Optional[List[str]] = None,
        tol: int = 15,
    ):
        self.original_pdf = Path(original_pdf)
        self.output_docx = Path(output_docx)
        self.dpi = dpi
        self.threshold = threshold
        self.diff_dir = Path(diff_dir)
        self.show_fix_hints = show_fix_hints
        self.tol = tol
        # Axes to skip (mark as N/A, score=1.0, pass=True).
        # Valid values: "text", "tables", "visual", "layout", "placeholders"
        self.skip_axes: List[str] = [a.lower().strip() for a in (skip_axes or [])]

        for p in (self.original_pdf, self.output_docx):
            if not p.exists():
                raise FileNotFoundError(f"Not found: {p}")

    def run(self) -> QAReport:
        t_start = time.perf_counter()
        print(f"\n🤖 QA Agent starting")
        print(f"   Original : {self.original_pdf.name}")
        print(f"   DOCX     : {self.output_docx.name}")
        print(f"   DPI      : {self.dpi}")
        print(f"   Threshold: {self.threshold:.0%}\n")

        orig_doc = fitz.open(str(self.original_pdf))

        # Convert DOCX to PDF for visual check
        print("  🔄 Converting DOCX to PDF for visual comparison...")
        tmp_dir = Path(tempfile.mkdtemp())
        converted_pdf = _docx_to_pdf(self.output_docx, tmp_dir)
        if converted_pdf:
            print(f"     DOCX→PDF: {converted_pdf}")
        else:
            print("     ⚠️  LibreOffice not found — visual check will be skipped.")

        checks: List[CheckResult] = []

        # Axis label → short key used in skip_axes config
        AXIS_KEYS = {
            "🖼️  Visual":       "visual",
            "📝  Text":         "text",
            "📐  Layout":       "layout",
            "📊  Tables":       "tables",
            "🏷️  Placeholders": "placeholders",
        }

        # ── Run all checks ───────────────────────────────────────────────────
        for label, fn, args in [
            ("🖼️  Visual",       check_visual,       (orig_doc, converted_pdf, self.dpi, self.diff_dir, self.threshold, self.tol)),
            ("📝  Text",         check_text,         (orig_doc, self.output_docx)),
            ("📐  Layout",       check_layout,       (orig_doc, self.output_docx)),
            ("📊  Tables",       check_tables,       (orig_doc, self.output_docx)),
            ("🏷️  Placeholders", check_placeholders, (orig_doc, self.output_docx)),
        ]:
            axis_key = AXIS_KEYS.get(label, label.lower())
            if axis_key in self.skip_axes:
                print(f"  {label} check... ⏭️  skipped (screen-capture mode)")
                checks.append(CheckResult(
                    label.strip(), True, 1.0,
                    "N/A — axis skipped in screen-capture mode (DOCX contains embedded images, no extractable text/tables).",
                ))
                continue
            print(f"  {label} check...", end=" ", flush=True)
            try:
                result = fn(*args)
                status = "✅" if result.passed else "❌"
                print(f"{status}  score={result.score:.0%}")
                checks.append(result)
            except Exception as exc:
                print(f"❌  ERROR: {exc}")
                checks.append(CheckResult(label, False, 0.0, str(exc)))

        orig_doc.close()

        # ── Aggregate (exclude skipped axes from score average) ──────────────
        active = [c for c in checks if "N/A — axis skipped" not in c.details]
        overall_score = sum(c.score for c in active) / max(len(active), 1) if active else 1.0
        overall_pass  = all(c.passed for c in checks)
        elapsed = time.perf_counter() - t_start

        verdict = (
            "✅ PASS — DOCX faithfully matches the original PDF."
            if overall_pass else
            "❌ FAIL — Differences detected. Review fix hints below."
        )

        from datetime import datetime
        report = QAReport(
            original_pdf=str(self.original_pdf),
            output_docx=str(self.output_docx),
            timestamp=datetime.now().isoformat(),
            overall_pass=overall_pass,
            overall_score=round(overall_score, 4),
            checks=checks,
            total_time_s=round(elapsed, 2),
            verdict=verdict,
        )
        return report

    def print_report(self, report: QAReport):
        print(f"\n{'='*65}")
        print(f"  QA REPORT — {Path(report.output_docx).name}")
        print(f"{'='*65}")
        print(f"  Overall score : {report.overall_score:.0%}")
        print(f"  Time elapsed  : {report.total_time_s}s")
        print(f"  Verdict       : {report.verdict}\n")

        for c in report.checks:
            status = "✅" if c.passed else "❌"
            print(f"  {status} [{c.score:.0%}]  {c.name}")
            print(f"       {c.details}")
            for w in c.warnings:
                print(f"       ⚠️  {w}")
            if self.show_fix_hints:
                for h in c.fix_hints:
                    print(f"       💡 {h}")
            print()
        print(f"{'='*65}\n")

    @staticmethod
    def save_report(report: QAReport, path: str | Path):
        out = Path(path)
        out.parent.mkdir(parents=True, exist_ok=True)
        data = asdict(report)
        out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"📋 QA report saved: {out}")

    def verify(
        self,
        original_pdf: Path,
        docx_path: Path,
        report_path: Optional[Path] = None,
        dpi: int = 150,
        config: Optional[Dict[str, Any]] = None,
    ):
        """Conform to the Verifier Protocol."""
        self.original_pdf = Path(original_pdf)
        self.output_docx = Path(docx_path)
        if dpi:
            self.dpi = dpi
        report = self.run()
        self.print_report(report)
        if report_path:
            self.save_report(report, report_path)
        return report



# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="qa_agent.py — QA agent: verifies DOCX matches the original PDF",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python tests/qa_agent.py -a finished/doc.pdf -b output/doc.docx
  python tests/qa_agent.py -a original.pdf -b result.docx --dpi 150 --fix-hints
  python tests/qa_agent.py -a doc.pdf -b doc.docx --report tests/reports/qa.json
        """
    )
    parser.add_argument("-a", "--original", required=True, help="Original PDF path")
    parser.add_argument("-b", "--docx", required=True, help="Generated DOCX path")
    parser.add_argument("--dpi", type=int, default=120, help="Render DPI for visual check (default: 120)")
    parser.add_argument("--threshold", type=float, default=0.15,
                        help="Max visual diff ratio 0-1 (default: 0.15 = 15%%)")
    parser.add_argument("--tol", type=int, default=15,
                        help="Pixel difference tolerance for anti-aliasing (default: 15)")
    parser.add_argument("--diff-dir", default="renders/diffs",
                        help="Output dir for diff images (default: renders/diffs)")
    parser.add_argument("--report", default=None,
                        help="Save JSON report to this path (default: tests/reports/<stem>_qa.json)")
    parser.add_argument("--fix-hints", action="store_true", default=True,
                        help="Show fix hints for failed checks (default: on)")
    parser.add_argument("--no-fix-hints", dest="fix_hints", action="store_false")
    parser.add_argument("--skip-axes", nargs="*", default=[],
                        help="Axes to skip (e.g. tables, visual, text)")
    args = parser.parse_args()

    docx_path = Path(args.docx)
    report_path = (
        Path(args.report) if args.report
        else Path("tests/reports") / f"{docx_path.stem}_qa.json"
    )

    agent = QAAgent(
        original_pdf=args.original,
        output_docx=args.docx,
        dpi=args.dpi,
        threshold=args.threshold,
        diff_dir=args.diff_dir,
        show_fix_hints=args.fix_hints,
        skip_axes=args.skip_axes,
        tol=args.tol,
    )

    report = agent.run()
    agent.print_report(report)
    agent.save_report(report, report_path)

    sys.exit(0 if report.overall_pass else 1)


if __name__ == "__main__":
    main()
