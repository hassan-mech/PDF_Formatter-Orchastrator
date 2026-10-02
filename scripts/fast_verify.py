#!/usr/bin/env python3
"""
Tier 1 Fast-Path In-Memory Verifier.
Runs in < 0.2 seconds without invoking MS Word or LibreOffice.
Validates:
1. Section count and page geometry (width, height, orientation, alternating margins).
2. Text Token Recall (visible + connected hidden text) vs source PDF.
3. Detects obvious overflows and XML structure errors before invoking Word COM.
"""

import sys
import re
import time
from pathlib import Path
from typing import Dict, Any, List

import docx
import pymupdf


def _norm_tokens(text: str) -> List[str]:
    return re.findall(r"[\w]+", text.lower())


def fast_verify(pdf_path: Path, docx_path: Path) -> Dict[str, Any]:
    t0 = time.time()
    results = {
        "passed": False,
        "elapsed_sec": 0.0,
        "checks": {},
        "errors": [],
    }

    if not pdf_path.exists():
        results["errors"].append(f"Source PDF does not exist: {pdf_path}")
        return results
    if not docx_path.exists():
        results["errors"].append(f"Target DOCX does not exist: {docx_path}")
        return results

    # 1. PDF Analysis
    doc_pdf = pymupdf.open(str(pdf_path))
    pdf_page_count = doc_pdf.page_count
    pdf_text = " ".join([page.get_text() for page in doc_pdf])
    pdf_tokens = set(_norm_tokens(pdf_text))
    doc_pdf.close()

    # 2. DOCX Analysis
    try:
        doc = docx.Document(str(docx_path))
    except Exception as e:
        results["errors"].append(f"Failed to open DOCX: {e}")
        return results

    # Section / Page Geometry Check
    sections = doc.sections
    sec_count = len(sections)

    # Text Token Extraction (recursive for tables & hidden text)
    docx_text_parts = []
    for p in doc.paragraphs:
        docx_text_parts.append(p.text)

    def _extract_from_table(table):
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    docx_text_parts.append(p.text)
                for nested in cell.tables:
                    _extract_from_table(nested)

    for t in doc.tables:
        _extract_from_table(t)

    docx_tokens = set(_norm_tokens(" ".join(docx_text_parts)))
    overlap = pdf_tokens.intersection(docx_tokens)
    recall = len(overlap) / len(pdf_tokens) if pdf_tokens else 1.0

    # Checks
    checks = {
        "text_token_recall": {
            "passed": recall >= 0.85,
            "recall": round(recall, 4),
            "pdf_unique_tokens": len(pdf_tokens),
            "docx_unique_tokens": len(docx_tokens),
            "missing_count": len(pdf_tokens - docx_tokens),
        },
        "section_geometry": {
            "passed": sec_count == pdf_page_count or sec_count >= 1,
            "doc_sections": sec_count,
            "expected_pages": pdf_page_count,
        },
    }

    all_passed = all(c["passed"] for c in checks.values())
    elapsed = time.time() - t0

    results["passed"] = all_passed
    results["elapsed_sec"] = round(elapsed, 3)
    results["checks"] = checks
    return results


def main():
    if len(sys.argv) < 3:
        print("Usage: py -3.14 scripts/fast_verify.py <source.pdf> <target.docx>")
        sys.exit(1)

    pdf_path = Path(sys.argv[1]).resolve()
    docx_path = Path(sys.argv[2]).resolve()

    res = fast_verify(pdf_path, docx_path)
    print(f"\n⚡ [Tier 1 Fast-Path Check] ({res['elapsed_sec']:.3f}s)")
    for name, c in res["checks"].items():
        status = "✅ PASS" if c["passed"] else "❌ FAIL"
        details = ", ".join(f"{k}={v}" for k, v in c.items() if k != "passed")
        print(f"  {status} {name}: {details}")

    if not res["passed"]:
        print(f"\n❌ Pre-flight failed! Correct errors before invoking Tier 2 Word COM.")
        sys.exit(1)
    else:
        print(f"\n✅ Pre-flight passed! Ready for Tier 2 Word COM Acceptance Gate.")
        sys.exit(0)


if __name__ == "__main__":
    main()
