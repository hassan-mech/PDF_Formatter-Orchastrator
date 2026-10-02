"""
scripts/core/pipeline.py — Master Orchestration Pipeline

Implements the single stable conversion pipeline.
Satisfies SOLID Dependency Inversion:
- Depends ONLY on abstract interfaces (Inspector, Extractor, OcrEngine, Builder, Verifier).
- All dependencies are injected via constructor.
- ZERO document-type if/else switching.
- ZERO imports of concrete libraries (fitz, docx, etc.).
"""

from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from .context import BuildContext, ExtractResult, InputContext, InspectResult, QaReport
from .interfaces import Builder, Extractor, Inspector, OcrEngine, Verifier


class Pipeline:
    """Orchestrates document conversion using 5 injected stage dependencies."""

    def __init__(
        self,
        inspector: Inspector,
        extractor: Extractor,
        ocr_engine: OcrEngine,
        builder: Builder,
        verifier: Verifier,
    ):
        self.inspector = inspector
        self.extractor = extractor
        self.ocr_engine = ocr_engine
        self.builder = builder
        self.verifier = verifier

    def run(self, ctx: InputContext) -> Tuple[Path, Optional[QaReport]]:
        """Run the full conversion pipeline for the given InputContext.

        Workflow:
          1. Inspect: analyze document structure
          2. Extract: extract text, tables, and page renders
          3. OCR: process image-only pages or crops if needed
          4. Build: assemble DOCX according to config and extracted context
          5. Verify: evaluate against 5 QA axes
        """
        stem = ctx.pdf_path.stem
        stem_temp_dir = ctx.temp_dir / stem
        stem_temp_dir.mkdir(parents=True, exist_ok=True)

        print(f"\n{'='*65}")
        print(f"  [PIPELINE START] Processing: {ctx.pdf_path.name}")
        print(f"  Target DOCX: {ctx.output_docx}")
        print(f"{'='*65}\n")

        # ── STAGE 1: INSPECT ──
        print("  ▶ [Stage 1/5] Inspecting document structure...")
        t0 = time.time()
        inspect_res = self.inspector.inspect(ctx)
        inspect_file = stem_temp_dir / "inspect.json"
        inspect_file.write_text(inspect_res.to_json(), encoding="utf-8")
        print(
            f"  ✓ Inspect complete ({time.time() - t0:.2f}s): "
            f"{inspect_res.page_count} pages, "
            f"image-only={inspect_res.image_only_pages or 'none'}"
        )

        # ── STAGE 2: EXTRACT ──
        print("  ▶ [Stage 2/5] Extracting text, tables, and renders...")
        t0 = time.time()
        extract_res = self.extractor.extract(ctx)
        extract_dir = stem_temp_dir / "extract"
        extract_dir.mkdir(parents=True, exist_ok=True)
        (extract_dir / "extract.json").write_text(extract_res.to_json(), encoding="utf-8")
        print(
            f"  ✓ Extract complete ({time.time() - t0:.2f}s): "
            f"{len(extract_res.text_by_page)} text pages, "
            f"{len(extract_res.tables_by_page)} table pages, "
            f"{len(extract_res.rendered_pages)} rendered pages"
        )

        # ── STAGE 3: OCR (if needed or forced) ──
        pages_to_ocr = inspect_res.image_only_pages
        if ctx.enable_ocr and not pages_to_ocr:
            pages_to_ocr = list(range(1, inspect_res.page_count + 1))

        if pages_to_ocr:
            print(f"  ▶ [Stage 3/5] Running OCR on {len(pages_to_ocr)} page(s)...")
            t0 = time.time()
            ocr_images: list[Path] = []
            for p_num in pages_to_ocr:
                img_path = extract_res.rendered_pages.get(p_num)
                if img_path and Path(img_path).exists():
                    ocr_images.append(Path(img_path))

            ocr_langs = ctx.config.get("ocr", {}).get("langs", ["en", "ar"])
            if ocr_images:
                ocr_out = self.ocr_engine.ocr(ocr_images, ocr_langs)
                extract_res.ocr_results = ocr_out
                ocr_dir = stem_temp_dir / "ocr"
                ocr_dir.mkdir(parents=True, exist_ok=True)
                (ocr_dir / "ocr.json").write_text(
                    json.dumps(ocr_out, ensure_ascii=False, indent=2), encoding="utf-8"
                )
                print(f"  ✓ OCR complete ({time.time() - t0:.2f}s)")
        else:
            print("  ▶ [Stage 3/5] OCR skipped (no image-only pages and --ocr not forced).")

        # ── STAGE 4: BUILD ──
        print("  ▶ [Stage 4/5] Building DOCX document...")
        t0 = time.time()
        ctx.output_docx.parent.mkdir(parents=True, exist_ok=True)
        build_ctx = BuildContext(
            input_context=ctx,
            inspect_result=inspect_res,
            extract_result=extract_res,
            output_docx=ctx.output_docx,
            config=ctx.config,
        )
        built_docx = self.builder.build(build_ctx)
        doc_size = built_docx.stat().st_size if built_docx.exists() else 0
        print(f"  ✓ Build complete ({time.time() - t0:.2f}s): {built_docx} ({doc_size:,} bytes)")

        # ── STAGE 5: VERIFY ──
        qa_report: Optional[QaReport] = None
        if ctx.enable_verify:
            print("  ▶ [Stage 5/5] Running QA verification (5 axes)...")
            t0 = time.time()
            report_path = Path("tests/reports") / f"{stem}_qa.json"
            qa_report = self.verifier.verify(
                original_pdf=ctx.pdf_path,
                docx_path=built_docx,
                report_path=report_path,
                dpi=ctx.config.get("verification", {}).get("dpi_for_render", 150),
                config=ctx.config,
            )
            print(f"  ✓ QA verification finished ({time.time() - t0:.2f}s): {qa_report.verdict}")
        else:
            print("  ▶ [Stage 5/5] Verification skipped per configuration.")

        return built_docx, qa_report
