"""
parallel_ocr.py — Parallel OCR engine for image-only PDF pages

Dispatches image rendering and OCR to a ThreadPoolExecutor (I/O-friendly).
Falls back gracefully between engines: Windows OCR → pytesseract → skip.

Usage:
    python parallel_ocr.py -i input/scanned.pdf -o extracted/ocr/ --workers 8
    python parallel_ocr.py -i renders/pages/ -o extracted/ocr/ --engine pytesseract
"""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

try:
    import fitz
except ImportError:
    raise ImportError("PyMuPDF required: pip install pymupdf")


# ────────────────────────────────────────────────────────────────────────────
# Data
# ────────────────────────────────────────────────────────────────────────────

@dataclass
class OcrTask:
    source: str          # PNG file path OR (pdf_path + page_index encoded)
    page_number: int     # 1-based
    output_path: str     # where to write the .txt result
    engine: str = "windows"
    lang: str = "en"


@dataclass
class OcrResult:
    page_number: int
    success: bool
    text: str = ""
    line_count: int = 0
    elapsed_ms: float = 0.0
    error: str = ""


# ────────────────────────────────────────────────────────────────────────────
# Per-page OCR worker (thread-safe — I/O bound, no GIL issues)
# ────────────────────────────────────────────────────────────────────────────

# Embedded PowerShell snippet for Windows OCR (called per image)
_PS_OCR_TEMPLATE = r"""
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Globalization.Language,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Storage.StorageFile,Windows.Foundation,ContentType=WindowsRuntime]

$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.IsGenericMethod } |
    Select-Object -First 1

function Await($op, $t) {
    $m = $asTask.MakeGenericMethod($t)
    $task = $m.Invoke($null, @($op))
    $task.Wait()
    return $task.Result
}

$lang   = [Windows.Globalization.Language]::new("__LANG__")
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
if (-not $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }

$file   = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync("__IMG_PATH__")) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$dec    = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bmp    = Await ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$res    = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
$res.Text
"""


def _ocr_with_windows(img_path: str, lang: str = "en") -> str:
    """Run Windows OCR on a PNG via PowerShell subprocess."""
    abs_path = str(Path(img_path).resolve())
    ps_code = _PS_OCR_TEMPLATE.replace("__IMG_PATH__", abs_path.replace("\\", "\\\\")).replace("__LANG__", lang)
    with tempfile.NamedTemporaryFile(suffix=".ps1", mode="w", encoding="utf-8", delete=False) as f:
        f.write(ps_code)
        ps_file = f.name
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps_file],
            capture_output=True, text=True, timeout=60,
        )
        return result.stdout.strip()
    except subprocess.TimeoutExpired:
        return ""
    finally:
        Path(ps_file).unlink(missing_ok=True)



def _ocr_with_pytesseract(img_path: str, lang: str = "eng") -> str:
    """Run tesseract OCR on a PNG."""
    try:
        import pytesseract
        from PIL import Image
        img = Image.open(img_path)
        return pytesseract.image_to_string(img, lang=lang)
    except ImportError:
        return ""


def _render_pdf_page_to_png(pdf_path: str, page_index: int, dpi: int = 200) -> Optional[str]:
    """Render a single PDF page to a temp PNG and return its path."""
    try:
        doc = fitz.open(pdf_path)
        page = doc[page_index]
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        doc.close()
        tmp = tempfile.NamedTemporaryFile(suffix=f"_p{page_index+1}.png", delete=False)
        pix.save(tmp.name)
        return tmp.name
    except Exception:
        return None


def _process_ocr_task(task: OcrTask, dpi: int = 200) -> OcrResult:
    """Worker: render page (if PDF source) → run OCR → save txt → return result."""
    t0 = time.perf_counter()
    result = OcrResult(page_number=task.page_number, success=False)
    tmp_png: Optional[str] = None

    try:
        # Resolve source image
        src = Path(task.source)
        if src.suffix.lower() == ".pdf":
            # "source" encoded as "pdf_path::page_index"
            parts = task.source.split("::")
            pdf_p, page_idx = parts[0], int(parts[1])
            tmp_png = _render_pdf_page_to_png(pdf_p, page_idx, dpi=dpi)
            if not tmp_png:
                raise RuntimeError(f"Could not render page {task.page_number}")
            img_path = tmp_png
        else:
            img_path = str(src)

        # Run OCR
        if task.engine == "windows":
            text = _ocr_with_windows(img_path, lang=task.lang)
        elif task.engine == "pytesseract":
            text = _ocr_with_pytesseract(img_path, lang=task.lang if task.lang != "en" else "eng")
        else:
            text = ""

        # Save output
        out = Path(task.output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text, encoding="utf-8")

        result.text = text
        result.line_count = len([l for l in text.splitlines() if l.strip()])
        result.success = True

    except Exception as exc:
        result.error = str(exc)
    finally:
        if tmp_png:
            Path(tmp_png).unlink(missing_ok=True)

    result.elapsed_ms = (time.perf_counter() - t0) * 1000
    return result


# ────────────────────────────────────────────────────────────────────────────
# High-level Parallel OCR API
# ────────────────────────────────────────────────────────────────────────────

class ParallelOCREngine:
    """
    Parallel OCR engine — processes multiple pages concurrently via threads.

    Example
    -------
    engine = ParallelOCREngine("input/scanned.pdf", output_dir="extracted/ocr/",
                                workers=8, engine="windows")
    results = engine.run(page_indices=[0,1,2,5])  # or None for all
    """

    def __init__(
        self,
        source: str | Path,           # PDF path or directory of PNGs
        output_dir: str | Path,
        workers: int = 4,
        engine: str = "windows",      # 'windows' | 'pytesseract'
        lang: str = "en",
        dpi: int = 200,
    ):
        self.source = Path(source)
        self.output_dir = Path(output_dir)
        self.workers = workers
        self.engine_name = engine
        self.lang = lang
        self.dpi = dpi

        self.output_dir.mkdir(parents=True, exist_ok=True)

    def _build_tasks(self, page_indices: Optional[List[int]] = None) -> List[OcrTask]:
        tasks = []
        if self.source.is_dir():
            # Directory of PNGs
            pngs = sorted(self.source.glob("*.png"))
            for i, png in enumerate(pngs):
                if page_indices is not None and i not in page_indices:
                    continue
                out_txt = self.output_dir / (png.stem + "_ocr.txt")
                tasks.append(OcrTask(
                    source=str(png),
                    page_number=i + 1,
                    output_path=str(out_txt),
                    engine=self.engine_name,
                    lang=self.lang,
                ))
        else:
            # PDF: build encoded source strings
            doc = fitz.open(str(self.source))
            total = doc.page_count
            doc.close()
            indices = page_indices if page_indices is not None else list(range(total))
            for idx in indices:
                out_txt = self.output_dir / f"page_{idx+1:03d}_ocr.txt"
                tasks.append(OcrTask(
                    source=f"{self.source}::{idx}",
                    page_number=idx + 1,
                    output_path=str(out_txt),
                    engine=self.engine_name,
                    lang=self.lang,
                ))
        return tasks

    def run(self, page_indices: Optional[List[int]] = None, verbose: bool = True) -> List[OcrResult]:
        tasks = self._build_tasks(page_indices)
        if not tasks:
            print("⚠️  No tasks to run.")
            return []

        results: List[Optional[OcrResult]] = [None] * len(tasks)
        t_start = time.perf_counter()

        if verbose:
            print(f"\n🔍 ParallelOCREngine starting")
            print(f"   Source   : {self.source}")
            print(f"   Pages    : {len(tasks)}")
            print(f"   Engine   : {self.engine_name}")
            print(f"   Workers  : {self.workers}")
            print(f"   DPI      : {self.dpi}")
            print(f"   Output   : {self.output_dir}")

        with ThreadPoolExecutor(max_workers=self.workers) as pool:
            future_to_idx = {
                pool.submit(_process_ocr_task, task, self.dpi): i
                for i, task in enumerate(tasks)
            }
            done_count = 0
            for future in as_completed(future_to_idx):
                i = future_to_idx[future]
                results[i] = future.result()
                done_count += 1
                if verbose:
                    r = results[i]
                    status = "✅" if r.success else "❌"
                    print(
                        f"  {status} Page {r.page_number:>4}  "
                        f"lines={r.line_count:<5}  {r.elapsed_ms:.0f}ms"
                        + (f"  ERR: {r.error[:60]}" if r.error else ""),
                        flush=True,
                    )

        ordered = [r for r in results if r is not None]
        ordered.sort(key=lambda r: r.page_number)

        elapsed = time.perf_counter() - t_start
        succeeded = sum(1 for r in ordered if r.success)

        if verbose:
            print(f"\n✅ OCR complete in {elapsed:.2f}s")
            print(f"   Succeeded : {succeeded} / {len(ordered)}")
            total_lines = sum(r.line_count for r in ordered)
            print(f"   Total lines extracted: {total_lines}")

        return ordered


# ────────────────────────────────────────────────────────────────────────────
# CLI
# ────────────────────────────────────────────────────────────────────────────

def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="parallel_ocr.py — parallel OCR engine for image-only PDF pages",
    )
    parser.add_argument("-i", "--input", required=True, help="PDF path or directory of PNGs")
    parser.add_argument("-o", "--output", default="extracted/ocr", help="Output directory for txt files")
    parser.add_argument("--workers", type=int, default=4, help="Parallel threads (default: 4)")
    parser.add_argument("--engine", choices=["windows", "pytesseract"], default="windows")
    parser.add_argument("--lang", default="en", help="OCR language (default: en)")
    parser.add_argument("--dpi", type=int, default=200, help="Render DPI (default: 200)")
    parser.add_argument("--pages", default=None, help="Comma-separated 1-based page numbers (default: all)")
    args = parser.parse_args()

    page_indices = None
    if args.pages:
        page_indices = [int(p.strip()) - 1 for p in args.pages.split(",")]

    engine = ParallelOCREngine(
        source=args.input,
        output_dir=args.output,
        workers=args.workers,
        engine=args.engine,
        lang=args.lang,
        dpi=args.dpi,
    )
    engine.run(page_indices=page_indices, verbose=True)


if __name__ == "__main__":
    main()
