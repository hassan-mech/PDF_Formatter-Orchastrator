"""
ocr_pages.py — OCR runner for image-only PDF pages or PNG directories (reusable, CLI-driven)

Replaces: run_win_ocr.ps1 + inspect_ocr_pages.py

Usage:
    python ocr_pages.py -i renders/ -o extracted/ocr/ --engine windows
    python ocr_pages.py -i input/file.pdf -o extracted/ocr/ --engine pytesseract --lang eng
"""

import argparse
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    import fitz  # PyMuPDF — used when input is a PDF
except ImportError:
    fitz = None  # optional if input is already a directory of PNGs


# ---------------------------------------------------------------------------
# OCR engines
# ---------------------------------------------------------------------------

WINDOWS_OCR_PS_TEMPLATE = r"""
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Foundation.IAsyncOperation`1, Windows.Foundation, ContentType=WindowsRuntime]

function Await($WinRtTask, $ResultType) {{
    $asTask = [System.WindowsRuntimeSystemExtensions]::AsTask($WinRtTask)
    $asTask.Wait(-1) | Out-Null
    $asTask.Result
}}

$lang = [Windows.Globalization.Language]::new("{lang}")
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
if (-not $engine) {{
    Write-Error "OCR engine not available for language: {lang}"
    exit 1
}}

$imgFiles = Get-ChildItem -Path "{input_dir}" -Filter "*.png" | Sort-Object Name
$total = $imgFiles.Count
$done = 0
foreach ($imgFile in $imgFiles) {{
    $outFile = Join-Path "{output_dir}" ($imgFile.BaseName + ".txt")
    try {{
        $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($imgFile.FullName)) ([Windows.Storage.StorageFile])
        $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
        $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
        $result.Text | Out-File -FilePath $outFile -Encoding UTF8
        $done++
        Write-Host "  OCR done ($done/$total): $($imgFile.Name)"
    }} catch {{
        Write-Warning "  OCR failed for $($imgFile.Name): $_"
    }}
}}
Write-Host "OCR complete: $done/$total files processed."
"""


def run_windows_ocr(input_dir: Path, output_dir: Path, lang: str) -> int:
    """Write a PowerShell script and run it for Windows OCR. Returns count of processed files."""
    output_dir.mkdir(parents=True, exist_ok=True)
    ps_script = WINDOWS_OCR_PS_TEMPLATE.format(
        lang=lang,
        input_dir=str(input_dir).replace("\\", "\\\\"),
        output_dir=str(output_dir).replace("\\", "\\\\"),
    )
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".ps1", delete=False, encoding="utf-8"
    ) as tmp:
        tmp.write(ps_script)
        tmp_path = Path(tmp.name)

    print(f"  🪟 Running Windows OCR via PowerShell (lang={lang})...")
    try:
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(tmp_path)],
            capture_output=False,
            text=True,
        )
        if result.returncode != 0:
            print(f"  ⚠️  PowerShell returned code {result.returncode}")
    except Exception as exc:
        print(f"  ❌ PowerShell execution failed: {exc}")
    finally:
        tmp_path.unlink(missing_ok=True)

    count = len(list(output_dir.glob("*.txt")))
    return count


def run_pytesseract_ocr(input_dir: Path, output_dir: Path, lang: str) -> int:
    """Run pytesseract OCR on all PNGs in input_dir. Returns count of processed files."""
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        print("[ERROR] pytesseract or Pillow not installed. Run: pip install pytesseract pillow")
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)
    png_files = sorted(input_dir.glob("*.png"))
    count = 0
    for i, png in enumerate(png_files, 1):
        out_txt = output_dir / (png.stem + ".txt")
        try:
            img = Image.open(png)
            text = pytesseract.image_to_string(img, lang=lang)
            out_txt.write_text(text, encoding="utf-8")
            print(f"  ✅ ({i}/{len(png_files)}) {png.name}")
            count += 1
        except Exception as exc:
            print(f"  ❌ {png.name}: {exc}")
    return count


# ---------------------------------------------------------------------------
# PDF → PNGs helper
# ---------------------------------------------------------------------------

def render_pdf_to_pngs(pdf_path: Path, out_dir: Path, dpi: int = 150) -> Path:
    """Render all pages of a PDF to PNG files in out_dir. Returns out_dir."""
    if fitz is None:
        print("[ERROR] PyMuPDF (fitz) required when --input is a PDF. Run: pip install pymupdf")
        sys.exit(1)
    out_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf_path))
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    for i in range(doc.page_count):
        pix = doc[i].get_pixmap(matrix=mat, alpha=False)
        out_file = out_dir / f"{pdf_path.stem}_page_{i+1:03d}.png"
        pix.save(str(out_file))
        print(f"  🖼  Rendered page {i+1}/{doc.page_count} → {out_file.name}")
    doc.close()
    return out_dir


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Run OCR on a directory of PNG images or a PDF file."
    )
    parser.add_argument(
        "-i", "--input", required=True,
        help="Directory of page PNGs or path to a PDF file"
    )
    parser.add_argument(
        "-o", "--output", default="extracted/ocr/",
        help="Output directory for OCR .txt files (default: extracted/ocr/)"
    )
    parser.add_argument(
        "--lang", default="en",
        help="OCR language code (default: en; use 'eng' for pytesseract)"
    )
    parser.add_argument(
        "--engine", choices=["windows", "pytesseract"], default="windows",
        help="OCR engine to use (default: windows)"
    )
    parser.add_argument(
        "--dpi", type=int, default=150,
        help="DPI when rendering PDF pages (default: 150, only used if --input is a PDF)"
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output)

    # Determine input directory of PNGs
    if input_path.is_file() and input_path.suffix.lower() == ".pdf":
        print(f"📄 Input is a PDF — rendering pages to PNGs first...")
        png_dir = output_dir.parent / "renders" / input_path.stem
        input_path = render_pdf_to_pngs(input_path, png_dir, dpi=args.dpi)
    elif input_path.is_dir():
        pass  # already a directory
    else:
        print(f"[ERROR] --input must be a PDF file or a directory of PNGs: {input_path}")
        sys.exit(1)

    png_count = len(list(input_path.glob("*.png")))
    print(f"\n🔤 Running OCR on {png_count} PNG file(s) in '{input_path}' (engine={args.engine})")

    if args.engine == "windows":
        done = run_windows_ocr(input_path, output_dir, args.lang)
    else:
        done = run_pytesseract_ocr(input_path, output_dir, args.lang)

    print(f"\n✅ OCR complete: {done}/{png_count} files processed → {output_dir}")


if __name__ == "__main__":
    main()
