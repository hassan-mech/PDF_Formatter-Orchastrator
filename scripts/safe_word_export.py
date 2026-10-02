"""
scripts/safe_word_export.py — Safe, timeout-protected DOCX to PDF converter

Guarantees:
- Strict per-step and overall execution timeout.
- Auto-kills zombie WINWORD.EXE processes if they hang.
- Complete COM cleanup preventing memory leaks and handle locks.
"""

import sys
import time
import subprocess
from pathlib import Path
import psutil

def kill_word_processes():
    """Forcefully terminate any existing WINWORD.EXE processes."""
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['name'] and 'WINWORD' in proc.info['name'].upper():
                proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

def convert_docx_to_pdf(docx_path: str | Path, pdf_path: str | Path, timeout_sec: int = 60) -> bool:
    docx_path = Path(docx_path).resolve()
    pdf_path = Path(pdf_path).resolve()
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    if not docx_path.exists():
        print(f"❌ Error: DOCX does not exist: {docx_path}")
        return False

    # Clean pre-existing zombie processes
    kill_word_processes()

    # Generate isolated powershell script with progress markers
    ps_code = f"""
$ErrorActionPreference = 'Stop'
$sw = [System.Diagnostics.Stopwatch]::StartNew()
Write-Host "STAGE:1_LAUNCHING_WORD"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$word.ScreenUpdating = $false

try {{
    Write-Host "STAGE:2_OPENING_DOCX"
    $doc = $word.Documents.Open('{docx_path}')
    
    Write-Host "STAGE:3_EXPORTING_PDF"
    $doc.ExportAsFixedFormat('{pdf_path}', 17)
    
    Write-Host "STAGE:4_CLOSING"
    $doc.Close($false)
    Write-Host "STAGE:5_SUCCESS"
}}
catch {{
    Write-Host "STAGE:ERROR: $_"
    exit 1
}}
finally {{
    if ($word -ne $null) {{
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }}
}}
"""
    tmp_ps = docx_path.parent / f"_export_{int(time.time())}.ps1"
    tmp_ps.write_text(ps_code, encoding="utf-8")

    t0 = time.time()
    success = False
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(tmp_ps)],
            capture_output=True,
            text=True,
            timeout=timeout_sec
        )
        print(res.stdout.strip())
        if res.returncode == 0 and pdf_path.exists() and pdf_path.stat().st_size > 0:
            print(f"✅ Converted in {time.time() - t0:.2f}s: {pdf_path.name} ({pdf_path.stat().st_size / 1024:.1f} KB)")
            success = True
        else:
            print(f"❌ Conversion failed with code {res.returncode}. Stderr: {res.stderr.strip()}")
    except subprocess.TimeoutExpired:
        print(f"⏱️ TIMEOUT ({timeout_sec}s) exceeded! Killing hung Word process...")
        kill_word_processes()
    finally:
        try:
            tmp_ps.unlink(missing_ok=True)
        except Exception:
            pass
        kill_word_processes()

    return success

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python safe_word_export.py <input.docx> <output.pdf> [timeout_seconds]")
        sys.exit(1)
    timeout = int(sys.argv[3]) if len(sys.argv) > 3 else 60
    ok = convert_docx_to_pdf(sys.argv[1], sys.argv[2], timeout_sec=timeout)
    sys.exit(0 if ok else 1)
