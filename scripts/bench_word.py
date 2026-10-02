import time
import subprocess
from pathlib import Path

docx = Path(r"output\263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.docx").resolve()
pdf = Path(r"temp\test_bench.pdf").resolve()

ps_lines = [
    "$sw = [System.Diagnostics.Stopwatch]::StartNew()",
    "Write-Host \"[Start: $($sw.ElapsedMilliseconds) ms]\"",
    "$w = New-Object -ComObject Word.Application",
    "$w.Visible = $false",
    "$w.DisplayAlerts = 0",
    "Write-Host \"[Word Launched: $($sw.ElapsedMilliseconds) ms]\"",
    f"$d = $w.Documents.Open('{docx}')",
    "Write-Host \"[Doc Opened: $($sw.ElapsedMilliseconds) ms]\"",
    f"$d.ExportAsFixedFormat('{pdf}', 17)",
    "Write-Host \"[PDF Exported: $($sw.ElapsedMilliseconds) ms]\"",
    "$d.Close($false)",
    "$w.Quit()",
    "[System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null",
    "Write-Host \"[Finished: $($sw.ElapsedMilliseconds) ms]\""
]

bench_ps1 = Path("temp/bench.ps1")
bench_ps1.write_text("\n".join(ps_lines), encoding="utf-8")

t0 = time.time()
try:
    res = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(bench_ps1)],
        capture_output=True,
        text=True,
        timeout=60
    )
    print("Returncode:", res.returncode)
    print("STDOUT:\n", res.stdout)
    print("STDERR:\n", res.stderr)
except subprocess.TimeoutExpired:
    print("TIMEOUT at 60s!")
print(f"Elapsed: {time.time() - t0:.2f}s")
