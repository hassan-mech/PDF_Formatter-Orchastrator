param(
    [string]$docxPath = "D:\Hassan\Amouna\Hasssan&Eman\Approved Judgment_Notarised+Legalised_SN 193344 (1)-Non-Parsable-en-US#SRECFMT_DXBRBK#.docx",
    [string]$pdfPath = "D:\Hassan\Amouna\Hasssan&Eman\generated_output.pdf"
)

$docxPath = (Resolve-Path $docxPath).Path
$pdfDir = Split-Path -Parent $pdfPath
if ($pdfDir -and -not (Test-Path $pdfDir)) {
    New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($docxPath)
    # 17 is wdExportFormatPDF
    $doc.ExportAsFixedFormat($pdfPath, 17)
    $doc.Close($false)
    Write-Host "Exported to $pdfPath successfully"
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
