$docxPath = "D:\Hassan\Amouna\Hasssan&Eman\output\263116 SDDD-Non-Parsable-en-US#FMT_GYRCPJ#-Non-Parsable-en-US#FPREP_DXCGLP#.docx"
$pdfOutPath = "D:\Hassan\Amouna\Hasssan&Eman\output_word_exported.pdf"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($docxPath)
    # Force pagination
    $pageCount = $doc.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)
    Write-Host "Page count in Word: $pageCount"
    
    # Export to PDF (wdFormatPDF = 17)
    $doc.SaveAs([ref]$pdfOutPath, [ref]17)
    Write-Host "Exported to $pdfOutPath"
    $doc.Close([ref]$false)
} finally {
    $word.Quit()
}
