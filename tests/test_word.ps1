$word = New-Object -ComObject Word.Application
$word.Visible = $false
Write-Host "Word Version: $($word.Version)"
$word.Quit()
