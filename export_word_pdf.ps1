$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Open("C:\Users\ravin\Downloads\RESUM\Ravindra_Chowdary_Jonnagadla_Resume.docx", $false, $true)
$doc.ExportAsFixedFormat("C:\Users\ravin\Downloads\RESUM\word_exported.pdf", 17)
$pages = $doc.ComputeStatistics(2)
Write-Output "Word Computed Pages: $pages"
$doc.Close(0)
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
