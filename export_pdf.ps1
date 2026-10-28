$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $docxPath = 'd:\AVLANCE\Advocate\final iwp report (1).docx'
    $pdfPath = 'd:\AVLANCE\Advocate\final iwp report (1).pdf'
    $doc = $word.Documents.Open($docxPath, $false, $true)
    $doc.SaveAs([ref]$pdfPath, [ref]17) # 17 = wdFormatPDF
    $doc.Close([ref]0)
    Write-Host "Exported PDF successfully to $pdfPath"
} finally {
    $word.Quit([ref]0)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
