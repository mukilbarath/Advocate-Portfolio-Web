$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $docxPath = 'd:\AVLANCE\Advocate\Project_Technical_Summary_Advocate_Portfolio.docx'
    $pdfPath = 'd:\AVLANCE\Advocate\Project_Technical_Summary_Advocate_Portfolio.pdf'
    Write-Host "Opening $docxPath..."
    $doc = $word.Documents.Open($docxPath, $false, $true)
    Write-Host "Exporting to $pdfPath..."
    $doc.SaveAs([ref]$pdfPath, [ref]17) # 17 is wdFormatPDF
    $doc.Close([ref]0) # 0 is wdDoNotSaveChanges
    Write-Host "Export completed successfully!"
} catch {
    Write-Error $_
} finally {
    $word.Quit([ref]0)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
