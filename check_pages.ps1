$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open('d:\AVLANCE\Advocate\KPR_Assignment_II_Client_Project_Report_Advocate_Kannan.docx')
    $pages = $doc.ComputeStatistics(2)
    $words = $doc.ComputeStatistics(0)
    Write-Host "DOCX Statistics -> Exact Page Count: $pages | Exact Word Count: $words"
    $doc.Close([ref]0)
} finally {
    $word.Quit([ref]0)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
