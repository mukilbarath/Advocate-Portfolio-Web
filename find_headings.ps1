$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open('d:\AVLANCE\Advocate\KPR_Assignment_II_Client_Project_Report_Advocate_Kannan.docx')
    foreach ($p in $doc.Paragraphs) {
        $txt = $p.Range.Text.Trim()
        if ($txt -match '^\d+\.\s+[A-Z]') {
            $page = $p.Range.Information(3) # wdActiveEndPageNumber
            Write-Host ('{0} -> Page {1}' -f $txt, $page)
        }
    }
    $doc.Close([ref]0)
} finally {
    $word.Quit([ref]0)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
