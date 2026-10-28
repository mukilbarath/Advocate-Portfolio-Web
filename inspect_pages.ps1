$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open('d:\AVLANCE\Advocate\KPR_Assignment_II_Client_Project_Report_Advocate_Kannan.docx')
    $numPages = $doc.ComputeStatistics(2)
    Write-Host "Total pages: $numPages"
    for ($i = 1; $i -le $numPages; $i++) {
        $r1 = $doc.GoTo(1, 1, $i)
        if ($i -lt $numPages) {
            $r2 = $doc.GoTo(1, 1, $i + 1)
            $text = $doc.Range($r1.Start, $r2.Start).Text
        } else {
            $text = $doc.Range($r1.Start, $doc.Content.End).Text
        }
        $words = ($text -split '\s+' | Where-Object { $_ -ne '' }).Count
        Write-Host "Page $i : $words words"
    }
    $doc.Close([ref]0)
} finally {
    $word.Quit([ref]0)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
