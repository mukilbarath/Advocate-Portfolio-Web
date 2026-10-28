$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open('d:\AVLANCE\Advocate\final iwp report (1).docx')
    $pages = $doc.ComputeStatistics(2)
    Write-Host "Total Pages: $pages"
    
    for ($p = 1; $p -le $pages; $p++) {
        $range = $doc.Range()
        $start = $doc.GoTo(1, 1, $p).Start # wdGoToPage, wdGoToAbsolute
        if ($p -lt $pages) {
            $end = $doc.GoTo(1, 1, $p + 1).Start
        } else {
            $end = $doc.Content.End
        }
        $pageRange = $doc.Range($start, $end)
        $text = $pageRange.Text.Trim()
        $firstLine = if ($text.Length -gt 0) { $text.Split([char]13)[0].Trim() } else { "[EMPTY PAGE]" }
        $wordCount = $pageRange.ComputeStatistics(0)
        Write-Host "Page $p : Words = $wordCount | First Line: $($firstLine.Substring(0, [Math]::Min(70, $firstLine.Length)))"
    }
    $doc.Close([ref]0)
} finally {
    $word.Quit([ref]0)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
