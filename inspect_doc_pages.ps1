$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open('d:\AVLANCE\Advocate\final iwp report (1).docx')
    $pages = $doc.ComputeStatistics(2)
    Write-Host "Total Pages: $pages"
    
    # Iterate through pages or headings to see distribution
    # Let's inspect headings and their page numbers
    $headings = $doc.Paragraphs
    for ($i = 1; $i -le $headings.Count; $i++) {
        $p = $headings.Item($i)
        $txt = $p.Range.Text.Trim()
        if ($txt -match '^[0-9]+\.' -or $txt -match '^[0-9]+\.[0-9]+' -or $txt -eq 'Table of Contents' -or $txt -eq 'CANDIDATE’S DECLARATION' -or $txt -match 'Annexure') {
            $pageNum = $p.Range.Information(3) # wdActiveEndPageNumber
            Write-Host ("Page " + $pageNum + ": " + $txt.Substring(0, [Math]::Min(60, $txt.Length)))
        }
    }
    $doc.Close([ref]0)
} finally {
    $word.Quit([ref]0)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
