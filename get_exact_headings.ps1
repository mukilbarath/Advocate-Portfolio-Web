$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open('d:\AVLANCE\Advocate\final iwp report (1).docx')
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $t = $p.Range.Text.Trim()
        if ($t -match '^[0-9]+\.\s+[A-Za-z]' -or $t -match '^CANDIDATE' -or $t -match '^CLIENT' -or $t -match '^ACKNOWLEDGEMENT' -or $t -match '^ABSTRACT' -or $t -match '^INDEX' -or $t -match '^Annexure') {
            $pageNum = $p.Range.Information(3) # wdActiveEndPageNumber
            Write-Host "Page $pageNum : $($t.Substring(0, [Math]::Min(65, $t.Length)))"
        }
    }
    $doc.Close([ref]0)
} finally {
    $word.Quit([ref]0)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
