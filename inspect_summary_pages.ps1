$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $docxPath = 'd:\AVLANCE\Advocate\Project_Technical_Summary_Advocate_Portfolio.docx'
    $doc = $word.Documents.Open($docxPath, $false, $true)
    $pageCount = $doc.ComputeStatistics(2)
    Write-Host "Total Pages: $pageCount"
    
    # Iterate through paragraphs to find headings and their page numbers
    foreach ($p in $doc.Paragraphs) {
        $txt = $p.Range.Text.Trim()
        if ($txt -match '^[0-9]\.' -or $txt -match '^Figure' -or $txt -match '^KPR' -or $txt -match '^HIGH COURT') {
            $pageNum = $p.Range.Information(1) # 1 = wdActiveEndPageNumber
            Write-Host "Page $pageNum : $txt"
        }
    }
    $doc.Close([ref]0)
} finally {
    $word.Quit([ref]0)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
