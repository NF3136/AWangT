param([Parameter(Mandatory=$true)][string]$Source,
      [Parameter(Mandatory=$true)][string]$Destination,
      [Parameter(Mandatory=$true)][string]$Report)
$ErrorActionPreference = 'Stop'
$wordConverter = $null
$wordInput = $null
$wordCheck = $null
$previousUpdateLinks = $null
try {
    $wordConverter = New-Object -ComObject Word.Application
    $wordConverter.Visible = $false
    $wordConverter.DisplayAlerts = 0
    $wordConverter.AutomationSecurity = 3
    $previousUpdateLinks = $wordConverter.Options.UpdateLinksAtOpen
    $wordConverter.Options.UpdateLinksAtOpen = $false
    # Read-only, no recent-files entry, no macros, no external-link refresh.
    $wordInput = $wordConverter.Documents.Open($Source, $false, $true, $false, 'HDX_NoPassword')
    $beforeText = $wordInput.Content.Text
    $beforeCounts = @($wordInput.Tables.Count, $wordInput.InlineShapes.Count,
        $wordInput.Shapes.Count, $wordInput.Footnotes.Count, $wordInput.Endnotes.Count)
    $wordInput.SaveAs2($Destination, 12)
    $wordInput.Close(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordInput)
    $wordInput = $null
    $wordCheck = $wordConverter.Documents.Open($Destination, $false, $true, $false)
    $afterCounts = @($wordCheck.Tables.Count, $wordCheck.InlineShapes.Count,
        $wordCheck.Shapes.Count, $wordCheck.Footnotes.Count, $wordCheck.Endnotes.Count)
    $textIntact = $beforeText -ceq $wordCheck.Content.Text
    $countsIntact = ($beforeCounts -join ',') -ceq ($afterCounts -join ',')
    @{ converter='Microsoft Word'; main_story_text_intact=$textIntact;
       object_counts_intact=$countsIntact; before=$beforeCounts; after=$afterCounts;
       visual_review='pending' } | ConvertTo-Json | Set-Content -LiteralPath $Report -Encoding UTF8
    if (-not $textIntact -or -not $countsIntact) {
        throw 'DOC to DOCX verification failed: text or object counts changed.'
    }
} finally {
    if ($wordCheck) { $wordCheck.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordCheck) }
    if ($wordInput) { $wordInput.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordInput) }
    if ($wordConverter) {
        if ($null -ne $previousUpdateLinks) { $wordConverter.Options.UpdateLinksAtOpen = $previousUpdateLinks }
        $wordConverter.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordConverter)
    }
}
