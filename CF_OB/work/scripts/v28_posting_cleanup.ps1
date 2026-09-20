# v28_posting_cleanup.ps1 — Neutralize swap + verify zero static refs
$ErrorActionPreference = 'Stop'
$vbaStage = 'g:\SAP\CF_OB\work\scratch\vba_posting_region'

# ═══════════════════════════════════════════════════════════════════════════════
# 1. Neutralize CF_SwapPostingSheets (no-op in REGION mode)
# ═══════════════════════════════════════════════════════════════════════════════
$postPath = Join-Path $vbaStage 'modCF_Posting.bas'
$postContent = Get-Content $postPath -Raw -Encoding UTF8

# Find CF_SwapPostingSheets and replace with no-op
$oldSwap = @'
Public Sub CF_SwapPostingSheets(ByVal direction As String)
    Dim postingSheet As Worksheet, rawSheet As Worksheet, tempName As String
    On Error Resume Next
    Set postingSheet = ThisWorkbook.Worksheets("POSTING")
    Set rawSheet = ThisWorkbook.Worksheets("POSTING_RAW")
    On Error GoTo 0
    If postingSheet Is Nothing Or rawSheet Is Nothing Then Exit Sub

    tempName = "TEMP_SWAP_" & Format$(Now, "yyyymmddhhnnss")
    postingSheet.Name = tempName
    rawSheet.Name = "POSTING"
    ThisWorkbook.Worksheets(tempName).Name = "POSTING_RAW"
End Sub
'@

$newSwap = @'
Public Sub CF_SwapPostingSheets(ByVal direction As String)
    ' V2.8 Phase4: POSTING moved to REGION band on DATA_CORE; swap logic obsolete.
    ' Kept as no-op stub for backward compatibility with callers.
End Sub
'@

if ($postContent -match [regex]::Escape($oldSwap)) {
    $postContent = $postContent.Replace($oldSwap, $newSwap)
    Write-Host "NEUTRALIZED_CF_SwapPostingSheets"
} else {
    Write-Host "WARNING: CF_SwapPostingSheets pattern not found (may be already modified)"
}

Set-Content $postPath -Value $postContent -Encoding UTF8 -NoNewline

# ═══════════════════════════════════════════════════════════════════════════════
# 2. Verify zero static refs to Worksheets("POSTING") / Sheets("POSTING")
# ═══════════════════════════════════════════════════════════════════════════════
Write-Host "`n=== STATIC REF VERIFICATION ==="

$postingRefs = @()
$patterns = @('Worksheets\("POSTING"\)', 'Sheets\("POSTING"\)')

foreach ($pat in $patterns) {
    $res = Select-String -Path "$vbaStage\*.bas","$vbaStage\*.cls" -Pattern $pat -CaseSensitive:$false
    foreach ($m in $res) {
        # Exclude SSR_SHEET_NAME constant definition and RT_Log/WriteLog string literals
        $line = $m.Line.Trim()
        if ($line -notmatch 'Const.*SSR_SHEET_NAME' -and
            $line -notmatch 'RT_Log|WriteLog' -and
            $line -notmatch '^\s*''') {  # exclude comments
            $postingRefs += "$($m.Filename):$($m.LineNumber) | $line"
        }
    }
}

if ($postingRefs.Count -eq 0) {
    Write-Host "ZERO_STATIC_REFS=TRUE (all consumers migrated to SSR)"
} else {
    Write-Host "WARNING: Found $($postingRefs.Count) remaining static refs:"
    $postingRefs | ForEach-Object { Write-Host "  $_" }
}

Write-Host "`nCLEANUP_COMPLETE"
