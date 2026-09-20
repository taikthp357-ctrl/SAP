# v28_runhistory_discovery.ps1 — Discover RUN_LOG + RUN_SUMMARY references for band merge
$ErrorActionPreference = 'Stop'
$vbaBase = 'g:\SAP\CF_OB\work\scratch\vba_v28'

Write-Host "`n=== RUN_LOG ==="
$rlHits = @()
$rlPatterns = @('Worksheets\("RUN_LOG"\)', 'Sheets\("RUN_LOG"\)', '"RUN_LOG"')
foreach ($pat in $rlPatterns) {
    $res = Select-String -Path "$vbaBase\*.bas","$vbaBase\*.cls" -Pattern $pat -CaseSensitive:$false
    foreach ($m in $res) {
        $rlHits += "$($m.Filename):$($m.LineNumber) | $($m.Line.Trim())"
    }
}
$rlHits | Sort-Object -Unique | ForEach-Object { Write-Host "  $_" }
Write-Host "  TOTAL: $($rlHits.Count) references"

Write-Host "`n=== RUN_SUMMARY ==="
$rsHits = @()
$rsPatterns = @('Worksheets\("RUN_SUMMARY"\)', 'Sheets\("RUN_SUMMARY"\)', '"RUN_SUMMARY"')
foreach ($pat in $rsPatterns) {
    $res = Select-String -Path "$vbaBase\*.bas","$vbaBase\*.cls" -Pattern $pat -CaseSensitive:$false
    foreach ($m in $res) {
        $rsHits += "$($m.Filename):$($m.LineNumber) | $($m.Line.Trim())"
    }
}
$rsHits | Sort-Object -Unique | ForEach-Object { Write-Host "  $_" }
Write-Host "  TOTAL: $($rsHits.Count) references"

Write-Host "`n=== Dynamic build check ==="
$dynPatterns = @('"RUN"', '"LOG"', '"SUMMARY"', '&.*RUN', '&.*LOG', '&.*SUMMARY')
foreach ($dp in $dynPatterns) {
    $dh = Select-String -Path "$vbaBase\*.bas","$vbaBase\*.cls" -Pattern $dp -CaseSensitive:$false | Select-Object -First 3
    if ($dh.Count -gt 0) {
        Write-Host "`nPattern: $dp (showing first 3)"
        $dh | ForEach-Object { Write-Host "  $($_.Filename):$($_.LineNumber) | $($_.Line.Trim())" }
    }
}

Write-Host "`nDISCOVERY_COMPLETE"
