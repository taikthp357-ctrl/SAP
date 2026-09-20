# v28_system_band_discovery.ps1 — Discover all CONFIG/SAP_ACTION_MAP/SAP_PROBE_LOG references
$ErrorActionPreference = 'Stop'
$vbaBase = 'g:\SAP\CF_OB\work\scratch\vba_v28'

$targets = @{
    'CONFIG' = @{
        patterns = @(
            'Worksheets\("CONFIG"\)',
            'Sheets\("CONFIG"\)',
            'CONFIG_SHEET',
            'SHEET_CONFIG'
        )
    }
    'SAP_ACTION_MAP' = @{
        patterns = @(
            'Worksheets\("SAP_ACTION_MAP"\)',
            'Sheets\("SAP_ACTION_MAP"\)',
            'SHEET_ACTION',
            '"SAP_ACTION_MAP"'
        )
    }
    'SAP_PROBE_LOG' = @{
        patterns = @(
            'Worksheets\("SAP_PROBE_LOG"\)',
            'Sheets\("SAP_PROBE_LOG"\)',
            'SP_LOG',
            '"SAP_PROBE_LOG"'
        )
    }
}

foreach ($sheetName in $targets.Keys) {
    Write-Host "`n=== $sheetName ==="
    $allHits = @()

    foreach ($pat in $targets[$sheetName].patterns) {
        $res = Select-String -Path "$vbaBase\*.bas","$vbaBase\*.cls" -Pattern $pat -CaseSensitive:$false
        foreach ($m in $res) {
            $allHits += "$($m.Filename):$($m.LineNumber) | $($m.Line.Trim())"
        }
    }

    $allHits | Sort-Object -Unique | ForEach-Object { Write-Host "  $_" }
    Write-Host "  TOTAL: $($allHits.Count) references"
}

Write-Host "`n=== Dynamic build check ==="
$dynPatterns = @(
    '"CONFIG"',
    '"SAP_ACTION_MAP"',
    '"SAP_PROBE_LOG"',
    '&.*CONFIG',
    '&.*ACTION.*MAP',
    '&.*PROBE.*LOG'
)

foreach ($dp in $dynPatterns) {
    $dh = Select-String -Path "$vbaBase\*.bas","$vbaBase\*.cls" -Pattern $dp -CaseSensitive:$false
    if ($dh.Count -gt 0) {
        Write-Host "`nPattern: $dp"
        $dh | ForEach-Object { Write-Host "  $($_.Filename):$($_.LineNumber) | $($_.Line.Trim())" }
    }
}

Write-Host "`nDISCOVERY_COMPLETE"
