$ErrorActionPreference = 'Stop'
$d = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\v28_sheet_refs.log'
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

# Read-only: for every sheet being consolidated in Phase 3, count string-literal
# references ("NAME") across all .bas, broken down per file. Anchor = heaviest.
$sheets = @('INPUT','WH_ACTUAL','OB_QUEUE','OB_WH_CACHE','POSTING_RAW','POSTING_SUMMARY','POSTING_CACHE',
            'PLAN','PLAN_SOURCE_LEGS','PLAN_RELEASE',
            'CONFIG','SAP_MAP','SAP_ACTION_MAP','SAP_PROBE_LOG',
            'RUN_HISTORY','RUN_LOG','RUN_SUMMARY',
            'ORDER_QUERY','REGRESSION_AUDIT','POSTING')
$files = Get-ChildItem $d -Filter *.bas -File
foreach ($s in $sheets) {
    $pat = '"' + [regex]::Escape($s) + '"'
    $total = 0
    $perFile = @()
    foreach ($f in $files) {
        $m = @(Select-String -Path $f.FullName -Pattern $pat -AllMatches)
        $cnt = 0; foreach ($x in $m) { $cnt += $x.Matches.Count }
        if ($cnt -gt 0) { $total += $cnt; $perFile += ($f.BaseName + ':' + $cnt) }
    }
    Log ("SHEET=" + $s.PadRight(18) + " TOTAL=" + $total + "  [" + ($perFile -join ', ') + "]")
}
Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
