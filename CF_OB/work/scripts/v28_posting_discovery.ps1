$ErrorActionPreference = 'Stop'
$d = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\v28_posting_discovery.log'
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

# Read-only static discovery for POSTING removal (CP5). Post-CP4 ground truth.
# Goal: enumerate every SURVIVING direct reference to the POSTING *sheet* so the
# SSR REGION flip + physical delete is provably safe. Siblings POSTING_RAW/
# _SUMMARY/_CACHE are DIFFERENT sheets and MUST be excluded (exact-quoted token
# "POSTING" with trailing " protects against matching "POSTING_RAW" etc).
$files = Get-ChildItem $d -Filter *.bas -File | Sort-Object Name
$tok = '"POSTING"'

# 1) Exact quoted "POSTING" (closing quote) — the real sheet-name references.
Log '=== (1) EXACT QUOTED "POSTING" (file:line: text) ==='
$litTotal = 0
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern ([regex]::Escape($tok)) -AllMatches)
    foreach ($x in $m) { $litTotal++; Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
}
Log ("EXACT_QUOTED_POSTING_TOTAL=" + $litTotal)

# 2) Direct sheet-access Worksheets("POSTING") / .Sheets("POSTING") (must be 0 —
#    Phase 2 redirected all data-binds through SSR).
Log ''
Log '=== (2) DIRECT Worksheets("POSTING") / .Sheets("POSTING") (should be redirected->SSR) ==='
$directTotal = 0
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern 'Worksheets\("POSTING"\)|\.Sheets\("POSTING"\)' -AllMatches)
    foreach ($x in $m) { $directTotal++; Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
}
Log ("DIRECT_SHEET_ACCESS_TOTAL=" + $directTotal)

# 3) The swap-rename coupling: CF_SwapPostingSheets and any .Name = "POSTING..." lines.
Log ''
Log '=== (3) CF_SwapPostingSheets + .Name = assignments touching POSTING ==='
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern 'CF_SwapPostingSheets|\.Name\s*=\s*"POSTING' -AllMatches)
    foreach ($x in $m) { Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
}

# 4) Name-constants whose VALUE is exactly "POSTING" (indirection holders).
Log ''
Log '=== (4) Const = "POSTING" (name-constant holders) ==='
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern 'Const\s+\w+\s+As\s+String\s*=\s*"POSTING"' -AllMatches)
    foreach ($x in $m) { Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
}

# 5) Sibling sheets that must NOT be confused with POSTING (context only).
Log ''
Log '=== (5) POSTING_RAW / POSTING_SUMMARY / POSTING_CACHE (siblings, exclude from delete) ==='
$sibTotal = 0
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern '"POSTING_(RAW|SUMMARY|CACHE)"' -AllMatches)
    foreach ($x in $m) { $sibTotal++; Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
}
Log ("SIBLING_TOTAL=" + $sibTotal)

# 6) SSR indirection anchor (the one legit remaining "POSTING" name, in SSR only).
Log ''
Log '=== (6) SSR_SHEET_NAME / SSR backing const (indirection point) ==='
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern 'SSR_SHEET_NAME|SSR_BACKING\b|SSR_REGION_(SHEET|ROW0|COL0)' -AllMatches)
    foreach ($x in $m) { Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
}

Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
Log ('LOG_WRITTEN=' + $outLog)
