$ErrorActionPreference = 'Stop'
$d = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\v28_whactual_discovery.log'
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

# Read-only static discovery for the WH_ACTUAL -> DATA_CORE anchor rename.
# Goal: enumerate EVERY reference site (literal + likely dynamic) so the rename
# is mechanical and provably complete, not a guess.
$files = Get-ChildItem $d -Filter *.bas -File | Sort-Object Name

# 1) All string-literal "WH_ACTUAL" occurrences, file:line:text
Log '=== (1) STRING-LITERAL "WH_ACTUAL" (file:line: text) ==='
$litTotal = 0
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern ([regex]::Escape('"WH_ACTUAL"')) -AllMatches)
    foreach ($x in $m) {
        $litTotal++
        Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim())
    }
}
Log ("LITERAL_TOTAL=" + $litTotal)

# 2) Any WH_ACTUAL token NOT inside the double-quoted literal (dynamic refs,
#    CodeName usage, constants, comments) so nothing is missed.
Log ''
Log '=== (2) WH_ACTUAL token WITHOUT surrounding quotes (potential dynamic/const/codename) ==='
$dynTotal = 0
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern 'WH_ACTUAL' -AllMatches)
    foreach ($x in $m) {
        if ($x.Line -notmatch '"WH_ACTUAL"') {
            $dynTotal++
            Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim())
        }
    }
}
Log ("NON_LITERAL_TOTAL=" + $dynTotal)

# 3) Sheet-name constants / Const decls that could carry the name indirectly.
Log ''
Log '=== (3) Const / sheet-name declarations mentioning WH or ACTUAL or SHEET_ ==='
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern '(?i)const\s+\w*\s*=\s*".*(WH|ACTUAL).*"' -AllMatches)
    foreach ($x in $m) { Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
}

# 4) CodeName-based access (Sheetnn.) is rename-immune but list any .CodeName usage.
Log ''
Log '=== (4) .CodeName usage (rename-immune access to confirm) ==='
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern '\.CodeName' -AllMatches)
    foreach ($x in $m) { Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
}

Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
