$ErrorActionPreference = 'Stop'
$d = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\v28_sapmap_discovery.log'
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

# Read-only static discovery for the SAP_MAP -> SYSTEM anchor rename (CP4).
# Delicate: module name modCF_SAP_MAP_MON contains substring SAP_MAP, and
# SAP_ACTION_MAP is a lookalike sibling sheet. Prove the quoted-literal rule is
# exact and cannot corrupt either.
$files = Get-ChildItem $d -Filter *.bas -File | Sort-Object Name

# 1) Exact quoted literal "SAP_MAP" (the rename target set), file:line:text
Log '=== (1) QUOTED LITERAL "SAP_MAP" (file:line: text) ==='
$litTotal = 0
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern ([regex]::Escape('"SAP_MAP"')) -AllMatches)
    foreach ($x in $m) { $litTotal++; Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
}
Log ("QUOTED_LITERAL_TOTAL=" + $litTotal)

# 2) Dynamic-construction hazard: is "SAP_MAP" ever assembled from pieces?
Log ''
Log '=== (2) DYNAMIC CONSTRUCTION probes (must be zero) ==='
$dynP=0
foreach ($f in $files) {
    foreach ($pat in @('"SAP"\s*&','&\s*"_MAP"','"SAP_"\s*&','"_MAP"\s*&','&\s*"MAP"')) {
        $m = @(Select-String -Path $f.FullName -Pattern $pat -AllMatches)
        foreach ($x in $m) { $dynPct++; $dynP=1; Log ('DYN ' + $f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
    }
}
Log ("DYNAMIC_HITS=" + $dynP)

# 3) SAP_MAP token WITHOUT surrounding quotes (module name / comments / codename)
Log ''
Log '=== (3) SAP_MAP token WITHOUT quotes (module name modCF_SAP_MAP_MON, comments) ==='
$nonLit=0
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern 'SAP_MAP' -AllMatches)
    foreach ($x in $m) {
        if ($x.Line -notmatch '"SAP_MAP"') { $nonLit++; Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
    }
}
Log ("NON_LITERAL_TOTAL=" + $nonLit)

# 4) Substring-safety: prove "SAP_MAP" (quoted) never sits inside "SAP_ACTION_MAP".
Log ''
Log '=== (4) SAP_ACTION_MAP occurrences (must NOT be touched by the rule) ==='
$actTotal=0
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern ([regex]::Escape('SAP_ACTION_MAP')) -AllMatches)
    foreach ($x in $m) { $actTotal++; Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
}
Log ("SAP_ACTION_MAP_TOTAL=" + $actTotal)

# 5) Collision: does any code already reference a sheet named "SYSTEM"?
Log ''
Log '=== (5) EXISTING "SYSTEM" references (rename-target collision check) ==='
$sysTotal=0
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern ([regex]::Escape('"SYSTEM"')) -AllMatches)
    foreach ($x in $m) { $sysTotal++; Log ($f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
}
Log ("SYSTEM_REF_TOTAL=" + $sysTotal)

# 6) Which of the quoted literals sit on a sheet-ref line (Worksheets/Sheets/Const/Array)?
Log ''
Log '=== (6) Quoted "SAP_MAP" on sheet-ref lines vs data-value lines ==='
$refLine=0; $dataLine=0
foreach ($f in $files) {
    $m = @(Select-String -Path $f.FullName -Pattern ([regex]::Escape('"SAP_MAP"')) -AllMatches)
    foreach ($x in $m) {
        if ($x.Line -match 'Worksheets\(|\.Sheets\(|\bConst\b|Array\(') { $refLine++ }
        else { $dataLine++; Log ('DATA-VALUE ' + $f.BaseName + ':' + $x.LineNumber + ': ' + $x.Line.Trim()) }
    }
}
Log ("SHEETREF_LINES=" + $refLine + " DATAVALUE_LINES=" + $dataLine)

Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
Log ('LOG_WRITTEN=' + $outLog)
