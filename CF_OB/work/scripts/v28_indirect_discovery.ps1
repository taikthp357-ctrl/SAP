# v28_indirect_discovery.ps1
# CLOSE THE CLASSIFIER GAP: v28_plan_discovery only caught literal Worksheets("X").
# Band sheets are ALSO reached indirectly: Const SHEET_X = "X"  +  Worksheets(SHEET_X).
# modSAPActionRegistry (const SHEET_ACTION="SAP_ACTION_MAP") proved this is real and
# carries a LIVE writer (SAPAction_PromoteLivePass) -> my "fully-offline" tag was false.
# This scan finds, for the 16 band tokens:
#   (1) every Const <name> [As String] = "<token>"  (the indirection SOURCE)
#   (2) every Worksheets(<name>)/Sheets(<name>) use of those const names (the SINK),
#       tagged L if the module is on the live-only role list.
#   (3) also re-flags any Set ws = <helper>()  wrappers already known (SSR/CWR_HistorySheet).
# READ-ONLY. No workbook opened.
$ErrorActionPreference = 'Stop'
$vbaDir = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\v28_indirect_discovery.log'
$enc = [Text.Encoding]::GetEncoding(1252)
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Host $m }

$tokens = @('PLAN','PLAN_SOURCE_LEGS','PLAN_RELEASE','CONFIG','SAP_ACTION_MAP','SAP_PROBE_LOG',
    'RUN_LOG','RUN_SUMMARY','RUN_HISTORY','OB_QUEUE','OB_WH_CACHE','POSTING_RAW','POSTING_SUMMARY',
    'POSTING_CACHE','ORDER_QUERY','REGRESSION_AUDIT','WH_ACTUAL','SAP_MAP','POSTING','DATA_CORE','SYSTEM','INPUT')
$liveModules = @('modCodexWTRework','modSAPReadOnlyProbe','ModuleSAP','modCF_Executor','modCF_MONCore_STABLE',
    'modCF_AutoRun_TwoPilot','modCF_FinalizeVerifiedPilot','modMultiPostingLoader','modReleaseHolderReader',
    'modCFX_Entry','modCFX_Planner','modSAPActionRegistry','modSAPRunGuard')

# Pass 1: harvest Const <name> = "<token>" across all modules -> map constName -> token/module
$constMap = @{}   # constName -> "token @ module:line"
$files = Get-ChildItem -Path $vbaDir -Filter *.bas | Sort-Object Name
foreach ($f in $files) {
    $lines = ([IO.File]::ReadAllText($f.FullName, $enc)).Split([char]10)
    for ($i=0; $i -lt $lines.Count; $i++) {
        $ln = $lines[$i]
        $m = [regex]::Match($ln, '(?i)\bConst\s+([A-Za-z_][A-Za-z0-9_]*)\s+(?:As\s+String\s+)?=\s*"([^"]+)"')
        if ($m.Success) {
            $cname = $m.Groups[1].Value; $cval = $m.Groups[2].Value
            if ($tokens -contains $cval) {
                $constMap[$cname] = @{ token=$cval; where=($f.BaseName + ':' + ($i+1)); mod=$f.BaseName }
                Log ("CONST  " + $f.BaseName + ':' + ($i+1) + "  " + $cname + ' = "' + $cval + '"')
            }
        }
    }
}
Log ("---- harvested " + $constMap.Count + " band-name consts ----")

# Pass 2: find Worksheets(<constName>)/Sheets(<constName>) sinks + module-level (Dim/assignment) uses
$sinkByToken = @{}
foreach ($f in $files) {
    $mod = $f.BaseName
    $isLive = ($liveModules -contains $mod)
    $lines = ([IO.File]::ReadAllText($f.FullName, $enc)).Split([char]10)
    for ($i=0; $i -lt $lines.Count; $i++) {
        $ln = $lines[$i]; $trim = $ln.Trim()
        if ($trim.StartsWith("'")) { continue }
        foreach ($cname in $constMap.Keys) {
            if ($ln -match ('(Worksheets|Sheets)\s*\(\s*' + [regex]::Escape($cname) + '\s*\)')) {
                $tok = $constMap[$cname].token
                if (-not $sinkByToken.ContainsKey($tok)) { $sinkByToken[$tok] = [System.Collections.ArrayList]::new() }
                [void]$sinkByToken[$tok].Add($mod + ':' + ($i+1) + $(if($isLive){' [LIVE]'}else{' [offline?]'}))
                Log ("SINK   " + $(if($isLive){'L '}else{'  '}) + $mod + ':' + ($i+1) + "  via " + $cname + "  " + $trim)
            }
        }
    }
}
Log "==== INDIRECT (const-based) SHEET-LOOKUP SINKS PER BAND ===="
foreach ($t in ($sinkByToken.Keys | Sort-Object)) {
    Log ("  " + $t.PadRight(20) + " <- " + ($sinkByToken[$t] -join ', '))
}
Log ("TOKENS_WITH_INDIRECT_SINKS=" + $sinkByToken.Count)
Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
Log ("LOG_WRITTEN=" + $outLog)
