# v28_plan_discovery.ps1
# Classify every PLAN-family reference on post-CP5a ground truth so the PLAN
# band-merge can be sequenced on measured facts (like CP4/CP5 discovery).
#   - exact quoted sheet tokens "PLAN" / "PLAN_SOURCE_LEGS" / "PLAN_RELEASE"
#   - split lookups (Worksheets(/Sheets() vs config-key/log-tag/InStr false positives
#   - dynamic name-building probes ("PLAN_" & , & "_LEGS", etc.)
# READ-ONLY. No workbook opened. Emits a per-site table + counts.
$ErrorActionPreference = 'Stop'
$vbaDir = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\v28_plan_discovery.log'
$enc = [Text.Encoding]::GetEncoding(1252)
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Host $m }

$tokens = @(
    '"PLAN"','"PLAN_SOURCE_LEGS"','"PLAN_RELEASE"',
    '"CONFIG"','"SAP_ACTION_MAP"','"SAP_PROBE_LOG"',
    '"RUN_LOG"','"RUN_SUMMARY"','"RUN_HISTORY"',
    '"OB_QUEUE"','"OB_WH_CACHE"','"POSTING_RAW"','"POSTING_SUMMARY"','"POSTING_CACHE"',
    '"ORDER_QUERY"','"REGRESSION_AUDIT"'
)
# module role map (from CP3/CP4/CP5 discovery): which consumers are live-SAP-only.
$liveModules = @('modCodexWTRework','modSAPReadOnlyProbe','ModuleSAP','modCF_Executor','modCF_MONCore_STABLE','modCF_AutoRun_TwoPilot','modCF_FinalizeVerifiedPilot','modMultiPostingLoader','modReleaseHolderReader','modCFX_Entry','modCFX_Planner')
# dynamic-build canaries: if ANY of these appear the literal-only rename is unsafe
$dynPats = @('"PLAN_" &','& "_LEGS"','& "_RELEASE"','& "_SOURCE_LEGS"','"PLAN" &','& "PLAN"')

$sheetLookup = 0; $configKey = 0; $logTag = 0; $instr = 0; $comment = 0; $other = 0
$dynHits = 0
$byToken = @{}
foreach ($t in $tokens) { $byToken[$t] = 0 }
$lookupByBand = @{}

$files = Get-ChildItem -Path $vbaDir -Filter *.bas | Sort-Object Name
foreach ($f in $files) {
    $lines = ([System.IO.File]::ReadAllText($f.FullName, $enc)).Split([char]10)
    for ($i=0; $i -lt $lines.Count; $i++) {
        $ln = $lines[$i]; $trim = $ln.Trim()
        foreach ($dp in $dynPats) { if ($ln.Contains($dp)) { $dynHits++; Log ("DYN  " + $f.Name + ":" + ($i+1) + "  " + $trim) } }
        foreach ($t in $tokens) {
            $n = ([regex]::Matches($ln, [regex]::Escape($t))).Count
            if ($n -lt 1) { continue }
            $byToken[$t] += $n
            $bare = $t.Trim('"')
            $mod  = $f.BaseName
            $isLive = ($liveModules -contains $mod)
            $cls = 'other'
            if ($trim.StartsWith("'")) { $cls = 'comment'; $comment += $n }
            elseif ($ln -match ('(Worksheets|Sheets)\s*\(\s*"' + [regex]::Escape($bare) + '"\s*\)')) {
                $cls='sheetLookup'; $sheetLookup += $n
                $k = $bare + ($(if($isLive){' [LIVE]'}else{' [offline?]'}))
                if (-not $lookupByBand.ContainsKey($k)) { $lookupByBand[$k] = [System.Collections.ArrayList]::new() }
                [void]$lookupByBand[$k].Add($mod + ':' + ($i+1))
            }
            elseif ($ln -match 'WriteLog|RT_Log|\.Log ') { $cls='logTag'; $logTag += $n }
            elseif ($ln -match 'InStr') { $cls='instr'; $instr += $n }
            elseif ($ln -match 'Const|Key|CONFIG|GetSetting|Setting') { $cls='configKey'; $configKey += $n }
            else { $other += $n }
            Log ($cls.PadRight(11) + ($(if($isLive){'L '}else{'  '})) + $f.Name + ":" + ($i+1) + "  " + $trim)
        }
    }
}
Log "==== TOKEN COUNTS ===="
foreach ($t in $tokens) { Log ("  " + $t.PadRight(20) + "= " + $byToken[$t]) }
Log "==== CLASS COUNTS ===="
Log ("  sheetLookup=$sheetLookup  configKey=$configKey  logTag=$logTag  instr=$instr  comment=$comment  other=$other")
Log ("  DYNAMIC_BUILD_HITS=$dynHits  (MUST be 0 for literal-only rename to be provably complete)")
Log "==== SHEET-LOOKUP CONSUMERS PER BAND (partition) ===="
foreach ($k in ($lookupByBand.Keys | Sort-Object)) {
    Log ("  " + $k.PadRight(26) + " <- " + ($lookupByBand[$k] -join ', '))
}
Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
Log ("LOG_WRITTEN=" + $outLog)
