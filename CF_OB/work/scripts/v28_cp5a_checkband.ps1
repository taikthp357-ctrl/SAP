# v28_cp5a_checkband.ps1
# CP5a: prove the CF_CheckActualPosting raw-coord -> SSR band-accessor refactor is
# byte-identical in SHEET mode, then ship it.
#   Phase A: run scratch probe on GOLDEN (seam only, raw coords) -> goldSig.
#   Phase B: run scratch probe on SSR    (seam + coord swap)     -> candSig.
#   GATE 1 : goldSig -ceq candSig  (identical write-coordinate signature).
#   Phase C: build CP5a clean from CP4 (import ssr modCF_Posting, NO probe, NO
#            seeding) + full green suite; save + promote ssr -> ground truth ONLY
#            on all-green. Probe copies are throwaways, deleted regardless.
# Live SAP paths remain NOT_LIVE_PROVEN (GUI Scripting absent offline).
$ErrorActionPreference = 'Stop'
$cp4    = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP4_SYSTEM.xlsm'
$goldWb = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP5a_probe_golden.xlsm'
$ssrWb  = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP5a_probe_ssr.xlsm'
$cp5a   = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP5a_CHECKBAND.xlsm'
$basGold= 'g:\SAP\CF_OB\work\scratch\vba_v28_cp5stage\golden\modCF_Posting.bas'
$basSsr = 'g:\SAP\CF_OB\work\scratch\vba_v28_cp5stage\ssr\modCF_Posting.bas'
$probe  = 'g:\SAP\CF_OB\work\scratch\vba_v28_cp5stage\modCFCheckProbe.bas'
$regMod = 'g:\SAP\CF_OB\work\scratch\modCFAllocRegression.bas'
$gtPost = 'g:\SAP\CF_OB\work\scratch\vba_v28\modCF_Posting.bas'
$sigGoldF='g:\SAP\CF_OB\work\REPORTS\v28_cp5a_sig_golden.txt'
$sigCandF='g:\SAP\CF_OB\work\REPORTS\v28_cp5a_sig_ssr.txt'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\v28_cp5a_checkband.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Host $m }

if (-not (Test-Path $cp4))    { throw "NO_CP4=$cp4" }
if (-not (Test-Path $basGold)){ throw "NO_GOLDEN=$basGold" }
if (-not (Test-Path $basSsr)) { throw "NO_SSR=$basSsr" }
if (-not (Test-Path $probe))  { throw "NO_PROBE=$probe" }
if (-not (Test-Path $regMod)) { throw "NO_REGMOD=$regMod" }
Log ("CP4_SHA256=" + (Get-FileHash -Path $cp4 -Algorithm SHA256).Hash)

function ImportModule($vbp, $name, $path){
    $comp = $null
    foreach ($c in $vbp.VBComponents) { if ($c.Name -eq $name) { $comp = $c; break } }
    if ($comp -ne $null) { $vbp.VBComponents.Remove($comp) | Out-Null }
    $vbp.VBComponents.Import($path) | Out-Null
}
function HasModule($vbp, $name){
    foreach ($c in $vbp.VBComponents) { if ($c.Name -eq $name) { return $true } }
    return $false
}
function ModuleHasText($vbp, $name, $needle){
    foreach ($c in $vbp.VBComponents) {
        if ($c.Name -eq $name) {
            $cm = $c.CodeModule
            if ($cm.CountOfLines -lt 1) { return $false }
            return $cm.Lines(1, $cm.CountOfLines).Contains($needle)
        }
    }
    return $false
}
function SigHash([string]$s){
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $h = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($s)))).Replace('-','')
    $sha.Dispose(); return $h
}
# RunProbe: open throwaway copy, import a modCF_Posting variant + the probe, run
# CFCheck_Probe headless, close WITHOUT saving, delete the copy. Returns the
# signature string (or a RUNERR:/PROBE_ERR: marker). Log() uses Write-Host so it
# never contaminates this scalar return (the pipeline-pollution footgun).
function RunProbe([string]$copyPath, [string]$postBas, [string]$tag){
    Copy-Item $cp4 $copyPath -Force
    Log ("PROBE[$tag] copy -> " + (Split-Path $copyPath -Leaf))
    $wb = $null; $sig = 'RUNERR:not_run'
    try {
        $wb = $xl.Workbooks.Open($copyPath, 0, $false)
        $vbp = $wb.VBProject
        ImportModule $vbp 'modCF_Posting'   $postBas
        ImportModule $vbp 'modCFCheckProbe' $probe
        Log ("PROBE[$tag] imported modCF_Posting($tag) + modCFCheckProbe")
        $macro = "'" + (Split-Path $copyPath -Leaf) + "'!CFCheck_Probe"
        try { $sig = [string]$xl.Run($macro) } catch { $sig = 'RUNERR:' + $_.Exception.Message }
        $wb.Close($false); $wb = $null            # NEVER save a seeded probe copy
    } finally {
        if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
        if (Test-Path $copyPath) { try { Remove-Item $copyPath -Force } catch {} }
    }
    return $sig
}


$goldSig = $null; $candSig = $null; $saved = $false
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    # ===== PHASE A: GOLDEN probe (seam only, raw coords) =====================
    $goldSig = RunProbe $goldWb $basGold 'golden'
    [System.IO.File]::WriteAllText($sigGoldF, $goldSig, [Text.Encoding]::UTF8)
    Log ("GOLD_SIG len=" + $goldSig.Length + " sha256=" + (SigHash $goldSig) + " head=" + $goldSig.Substring(0,[Math]::Min(6,$goldSig.Length)))

    # ===== PHASE B: SSR probe (seam + coord swap) ===========================
    $candSig = RunProbe $ssrWb $basSsr 'ssr'
    [System.IO.File]::WriteAllText($sigCandF, $candSig, [Text.Encoding]::UTF8)
    Log ("CAND_SIG len=" + $candSig.Length + " sha256=" + (SigHash $candSig) + " head=" + $candSig.Substring(0,[Math]::Min(6,$candSig.Length)))

    # ===== GATE 1: signatures byte-identical, both real (not error markers) ==
    $gErr    = ($goldSig -match '^(PROBE_ERR|RUNERR):') -or ($candSig -match '^(PROBE_ERR|RUNERR):')
    $gShape  = ($goldSig -like 'SIG_V1*') -and ($candSig -like 'SIG_V1*')
    $gIdent  = ($goldSig -ceq $candSig)
    Log ("GATE1 noError=" + (-not $gErr) + " shapeOK=" + $gShape + " identical=" + $gIdent)
    if ($gErr)        { throw "GATE1_PROBE_ERROR gold='$goldSig' cand='$candSig'" }
    if (-not $gShape) { throw "GATE1_BAD_SHAPE (missing SIG_V1 header)" }
    if (-not $gIdent) { throw "GATE1_SIGNATURE_DIVERGENCE (refactor is NOT byte-identical in SHEET mode)" }
    Log "GATE1_PASS=SIGNATURES_IDENTICAL"

    # ===== PHASE C: build CP5a clean (ssr module, NO probe, NO seeding) ======
    Copy-Item $cp4 $cp5a -Force
    Log ("CP5a copy -> " + (Split-Path $cp5a -Leaf))
    $wbC = $xl.Workbooks.Open($cp5a, 0, $false)
    $vbpC = $wbC.VBProject
    ImportModule $vbpC 'modCF_Posting' $basSsr
    Log "CP5a imported ssr modCF_Posting"
    # purge any probe module that must NOT ship
    $probeComp = $null
    foreach ($c in $vbpC.VBComponents) { if ($c.Name -eq 'modCFCheckProbe') { $probeComp = $c; break } }
    if ($probeComp -ne $null) { $vbpC.VBComponents.Remove($probeComp) | Out-Null; Log 'CP5a removed stray modCFCheckProbe' }
    # structural gates: probe absent, seam present
    $probeAbsent = -not (HasModule $vbpC 'modCFCheckProbe')
    $seamPresent = ModuleHasText $vbpC 'modCF_Posting' 'Public CF_CheckSilent As Boolean'
    Log ("CP5a_STRUCT probeAbsent=$probeAbsent seamPresent=$seamPresent")
    if (-not $probeAbsent) { throw 'CP5a_PROBE_LEAKED' }
    if (-not $seamPresent) { throw 'CP5a_SEAM_MISSING' }
    # regression module: remove-then-import (lives in scratch\, not vba_v28\)
    ImportModule $vbpC 'modCFAllocRegression' $regMod
    Log 'CP5a import modCFAllocRegression'

    # FULL GREEN SUITE (same 7-gate contract as CP4)
    function RunS($name){ try { return [string]$xl.Run($name) } catch { return 'ERR:' + $_.Exception.Message } }
    $alloc = RunS 'CFAlloc_SelfTest';        Log ('CFAlloc_SelfTest=' + $alloc)
    $sscp  = RunS 'SSR_CompileProbe';         Log ('SSR_CompileProbe=' + $sscp + '<END>')
    $ssst  = RunS 'SSR_SelfTest';             Log ('SSR_SelfTest=' + $ssst)
    $reg   = RunS 'CFReg_RunAll';             Log ('CFReg_RunAll=' + $reg)
    $fast  = RunS 'CF_FastModuleRegression';  Log ('CF_FastModuleRegression=' + $fast)
    $reset = RunS 'CFResetUI_SelfTest';       Log ('CFResetUI_SelfTest=' + $reset)
    try { $xl.Run('CF_Phase3_RunUnitTests') | Out-Null } catch { Log ('Phase3_run_err=' + $_.Exception.Message) }
    $p3 = 'NA'
    foreach ($s in $wbC.Worksheets) { if ($s.Name -eq 'PHASE3_TESTS') { $p3 = [string]$s.Range('H1').Value2; break } }
    Log ('PHASE3_TESTS!H1=' + $p3)

    $gAlloc = ($alloc -match '(?i)PASS')
    $gSscp  = ($sscp  -notmatch '^ERR:')
    $gSsst  = ($ssst  -match '(?i)PASS')
    $gReg   = ($reg   -match 'REG_ALL_PASS' -and $reg -match 'FAIL=0')
    $gFast  = ($fast  -match '^(?i)PASS')
    $gReset = ($reset -match '^(?i)PASS')
    $gP3    = ($p3    -match '11/11' -and $p3 -match '(?i)SAP\s*Write\s*=\s*0')
    Log ("GATES alloc=$gAlloc ssrCompile=$gSscp ssrSelf=$gSsst reg=$gReg fast=$gFast reset=$gReset phase3=$gP3")
    $allGreen = $gAlloc -and $gSscp -and $gSsst -and $gReg -and $gFast -and $gReset -and $gP3

    if ($allGreen) { $wbC.Save(); $saved = $true; Log 'SAVED_CP5a=TRUE' }
    else { Log 'SAVED_CP5a=FALSE (gate red -> not saved)' }
    $wbC.Close($saved); $wbC = $null
} catch {
    Log ('FATAL=' + $_.Exception.Message)
    try { if ($wbC -ne $null) { $wbC.Close($false) } } catch {}
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if ($preExcel -notcontains $p.Id) {
            try { if ($p.StartTime -ge $scriptStart) { Stop-Process -Id $p.Id -Force; Log ('KILLED_EXCEL_PID=' + $p.Id) } } catch {}
        }
    }
    foreach ($tmp in @($goldWb,$ssrWb)) { if (Test-Path $tmp) { try { Remove-Item $tmp -Force; Log ('CLEANED ' + (Split-Path $tmp -Leaf)) } catch {} } }
}

# ===== PROMOTE ssr modCF_Posting -> ground truth ONLY on a saved green build =
if ($saved) {
    Copy-Item $basSsr $gtPost -Force
    Log ('PROMOTED_GROUNDTRUTH modCF_Posting <- ssr; sha256=' + (Get-FileHash -Path $gtPost -Algorithm SHA256).Hash)
    Log ('CP5a_SHA256=' + (Get-FileHash -Path $cp5a -Algorithm SHA256).Hash)
    Log 'CP5A_OK'
} else {
    Log 'PROMOTE_SKIPPED (no green save)'
    if (Test-Path $cp5a) { Remove-Item $cp5a -Force; Log 'REMOVED_PARTIAL_CP5a' }
    Log 'CP5A_FAILED'
}
Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
Log ('LOG_WRITTEN=' + $outLog)
