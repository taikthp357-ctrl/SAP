# v28_cp3_datacore.ps1
# CP3 — rename anchor WH_ACTUAL -> DATA_CORE (coordinate-preserving) + update the
# 58 sheet-reference sites, protecting the 1 provenance data-value at CheckPlanner:875.
# Fail-closed: saves CP3 ONLY when every gate is green. Never touches baselines/CP2.
$ErrorActionPreference = 'Stop'
$src    = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP2_REPO.xlsm'
$dst    = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP3_DATACORE.xlsm'
$vbaDir = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$stgDir = 'g:\SAP\CF_OB\work\scratch\vba_v28_cp3stage'
$regMod = 'g:\SAP\CF_OB\work\scratch\modCFAllocRegression.bas'   # lives in scratch\, not vba_v28\
$outLog = 'g:\SAP\CF_OB\work\REPORTS\v28_cp3_datacore.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

$OLD='WH_ACTUAL'; $NEW='DATA_CORE'
$tok = '"' + $OLD + '"'; $rep = '"' + $NEW + '"'
$enc = [Text.Encoding]::GetEncoding(1252)   # single-byte => byte-preserving round-trip
$EXPECT_TOTAL = 58                           # 59 literals minus the 1 protected data-value

# ---- 1) STAGE: byte-safe, line-scoped rename in the 18 target modules -----------
if (Test-Path $stgDir) { Remove-Item $stgDir -Recurse -Force }
New-Item -ItemType Directory -Path $stgDir | Out-Null
$targets = @('modCFX_Entry','modCF_ActualPicking','modCF_CheckPlanner','modCF_AutoRun_TwoPilot',
  'modCF_AutoSetActualVerified','modCF_Entry','modCF_Executor','modCF_FinalizeVerifiedPilot',
  'modCF_MONCore_STABLE','modCF_MultiOB','modCF_Posting','modCF_PreLiveAudit','modCF_Phase3Planner',
  'modCodexNewOBSnapshot','modCodexWTRework','modMultiPostingLoader','modReleaseHolderReader',
  'modSAPReadOnlyProbe')
$totalReplaced = 0
$changed = @()
foreach ($t in $targets) {
    $p = Join-Path $vbaDir ($t + '.bas')
    if (-not (Test-Path $p)) { Log "MISS_SRC=$t"; continue }
    $text = [System.IO.File]::ReadAllText($p, $enc)
    $lines = $text.Split([char]10)
    $cnt = 0
    for ($i=0; $i -lt $lines.Count; $i++) {
        $ln = $lines[$i]
        if ($ln -match 'Worksheets\(|\.Sheets\(|\bConst\b|Array\(') {
            if ($ln.Contains($tok)) {
                $cnt += ([regex]::Matches($ln, [regex]::Escape($tok))).Count
                $lines[$i] = $ln.Replace($tok, $rep)
            }
        }
    }
    if ($cnt -gt 0) {
        [System.IO.File]::WriteAllText((Join-Path $stgDir ($t + '.bas')), [string]::Join([char]10,$lines), $enc)
        $changed += $t; $totalReplaced += $cnt
        Log ("STAGE " + $t.PadRight(28) + " replaced=" + $cnt)
    } else { Log ("NOCHG " + $t) }
}
Log ("TOTAL_REPLACED=" + $totalReplaced + " EXPECT=" + $EXPECT_TOTAL + " CHANGED=" + $changed.Count)
# Fail-closed BEFORE opening Excel: the rename must hit exactly the 58 known sites.
if ($totalReplaced -ne $EXPECT_TOTAL) {
    Log ("ABORT_STAGE_COUNT_MISMATCH got=" + $totalReplaced + " expect=" + $EXPECT_TOTAL)
    Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
    throw ("Stage count mismatch: " + $totalReplaced + " != " + $EXPECT_TOTAL)
}

# ---- 2) COPY CP2 -> CP3 (never mutate CP2 in place) -----------------------------
Copy-Item $src $dst -Force
Log ("COPIED -> " + $dst)

# ---- 3) OPEN, reimport changed modules, physical sheet rename, gates ------------
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
$saved = $false
try {
    $wb = $xl.Workbooks.Open($dst, 0, $false)   # writable
    $vbp = $wb.VBProject

    # 3a) Reimport each changed module from the staged .bas (import restores VB_Name)
    foreach ($t in $changed) {
        $comp = $null
        foreach ($c in $vbp.VBComponents) { if ($c.Name -eq $t) { $comp = $c; break } }
        if ($comp -ne $null) { $vbp.VBComponents.Remove($comp) }
        $vbp.VBComponents.Import((Join-Path $stgDir ($t + '.bas'))) | Out-Null
        Log ("REIMPORT " + $t)
    }

    # 3b) Physical, coordinate-preserving sheet rename WH_ACTUAL -> DATA_CORE
    $wsOld = $null
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq $OLD) { $wsOld = $s; break } }
    if ($wsOld -eq $null) { throw ("NO_SHEET_" + $OLD) }
    $dimBefore = $wsOld.UsedRange.Address($true,$true,1,$false)
    $wsOld.Name = $NEW
    $wsNew = $null; $wsStale = $null
    foreach ($s in $wb.Worksheets) {
        if ($s.Name -eq $NEW) { $wsNew = $s }
        if ($s.Name -eq $OLD) { $wsStale = $s }
    }
    $dimAfter = $wsNew.UsedRange.Address($true,$true,1,$false)
    Log ("SHEET_RENAME ok=" + ($wsNew -ne $null -and $wsStale -eq $null) + " dimBefore=" + $dimBefore + " dimAfter=" + $dimAfter)
    if ($wsNew -eq $null -or $wsStale -ne $null -or $dimBefore -ne $dimAfter) { throw "SHEET_RENAME_FAILED" }
    # 3c) STATIC zero-reference proof across the WHOLE project (all components),
    #     scanning CodeModule text for any surviving sheet-ref to the old name.
    $staleRefs = 0
    foreach ($c in $vbp.VBComponents) {
        $cm = $c.CodeModule
        if ($cm.CountOfLines -lt 1) { continue }
        $code = $cm.Lines(1, $cm.CountOfLines)
        foreach ($ln in $code.Split([char]10)) {
            if ($ln -match 'Worksheets\(|\.Sheets\(|\bConst\b|Array\(') {
                if ($ln.Contains($tok)) { $staleRefs++; Log ("STALE_REF " + $c.Name + ": " + $ln.Trim()) }
            }
        }
    }
    Log ("STATIC_ZERO_REF stale=" + $staleRefs)
    if ($staleRefs -ne 0) { throw ("STALE_SHEET_REFS=" + $staleRefs) }

    # 3d) Ensure regression module present & fresh (Phase 2 imported at runtime; may
    #     not have persisted into CP2). Remove-then-import so CFReg_RunAll is callable.
    if (-not (Test-Path $regMod)) { throw ("NO_REGMOD=" + $regMod) }
    $regComp = $null
    foreach ($c in $vbp.VBComponents) { if ($c.Name -eq 'modCFAllocRegression') { $regComp = $c; break } }
    if ($regComp -ne $null) { $vbp.VBComponents.Remove($regComp); Log 'REMOVED_STALE modCFAllocRegression' }
    $vbp.VBComponents.Import($regMod) | Out-Null
    Log 'IMPORT modCFAllocRegression'

    # 3e) FULL GREEN SUITE (compile-forcing self-tests + regressions)
    function RunS($name){ try { return [string]$xl.Run($name) } catch { return 'ERR:' + $_.Exception.Message } }
    $alloc = RunS 'CFAlloc_SelfTest';           Log ('CFAlloc_SelfTest=' + $alloc)
    $sscp  = RunS 'SSR_CompileProbe';            Log ('SSR_CompileProbe=' + $sscp + '<END>')
    $ssst  = RunS 'SSR_SelfTest';                Log ('SSR_SelfTest=' + $ssst)
    $reg   = RunS 'CFReg_RunAll';                Log ('CFReg_RunAll=' + $reg)
    $fast  = RunS 'CF_FastModuleRegression';     Log ('CF_FastModuleRegression=' + $fast)
    $reset = RunS 'CFResetUI_SelfTest';          Log ('CFResetUI_SelfTest=' + $reset)
    # Phase3 unit tests are a Sub -> read the result cell PHASE3_TESTS!H1
    try { $xl.Run('CF_Phase3_RunUnitTests') | Out-Null } catch { Log ('Phase3_run_err=' + $_.Exception.Message) }
    $p3 = 'NA'
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq 'PHASE3_TESTS') { $p3 = [string]$s.Range('H1').Value2; break } }
    Log ('PHASE3_TESTS!H1=' + $p3)

    $gAlloc = ($alloc -match '(?i)PASS')
    # SSR_CompileProbe is a Sub (returns nothing). "compiled" == ran without ERR:.
    # SSR_SelfTest=PASS is the real functional proof of modSourceStockRepo.
    $gSscp  = ($sscp  -notmatch '^ERR:')
    $gSsst  = ($ssst  -match '(?i)PASS')
    $gReg   = ($reg   -match 'REG_ALL_PASS' -and $reg -match 'FAIL=0')
    $gFast  = ($fast  -match '^(?i)PASS')
    $gReset = ($reset -match '^(?i)PASS')
    $gP3    = ($p3    -match '11/11' -and $p3 -match '(?i)SAP\s*Write\s*=\s*0')
    Log ("GATES alloc=$gAlloc ssrCompile=$gSscp ssrSelf=$gSsst reg=$gReg fast=$gFast reset=$gReset phase3=$gP3")
    $allGreen = $gAlloc -and $gSscp -and $gSsst -and $gReg -and $gFast -and $gReset -and $gP3

    if ($allGreen) {
        $wb.Save()
        $saved = $true
        Log 'SAVED_CP3=TRUE'
    } else {
        Log 'SAVED_CP3=FALSE (gate red -> not saved)'
    }
    $wb.Close($saved)
} catch {
    Log ('FATAL=' + $_.Exception.Message)
    try { $wb.Close($false) } catch {}
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if ($preExcel -notcontains $p.Id) {
            try { if ($p.StartTime -ge $scriptStart) { Stop-Process -Id $p.Id -Force; Log ('KILLED_EXCEL_PID=' + $p.Id) } } catch {}
        }
    }
}

# ---- 4) Promote staged modules to ground truth ONLY on a saved green build -------
if ($saved) {
    foreach ($t in $changed) { Copy-Item (Join-Path $stgDir ($t + '.bas')) (Join-Path $vbaDir ($t + '.bas')) -Force }
    Log ('PROMOTED_TO_GROUNDTRUTH=' + $changed.Count)
    if (Test-Path $dst) {
        $h = (Get-FileHash -Path $dst -Algorithm SHA256).Hash
        Log ('CP3_SHA256=' + $h)
    }
} else {
    Log 'PROMOTE_SKIPPED (no green save)'
    if (Test-Path $dst) { Remove-Item $dst -Force; Log 'REMOVED_PARTIAL_CP3' }
}
Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
Log ('LOG_WRITTEN=' + $outLog)
