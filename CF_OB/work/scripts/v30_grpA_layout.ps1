# v30_grpA_layout.ps1 — HARDENING Group A (layout) on a COPY of the V3.0 baseline.
# Re-applies the two CP1 worklist UI modules (WT+Material cols; status blank/A),
# adds the §4 regression TEST_UI_V3_LAYOUT, forces a full VBA compile, runs the
# V3.0 offline gate suite. Baseline is NEVER opened writable — copy first, only
# ever touch the WIP. Offline only: EnableEvents=$false, no SAP, SAP_WRITE=0.
$ErrorActionPreference = 'Stop'

$baseline   = 'g:\SAP\CF_OB\260821_CF_OB_V3.0.xlsm'
$wip        = 'g:\SAP\CF_OB\work\WORKING\260822_CF_OB_V3.0_GRPA_LAYOUT_WIP.xlsm'
$final      = 'g:\SAP\CF_OB\work\WORKING\260822_CF_OB_V3.0_GRPA_LAYOUT.xlsm'
$staging    = 'g:\SAP\CF_OB\work\scratch\vba_v30_analysis'
$logPath    = 'g:\SAP\CF_OB\work\REPORTS\v30_grpA_layout.log'
$modules    = @('modCF_InputWorklist.bas','modCF_InputLayout.bas')

$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

if (-not (Test-Path $baseline)) { throw "NO_BASELINE=$baseline" }
if (-not (Test-Path $staging))  { throw "NO_STAGING=$staging" }
foreach ($m in $modules) { if (-not (Test-Path (Join-Path $staging $m))) { throw "NO_MODULE=$m" } }
New-Item -ItemType Directory -Force -Path (Split-Path $wip)     | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $logPath) | Out-Null

Copy-Item $baseline $wip -Force
Log "WIP_CREATED=$wip"
Log "BASELINE_SHA=$((Get-FileHash $baseline -Algorithm SHA256).Hash)"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
$wb = $null
$results = [ordered]@{}
$passAll = $true

function RunStr($name){
    try { $v = $xl.Run($name); if ($null -eq $v) { return "<null>" } else { return "$v" } }
    catch { return "RUN_ERROR: " + $_.Exception.Message }
}
function RunVoid($name){
    try { $xl.Run($name) | Out-Null; return "RAN_NO_EXCEPTION" }
    catch { return "RUN_ERROR: " + $_.Exception.Message }
}

try {
    $wb = $xl.Workbooks.Open($wip, 0, $false)
    Log "OPENED_WIP"
    $proj = $wb.VBProject
    Log "VBA_COMPONENTS_BEFORE=$($proj.VBComponents.Count)"

    # --- Import the two edited UI modules (remove old, then import) ---
    foreach ($mod in $modules) {
        $modName = $mod -replace '\.(bas|cls)$',''
        foreach ($c in @($proj.VBComponents)) {
            if ($c.Name -eq $modName) {
                Log ("PRE_LINES[$modName]=" + $c.CodeModule.CountOfLines)
                $proj.VBComponents.Remove($c); Log "REMOVED_OLD=$modName"; break
            }
        }
        $proj.VBComponents.Import((Join-Path $staging $mod)) | Out-Null
        Log "IMPORTED=$modName"
    }
    foreach ($mod in $modules) {
        $modName = $mod -replace '\.(bas|cls)$',''
        foreach ($c in @($proj.VBComponents)) {
            if ($c.Name -eq $modName) { Log ("POST_LINES[$modName]=" + $c.CodeModule.CountOfLines) }
        }
    }
    Log "VBA_COMPONENTS_AFTER=$($proj.VBComponents.Count)"
    $wb.Save(); Log "WIP_SAVED_WITH_MODULES"

    # --- Gate suite (first Run forces a full-project compile) ---
    $results['CFAlloc_SelfTest']          = RunStr  'CFAlloc_SelfTest'
    $results['SSR_CompileProbe']          = RunVoid 'SSR_CompileProbe'
    $results['SSR_SelfTest']              = RunStr  'SSR_SelfTest'
    $results['CF_FastModuleRegression']   = RunStr  'CF_FastModuleRegression'
    $results['CFX_RunAllReplay']          = RunStr  'CFX_RunAllReplay'
    $results['CF_Offline_CompileProbe']   = RunStr  'CF_Offline_CompileProbe'
    $results['CF_InputWorklist_Probe']    = RunVoid 'CF_InputWorklist_CompileProbe'
    $results['CF_InputLayout_Probe']      = RunVoid 'CF_InputLayout_CompileProbe'

    # Phase3 unit tests write to a sheet; read the summary cell (H1)
    $p3 = RunVoid 'CF_Phase3_RunUnitTests'
    if ($p3 -eq 'RAN_NO_EXCEPTION') {
        $p3sum = '<no summary cell found>'
        foreach ($ws in $wb.Worksheets) {
            $h1 = $null; try { $h1 = $ws.Range('H1').Value2 } catch {}
            if ($h1 -ne $null -and ("$h1" -match 'UNIT TEST')) { $p3sum = "$h1"; break }
        }
        $results['CF_Phase3_RunUnitTests'] = $p3sum
    } else { $results['CF_Phase3_RunUnitTests'] = $p3 }

    # §4 layout regression LAST — it invokes CF_ApplyInputLayout (bakes the
    # canonical E:P worklist + single-visible CF_CONTROL into the saved snapshot).
    $results['TEST_UI_V3_LAYOUT']         = RunStr  'TEST_UI_V3_LAYOUT'

    foreach ($k in $results.Keys) { Log ("GATE|$k=" + $results[$k]) }

    # --- Evaluate promotion criteria ---
    if ($results['CFAlloc_SelfTest'] -ne 'PASS') { $passAll = $false }
    if ($results['CF_FastModuleRegression'] -notmatch '^PASS\|') { $passAll = $false }
    if ($results['CFX_RunAllReplay'] -notmatch '^REPLAY_ALL_PASS') { $passAll = $false }
    if ($results['CF_Offline_CompileProbe'] -notmatch '^PASS\|') { $passAll = $false }
    if ($results['SSR_SelfTest'] -match 'RUN_ERROR') { $passAll = $false }
    if ($results['TEST_UI_V3_LAYOUT'] -notmatch '^PASS\|') { $passAll = $false }
    foreach ($k in @('SSR_CompileProbe','CF_InputWorklist_Probe','CF_InputLayout_Probe')) {
        if ($results[$k] -ne 'RAN_NO_EXCEPTION') { $passAll = $false }
    }
    Log "PASS_ALL=$passAll"

    $wb.Save(); $wb.Close($true); $wb = $null
    Log "WIP_SHA=$((Get-FileHash $wip -Algorithm SHA256).Hash)"
    if ($passAll) {
        Move-Item $wip $final -Force
        Log "CHECKPOINT=$final"
        Log "CHECKPOINT_SHA=$((Get-FileHash $final -Algorithm SHA256).Hash)"
    } else {
        Log "PROMOTION_WITHHELD|WIP_KEPT_FOR_INSPECTION=$wip"
    }
} catch {
    Log "FATAL=$($_.Exception.Message)"
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if ($preExcel -notcontains $p.Id) {
            try { if ($p.StartTime -ge $scriptStart) { Stop-Process -Id $p.Id -Force; Log ("KILLED_EXCEL_PID=" + $p.Id) } } catch {}
        }
    }
    Log "POST_BASELINE_SHA=$((Get-FileHash $baseline -Algorithm SHA256).Hash)"
    Set-Content -Path $logPath -Value ($L -join "`r`n") -Encoding UTF8
    Log "DONE"
}
