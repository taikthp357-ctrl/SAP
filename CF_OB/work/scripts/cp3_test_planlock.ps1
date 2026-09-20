$ErrorActionPreference = 'Stop'
$wbPath = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$shim   = 'g:\SAP\CF_OB\work\scratch\modCFPlanLockProbe.bas'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\cp3_planlock_test.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open($wbPath)   # writable in memory; NEVER saved
    $proj = $wb.VBProject
    foreach ($c in @($proj.VBComponents)) { if ($c.Name -eq 'modCFPlanLockProbe') { $proj.VBComponents.Remove($c) } }
    $proj.VBComponents.Import($shim) | Out-Null
    Log "PROBE_IMPORTED"

    # Run the probe: it tests the real CWR_AssertQueuePlanLock against the live locked
    # PLAN (expect PASS) and against a tampered plan-hash (expect STOP), restoring after.
    $res = "NOT_RUN"
    try { $res = $xl.Run('CFProbe_PlanLock') } catch { $res = "RUN_ERROR: " + $_.Exception.Message }
    Log ("PLANLOCK_PROBE=" + $res)

    foreach ($c in @($proj.VBComponents)) { if ($c.Name -eq 'modCFPlanLockProbe') { $proj.VBComponents.Remove($c) } }
    $wb.Close($false)   # discard: no tamper persists, no probe persists
    Log "CLOSED_NO_SAVE"
} catch {
    Log ("FATAL=" + $_.Exception.Message)
    try { $wb.Close($false) } catch {}
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if ($preExcel -notcontains $p.Id) {
            try { if ($p.StartTime -ge $scriptStart) { Stop-Process -Id $p.Id -Force; Log ("KILLED_EXCEL_PID=" + $p.Id) } } catch {}
        }
    }
    Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
}
