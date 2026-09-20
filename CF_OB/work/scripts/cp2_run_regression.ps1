$ErrorActionPreference = 'Stop'
$wbPath   = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$origPath = 'g:\SAP\CF_OB\260819_CF_OB_V2.6_LICH_SU.xlsm'
$regMod   = 'g:\SAP\CF_OB\work\scratch\modCFAllocRegression.bas'
$outLog   = 'g:\SAP\CF_OB\work\REPORTS\cp2_run_regression.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
try {
    # --- integrity: open COPY writable (need to import reg module), but never Save ---
    $wb = $xl.Workbooks.Open($wbPath)
    Log ("COPY_OPENED=" + $wb.Name)
    Log ("SHEETS=" + $wb.Worksheets.Count)
    $proj = $wb.VBProject
    $compCount = $proj.VBComponents.Count
    Log ("VBA_COMPONENTS=" + $compCount)
    $hasAlloc = $false
    foreach ($c in $proj.VBComponents) { if ($c.Name -eq 'modCFGroupAllocator') { $hasAlloc = $true } }
    Log ("ALLOC_PRESENT=" + $hasAlloc)

    # --- import regression module (remove prior copy first) ---
    foreach ($c in @($proj.VBComponents)) {
        if ($c.Name -eq 'modCFAllocRegression') { $proj.VBComponents.Remove($c) }
    }
    $proj.VBComponents.Import($regMod) | Out-Null
    Log "REG_MODULE_IMPORTED"

    # --- run allocator self-test (proves full compile) ---
    $selfTest = "NOT_RUN"
    try { $selfTest = $xl.Run('CFAlloc_SelfTest') } catch { $selfTest = "RUN_ERROR: " + $_.Exception.Message }
    Log ("SELFTEST=" + $selfTest)

    # --- run full Qty regression suite ---
    $reg = "NOT_RUN"
    try { $reg = $xl.Run('CFReg_RunAll') } catch { $reg = "RUN_ERROR: " + $_.Exception.Message }
    Log ("REGRESSION=" + $reg)

    # --- run V2.6 native scheduler-level harnesses (prove patch didn't break end-to-end) ---
    # CF_FastModuleRegression: compile-probes the whole executor chain + asserts SAP maps LOCKED (offline).
    $fastReg = "NOT_RUN"
    try { $fastReg = $xl.Run('CF_FastModuleRegression') } catch { $fastReg = "RUN_ERROR: " + $_.Exception.Message }
    Log ("FAST_MODULE_REGRESSION=" + $fastReg)

    # CF_Phase3_RunUnitTests: 11-case decision matrix (P3_Decide), writes to a sheet then read H1.
    $p3 = "NOT_RUN"
    try {
        $xl.Run('CF_Phase3_RunUnitTests') | Out-Null
        # read the summary cell from the test sheet
        foreach ($ws in $wb.Worksheets) {
            $h1 = $null
            try { $h1 = $ws.Range('H1').Value2 } catch {}
            if ($h1 -ne $null -and ("$h1" -match 'UNIT TESTS')) { $p3 = "$h1"; break }
        }
    } catch { $p3 = "RUN_ERROR: " + $_.Exception.Message }
    Log ("PHASE3_UNIT_TESTS=" + $p3)

    # --- remove reg module so it never persists; close WITHOUT saving ---
    foreach ($c in @($proj.VBComponents)) {
        if ($c.Name -eq 'modCFAllocRegression') { $proj.VBComponents.Remove($c) }
    }
    $wb.Close($false)
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
