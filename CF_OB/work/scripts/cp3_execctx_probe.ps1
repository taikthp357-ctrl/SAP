$ErrorActionPreference = 'Stop'
$wbPath   = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$ctxMod   = 'g:\SAP\CF_OB\work\scratch\modCFExecContext.bas'
$outLog   = 'g:\SAP\CF_OB\work\REPORTS\cp3_execctx_probe.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open($wbPath)   # writable (import temp module), never Save
    $proj = $wb.VBProject
    foreach ($c in @($proj.VBComponents)) {
        if ($c.Name -eq 'modCFExecContext') { $proj.VBComponents.Remove($c) }
    }
    $proj.VBComponents.Import($ctxMod) | Out-Null
    Log "CTX_MODULE_IMPORTED"

    $selfTest = "NOT_RUN"
    try { $selfTest = $xl.Run('CFExecCtx_SelfTest') } catch { $selfTest = "RUN_ERROR: " + $_.Exception.Message }
    Log ("EXECCTX_SELFTEST=" + $selfTest)

    # also re-confirm the allocator self-test still passes alongside the new module
    $alloc = "NOT_RUN"
    try { $alloc = $xl.Run('CFAlloc_SelfTest') } catch { $alloc = "RUN_ERROR: " + $_.Exception.Message }
    Log ("ALLOC_SELFTEST=" + $alloc)

    foreach ($c in @($proj.VBComponents)) {
        if ($c.Name -eq 'modCFExecContext') { $proj.VBComponents.Remove($c) }
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
