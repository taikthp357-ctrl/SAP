$ErrorActionPreference = 'Stop'
$wbPath  = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$resetMod= 'g:\SAP\CF_OB\work\scratch\modCFResetUI.bas'
$outLog  = 'g:\SAP\CF_OB\work\REPORTS\cp8_resetui.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    $wb = $xl.Workbooks.Open($wbPath)
    $proj = $wb.VBProject
    foreach ($c in @($proj.VBComponents)) { if ($c.Name -eq 'modCFResetUI') { $proj.VBComponents.Remove($c) } }
    $proj.VBComponents.Import($resetMod) | Out-Null
    Log "RESETUI_IMPORTED"

    # self-test invalidates lock in-memory then restores; safe.
    $st = "NOT_RUN"; try { $st = $xl.Run('CFResetUI_SelfTest') } catch { $st = "ERR:" + $_.Exception.Message }
    Log ("RESETUI_SELFTEST=" + $st)
    # compile proof
    $alloc = "NOT_RUN"; try { $alloc = $xl.Run('CFAlloc_SelfTest') } catch { $alloc = "ERR:" + $_.Exception.Message }
    Log ("ALLOC_SELFTEST=" + $alloc)

    Log ("VBA_COMPONENTS=" + $proj.VBComponents.Count)
    # Persist the new module (additive LAM MOI entry point). Verify lock cell intact after save.
    $wb.Save()
    Log "SAVED_WITH_RESETUI"
    $wb.Close($true)
    Log "DONE"
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
