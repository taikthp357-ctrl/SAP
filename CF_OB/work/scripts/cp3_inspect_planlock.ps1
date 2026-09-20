$ErrorActionPreference = 'Stop'
$wbPath = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\cp3_planlock_state.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open($wbPath, 0, $true)  # READ-ONLY: just inspect current PLAN lock cells
    $plan = $null; $wh = $null
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -eq 'PLAN') { $plan = $ws }
        if ($ws.Name -eq 'WH_ACTUAL') { $wh = $ws }
    }
    if ($plan -ne $null) {
        Log ("PLAN_AJ1=" + $plan.Range('AJ1').Value2)
        Log ("PLAN_AJ2_LOCK=" + $plan.Range('AJ2').Value2)
        Log ("PLAN_AK2_RUNID=" + $plan.Range('AK2').Value2)
        Log ("PLAN_AN2_PLANHASH=" + $plan.Range('AN2').Value2)
        Log ("PLAN_AO2_WHHASH=" + $plan.Range('AO2').Value2)
        Log ("PLAN_AT2_CONTRACT=" + $plan.Range('AT2').Value2)
        $lastRow = $plan.Cells($plan.Rows.Count, 1).End(-4162).Row  # xlUp
        Log ("PLAN_LASTROW=" + $lastRow)
    } else { Log "PLAN_SHEET_MISSING" }
    if ($wh -ne $null) {
        Log ("WH_W2_RUNID=" + $wh.Range('W2').Value2)
        Log ("WH_X2_GATE=" + $wh.Range('X2').Value2)
    } else { Log "WH_SHEET_MISSING" }
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
