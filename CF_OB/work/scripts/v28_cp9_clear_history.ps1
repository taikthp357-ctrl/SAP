# v28_cp9_clear_history.ps1 — Clear all history data for fresh V2.8 start
$ErrorActionPreference = 'Stop'
$cp8 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP8_SYSTEM_BAND.xlsm'
$wip = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP9_WIP.xlsm'
$logPath = 'g:\SAP\CF_OB\work\REPORTS\v28_cp9_clear_history.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m) { [void]$L.Add($m); Write-Host $m }

if (-not (Test-Path $cp8)) { throw "NO_CP8=$cp8" }

Copy-Item $cp8 $wip -Force
Log "WIP_CREATED"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$wb = $null

try {
    $wb = $xl.Workbooks.Open($wip, 0, $false)
    Log 'OPENED_WIP'

    # Clear RUN_HISTORY (keep headers row 1, clear data rows 2+)
    $wsRH = $wb.Worksheets.Item('RUN_HISTORY')
    $rhLR = $wsRH.Cells.Item($wsRH.Rows.Count, 1).End(-4162).Row
    if ($rhLR -gt 1) {
        $wsRH.Range("A2:AA$rhLR").ClearContents()
        Log "CLEARED_RUN_HISTORY|rows=2..$rhLR"
    }

    # Clear RUN_LOG (keep headers row 1-8, clear data rows 9+)
    $wsRL = $wb.Worksheets.Item('RUN_LOG')
    $rlLR = $wsRL.Cells.Item($wsRL.Rows.Count, 1).End(-4162).Row
    if ($rlLR -gt 8) {
        $wsRL.Range("A9:X$rlLR").ClearContents()
        Log "CLEARED_RUN_LOG|rows=9..$rlLR"
    }

    # Clear RUN_SUMMARY (keep headers row 1, clear data rows 2+)
    $wsRS = $wb.Worksheets.Item('RUN_SUMMARY')
    $rsLR = $wsRS.Cells.Item($wsRS.Rows.Count, 1).End(-4162).Row
    if ($rsLR -gt 1) {
        $wsRS.Range("A2:N$rsLR").ClearContents()
        Log "CLEARED_RUN_SUMMARY|rows=2..$rsLR"
    }

    # Clear SAP_PROBE_LOG in SYSTEM band (keep header row 1 at col 37, clear data rows 2+)
    $wsSys = $wb.Worksheets.Item('SYSTEM')
    $splLR = $wsSys.Cells.Item($wsSys.Rows.Count, 37).End(-4162).Row
    if ($splLR -gt 1) {
        $wsSys.Range("AK2:BA$splLR").ClearContents()
        Log "CLEARED_SAP_PROBE_LOG|rows=2..$splLR"
    }

    $wb.Save()
    Log 'ALL_HISTORY_CLEARED'

    # Run regression gates to verify system still works
    Log 'RUNNING_GATES'

    $r1 = $xl.Run('CFAlloc_SelfTest')
    if ($r1 -ne 'PASS') { throw "GATE_CFAlloc_FAIL|$r1" }
    Log "GATE_CFAlloc=PASS"

    $xl.Run('SSR_CompileProbe')
    Log "GATE_SSR_CompileProbe=RAN"

    $r3 = $xl.Run('SSR_SelfTest')
    if ($r3 -ne 'PASS') { throw "GATE_SSR_SelfTest_FAIL|$r3" }
    Log "GATE_SSR_SelfTest=PASS"

    $wsTest = $wb.Worksheets.Item('PHASE3_TESTS')
    $r6 = $wsTest.Range('H1').Value2
    if ($r6 -notmatch '11/11.*SAP Write=0') { throw "GATE_Phase3_FAIL|$r6" }
    Log "GATE_Phase3=PASS|$r6"

    Log 'ALL_GATES_GREEN'

    $wb.Save()
    $wb.Close($false)
    $wb = $null
    Log "WIP_HASH=$((Get-FileHash $wip -Algorithm SHA256).Hash)"

    $cp9 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP9_HISTORY_CLEARED.xlsm'
    Move-Item $wip $cp9 -Force
    Log "CP9_CHECKPOINT=$cp9"
    Log "CP9_HASH=$((Get-FileHash $cp9 -Algorithm SHA256).Hash)"

} catch {
    Log "FATAL=$($_.Exception.Message)"
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    throw
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $postPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $newPids = $postPids | Where-Object { $_ -notin $preExcel }
    if ($newPids -and ((Get-Date) - $scriptStart -gt [TimeSpan]::FromMinutes(5))) {
        $newPids | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue; Log "KILLED_HUNG=$_" }
    }
}

Set-Content -Path $logPath -Value ($L -join "`r`n") -Encoding UTF8
Log 'DONE'
