$ErrorActionPreference = 'Stop'
$cp2 = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP2_REPO.xlsm'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\v28_dump_layout.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

# Read-only: dump WH_ACTUAL header row (all 32 cols) + 1 sample data row, and
# INPUT's current top-left occupancy (rows 1-12, cols 1-16). Grounds Phase-1b
# worklist mirror design in the REAL column contract, not guesses.
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    $wb = $xl.Workbooks.Open($cp2, 0, $true)   # ReadOnly

    $wh = $wb.Worksheets('WH_ACTUAL')
    Log "=== WH_ACTUAL header row 1 (col: value) ==="
    for ($c=1; $c -le 32; $c++) {
        $v = [string]$wh.Cells(1,$c).Value2
        Log ("WH_H c=" + $c.ToString().PadLeft(2) + " : " + $v)
    }
    Log "=== WH_ACTUAL sample data row 2 ==="
    for ($c=1; $c -le 32; $c++) {
        $v = [string]$wh.Cells(2,$c).Value2
        if ($v -ne '') { Log ("WH_D2 c=" + $c.ToString().PadLeft(2) + " : " + $v) }
    }

    $inp = $wb.Worksheets('INPUT')
    Log "=== INPUT current occupancy rows 1-14 x cols 1-16 (nonblank only) ==="
    for ($r=1; $r -le 14; $r++) {
        for ($c=1; $c -le 16; $c++) {
            $v = [string]$inp.Cells($r,$c).Value2
            if ($v -ne '') { Log ("INP r=" + $r.ToString().PadLeft(2) + " c=" + $c.ToString().PadLeft(2) + " : " + $v) }
        }
    }
    Log ("INPUT_USED rows=" + $inp.UsedRange.Rows.Count + " cols=" + $inp.UsedRange.Columns.Count)

    $wb.Close($false)
} catch {
    Log ("COM_FATAL=" + $_.Exception.Message)
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
