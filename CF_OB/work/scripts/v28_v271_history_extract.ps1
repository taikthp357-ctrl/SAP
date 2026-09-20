# v28_v271_history_extract.ps1
# Extract RUN_HISTORY from V2.71 to analyze live execution patterns for V2.8 optimization.
# READ-ONLY. Exports history rows + metadata → JSON for offline analysis.
$ErrorActionPreference = 'Stop'
$v271   = 'g:\SAP\CF_OB\260820_CF_OB_V2.71.xlsm'
$outJson = 'g:\SAP\CF_OB\work\REPORTS\v271_history_extract.json'
$outLog  = 'g:\SAP\CF_OB\work\REPORTS\v271_history_extract.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Host $m }

if (-not (Test-Path $v271)) { throw "NO_V271=$v271" }
Log ("V271_SHA256=" + (Get-FileHash -Path $v271 -Algorithm SHA256).Hash)
Log ("V271_SIZE=" + (Get-Item $v271).Length + " bytes")

$history = @()
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    $wb = $xl.Workbooks.Open($v271, 0, $true)  # ReadOnly=TRUE
    Log ("Opened V2.71 ReadOnly")

    # Find RUN_HISTORY sheet
    $sh = $null
    foreach ($s in $wb.Worksheets) {
        if ($s.Name -eq 'RUN_HISTORY') { $sh = $s; break }
    }
    if ($sh -eq $null) { throw "NO_RUN_HISTORY_SHEET" }

    $lastRow = $sh.Cells($sh.Rows.Count, 1).End(-4162).Row  # xlUp=-4162
    Log ("RUN_HISTORY rows: $lastRow")

    if ($lastRow -lt 2) {
        Log "RUN_HISTORY empty (header only or no data)"
    } else {
        # Read header row (assume row 1)
        $hdrRange = $sh.Range("A1:Z1")
        $hdrValues = $hdrRange.Value2
        $headers = @()
        for ($c=1; $c -le 26; $c++) {
            $h = $hdrValues[1,$c]
            if ($h -ne $null -and $h -ne '') { $headers += [string]$h } else { break }
        }
        Log ("HEADERS=" + ($headers -join '|'))

        # Read data rows (limit to last 500 for performance; V2.71 likely has ~3304 rows)
        $startRow = [Math]::Max(2, $lastRow - 499)
        Log ("Reading rows $startRow to $lastRow (last 500 max)")

        $dataRange = $sh.Range("A$startRow`:Z$lastRow")
        $dataValues = $dataRange.Value2

        for ($r=1; $r -le ($lastRow - $startRow + 1); $r++) {
            $row = @{}
            for ($c=1; $c -le $headers.Count; $c++) {
                $val = $dataValues[$r,$c]
                if ($val -eq $null) { $val = '' }
                $row[$headers[$c-1]] = [string]$val
            }
            $history += $row
        }
        Log ("Extracted $($history.Count) rows")
    }

    $wb.Close($false)
} catch {
    Log ('FATAL=' + $_.Exception.Message)
    try { if ($wb -ne $null) { $wb.Close($false) } } catch {}
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

# Write JSON
$jsonOut = @{ extracted=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); rowCount=$history.Count; rows=$history } | ConvertTo-Json -Depth 3
[System.IO.File]::WriteAllText($outJson, $jsonOut, [Text.Encoding]::UTF8)
Log ("JSON_WRITTEN=$outJson size=" + (Get-Item $outJson).Length)

Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
Log ("LOG_WRITTEN=$outLog")
