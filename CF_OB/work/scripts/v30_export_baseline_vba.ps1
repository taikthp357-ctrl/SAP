# v30_export_baseline_vba.ps1 - READ-ONLY export of all VBA components from the
# V3.0 baseline, so the hardening audit targets the true shipped source (not the
# possibly-stale CP1 scratch copies). Baseline is opened ReadOnly and NEVER saved.
$ErrorActionPreference = 'Stop'

$baseline = 'g:\SAP\CF_OB\260821_CF_OB_V3.0.xlsm'
$outDir   = 'g:\SAP\CF_OB\work\scratch\vba_v30_baseline_export'
$logPath  = 'g:\SAP\CF_OB\work\REPORTS\v30_export_baseline_vba.log'

$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

if (-not (Test-Path $baseline)) { throw "NO_BASELINE=$baseline" }
New-Item -ItemType Directory -Force -Path $outDir            | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $logPath) | Out-Null
# clean stale exports so the folder is an exact mirror of the baseline
Get-ChildItem -Path $outDir -Filter '*.*' -File -ErrorAction SilentlyContinue | Remove-Item -Force

Log "BASELINE_SHA=$((Get-FileHash $baseline -Algorithm SHA256).Hash)"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
$wb = $null
try {
    # ReadOnly:=$true, UpdateLinks:=0 ; events off so Workbook_Open does not run
    $wb = $xl.Workbooks.Open($baseline, 0, $true)
    Log "OPENED_READONLY=$($wb.ReadOnly)"
    $proj = $wb.VBProject
    Log "VBA_COMPONENTS=$($proj.VBComponents.Count)"

    foreach ($c in @($proj.VBComponents)) {
        switch ($c.Type) {
            1   { $ext = 'bas' }   # std module
            2   { $ext = 'cls' }   # class module
            3   { $ext = 'frm' }   # MSForm
            100 { $ext = 'cls' }   # document (ThisWorkbook / sheet)
            default { $ext = 'txt' }
        }
        $lines = 0; try { $lines = $c.CodeModule.CountOfLines } catch {}
        $path = Join-Path $outDir ($c.Name + '.' + $ext)
        $c.Export($path)
        Log ("EXPORTED|{0}|type={1}|lines={2}" -f $c.Name, $c.Type, $lines)
    }

    $wb.Close($false); $wb = $null   # NEVER save the baseline
    Log "CLOSED_NO_SAVE"
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
