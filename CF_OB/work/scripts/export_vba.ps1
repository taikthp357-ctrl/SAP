$ErrorActionPreference = 'Stop'
$wbPath = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$outDir = 'g:\SAP\CF_OB\work\scratch\vba_export'
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open($wbPath, 0, $true)
    foreach ($c in $wb.VBProject.VBComponents) {
        $cm = $c.CodeModule
        $lc = $cm.CountOfLines
        if ($lc -gt 0) {
            $body = $cm.Lines(1, $lc)
            $ext = switch ($c.Type) { 1 {'.bas'} 2 {'.cls'} 100 {'.doc.txt'} default {'.txt'} }
            Set-Content -Path (Join-Path $outDir ($c.Name + $ext)) -Value $body -Encoding UTF8
        }
    }
    $wb.Close($false)
    Write-Output "EXPORT_DONE"
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if ($preExcel -notcontains $p.Id) {
            try { if ($p.StartTime -ge $scriptStart) { Stop-Process -Id $p.Id -Force } } catch {}
        }
    }
}