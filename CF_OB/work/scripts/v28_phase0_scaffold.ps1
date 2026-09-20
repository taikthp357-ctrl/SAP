$ErrorActionPreference = 'Stop'
$src     = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$v28     = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_WIP.xlsm'
$expDir  = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$outLog  = 'g:\SAP\CF_OB\work\REPORTS\v28_phase0.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

# --- 1. Non-destructive copy ---
Copy-Item -Path $src -Destination $v28 -Force
$h = (Get-FileHash -Path $v28 -Algorithm SHA256).Hash
Log ("V28_WIP_CREATED sha256=" + $h + " bytes=" + (Get-Item $v28).Length)

# --- 2. Clean export dir ---
if (Test-Path $expDir) { Remove-Item -Path $expDir -Recurse -Force }
New-Item -ItemType Directory -Path $expDir | Out-Null

# --- 3. Fresh full VBA export via COM ---
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    $wb = $xl.Workbooks.Open($v28, 0, $false)  # RW (WIP copy, safe to touch; we won't save here)
    $proj = $wb.VBProject
    Log ("VBA_COMPONENTS=" + $proj.VBComponents.Count)
    # ext map: 1=StdModule(.bas) 2=ClassModule(.cls) 3=MSForm(.frm) 100=Document(.cls)
    $n = 0
    foreach ($c in $proj.VBComponents) {
        $ext = switch ($c.Type) { 1 {'.bas'} 2 {'.cls'} 3 {'.frm'} 100 {'.cls'} default {'.txt'} }
        $path = Join-Path $expDir ($c.Name + $ext)
        $c.Export($path)
        $n++
    }
    Log ("EXPORTED=" + $n)
    $wb.Close($false)
    Log "CLOSED_NO_SAVE"
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
    Write-Output ("EXPORT_DIR_FILES=" + (Get-ChildItem $expDir -File | Measure-Object).Count)
}
