$ErrorActionPreference = 'Stop'
$cp2 = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP2_REPO.xlsm'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\v28_census.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

# Read-only census of CP2: enumerate VBComponents (Name/Type) + all worksheets
# (Name/Visible/rows/cols). Dual purpose: (a) confirm the 68->69 count is only
# the new repo + expected doc-modules, no surprise component; (b) capture the
# sheet inventory that Phase 3 (20->5 consolidation) will consolidate.
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    $wb = $xl.Workbooks.Open($cp2, 0, $true)   # ReadOnly
    $proj = $wb.VBProject
    Log ("OPENED=" + $wb.Name + " VBComponents=" + $proj.VBComponents.Count + " Worksheets=" + $wb.Worksheets.Count)

    Log "--- VBComponents (Type: 1=std 2=class 3=form 100=document) ---"
    $comps = @()
    foreach ($c in $proj.VBComponents) { $comps += [pscustomobject]@{ Name=$c.Name; Type=$c.Type } }
    foreach ($c in ($comps | Sort-Object Type,Name)) { Log ("COMP type=" + $c.Type + " " + $c.Name) }

    Log "--- Worksheets (visible: -1=visible 0=hidden 2=veryhidden) ---"
    foreach ($ws in $wb.Worksheets) {
        $ur = $ws.UsedRange
        $rows = 0; $cols = 0
        try { $rows = $ur.Rows.Count; $cols = $ur.Columns.Count } catch {}
        Log ("SHEET vis=" + $ws.Visible + " rows=" + $rows + " cols=" + $cols + " name=" + $ws.Name)
    }

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
