# V2.7 BASELINE INVENTORY - read-only Excel COM snapshot of the V2.6 workbook.
#   - Opens the V2.6 workbook READ-ONLY in a background job (parent watchdog + hard timeout).
#   - Enumerates VBA components (name, type, line count).
#   - Enumerates public Sub/Function entry points per standard module.
#   - Searches every module for legacy-sheet references (POSTING etc. - Section 74).
#   - Enumerates defined names and INPUT-sheet shapes/button macro bindings.
#   - NEVER saves. Original workbook untouched. Kills only the EXCEL PID it spawned.
$ErrorActionPreference = 'Stop'
$wbPath   = 'g:\SAP\CF_OB\260819_CF_OB_V2.6_LICH_SU.xlsm'
$outJson  = 'g:\SAP\CF_OB\work\REPORTS\v27_baseline_inventory.json'
$outLog   = 'g:\SAP\CF_OB\work\REPORTS\v27_baseline_inventory.log'
$timeoutSec = 180

$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date

$job = Start-Job -ScriptBlock {
    param($wbPath, $outJson, $outLog)
    $log = [System.Collections.ArrayList]::new()
    function L($m){ [void]$log.Add($m) }
    $legacySheets = @('ORDER_QUERY','POSTING_RAW','WH_ACTUAL','POSTING','RUN_LOG','SAP_MAP',
        'CONFIG','OB_QUEUE','OB_WH_CACHE','POSTING_SUMMARY','POSTING_CACHE','PLAN_SOURCE_LEGS',
        'PLAN_RELEASE','SAP_PROBE_LOG','SAP_ACTION_MAP','RUN_SUMMARY','REGRESSION_AUDIT',
        'ORDER_QUERY','PLAN','RUN_HISTORY','INPUT','WH_ACTUAL')
    $legacySheets = $legacySheets | Select-Object -Unique
    $result = [ordered]@{}
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
    try {
        $wb = $xl.Workbooks.Open($wbPath, 0, $true)  # ReadOnly
        L("OPENED_READONLY " + $wb.Name)

        # --- Sheets ---
        $sheets = @()
        foreach ($ws in $wb.Worksheets) {
            $sheets += [ordered]@{
                name = $ws.Name; visible = [int]$ws.Visible
                usedRows = $ws.UsedRange.Rows.Count; usedCols = $ws.UsedRange.Columns.Count
            }
        }
        $result.sheets = $sheets
        L("SHEETS=" + $sheets.Count)

        # --- Defined names ---
        $names = @()
        foreach ($n in $wb.Names) { try { $names += [ordered]@{ name=$n.Name; refersTo=$n.RefersTo } } catch {} }
        $result.definedNames = $names
        L("DEFINED_NAMES=" + $names.Count)

        # --- VBA components ---
        $comps = @(); $refHits = @(); $entryPoints = @()
        $proj = $wb.VBProject
        foreach ($c in $proj.VBComponents) {
            $cm = $c.CodeModule
            $lc = $cm.CountOfLines
            $typeName = switch ($c.Type) { 1 {'stdModule'} 2 {'classModule'} 3 {'msForm'} 100 {'document'} default {"t$($c.Type)"} }
            $comps += [ordered]@{ name=$c.Name; type=$typeName; lines=$lc }
            $body = if ($lc -gt 0) { $cm.Lines(1, $lc) } else { '' }
            # legacy sheet references
            foreach ($sh in $legacySheets) {
                $pat = 'Worksheets\("' + [regex]::Escape($sh) + '"\)|Sheets\("' + [regex]::Escape($sh) + '"\)'
                $mm = [regex]::Matches($body, $pat)
                if ($mm.Count -gt 0) { $refHits += [ordered]@{ module=$c.Name; sheet=$sh; count=$mm.Count } }
            }
            # public entry points (Public Sub/Function or bare Sub/Function at col 0)
            foreach ($line in ($body -split "`r?`n")) {
                if ($line -match '^\s*(Public\s+)?(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)') {
                    if ($line -notmatch '^\s*Private') {
                        $entryPoints += [ordered]@{ module=$c.Name; kind=$matches[2]; name=$matches[3] }
                    }
                }
            }
        }
        $result.components = $comps
        $result.legacyRefHits = $refHits
        $result.entryPoints = $entryPoints
        L("VBA_COMPONENTS=" + $comps.Count + " ENTRY_POINTS=" + $entryPoints.Count + " LEGACY_REF_HITS=" + $refHits.Count)

        # --- Shapes / button macro bindings on visible sheets ---
        $shapes = @()
        foreach ($ws in $wb.Worksheets) {
            if ($ws.Visible -ne -1) { continue }  # only visible
            try {
                foreach ($shp in $ws.Shapes) {
                    $macro = ''
                    try { $macro = $shp.OnAction } catch {}
                    if ($macro -ne '' -or $shp.Type -eq 8 -or $shp.Type -eq 1) {
                        $txt = ''
                        try { $txt = $shp.TextFrame2.TextRange.Text } catch {}
                        $shapes += [ordered]@{ sheet=$ws.Name; shape=$shp.Name; onAction=$macro; text=$txt }
                    }
                }
            } catch {}
        }
        $result.shapes = $shapes
        L("SHAPES_WITH_BINDING_OR_BUTTONS=" + $shapes.Count)

        $wb.Close($false)
    } catch {
        L("FATAL=" + $_.Exception.Message)
        try { $wb.Close($false) } catch {}
    } finally {
        $xl.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    $result.log = $log
    $result | ConvertTo-Json -Depth 6 | Set-Content -Path $outJson -Encoding UTF8
    Set-Content -Path $outLog -Value ($log -join "`r`n") -Encoding UTF8
    "DONE"
} -ArgumentList $wbPath, $outJson, $outLog

$deadline = (Get-Date).AddSeconds($timeoutSec)
while ($job.State -eq 'Running' -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 500 }
$timedOut = ($job.State -eq 'Running')
if ($timedOut) { Stop-Job $job; Write-Output ("TIMEOUT after ${timeoutSec}s") }
$out = Receive-Job $job -ErrorAction SilentlyContinue
Remove-Job $job -Force -ErrorAction SilentlyContinue
Write-Output ("JOB_RESULT=" + ($out -join '|'))

foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
    if ($preExcel -notcontains $p.Id) {
        try { if ($p.StartTime -ge $scriptStart) { Stop-Process -Id $p.Id -Force; Write-Output ("KILLED_EXCEL_PID=" + $p.Id) } } catch {}
    }
}
if (Test-Path $outJson) { Write-Output ("INVENTORY_JSON=" + (Get-Item $outJson).Length + " bytes") } else { Write-Output "INVENTORY_JSON=NOT_WRITTEN" }