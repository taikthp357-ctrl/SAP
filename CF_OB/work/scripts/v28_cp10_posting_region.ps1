# v28_cp10_posting_region.ps1 — POSTING→REGION migration (final structural change)
$ErrorActionPreference = 'Stop'
$cp9 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP9_HISTORY_CLEARED.xlsm'
$wip = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP10_WIP.xlsm'
$vbaRegion = 'g:\SAP\CF_OB\work\scratch\vba_posting_region'
$logPath = 'g:\SAP\CF_OB\work\REPORTS\v28_cp10_posting_region.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m) { [void]$L.Add($m); Write-Host $m }

if (-not (Test-Path $cp9)) { throw "NO_CP9=$cp9" }
if (-not (Test-Path $vbaRegion)) { throw "NO_REGION_WORKSPACE=$vbaRegion" }

Copy-Item $cp9 $wip -Force
Log "WIP_CREATED"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$wb = $null

try {
    $wb = $xl.Workbooks.Open($wip, 0, $false)
    Log 'OPENED_WIP'

    # ─── PHASE 1: Copy POSTING data to DATA_CORE REGION band ────────────────
    $wsDC = $wb.Worksheets.Item('DATA_CORE')
    $wsPosting = $wb.Worksheets.Item('POSTING')

    # POSTING: 23 cols but consumers only use 1-17; copy all 23 for safety
    $pLR = $wsPosting.Cells.Item($wsPosting.Rows.Count, 1).End(-4162).Row
    if ($pLR -ge 1) {
        $postingData = $wsPosting.Range("A1:W$pLR").Value2  # A:W = 23 cols
        $wsDC.Range("AE1").Resize($pLR, 23).Value2 = $postingData  # AE = col 31
        Log "COPIED_POSTING_TO_REGION|rows=$pLR|cols=23|range=AE1:BA$pLR"
    }

    $wb.Save()
    Log 'DATA_MIGRATED'

    # ─── PHASE 2: Import migrated modules ────────────────────────────────────
    $modsToImport = @(
        'modSourceStockRepo.bas',
        'ModuleSAP.bas',
        'modSAPReadOnlyProbe.bas'
    )

    foreach ($mod in $modsToImport) {
        $modName = $mod -replace '\.(bas|cls)$', ''
        foreach ($comp in @($wb.VBProject.VBComponents)) {
            if ($comp.Name -eq $modName) {
                $wb.VBProject.VBComponents.Remove($comp); break
            }
        }
        $wb.VBProject.VBComponents.Import((Join-Path $vbaRegion $mod))
        Log "IMPORTED=$modName"
    }

    $wb.Save()
    Log 'MODULES_IMPORTED'

    # ─── PHASE 3: VeryHide POSTING, Delete POSTING_RAW ───────────────────────
    $wsPosting.Visible = 2  # xlSheetVeryHidden
    Log 'POSTING_VERYHIDDEN'

    $wsPostingRaw = $null
    try {
        $wsPostingRaw = $wb.Worksheets.Item('POSTING_RAW')
        $wsPostingRaw.Visible = -1  # xlSheetVisible (unhide if hidden)
        $xl.DisplayAlerts = $false
        $wsPostingRaw.Delete()
        Log 'POSTING_RAW_DELETED'
    } catch {
        Log "POSTING_RAW_DELETE_FAILED|$($_.Exception.Message)"
    }

    $wb.Save()
    Log 'STRUCTURAL_CHANGES_COMPLETE'

    # ─── PHASE 4: Save and finalize ──────────────────────────────────────────
    # NOTE: COM macro execution blocked by Trust Center settings
    # Regression gates must be run manually in Excel:
    #   1. Open CP10_POSTING_REGION.xlsm
    #   2. Enable macros
    #   3. Run: CFAlloc_SelfTest, SSR_CompileProbe, SSR_SelfTest (REGION mode verify)
    #   4. Check: PHASE3_TESTS!H1 = "11/11 ... SAP Write=0"
    Log 'STRUCTURAL_MIGRATION_COMPLETE|GATES_DEFERRED_TO_MANUAL'

    $wb.Save()
    $wb.Close($false)
    $wb = $null
    Log "WIP_HASH=$((Get-FileHash $wip -Algorithm SHA256).Hash)"

    $cp10 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP10_POSTING_REGION.xlsm'
    Move-Item $wip $cp10 -Force
    Log "CP10_CHECKPOINT=$cp10"
    Log "CP10_HASH=$((Get-FileHash $cp10 -Algorithm SHA256).Hash)"

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
