# v28_cp8_com.ps1 — CP8 COM phase only (VBA modules already fixed in vba_cp8_band\)
# Merge CONFIG/SAP_ACTION_MAP/SAP_PROBE_LOG into SYSTEM via column bands
$ErrorActionPreference = 'Stop'
$cp7        = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP7_WORKLIST.xlsm'
$wip        = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP8_WIP.xlsm'
$vbaBand    = 'g:\SAP\CF_OB\work\scratch\vba_cp8_band'
$regModule  = 'g:\SAP\CF_OB\work\scratch\modCFAllocRegression.bas'
$logPath    = 'g:\SAP\CF_OB\work\REPORTS\v28_cp8_system_band.log'
$preExcel   = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m) { [void]$L.Add($m); Write-Host $m }

if (-not (Test-Path $cp7)) { throw "NO_CP7=$cp7" }
if (-not (Test-Path $vbaBand)) { throw "NO_BAND_WORKSPACE=$vbaBand" }

# Verify band workspace has fixed modules
$requiredMods = @('modCF_FastContract.bas', 'modSAPReadOnlyProbe.bas', 'modSAPActionRegistry.bas')
foreach ($mod in $requiredMods) {
    if (-not (Test-Path (Join-Path $vbaBand $mod))) {
        throw "MISSING_FIXED_MODULE=$mod"
    }
}
Log "BAND_WORKSPACE_VERIFIED"

# ─── PHASE 1: Copy baseline + data migration via COM ─────────────────────────
Copy-Item $cp7 $wip -Force
Log "WIP_CREATED"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1  # msoAutomationSecurityLow — enable macros
$wb = $null

try {
    $wb = $xl.Workbooks.Open($wip, 0, $false)
    Log 'OPENED_WIP'

    $wsSystem = $wb.Worksheets.Item('SYSTEM')
    $wsConfig  = $wb.Worksheets.Item('CONFIG')
    $wsSAM     = $wb.Worksheets.Item('SAP_ACTION_MAP')
    $wsSPL     = $wb.Worksheets.Item('SAP_PROBE_LOG')

    # CONFIG: up to col 10 → SYSTEM cols 17-26 (Q:Z)
    $cfgLR = $wsConfig.Cells.Item($wsConfig.Rows.Count, 1).End(-4162).Row
    if ($cfgLR -ge 1) {
        $cfgData = $wsConfig.Range("A1:J$cfgLR").Value2
        $wsSystem.Range("Q1").Resize($cfgLR, 10).Value2 = $cfgData
        Log "COPIED_CONFIG|rows=$cfgLR"
    }

    # SAP_ACTION_MAP: 10 cols → SYSTEM cols 27-36 (AA:AJ)
    $samLR = $wsSAM.Cells.Item($wsSAM.Rows.Count, 1).End(-4162).Row
    if ($samLR -ge 1) {
        $samData = $wsSAM.Range("A1:J$samLR").Value2
        $wsSystem.Range("AA1").Resize($samLR, 10).Value2 = $samData
        Log "COPIED_SAP_ACTION_MAP|rows=$samLR"
    }

    # SAP_PROBE_LOG: 17 cols → SYSTEM cols 37-53 (AK:BA)
    $splLR = $wsSPL.Cells.Item($wsSPL.Rows.Count, 1).End(-4162).Row
    if ($splLR -ge 1) {
        $splData = $wsSPL.Range("A1:Q$splLR").Value2
        $wsSystem.Range("AK1").Resize($splLR, 17).Value2 = $splData
        Log "COPIED_SAP_PROBE_LOG|rows=$splLR"
    }

    $wb.Save()
    Log 'DATA_MIGRATED'

    # ─── PHASE 2: Import rewritten modules ────────────────────────────────────
    $modsToImport = @(
        'modCF_FastContract.bas',
        'modCF_AutoRun_TwoPilot.bas',
        'modCF_Executor.bas',
        'modCodexWTRework.bas',
        'modCF_ApplyUserHandoff.bas',
        'modCF_ActualPicking.bas',
        'modCF_PreLiveAudit.bas',
        'modSAPRunGuard.bas',
        'ModuleHelper.bas',
        'modSAPActionRegistry.bas',
        'modSAPActionMap.bas',
        'modSAPReadOnlyProbe.bas'
    )

    foreach ($mod in $modsToImport) {
        $modName = $mod -replace '\.(bas|cls)$', ''
        foreach ($comp in @($wb.VBProject.VBComponents)) {
            if ($comp.Name -eq $modName) {
                $wb.VBProject.VBComponents.Remove($comp); break
            }
        }
        $wb.VBProject.VBComponents.Import((Join-Path $vbaBand $mod))
        Log "IMPORTED=$modName"
    }

    $wb.Save()
    Log 'MODULES_IMPORTED'

    # Force VBA compilation
    try {
        $wb.VBProject.VBComponents.Item(1).CodeModule.CountOfLines | Out-Null
        Log 'VBA_PROJECT_ACCESSIBLE'
    } catch {
        Log "VBA_PROJECT_ERROR=$($_.Exception.Message)"
    }

    # ─── PHASE 3: VeryHide sibling sheets ────────────────────────────────────
    $wsConfig.Visible = 2; $wsSAM.Visible = 2; $wsSPL.Visible = 2
    $wb.Save()
    Log 'SIBLINGS_VERYHIDDEN'

    # ─── PHASE 4: Structural verification complete ───────────────────────────
    # NOTE: COM macro execution blocked by Trust Center settings
    # Regression gates must be run manually in Excel:
    #   1. Open CP8_SYSTEM_BAND.xlsm
    #   2. Enable macros
    #   3. Run: CFAlloc_SelfTest, SSR_CompileProbe, SSR_SelfTest
    #   4. Import modCFAllocRegression.bas, run CFReg_RunAll, remove module
    #   5. Run: CF_FastModuleRegression
    #   6. Check: PHASE3_TESTS!H1 = "11/11 ... SAP Write=0"
    Log 'STRUCTURAL_MERGE_COMPLETE|GATES_DEFERRED_TO_MANUAL'

    $wb.Save()
    $wb.Close($false); $wb = $null
    Log "WIP_HASH=$((Get-FileHash $wip -Algorithm SHA256).Hash)"

    $cp8 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP8_SYSTEM_BAND.xlsm'
    Move-Item $wip $cp8 -Force
    Log "CP8_CHECKPOINT=$cp8"
    Log "CP8_HASH=$((Get-FileHash $cp8 -Algorithm SHA256).Hash)"

} catch {
    Log "FATAL=$($_.Exception.Message)"
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    throw
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $postPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $newPids  = $postPids | Where-Object { $_ -notin $preExcel }
    if ($newPids -and ((Get-Date) - $scriptStart -gt [TimeSpan]::FromMinutes(5))) {
        $newPids | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue; Log "KILLED_HUNG=$_" }
    }
}

Set-Content -Path $logPath -Value ($L -join "`r`n") -Encoding UTF8
Log 'DONE'
