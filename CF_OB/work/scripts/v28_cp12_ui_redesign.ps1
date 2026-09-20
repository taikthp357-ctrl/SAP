# v28_cp12_ui_redesign.ps1 — Build CP12 with 28-point UI redesign
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# Paths
# ─────────────────────────────────────────────────────────────────────────────
$cp10 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP10_POSTING_REGION.xlsm'
$wip = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP12_WIP.xlsm'
$vbaStaging = 'g:\SAP\CF_OB\work\scratch\vba_v28_ui'
$logPath = 'g:\SAP\CF_OB\work\REPORTS\v28_cp12_ui_redesign.log'

$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m) { [void]$L.Add($m); Write-Host $m }

# ─────────────────────────────────────────────────────────────────────────────
# Validation
# ─────────────────────────────────────────────────────────────────────────────
if (-not (Test-Path $cp10)) { throw "NO_CP10=$cp10" }
if (-not (Test-Path $vbaStaging)) { throw "NO_STAGING=$vbaStaging" }

# ─────────────────────────────────────────────────────────────────────────────
# Phase 1: Copy baseline and open
# ─────────────────────────────────────────────────────────────────────────────
Copy-Item $cp10 $wip -Force
Log "WIP_CREATED"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$wb = $null

try {
    $wb = $xl.Workbooks.Open($wip, 0, $false)
    Log 'OPENED_WIP'

    # ─────────────────────────────────────────────────────────────────────────
    # Phase 2: Remove CP11 legacy UI modules (if present)
    # ─────────────────────────────────────────────────────────────────────────
    $cp11Modules = @('modCF_ControlLayout', 'modCF_ControlButtons', 'modCF_ControlData')
    foreach ($modName in $cp11Modules) {
        foreach ($comp in @($wb.VBProject.VBComponents)) {
            if ($comp.Name -eq $modName) {
                $wb.VBProject.VBComponents.Remove($comp)
                Log "REMOVED_CP11_MODULE=$modName"
                break
            }
        }
    }

    # ─────────────────────────────────────────────────────────────────────────
    # Phase 3: Import updated modules from vba_v28_ui staging
    # ─────────────────────────────────────────────────────────────────────────
    $modulesToImport = @(
        'modCF_InputLayout.bas',
        'modCF_InputWorklist.bas',
        'modCF_ControlButtons.bas',
        'modCFResetUI.bas'
    )

    foreach ($mod in $modulesToImport) {
        $modPath = Join-Path $vbaStaging $mod
        if (-not (Test-Path $modPath)) {
            Log "WARN_MISSING=$mod"
            continue
        }

        $modName = $mod -replace '\.(bas|cls)$', ''

        # Remove old version if exists
        foreach ($comp in @($wb.VBProject.VBComponents)) {
            if ($comp.Name -eq $modName) {
                $wb.VBProject.VBComponents.Remove($comp)
                Log "REMOVED_OLD=$modName"
                break
            }
        }

        # Import new version
        $wb.VBProject.VBComponents.Import($modPath)
        Log "IMPORTED=$modName"
    }

    $wb.Save()
    Log 'MODULES_UPDATED'

    # ─────────────────────────────────────────────────────────────────────────
    # Phase 3b: Update ThisWorkbook code module
    # ─────────────────────────────────────────────────────────────────────────
    $thiswbPath = Join-Path $vbaStaging 'ThisWorkbook.cls'
    if (Test-Path $thiswbPath) {
        $thiswbCode = Get-Content $thiswbPath -Raw -Encoding UTF8
        $codeStart = $thiswbCode.IndexOf('Private Sub Workbook_Open()')
        if ($codeStart -ge 0) {
            $codeToInsert = $thiswbCode.Substring($codeStart)
            $thiswb = $wb.VBProject.VBComponents.Item('ThisWorkbook')
            $lineCount = $thiswb.CodeModule.CountOfLines
            if ($lineCount -gt 0) {
                $thiswb.CodeModule.DeleteLines(1, $lineCount)
            }
            $thiswb.CodeModule.AddFromString($codeToInsert)
            Log 'THISWORKBOOK_UPDATED'
        }
    }

    $wb.Save()

    # ─────────────────────────────────────────────────────────────────────────
    # Phase 4: Rename INPUT → CF_CONTROL (if not already renamed)
    # ─────────────────────────────────────────────────────────────────────────
    $wsInput = $null
    try {
        $wsInput = $wb.Worksheets.Item('INPUT')
        $wsInput.Name = 'CF_CONTROL'
        Log 'RENAMED_INPUT_TO_CF_CONTROL'
    } catch {
        # Already renamed or doesn't exist
        try {
            $wsInput = $wb.Worksheets.Item('CF_CONTROL')
            Log 'CF_CONTROL_ALREADY_EXISTS'
        } catch {
            Log 'WARN_NO_INPUT_OR_CF_CONTROL_SHEET'
        }
    }

    # ─────────────────────────────────────────────────────────────────────────
    # Phase 5: VeryHide technical sheets (keep only CF_CONTROL visible)
    # ─────────────────────────────────────────────────────────────────────────
    $technicalSheets = @(
        'DATA_CORE', 'PLAN', 'RUN_LOG', 'RUN_HISTORY',
        'OB_QUEUE', 'OB_WH_CACHE', 'PHASE3_TESTS', 'REGION_POSTING'
    )

    foreach ($sheetName in $technicalSheets) {
        try {
            $ws = $wb.Worksheets.Item($sheetName)
            $ws.Visible = [Microsoft.Office.Interop.Excel.XlSheetVisibility]::xlSheetVeryHidden
            Log "VERYHIDDEN=$sheetName"
        } catch {
            # Sheet doesn't exist, skip
        }
    }

    # Ensure CF_CONTROL is visible
    try {
        $wsCF = $wb.Worksheets.Item('CF_CONTROL')
        $wsCF.Visible = [Microsoft.Office.Interop.Excel.XlSheetVisibility]::xlSheetVisible
        Log 'CF_CONTROL_VISIBLE'
    } catch {
        Log 'WARN_CF_CONTROL_NOT_FOUND'
    }

    $wb.Save()

    # ─────────────────────────────────────────────────────────────────────────
    # Phase 4b: Update CF_CONTROL sheet code module (event handlers)
    # ─────────────────────────────────────────────────────────────────────────
    $sheetCodePath = 'g:\SAP\CF_OB\work\scratch\vba_cf_control\Sheet_CF_CONTROL.cls'
    if (Test-Path $sheetCodePath) {
        $sheetCode = Get-Content $sheetCodePath -Raw -Encoding UTF8
        $codeStart = $sheetCode.IndexOf('Private Sub Worksheet_')
        if ($codeStart -ge 0) {
            $codeToInsert = $sheetCode.Substring($codeStart)

            # Get sheet code module (use CodeName, not sheet name)
            $sheetComp = $null
            foreach ($comp in $wb.VBProject.VBComponents) {
                if ($comp.Name -like 'Sheet*') {
                    # Check if this sheet object references CF_CONTROL
                    try {
                        $sheetObj = $wb.Worksheets.Item($comp.Properties.Item('Name').Value)
                        if ($sheetObj.Name -eq 'CF_CONTROL') {
                            $sheetComp = $comp
                            break
                        }
                    } catch {
                        # Sheet doesn't exist or not the right one
                    }
                }
            }

            if ($sheetComp -ne $null) {
                $lineCount = $sheetComp.CodeModule.CountOfLines
                if ($lineCount -gt 0) {
                    $sheetComp.CodeModule.DeleteLines(1, $lineCount)
                }
                $sheetComp.CodeModule.AddFromString($codeToInsert)
                Log "CF_CONTROL_SHEET_CODE_UPDATED=$($sheetComp.Name)"
            } else {
                Log 'WARN_CF_CONTROL_SHEET_NOT_FOUND'
            }
        }
    }

    $wb.Save()
    Log 'SHEET_CODE_UPDATED'

    # ─────────────────────────────────────────────────────────────────────────
    # Phase 5: VeryHide technical sheets (keep only CF_CONTROL visible)
    # ─────────────────────────────────────────────────────────────────────────
    # NOTE: COM macro execution is blocked by Trust Center in automation mode.
    # User must open CP12 manually and run: CF_ApplyInputLayout
    Log 'UI_MODULES_IMPORTED|BUILD_DEFERRED_TO_MANUAL'

    # ─────────────────────────────────────────────────────────────────────────
    # Phase 7: Finalize and rename to CP12
    # ─────────────────────────────────────────────────────────────────────────
    $wb.Save()
    $wb.Close($false)
    $wb = $null
    Log "WIP_HASH=$((Get-FileHash $wip -Algorithm SHA256).Hash)"

    $cp12 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP12_UI_REDESIGN.xlsm'
    Move-Item $wip $cp12 -Force
    Log "CP12_CHECKPOINT=$cp12"
    Log "CP12_HASH=$((Get-FileHash $cp12 -Algorithm SHA256).Hash)"

} catch {
    Log "FATAL=$($_.Exception.Message)"
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    throw
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # Kill hung Excel processes (only if script ran > 5 min)
    $postPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $newPids = $postPids | Where-Object { $_ -notin $preExcel }
    if ($newPids -and ((Get-Date) - $scriptStart -gt [TimeSpan]::FromMinutes(5))) {
        $newPids | ForEach-Object {
            Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue
            Log "KILLED_HUNG=$_"
        }
    }
}

Set-Content -Path $logPath -Value ($L -join "`r`n") -Encoding UTF8
Log 'DONE'
