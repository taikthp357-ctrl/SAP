# v28_cp7_worklist.ps1 — CP7: Phase 1b INPUT worklist mirror with Vietnamese feedback
# Imports 4 modules (1 new + 3 modified) + patches ThisWorkbook.Workbook_Open
$ErrorActionPreference = 'Stop'
$cp6        = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP6_GRIDFIX.xlsm'
$wip        = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP7_WIP.xlsm'
$vbaBase    = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$regModule  = 'g:\SAP\CF_OB\work\scratch\modCFAllocRegression.bas'
$logPath    = 'g:\SAP\CF_OB\work\REPORTS\v28_cp7_worklist.log'
$preExcel   = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m) { [void]$L.Add($m); Write-Host $m }

if (-not (Test-Path $cp6)) { throw "NO_CP6=$cp6" }

# ── PHASE 1: Fresh WIP from CP6 ──────────────────────────────────────────────
Copy-Item $cp6 $wip -Force
Log "WIP_CREATED"

# ── PHASE 2: Import modules via COM ──────────────────────────────────────────
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($wip, 0, $false)
    Log 'OPENED_WIP'

    # Remove old versions of the 4 modules
    foreach ($name in @('modCF_InputWorklist', 'modCF_InputLayout', 'modCFResetUI', 'modCF_CheckPlanner')) {
        foreach ($comp in @($wb.VBProject.VBComponents)) {
            if ($comp.Name -eq $name) {
                $wb.VBProject.VBComponents.Remove($comp)
                Log "REMOVED=$name"
                break
            }
        }
    }

    # Import modified/new modules
    $wb.VBProject.VBComponents.Import((Join-Path $vbaBase 'modCF_InputWorklist.bas'))
    Log 'IMPORTED_modCF_InputWorklist'
    $wb.VBProject.VBComponents.Import((Join-Path $vbaBase 'modCF_InputLayout.bas'))
    Log 'IMPORTED_modCF_InputLayout'
    $wb.VBProject.VBComponents.Import((Join-Path $vbaBase 'modCFResetUI.bas'))
    Log 'IMPORTED_modCFResetUI'
    $wb.VBProject.VBComponents.Import((Join-Path $vbaBase 'modCF_CheckPlanner.bas'))
    Log 'IMPORTED_modCF_CheckPlanner'

    # ── PHASE 3: Patch ThisWorkbook.Workbook_Open ────────────────────────────
    $twComp = $null
    foreach ($c in @($wb.VBProject.VBComponents)) {
        if ($c.Name -eq 'ThisWorkbook') { $twComp = $c; break }
    }
    if ($null -eq $twComp) { throw "THISWORKBOOK_NOT_FOUND" }
    $cm = $twComp.CodeModule
    $n = $cm.CountOfLines
    for ($li = 1; $li -le $n; $li++) {
        if ($cm.Lines($li, 1) -match 'On Error GoTo 0') {
            if ($li -gt 1 -and $cm.Lines($li - 1, 1) -match 'RT_InitializeUI') {
                # Insert after this "On Error GoTo 0"
                $cm.InsertLines($li + 1, '    If ThisWorkbook.Worksheets("DATA_CORE").Cells(5, 1).Value2 <> "" Then CF_RefreshInputWorklist')
                Log "PATCHED_WORKBOOK_OPEN"
                break
            }
        }
    }

    # ── PHASE 4: Apply layout (triggers CF_ApplyInputLayout) ─────────────────
    $wb.Save()
    Log 'SAVED_AFTER_IMPORT'
    $xl.Run('RT_InitializeUI')
    Log 'APPLIED_LAYOUT'

    # ── PHASE 5: Offline probe (seed DATA_CORE, call CF_RefreshInputWorklist)
    $wsWH = $wb.Worksheets.Item('DATA_CORE')
    $wsInput = $wb.Worksheets.Item('INPUT')

    # Clear existing DATA_CORE rows 5-12
    $wsWH.Range('A5:AT12').ClearContents

    # Seed 8 synthetic rows
    $wsWH.Cells.Item(5, 1).Value2 = 'TEST_OB_001'
    $wsWH.Cells.Item(5, 2).Value2 = 'DOC001'
    $wsWH.Cells.Item(5, 6).Value2 = 'MAT001'
    $wsWH.Cells.Item(5, 10).Value2 = ''       # Status blank (pending)
    $wsWH.Cells.Item(5, 18).Value2 = 'PASS'
    $wsWH.Cells.Item(5, 19).Value2 = ''

    $wsWH.Cells.Item(6, 1).Value2 = 'TEST_OB_002'
    $wsWH.Cells.Item(6, 2).Value2 = 'DOC002'
    $wsWH.Cells.Item(6, 6).Value2 = 'MAT002'
    $wsWH.Cells.Item(6, 10).Value2 = 'C'      # Status C (completed, should be hidden)
    $wsWH.Cells.Item(6, 18).Value2 = 'DONE'
    $wsWH.Cells.Item(6, 19).Value2 = ''

    $wsWH.Cells.Item(7, 1).Value2 = 'TEST_OB_003'
    $wsWH.Cells.Item(7, 2).Value2 = 'DOC003'
    $wsWH.Cells.Item(7, 6).Value2 = 'MAT003'
    $wsWH.Cells.Item(7, 10).Value2 = ''
    $wsWH.Cells.Item(7, 18).Value2 = 'BLOCKED_SOURCE_PLAN_INCOMPLETE'
    $wsWH.Cells.Item(7, 19).Value2 = 'Multi-batch stock requires specific batch selection'

    $wsWH.Cells.Item(8, 1).Value2 = 'TEST_OB_004'
    $wsWH.Cells.Item(8, 2).Value2 = 'DOC004'
    $wsWH.Cells.Item(8, 6).Value2 = 'MAT004'
    $wsWH.Cells.Item(8, 10).Value2 = ''
    $wsWH.Cells.Item(8, 18).Value2 = 'BLOCKED_QTY_GROUP_BASELINE'
    $wsWH.Cells.Item(8, 19).Value2 = ''

    Log 'SEEDED_DATA_CORE'

    # Call CF_RefreshInputWorklist
    $xl.Run('CF_RefreshInputWorklist')
    Log 'RAN_CF_RefreshInputWorklist'

    # Assert: INPUT col N (col 14) has Vietnamese strings
    $vnResult = $wsInput.Cells.Item(5, 14).Value2
    if ($vnResult -notmatch 'sàng|S.*n s.*ng') { throw "PROBE_FAIL_VN_NOT_FOUND|got=$vnResult" }
    Log "PROBE_VN_OK|row5_col14=$vnResult"

    # Assert: AutoFilter active on INPUT
    if (-not $wsInput.AutoFilterMode) { throw "PROBE_FAIL_NO_AUTOFILTER" }
    Log 'PROBE_AUTOFILTER_OK'

    # Clear probe data
    $wsWH.Range('A5:AT12').ClearContents
    Log 'CLEARED_PROBE_DATA'

    $wb.Save()
    Log 'SAVED_WIP_AFTER_PROBE'

    # ── PHASE 6: Regression gates ────────────────────────────────────────────
    # Gate 1: CFAlloc_SelfTest
    $r1 = $xl.Run('CFAlloc_SelfTest')
    if ($r1 -ne 'PASS') { throw "GATE_CFAlloc_FAIL|$r1" }
    Log "GATE_CFAlloc=PASS"

    # Gate 2: SSR_CompileProbe (Sub, no return)
    $xl.Run('SSR_CompileProbe')
    Log "GATE_SSR_CompileProbe=RAN"

    # Gate 3: SSR_SelfTest
    $r3 = $xl.Run('SSR_SelfTest')
    if ($r3 -ne 'PASS') { throw "GATE_SSR_SelfTest_FAIL|$r3" }
    Log "GATE_SSR_SelfTest=PASS"

    # Gate 4: CFReg_RunAll (need to import modCFAllocRegression first)
    if (Test-Path $regModule) {
        foreach ($comp in @($wb.VBProject.VBComponents)) {
            if ($comp.Name -eq 'modCFAllocRegression') {
                $wb.VBProject.VBComponents.Remove($comp)
                break
            }
        }
        $wb.VBProject.VBComponents.Import($regModule)
        Log 'IMPORTED_modCFAllocRegression'
    }
    $r4 = $xl.Run('CFReg_RunAll')
    if ($r4 -notmatch 'REG_ALL_PASS.*FAIL=0') { throw "GATE_CFReg_FAIL|$r4" }
    Log "GATE_CFReg=PASS|$r4"
    # Remove modCFAllocRegression after test
    foreach ($comp in @($wb.VBProject.VBComponents)) {
        if ($comp.Name -eq 'modCFAllocRegression') {
            $wb.VBProject.VBComponents.Remove($comp)
            Log 'REMOVED_modCFAllocRegression'
            break
        }
    }

    # Gate 5: CF_FastModuleRegression
    $r5 = $xl.Run('CF_FastModuleRegression')
    if (-not ($r5 -like 'PASS*')) { throw "GATE_FastReg_FAIL|$r5" }
    Log "GATE_FastReg=PASS"

    # Gate 6: PHASE3_TESTS
    $wsTest = $wb.Worksheets.Item('PHASE3_TESTS')
    $r6 = $wsTest.Range('H1').Value2
    if ($r6 -notmatch '11/11.*SAP Write=0') { throw "GATE_Phase3_FAIL|$r6" }
    Log "GATE_Phase3=PASS|$r6"

    Log 'ALL_GATES_GREEN'

    $wb.Save()
    Log 'SAVED_WIP'
    $wb.Close($false)
    $wb = $null
    Log "WIP_HASH=$((Get-FileHash $wip -Algorithm SHA256).Hash)"

    # Rename WIP to final checkpoint
    $cp7 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP7_WORKLIST.xlsm'
    Move-Item $wip $cp7 -Force
    Log "CP7_CHECKPOINT=$cp7"
    Log "CP7_HASH=$((Get-FileHash $cp7 -Algorithm SHA256).Hash)"

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
