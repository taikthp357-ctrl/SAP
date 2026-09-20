# v28_cf_control_build.ps1 — Build CF_CONTROL UI on CP10
$ErrorActionPreference = 'Stop'
$cp10 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP10_POSTING_REGION.xlsm'
$wip = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP11_WIP.xlsm'
$vbaControl = 'g:\SAP\CF_OB\work\scratch\vba_cf_control'
$logPath = 'g:\SAP\CF_OB\work\REPORTS\v28_cf_control_build.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m) { [void]$L.Add($m); Write-Host $m }

if (-not (Test-Path $cp10)) { throw "NO_CP10=$cp10" }
if (-not (Test-Path $vbaControl)) { throw "NO_CONTROL_WORKSPACE=$vbaControl" }

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

    # ─── PHASE 1: Import CF_CONTROL modules ──────────────────────────────────
    $modsToImport = @(
        'modCF_ControlLayout.bas',
        'modCF_ControlButtons.bas',
        'modCF_ControlData.bas'
    )

    foreach ($mod in $modsToImport) {
        $modName = $mod -replace '\.(bas|cls)$', ''
        # Remove old module if exists
        foreach ($comp in @($wb.VBProject.VBComponents)) {
            if ($comp.Name -eq $modName) {
                $wb.VBProject.VBComponents.Remove($comp)
                Log "REMOVED_OLD=$modName"
                break
            }
        }
        $wb.VBProject.VBComponents.Import((Join-Path $vbaControl $mod))
        Log "IMPORTED=$modName"
    }

    $wb.Save()
    Log 'MODULES_IMPORTED'

    # ─── PHASE 2: Update ThisWorkbook code module ────────────────────────────
    $thiswbCode = Get-Content (Join-Path $vbaControl 'ThisWorkbook.cls') -Raw -Encoding UTF8
    # Extract only the code part (skip Attribute lines)
    $codeStart = $thiswbCode.IndexOf("Private Sub Workbook_Open()")
    if ($codeStart -ge 0) {
        $codeToInsert = $thiswbCode.Substring($codeStart)

        $thiswb = $wb.VBProject.VBComponents.Item('ThisWorkbook')
        $lineCount = $thiswb.CodeModule.CountOfLines

        # Check if Workbook_Open already exists
        $existingCode = ""
        if ($lineCount -gt 0) {
            $existingCode = $thiswb.CodeModule.Lines(1, $lineCount)
        }

        if ($existingCode -notmatch 'Sub Workbook_Open\(\)') {
            # Add new Workbook_Open
            $thiswb.CodeModule.AddFromString($codeToInsert)
            Log 'ADDED_WORKBOOK_OPEN'
        } else {
            Log 'WORKBOOK_OPEN_EXISTS_SKIPPED'
        }
    }

    $wb.Save()
    Log 'THISWORKBOOK_UPDATED'

    # ─── PHASE 3: Rename INPUT → CF_CONTROL and update sheet code ────────────
    $wsInput = $null
    try {
        $wsInput = $wb.Worksheets.Item('INPUT')
        $wsInput.Name = 'CF_CONTROL'
        Log 'RENAMED_INPUT_TO_CF_CONTROL'
    } catch {
        # Maybe already renamed
        $wsInput = $wb.Worksheets.Item('CF_CONTROL')
        Log 'CF_CONTROL_ALREADY_EXISTS'
    }

    # Update sheet code module
    $sheetCode = Get-Content (Join-Path $vbaControl 'Sheet_CF_CONTROL.cls') -Raw -Encoding UTF8
    $codeStart = $sheetCode.IndexOf("Private Sub Worksheet_")
    if ($codeStart -ge 0) {
        $codeToInsert = $sheetCode.Substring($codeStart)

        $sheetModule = $wsInput.CodeName
        $sheetComp = $wb.VBProject.VBComponents.Item($sheetModule)

        # Clear existing and add new
        $lineCount = $sheetComp.CodeModule.CountOfLines
        if ($lineCount -gt 0) {
            $sheetComp.CodeModule.DeleteLines(1, $lineCount)
        }
        $sheetComp.CodeModule.AddFromString($codeToInsert)
        Log 'SHEET_CODE_UPDATED'
    }

    $wb.Save()
    Log 'SHEET_RENAMED_AND_CODE_UPDATED'

    # ─── PHASE 4: Build CF_CONTROL layout ─────────────────────────────────────
    # NOTE: COM macro execution blocked by Trust Center
    # User must open CP11 manually and run: CF_BuildControlLayout
    Log 'CF_CONTROL_MODULES_IMPORTED|BUILD_DEFERRED_TO_MANUAL'

    $wb.Save()
    $wb.Close($false)
    $wb = $null
    Log "WIP_HASH=$((Get-FileHash $wip -Algorithm SHA256).Hash)"

    $cp11 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP11_CF_CONTROL.xlsm'
    Move-Item $wip $cp11 -Force
    Log "CP11_CHECKPOINT=$cp11"
    Log "CP11_HASH=$((Get-FileHash $cp11 -Algorithm SHA256).Hash)"

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
