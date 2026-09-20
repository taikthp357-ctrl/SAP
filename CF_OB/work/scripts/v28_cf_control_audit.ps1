# v28_cf_control_audit.ps1 — Audit current CF_CONTROL structure
$ErrorActionPreference = 'Stop'
$cp10 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP10_POSTING_REGION.xlsm'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

try {
    $wb = $xl.Workbooks.Open($cp10, 0, $true)  # ReadOnly

    Write-Host "=== CURRENT SHEET STRUCTURE ==="
    $sheets = @()
    foreach ($ws in $wb.Worksheets) {
        $visible = switch ($ws.Visible) {
            -1 { "VISIBLE" }
            0 { "HIDDEN" }
            2 { "VERYHIDDEN" }
            default { "UNKNOWN" }
        }
        $sheets += "$($ws.Name) | $visible"
    }
    $sheets | Sort-Object | ForEach-Object { Write-Host "  $_" }

    Write-Host "`n=== INPUT SHEET (if exists) ==="
    $wsInput = $null
    try {
        $wsInput = $wb.Worksheets.Item('INPUT')
        $inputLR = $wsInput.Cells.Item($wsInput.Rows.Count, 1).End(-4162).Row
        $inputLC = $wsInput.Cells.Item(1, $wsInput.Columns.Count).End(-4159).Column
        Write-Host "  UsedRange: rows=$inputLR, cols=$inputLC"
        Write-Host "  Control panel location: checking..."

        # Check for buttons/shapes
        $shapes = @($wsInput.Shapes)
        Write-Host "  Shapes/Buttons: $($shapes.Count)"
        foreach ($shp in $shapes | Select-Object -First 10) {
            Write-Host "    $($shp.Name) | Type=$($shp.Type) | Top=$([int]$shp.Top) | Left=$([int]$shp.Left)"
        }
    } catch {
        Write-Host "  INPUT sheet not found or error: $($_.Exception.Message)"
    }

    Write-Host "`n=== DATA_CORE STRUCTURE ==="
    $wsDC = $wb.Worksheets.Item('DATA_CORE')
    $dcLR = $wsDC.Cells.Item($wsDC.Rows.Count, 1).End(-4162).Row
    $dcLC = $wsDC.Cells.Item(4, $wsDC.Columns.Count).End(-4159).Column
    Write-Host "  UsedRange: rows=$dcLR, cols=$dcLC"
    Write-Host "  Headers (row 4, first 15 cols):"
    for ($c = 1; $c -le 15; $c++) {
        $hdr = $wsDC.Cells.Item(4, $c).Value2
        if ($hdr) { Write-Host "    Col $c = $hdr" }
    }

    Write-Host "`n=== BUTTONS/MACROS CHECK ==="
    # Check VBA modules for button handlers
    $buttonProcs = @('CF_Check', 'CF_Execute', 'CF_ResetUI', 'CF_RefreshInputWorklist')
    foreach ($proc in $buttonProcs) {
        $found = $false
        foreach ($comp in $wb.VBProject.VBComponents) {
            if ($comp.Type -eq 1) {  # Standard module
                $codeLines = $comp.CodeModule.CountOfLines
                if ($codeLines -gt 0) {
                    $code = $comp.CodeModule.Lines(1, $codeLines)
                    if ($code -match "Sub\s+$proc\b") {
                        Write-Host "  $proc → $($comp.Name)"
                        $found = $true
                        break
                    }
                }
            }
        }
        if (-not $found) { Write-Host "  $proc → NOT FOUND" }
    }

    $wb.Close($false)
    Write-Host "`nAUDIT_COMPLETE"

} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    throw
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
