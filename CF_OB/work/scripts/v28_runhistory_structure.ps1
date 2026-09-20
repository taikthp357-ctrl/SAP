# v28_runhistory_structure.ps1 — Read RUN_LOG and RUN_SUMMARY structure from CP8
$ErrorActionPreference = 'Stop'
$cp8 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP8_SYSTEM_BAND.xlsm'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

try {
    $wb = $xl.Workbooks.Open($cp8, 0, $true)  # ReadOnly

    $wsRL = $wb.Worksheets.Item('RUN_LOG')
    $wsRS = $wb.Worksheets.Item('RUN_SUMMARY')
    $wsRH = $wb.Worksheets.Item('RUN_HISTORY')

    Write-Host "=== RUN_LOG ==="
    $rlLR = $wsRL.Cells.Item($wsRL.Rows.Count, 1).End(-4162).Row
    $rlLC = $wsRL.Cells.Item(8, $wsRL.Columns.Count).End(-4159).Column
    Write-Host "  UsedRange: rows=$rlLR, cols=$rlLC"
    Write-Host "  Headers (row 8):"
    for ($c = 1; $c -le [Math]::Min($rlLC, 20); $c++) {
        $hdr = $wsRL.Cells.Item(8, $c).Value2
        if ($hdr) { Write-Host "    Col $c = $hdr" }
    }

    Write-Host "`n=== RUN_SUMMARY ==="
    $rsLR = $wsRS.Cells.Item($wsRS.Rows.Count, 1).End(-4162).Row
    $rsLC = $wsRS.Cells.Item(1, $wsRS.Columns.Count).End(-4159).Column
    Write-Host "  UsedRange: rows=$rsLR, cols=$rsLC"
    Write-Host "  Headers (row 1):"
    for ($c = 1; $c -le [Math]::Min($rsLC, 20); $c++) {
        $hdr = $wsRS.Cells.Item(1, $c).Value2
        if ($hdr) { Write-Host "    Col $c = $hdr" }
    }

    Write-Host "`n=== RUN_HISTORY (anchor) ==="
    $rhLR = $wsRH.Cells.Item($wsRH.Rows.Count, 1).End(-4162).Row
    $rhLC = $wsRH.Cells.Item(1, $wsRH.Columns.Count).End(-4159).Column
    Write-Host "  UsedRange: rows=$rhLR, cols=$rhLC"
    Write-Host "  Headers (row 1):"
    for ($c = 1; $c -le [Math]::Min($rhLC, 27); $c++) {
        $hdr = $wsRH.Cells.Item(1, $c).Value2
        if ($hdr) { Write-Host "    Col $c = $hdr" }
    }

    $wb.Close($false)
    Write-Host "`nSTRUCTURE_READ_COMPLETE"

} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    throw
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
