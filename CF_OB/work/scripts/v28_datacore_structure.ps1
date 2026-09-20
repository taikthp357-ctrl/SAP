# v28_datacore_structure.ps1 — Check DATA_CORE structure for REGION band planning
$ErrorActionPreference = 'Stop'
$cp9 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP9_HISTORY_CLEARED.xlsm'

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1

try {
    $wb = $xl.Workbooks.Open($cp9, 0, $true)  # ReadOnly

    $wsDC = $wb.Worksheets.Item('DATA_CORE')
    $wsPosting = $wb.Worksheets.Item('POSTING')

    Write-Host "=== DATA_CORE (anchor for REGION band) ==="
    $dcLR = $wsDC.Cells.Item($wsDC.Rows.Count, 1).End(-4162).Row
    $dcLC = $wsDC.Cells.Item(4, $wsDC.Columns.Count).End(-4159).Column
    Write-Host "  UsedRange: rows=$dcLR, cols=$dcLC"
    Write-Host "  Headers (row 4, showing cols 1-32):"
    for ($c = 1; $c -le 32; $c++) {
        $hdr = $wsDC.Cells.Item(4, $c).Value2
        Write-Host "    Col $c = $hdr"
    }

    Write-Host "`n=== POSTING (to be merged as REGION) ==="
    $pLR = $wsPosting.Cells.Item($wsPosting.Rows.Count, 1).End(-4162).Row
    $pLC = $wsPosting.Cells.Item(1, $wsPosting.Columns.Count).End(-4159).Column
    Write-Host "  UsedRange: rows=$pLR, cols=$pLC"
    Write-Host "  Headers (row 1):"
    for ($c = 1; $c -le $pLC; $c++) {
        $hdr = $wsPosting.Cells.Item(1, $c).Value2
        if ($hdr) { Write-Host "    Col $c = $hdr" }
    }

    Write-Host "`n=== REGION band layout decision ==="
    Write-Host "  DATA_CORE ends at col $dcLC (column $([char](64 + $dcLC)))"
    Write-Host "  POSTING has $pLC cols"
    $regionStart = $dcLC + 1
    $regionEnd = $regionStart + $pLC - 1
    Write-Host "  REGION band: cols $regionStart-$regionEnd (offset +$dcLC)"
    Write-Host "  REGION range: $([char](64 + $regionStart))1:$([char](64 + $regionEnd))N"

    $wb.Close($false)

} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    throw
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
