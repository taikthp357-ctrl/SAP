$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wip = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP7_WIP.xlsm'
$wb = $xl.Workbooks.Open($wip, 0, $false)
$wsInput = $wb.Worksheets.Item('INPUT')

Write-Host "Protected=$($wsInput.ProtectContents)"

# Check rows 1-8, cols E-O for existing content
for ($r = 1; $r -le 8; $r++) {
    for ($c = 5; $c -le 15; $c++) {
        $v = $wsInput.Cells.Item($r, $c).Value2
        if ("$v" -ne '') { Write-Host "  R${r}C${c}=[$v]" }
    }
}

# Check E5 properties
Write-Host "E5.MergeCells=$($wsInput.Cells.Item(5,5).MergeCells)"
Write-Host "E5.Locked=$($wsInput.Cells.Item(5,5).Locked)"

# Try direct write
$err = $null
try {
    $wsInput.Cells.Item(5, 5).Value2 = 'TESTWRITE'
} catch {
    $err = $_.Exception.Message
}
$rb = $wsInput.Cells.Item(5, 5).Value2
Write-Host "DirectWrite_err=[$err] readback=[$rb]"

$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
