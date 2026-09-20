$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open('g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP7_WIP.xlsm', 0, $false)

Write-Host "Attempting to import modCF_InputWorklist..."
try {
    $comp = $wb.VBProject.VBComponents.Import('g:\SAP\CF_OB\work\scratch\vba_v28\modCF_InputWorklist.bas')
    Write-Host "IMPORT_OK: Name=$($comp.Name)"
} catch {
    Write-Host "IMPORT_FAILED: $($_.Exception.Message)"
}

Write-Host "`nModules in workbook:"
$wb.VBProject.VBComponents | ForEach-Object { Write-Host "  $($_.Name)" } | Select-String -Pattern "Input|Reset|Check|Vietnamese"

$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
