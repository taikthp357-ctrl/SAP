$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open('g:\SAP\v182_extracted\260803_V1.8.2_BBBG_VAT_SMART_GOS_ATT_REBUILT.xlsm')
    $outDir = 'g:\SAP\v182_extracted\vba_exported'
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
    foreach ($comp in $wb.VBProject.VBComponents) {
        $file = Join-Path $outDir ($comp.Name + '.bas')
        $comp.Export($file)
    }
    Write-Host 'Exported components successfully'
    $wb.Close($false)
} finally {
    $excel.Quit()
}
