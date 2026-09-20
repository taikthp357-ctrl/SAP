# Export VBA modules from original V1.7
$ErrorActionPreference = 'Stop'
$src = "G:\SAP\BBBG_VAT\260803_V1.7_BBBG_VAT_SMART_GOS_ATT.xlsm"
$outDir = "g:\SAP\v17_original_exported"

if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false

try {
    $wb = $excel.Workbooks.Open($src, $false, $true)
    $vbp = $wb.VBProject
    $count = 0
    foreach ($comp in $vbp.VBComponents) {
        $name = $comp.Name
        $type = $comp.Type  # 1=Module, 2=ClassModule, 3=Form, 100=Document
        switch ($type) {
            1 { $ext = ".bas" }
            2 { $ext = ".cls" }
            3 { $ext = ".frm" }
            100 { $ext = ".bas" }
            default { $ext = ".bas" }
        }
        $outFile = Join-Path $outDir ($name + $ext)
        $comp.Export($outFile)
        $count++
        Write-Host "[OK] Exported: $name ($type) -> $outFile"
    }
    Write-Host "`nTotal exported: $count components"

    # Sheet info
    Write-Host "`nWorksheets:"
    foreach ($ws in $wb.Worksheets) {
        Write-Host "  - $($ws.Name)"
    }

    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
