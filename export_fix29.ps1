$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wbPath = "g:\SAP\BBBG_VAT\260804_V1.7.1_BBBG_VAT_PORTABLE_GOS_MULTI_USER_FIX29.xlsm"
    $outDir = "g:\SAP\fix29_exported"
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
    
    $wb = $excel.Workbooks.Open($wbPath)
    foreach ($comp in $wb.VBProject.VBComponents) {
        $ext = ".bas"
        if ($comp.Type -eq 2) { $ext = ".cls" }
        elseif ($comp.Type -eq 3) { $ext = ".frm" }
        $file = Join-Path $outDir ($comp.Name + $ext)
        $comp.Export($file)
    }
    Write-Host "Exported FIX29 VBA components successfully!" -ForegroundColor Green
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
