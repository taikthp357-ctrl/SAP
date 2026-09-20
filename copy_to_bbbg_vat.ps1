$srcXlsm = "g:\SAP\v183_output\260803_V1.8.3_BBBG_VAT_SMART_GOS_ATT_REBUILT.xlsm"
$srcZip  = "g:\SAP\v183_output\260803_V1.8.3_BBBG_VAT_SMART_GOS_ATT_REBUILT.zip"
$destDir = "g:\SAP\BBBG_VAT"

Copy-Item -Path $srcXlsm -Destination (Join-Path $destDir "260803_V1.8.3_BBBG_VAT_SMART_GOS_ATT_REBUILT.xlsm") -Force
Copy-Item -Path $srcZip  -Destination (Join-Path $destDir "260803_V1.8.3_BBBG_VAT_SMART_GOS_ATT_REBUILT.zip") -Force

Write-Host "Successfully copied V1.8.3 files to G:\SAP\BBBG_VAT" -ForegroundColor Green
Get-ChildItem -Path $destDir
