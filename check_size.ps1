Add-Type -AssemblyName System.IO.Compression.FileSystem
$z = [System.IO.Compression.ZipFile]::OpenRead('g:\SAP\v182_extracted\260803_V1.8.2_BBBG_VAT_SMART_GOS_ATT_REBUILT.xlsm')
$e = $z.Entries | Where-Object { $_.FullName -eq 'xl/vbaProject.bin' }
Write-Host "V1.8.2 vbaProject.bin size: $($e.Length)"
$z.Dispose()
