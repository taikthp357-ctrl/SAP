# Fix UTF-8 BOM on all .bas files in patch directories
$patchDirs = @("g:\SAP\v183_patch", "g:\SAP\v17_patch")

foreach ($dir in $patchDirs) {
    if (Test-Path $dir) {
        $files = Get-ChildItem -Path $dir -Filter "*.bas"
        foreach ($f in $files) {
            $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
            if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
                Write-Host "Found UTF-8 BOM in $($f.Name)! Removing BOM..." -ForegroundColor Yellow
                $cleanBytes = $bytes[3..($bytes.Length - 1)]
                [System.IO.File]::WriteAllBytes($f.FullName, $cleanBytes)
                Write-Host "Removed BOM from $($f.Name)" -ForegroundColor Green
            } else {
                Write-Host "$($f.Name) has no BOM. [OK]" -ForegroundColor Green
            }
        }
    }
}
