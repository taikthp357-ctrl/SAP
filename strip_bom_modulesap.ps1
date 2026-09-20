$f = "g:\SAP\v184_patch\ModuleSAP.bas"
$bytes = [System.IO.File]::ReadAllBytes($f)
if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    $bytes = $bytes[3..($bytes.Length-1)]
    [System.IO.File]::WriteAllBytes($f, $bytes)
    Write-Host "BOM removed from ModuleSAP"
} else {
    Write-Host "No BOM in ModuleSAP [OK]"
}
