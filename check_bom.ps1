$bytes = [System.IO.File]::ReadAllBytes('g:\SAP\v183_patch\MIGO311_RunPost_V24_ErrorGuard.bas')
Write-Host "Byte 0: $($bytes[0]) (0x$('{0:X2}' -f $bytes[0]))"
Write-Host "Byte 1: $($bytes[1]) (0x$('{0:X2}' -f $bytes[1]))"
Write-Host "Byte 2: $($bytes[2]) (0x$('{0:X2}' -f $bytes[2]))"
