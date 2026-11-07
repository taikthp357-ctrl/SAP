$srcDir = "C:\Users\taikt\.gemini\antigravity\brain\09dccbf3-c1cd-456a-bbe5-1e523440b2c4\scratch\SAP\SAP_ZONE"
$distExe = Join-Path $srcDir "dist\SAP_ZONE.exe"
$rootExe = Join-Path $srcDir "SAP_ZONE.exe"
$outDir = "G:\AI-Manager\outputs\sap-zone-v1.6"
$zipRel = Join-Path $srcDir "releases\SAP_ZONE_v1.6_PORTABLE.zip"
$zipOut = Join-Path $outDir "SAP_ZONE_v1.6_PORTABLE.zip"

if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}
if (-not (Test-Path (Join-Path $srcDir "releases"))) {
    New-Item -ItemType Directory -Path (Join-Path $srcDir "releases") -Force | Out-Null
}

Copy-Item $distExe $rootExe -Force
Copy-Item $distExe (Join-Path $outDir "SAP_ZONE.exe") -Force

$tempZip = Join-Path $srcDir "temp_pack"
if (Test-Path $tempZip) { Remove-Item $tempZip -Recurse -Force }
New-Item -ItemType Directory -Path $tempZip | Out-Null

Copy-Item $distExe (Join-Path $tempZip "SAP_ZONE.exe")
Copy-Item (Join-Path $srcDir "zone_map.json") (Join-Path $tempZip "zone_map.json")
Copy-Item (Join-Path $srcDir "HUONG_DAN_SU_DUNG.md") (Join-Path $tempZip "HUONG_DAN_SU_DUNG.md")
Copy-Item (Join-Path $srcDir "RELEASE_NOTES.md") (Join-Path $tempZip "RELEASE_NOTES.md")

if (Test-Path $zipRel) { Remove-Item $zipRel -Force }
Compress-Archive -Path (Join-Path $tempZip "*") -DestinationPath $zipRel -Force
Copy-Item $zipRel $zipOut -Force
Remove-Item $tempZip -Recurse -Force

Write-Host "PACKAGE_SUCCESS"
