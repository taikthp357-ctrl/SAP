$srcDir = "C:\Users\taikt\.gemini\antigravity\brain\09dccbf3-c1cd-456a-bbe5-1e523440b2c4\scratch\SAP\SAP_ZONE"
$distExe = Join-Path $srcDir "dist\SAP_ZONE.exe"
$rootExe = Join-Path $srcDir "SAP_ZONE.exe"

# Cập nhật cả thư mục v1.6 và v1.7 để người dùng dùng đường dẫn nào cũng có bản mới nhất
$outDirs = @(
    "G:\AI-Manager\outputs\sap-zone-v1.6",
    "G:\AI-Manager\outputs\sap-zone-v1.7"
)

foreach ($outDir in $outDirs) {
    if (-not (Test-Path $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }
}

if (-not (Test-Path (Join-Path $srcDir "releases"))) {
    New-Item -ItemType Directory -Path (Join-Path $srcDir "releases") -Force | Out-Null
}

Copy-Item $distExe $rootExe -Force

foreach ($outDir in $outDirs) {
    Copy-Item $distExe (Join-Path $outDir "SAP_ZONE.exe") -Force
}

$tempZip = Join-Path $srcDir "temp_pack"
if (Test-Path $tempZip) { Remove-Item $tempZip -Recurse -Force }
New-Item -ItemType Directory -Path $tempZip | Out-Null

Copy-Item $distExe (Join-Path $tempZip "SAP_ZONE.exe")
Copy-Item (Join-Path $srcDir "zone_map.json") (Join-Path $tempZip "zone_map.json")
Copy-Item (Join-Path $srcDir "HUONG_DAN_SU_DUNG.md") (Join-Path $tempZip "HUONG_DAN_SU_DUNG.md")
Copy-Item (Join-Path $srcDir "RELEASE_NOTES.md") (Join-Path $tempZip "RELEASE_NOTES.md")

# Tạo zip v1.7
$zipV17 = Join-Path $srcDir "releases\SAP_ZONE_v1.7_PORTABLE.zip"
if (Test-Path $zipV17) { Remove-Item $zipV17 -Force }
Compress-Archive -Path (Join-Path $tempZip "*") -DestinationPath $zipV17 -Force
Copy-Item $zipV17 (Join-Path "G:\AI-Manager\outputs\sap-zone-v1.7" "SAP_ZONE_v1.7_PORTABLE.zip") -Force

# Cập nhật cả zip v1.6 cho tương thích ngược
$zipV16 = Join-Path $srcDir "releases\SAP_ZONE_v1.6_PORTABLE.zip"
if (Test-Path $zipV16) { Remove-Item $zipV16 -Force }
Compress-Archive -Path (Join-Path $tempZip "*") -DestinationPath $zipV16 -Force
Copy-Item $zipV16 (Join-Path "G:\AI-Manager\outputs\sap-zone-v1.6" "SAP_ZONE_v1.6_PORTABLE.zip") -Force

Remove-Item $tempZip -Recurse -Force

Write-Host "PACKAGE_SUCCESS: v1.7 and v1.6 portable zip generated successfully!"
