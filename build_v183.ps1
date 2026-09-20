# =============================================================================
#  build_v183.ps1  -  Build V1.8.3 SMART GOS ATT REBUILT tu baseline V1.8.2
# =============================================================================

param(
    [string]$BaselineXlsm = "g:\SAP\v182_extracted\260803_V1.8.2_BBBG_VAT_SMART_GOS_ATT_REBUILT.xlsm",
    [string]$PatchDir     = "g:\SAP\v183_patch",
    [string]$OutputDir    = "g:\SAP\v183_output",
    [string]$OutputName   = "260803_V1.8.3_BBBG_VAT_SMART_GOS_ATT_REBUILT.xlsm"
)

$ErrorActionPreference = "Stop"
$VbaModules = @(
    "modWinFileDialog",
    "MIGO311_RunPost_V24_ErrorGuard"
)
$RequiredMacros = @(
    "SelectSourceFolder_NoOCR",
    "CheckFiles_NoOCR",
    "SAP_FillHeadTextOnly",
    "SAP_AttachCheckedFiles"
)
$MinVbaSize = 400000   # 400 KB (Baseline V1.8.2 clean size is ~517 KB)

function Write-Step { param($msg) Write-Host ("[STEP] " + $msg) -ForegroundColor Cyan }
function Write-Ok   { param($msg) Write-Host ("[OK  ] " + $msg) -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host ("[WARN] " + $msg) -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host ("[FAIL] " + $msg) -ForegroundColor Red; exit 1 }

# 1. Check baseline
Write-Step "Kiem tra baseline V1.8.2: $BaselineXlsm"
if (-not (Test-Path $BaselineXlsm)) { Write-Fail "Khong tim thay baseline: $BaselineXlsm" }

if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }
$outputXlsm = Join-Path $OutputDir $OutputName
Copy-Item -Path $BaselineXlsm -Destination $outputXlsm -Force
Write-Ok "Da copy baseline -> $outputXlsm"

# 2. Mo Excel COM & patch VBA
Write-Step "Mo Excel COM de patch VBA modules..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($outputXlsm)
    $vbaProj = $wb.VBProject
    
    foreach ($modName in $VbaModules) {
        $patchFile = Join-Path $PatchDir "$modName.bas"
        if (-not (Test-Path $patchFile)) { Write-Fail "Thieu patch file: $patchFile" }
        
        $existingComp = $null
        foreach ($comp in $vbaProj.VBComponents) {
            if ($comp.Name -eq $modName) {
                $existingComp = $comp
                break
            }
        }
        if ($null -ne $existingComp) {
            $vbaProj.VBComponents.Remove($existingComp)
            Write-Ok "Da xoa module cu: $modName"
            Start-Sleep -Milliseconds 200
        }
        
        $vbaProj.VBComponents.Import($patchFile) | Out-Null
        Write-Ok "Da import module V1.8.3 moi: $modName"
        Start-Sleep -Milliseconds 300
    }
    
    Write-Step "Kiem tra Macro bat buoc"
    foreach ($macro in $RequiredMacros) {
        $found = $false
        foreach ($comp in $vbaProj.VBComponents) {
            if ($comp.CodeModule.CountOfLines -gt 0) {
                $code = $comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines)
                if ($code -match "(?i)Sub\s+$macro\s*\(") {
                    $found = $true
                    Write-Ok "  Macro ton tai: $macro (trong $($comp.Name))"
                    break
                }
            }
        }
        if (-not $found) { Write-Fail "Thieu Macro: $macro" }
    }
    
    $wb.Save()
    $wb.Close($false)
    Write-Ok "Da luu V1.8.3: $outputXlsm"
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# 3. Kiem tra kich thuoc vbaProject.bin
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($outputXlsm)
$vbaEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/vbaProject.bin" }
$vbaSize = $vbaEntry.Length
$zip.Dispose()
Write-Ok "vbaProject.bin V1.8.3 size: $vbaSize bytes (~$([math]::Round($vbaSize/1MB,2)) MB)"

if ($vbaSize -lt $MinVbaSize) {
    Write-Fail "HARD GATE FAIL: vbaProject.bin qua nho ($vbaSize bytes < $MinVbaSize bytes)."
}

# 4. Pack ZIP & Manifest
$zipPath = Join-Path $OutputDir ($OutputName -replace "\.xlsm$", ".zip")
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path $outputXlsm -DestinationPath $zipPath -Force
Write-Ok "Da tao ZIP V1.8.3: $zipPath"

$manifestPath = Join-Path $OutputDir "BUILD_MANIFEST_V1.8.3.txt"
$xlsmInfo = Get-Item $outputXlsm
$zipInfo = Get-Item $zipPath

$manifestContent = @"
BUILD MANIFEST - V1.8.3 SMART GOS ATT REBUILT
==============================================
Build Date    : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Baseline      : $BaselineXlsm
Output XLSM   : $outputXlsm
Output ZIP    : $zipPath

File Sizes:
  XLSM : $($xlsmInfo.Length) bytes
  ZIP  : $($zipInfo.Length) bytes
  vbaProject.bin : $vbaSize bytes

V1.8.3 Upgrades Added:
  - GOS Adapter modularization (ATT_WaitDocReady, ATT_FindGosShell, ATT_OpenGosWithRetry,
    ATT_DetectSecurity, ATT_WaitSecurityClosed, ATT_RebindSessionAfterSecurity, ATT_SaveSuccessfulGosMap).
  - Automatic Session Re-binding after SAP GUI Security Popup closed.
  - Auto-Discovery via GuiShell control tree scan when Technical ID fails.
  - Dynamic GOS Map storage in CONFIG sheet (System_Client_User_MIGO).
  - Standardized logging stages (GOS_SEARCHING, SAP_SECURITY_WAIT_USER, ATT_VERIFIED, etc.).
"@
Set-Content -Path $manifestPath -Value $manifestContent -Encoding UTF8
Write-Ok "Da tao Manifest V1.8.3: $manifestPath"

Write-Host "==================================================" -ForegroundColor Green
Write-Host "  BUILD COMPLETE: V1.8.3 SMART GOS ATT REBUILT  " -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
