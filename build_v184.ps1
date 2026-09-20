# =============================================================================
#  build_v184.ps1  -  Build V1.8.4 MERGED (FIX29 multi-session + V1.8.3 GOS)
#  Baseline: V1.8.3 (dung V1.8.3 output lam co so, patch 4 module V1.8.4)
# =============================================================================

param(
    [string]$BaselineXlsm = "g:\SAP\v183_output\260803_V1.8.3_BBBG_VAT_SMART_GOS_ATT_REBUILT.xlsm",
    [string]$PatchDir     = "g:\SAP\v184_patch",
    [string]$OutputDir    = "g:\SAP\v184_output",
    [string]$OutputName   = "260804_V1.8.4_BBBG_VAT_MERGED_FIX29_GOS.xlsm"
)

$ErrorActionPreference = "Stop"
$VbaModules = @(
    "modSupports",
    "ModuleSAP",
    "modWinFileDialog",
    "MIGO311_RunPost_V24_ErrorGuard"
)
$RequiredMacros = @(
    "SelectSourceFolder_NoOCR",
    "CheckFiles_NoOCR",
    "SAP_FillHeadTextOnly",
    "SAP_AttachCheckedFiles"
)
$MinVbaSize = 400000   # 400 KB

function Write-Step { param($msg) Write-Host ("[STEP] " + $msg) -ForegroundColor Cyan }
function Write-Ok   { param($msg) Write-Host ("[OK  ] " + $msg) -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host ("[WARN] " + $msg) -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host ("[FAIL] " + $msg) -ForegroundColor Red; exit 1 }

# 1. Check baseline
Write-Step "Kiem tra baseline V1.8.3: $BaselineXlsm"
if (-not (Test-Path $BaselineXlsm)) { Write-Fail "Khong tim thay baseline: $BaselineXlsm" }

# Check all patch files exist
Write-Step "Kiem tra patch files..."
foreach ($modName in $VbaModules) {
    $patchFile = Join-Path $PatchDir "$modName.bas"
    if (-not (Test-Path $patchFile)) { Write-Fail "Thieu patch file: $patchFile" }
    $bytes = [System.IO.File]::ReadAllBytes($patchFile)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        Write-Fail "Patch file co BOM (se gay loi VBA import): $patchFile"
    }
    Write-Ok "  Patch OK: $modName.bas ($($bytes.Length) bytes, no BOM)"
}

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
        
        # Remove existing module
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
        } else {
            Write-Warn "Module $modName chua co trong baseline (se them moi)"
        }
        
        $vbaProj.VBComponents.Import($patchFile) | Out-Null
        Write-Ok "Da import module V1.8.4: $modName"
        Start-Sleep -Milliseconds 300
    }
    
    # Verify required macros
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

    # Verify key V1.8.4 patterns
    Write-Step "Kiem tra V1.8.4 patterns"
    $allCode = ""
    foreach ($comp in $vbaProj.VBComponents) {
        if ($comp.CodeModule.CountOfLines -gt 0) {
            $allCode += $comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines) + "`n"
        }
    }
    
    $patterns = @(
        @{ Name = "ActiveSession (multi-user)"; Pattern = "ActiveSession" },
        @{ Name = "EnumWindows (FIX29 security)"; Pattern = "EnumWindows" },
        @{ Name = "WinHandleSapSecurity"; Pattern = "WinHandleSapSecurity" },
        @{ Name = "ATT_FileExistsW (Unicode)"; Pattern = "ATT_FileExistsW" },
        @{ Name = "GetFileAttributesW (Unicode API)"; Pattern = "GetFileAttributesW" },
        @{ Name = "SAPX_WaitIdle"; Pattern = "SAPX_WaitIdle" }
    )
    
    foreach ($p in $patterns) {
        if ($allCode -match $p.Pattern) {
            Write-Ok "  V1.8.4 pattern found: $($p.Name)"
        } else {
            Write-Fail "  V1.8.4 pattern MISSING: $($p.Name)"
        }
    }
    
    $wb.Save()
    $wb.Close($false)
    Write-Ok "Da luu V1.8.4: $outputXlsm"
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
Write-Ok "vbaProject.bin V1.8.4 size: $vbaSize bytes (~$([math]::Round($vbaSize/1MB,2)) MB)"

if ($vbaSize -lt $MinVbaSize) {
    Write-Fail "HARD GATE FAIL: vbaProject.bin qua nho ($vbaSize bytes < $MinVbaSize bytes)."
}

# 4. Pack ZIP & Manifest
$zipPath = Join-Path $OutputDir ($OutputName -replace "\.xlsm$", ".zip")
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path $outputXlsm -DestinationPath $zipPath -Force
Write-Ok "Da tao ZIP V1.8.4: $zipPath"

$manifestPath = Join-Path $OutputDir "BUILD_MANIFEST_V1.8.4.txt"
$xlsmInfo = Get-Item $outputXlsm
$zipInfo = Get-Item $zipPath

$manifestContent = @"
BUILD MANIFEST - V1.8.4 MERGED (FIX29 + V1.8.3 GOS)
=====================================================
Build Date    : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Baseline      : $BaselineXlsm
Output XLSM   : $outputXlsm
Output ZIP    : $zipPath

File Sizes:
  XLSM : $($xlsmInfo.Length) bytes
  ZIP  : $($zipInfo.Length) bytes
  vbaProject.bin : $vbaSize bytes

V1.8.4 MERGED Upgrades:
  FROM FIX29:
  - modSupports: ActiveSession + multi-connection fallback
  - ModuleSAP: SAPX_GetSessionForTcode searches ALL connections, WaitIdle after Navigate
  - modWinFileDialog: EnumWindows + WinHandleSapSecurity auto-click (Allow/Permit/Yes/OK)
  - MIGO: ATT_RebindSessionAfterSecurity uses ActiveSession + fallback

  FROM V1.8.3:
  - modWinFileDialog: ATT_FileExistsW / ATT_FolderExistsW via GetFileAttributesW (Unicode)
  - modWinFileDialog: WFD_FindEditControl (ComboBoxEx32 hierarchy)
  - MIGO: GOS Adapter (ATT_WaitDocReady, ATT_FindGosShell, ATT_OpenGosWithRetry)
  - MIGO: Dynamic GOS Map storage in CONFIG sheet
  - MIGO: Standardized logging stages

  MERGED (NEW in V1.8.4):
  - WinImportFile: Unicode file check + EnumWindows + auto security + multi-method edit control
  - ATT_RebindSessionAfterSecurity: ActiveSession priority + multi-connection fallback
  - Trim backtick bug fix (Trim$ used correctly)
"@
Set-Content -Path $manifestPath -Value $manifestContent -Encoding UTF8
Write-Ok "Da tao Manifest V1.8.4: $manifestPath"

Write-Host "==================================================" -ForegroundColor Green
Write-Host "  BUILD COMPLETE: V1.8.4 MERGED (FIX29 + GOS)  " -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
