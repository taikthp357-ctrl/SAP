# =============================================================================
#  build_and_release_v17.ps1
#  Build, Verify & Package 260803_V1.7_BBBG_VAT_SMART_GOS_ATT
# =============================================================================

$ErrorActionPreference = "Stop"

$PatchDir = "g:\SAP\v184_patch"
$V17Xlsm = "g:\SAP\BBBG_VAT\260803_V1.7_BBBG_VAT_SMART_GOS_ATT.xlsm"
$V17Zip = "g:\SAP\BBBG_VAT\260803_V1.7_BBBG_VAT_SMART_GOS_ATT.zip"
$V165Xlsm = "g:\SAP\CF_OB\260803_V1.6.5_BBBG_VAT.xlsm"
$RepoDir = "g:\SAP\repo_push"

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

function Write-Step { param($msg) Write-Host ("[STEP] " + $msg) -ForegroundColor Cyan }
function Write-Ok   { param($msg) Write-Host ("[OK  ] " + $msg) -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host ("[WARN] " + $msg) -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host ("[FAIL] " + $msg) -ForegroundColor Red; exit 1 }

# 1. Kiem tra patch files
Write-Step "Kiem tra patch files trong $PatchDir"
foreach ($modName in $VbaModules) {
    $patchFile = Join-Path $PatchDir "$modName.bas"
    if (-not (Test-Path $patchFile)) { Write-Fail "Thieu patch file: $patchFile" }
    $bytes = [System.IO.File]::ReadAllBytes($patchFile)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        Write-Fail "Patch file co BOM: $patchFile"
    }
    Write-Ok "Patch file hop le: $modName.bas ($($bytes.Length) bytes, no BOM)"
}

# 2. Patch Excel Workbooks
Write-Step "Khoi dong Excel COM de patch VBA..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # --- Patch V1.7 ---
    Write-Step "Patching: $V17Xlsm"
    $wb = $excel.Workbooks.Open($V17Xlsm)
    $vbaProj = $wb.VBProject
    
    foreach ($modName in $VbaModules) {
        $patchFile = Join-Path $PatchDir "$modName.bas"
        $existingComp = $null
        foreach ($comp in $vbaProj.VBComponents) {
            if ($comp.Name -eq $modName) {
                $existingComp = $comp
                break
            }
        }
        if ($null -ne $existingComp) {
            $vbaProj.VBComponents.Remove($existingComp)
            Write-Ok "  Da go module cu: $modName"
            Start-Sleep -Milliseconds 200
        }
        $vbaProj.VBComponents.Import($patchFile) | Out-Null
        Write-Ok "  Da import module moi: $modName"
        Start-Sleep -Milliseconds 200
    }
    
    # Verify macros
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
    Write-Ok "Da luu thanh cong V1.7: $V17Xlsm"

    # --- Patch V1.6.5 ---
    if (Test-Path $V165Xlsm) {
        Write-Step "Patching: $V165Xlsm"
        $wb2 = $excel.Workbooks.Open($V165Xlsm)
        $vbaProj2 = $wb2.VBProject
        
        foreach ($modName in $VbaModules) {
            $patchFile = Join-Path $PatchDir "$modName.bas"
            $existingComp = $null
            foreach ($comp in $vbaProj2.VBComponents) {
                if ($comp.Name -eq $modName) {
                    $existingComp = $comp
                    break
                }
            }
            if ($null -ne $existingComp) {
                $vbaProj2.VBComponents.Remove($existingComp)
                Write-Ok "  Da go module cu: $modName"
                Start-Sleep -Milliseconds 200
            }
            $vbaProj2.VBComponents.Import($patchFile) | Out-Null
            Write-Ok "  Da import module moi: $modName"
            Start-Sleep -Milliseconds 200
        }
        $wb2.Save()
        $wb2.Close($false)
        Write-Ok "Da luu thanh cong V1.6.5: $V165Xlsm"
    }
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# 3. Kiem tra vbaProject.bin
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($V17Xlsm)
$vbaEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/vbaProject.bin" }
$vbaSize = $vbaEntry.Length
$zip.Dispose()
Write-Ok "vbaProject.bin size: $vbaSize bytes (~$([math]::Round($vbaSize/1MB,2)) MB)"
if ($vbaSize -lt 300000) { Write-Fail "HARD GATE FAIL: vbaProject.bin qua nho ($vbaSize bytes)" }

# 4. Pack ZIP
Write-Step "Dong goi ZIP: $V17Zip"
if (Test-Path $V17Zip) { Remove-Item $V17Zip -Force }
Compress-Archive -Path $V17Xlsm -DestinationPath $V17Zip -Force
Write-Ok "Da tao ZIP: $V17Zip"

# 5. Cap nhat Changelog & Manifest
$changelogPath = Join-Path $RepoDir "CHANGELOG_V1.7.txt"
$manifestPath = Join-Path $RepoDir "BUILD_MANIFEST.txt"
$xlsmInfo = Get-Item $V17Xlsm
$zipInfo = Get-Item $V17Zip

$changelogContent = @"
CHANGELOG - V1.7 SMART GOS ATT (HARDENED RELEASE)
==================================================
Phien ban: V1.7 (Hardened with V1.8.4 Core Engine)
Ngay build: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Baseline  : 260803_V1.7_BBBG_VAT_SMART_GOS_ATT.xlsm / 260803_V1.6.5_BBBG_VAT.xlsm

--- CAC LOI DA DUOC XU LY TRIET DE ---

1. [modSupports.bas] - KHAC PHUC LOI TIM SAP SESSION (Error 7021/7804)
   - Uu tien 'sapApp.ActiveSession' truoc tien (ho tro moi truong da phien / multi-session).
   - Fallback duyet qua tat ca cac connection va session trong 'sapApp.Children'.
   - Loai bo hoan toan loi treo/khoa cung session khi user mo nhieu cua so SAP.

2. [ModuleSAP.bas] - NANG CAP DIEU HUONG SAP & MIGO TCODE
   - 'SAPX_GetSessionForTcode': Tu dong duyet tat ca connection tren toan bo cay SAP GUI ('seed.parent.parent.Children').
   - Them 'SAPX_WaitIdle' ngay sau khi dieu huong dam bao MIGO da load 100% truoc khi thao tac tiep theo.

3. [modWinFileDialog.bas] - WIN32 UNICODE & ENUMWINDOWS AUTO SECURITY HANDLER
   - Loai bo hoan toan ANSI API 32-bit (FindWindowA, SetWindowTextA) gay loi tren Windows 10/11 64-bit.
   - Dung Win32 Unicode API (FindWindowW, SendMessageW, GetFileAttributesW).
   - Tich hop 'EnumWindows' va 'WinHandleSapSecurity' tu dong phat hien va nhan nut Allow/Permit/Yes/OK khi SAP bat hop thoai bao mat (SAP GUI Security).
   - Ham 'ATT_FileExistsW' va 'ATT_FolderExistsW' dung 'GetFileAttributesW' khong bao gio bi loi ky tu Tieng Viet co dau hay duong dan mang UNC nhu ham Dir$ cu.
   - Duyet cay control ComboBoxEx32 tim chinh xac Edit control trong hop thoai Windows File Open Dialog.

4. [MIGO311_RunPost_V24_ErrorGuard.bas] - GOS SMART ADAPTER LAYER & 3-STEP RETRY
   - 'ATT_FindGosShell': Tu dong do tim GOS Shell control tren moi theme SAP GUI (Belize, Quartz, Blue Crystal) va luu cache tai CONFIG!B16.
   - 'ATT_OpenGosWithRetry': Co che thu lai 3 buoc thong minh:
       + Buoc 1: Technical ID (%GOS_TOOLBOX / %GOS_PCATTA_CREA)
       + Buoc 2: Dua SAP len foreground + refresh lai control shell
       + Buoc 3: Menu fallback (ATT_FindGosButton / Context menu)
   - 'ATT_WaitDocReady': Cho MIGO load xong truoc khi bat dau attach file.
   - 'ATT_RebindSessionAfterSecurity': Tu dong khoi phuc session neu hop thoai bao mat SAP ngat ket noi tam thoi.

--- HUONG DAN SU DUNG 4 NUT MACRO CHINH ---
   1. [1. CHON THU MUC]         -> SelectSourceFolder_NoOCR
   2. [2. SO TEN FILE & HEADTEXT] -> CheckFiles_NoOCR
   3. [3. DIEN HEADTEXT (SAP)]   -> SAP_FillHeadTextOnly
   4. [3. ATTACH SAP]            -> SAP_AttachCheckedFiles
"@
Set-Content -Path $changelogPath -Value $changelogContent -Encoding UTF8
Write-Ok "Da cap nhat Changelog: $changelogPath"

$manifestContent = @"
BUILD MANIFEST - V1.7 SMART GOS ATT (HARDENED RELEASE)
======================================================
Build Date    : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Output XLSM   : $V17Xlsm
Output ZIP    : $V17Zip
XLSM Size     : $($xlsmInfo.Length) bytes
ZIP Size      : $($zipInfo.Length) bytes
vbaProject.bin: $vbaSize bytes

Modules Integrated:
  - modSupports.bas                   : ActiveSession Priority + Multi-Connection Fallback
  - ModuleSAP.bas                      : All-Connection Search + WaitIdle
  - modWinFileDialog.bas               : Win32 Unicode + EnumWindows Auto SAP Security Click
  - MIGO311_RunPost_V24_ErrorGuard.bas : Dynamic GOS Adapter + 3-Tier Retry + Unicode Checks
"@
Set-Content -Path $manifestPath -Value $manifestContent -Encoding UTF8
Write-Ok "Da cap nhat Manifest: $manifestPath"

# 6. Copy Deliverables sang repo_push
Write-Step "Sao chep thanh pham sang $RepoDir..."
Copy-Item -Path $V17Xlsm -Destination (Join-Path $RepoDir "260803_V1.7_BBBG_VAT_SMART_GOS_ATT.xlsm") -Force
Copy-Item -Path $V17Zip -Destination (Join-Path $RepoDir "260803_V1.7_BBBG_VAT_SMART_GOS_ATT.zip") -Force
Write-Ok "Da sao chep XLSM va ZIP vao repo_push"

Write-Host "==================================================" -ForegroundColor Green
Write-Host "  BUILD & PACKAGING V1.7 HOAN TAT XUAT SAC!       " -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
