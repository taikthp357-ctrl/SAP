# =============================================================================
#  build_v17.ps1  -  Build V1.7 SMART GOS ATT tu baseline V1.6.4
#  Quy trinh: Copy V1.6.4 -> Import 2 module nang cap -> Compile -> Luu
#  Khong tao lai VBA project moi. Khong thay the vbaProject.bin.
# =============================================================================

param(
    [string]$BaselineXlsm = "g:\SAP\v164_extracted\260802_V1.6.4_BBBG_VAT.xlsm",
    [string]$PatchDir     = "g:\SAP\v17_patch",
    [string]$OutputDir    = "g:\SAP\v17_output",
    [string]$OutputName   = "260803_V1.7_BBBG_VAT_SMART_GOS_ATT.xlsm"
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
$RequiredSheets = 10
$MinVbaSize = 1000000   # 1 MB
$MaxVbaSize = 1400000   # 1.4 MB

function Write-Step { param($msg) Write-Host ("[STEP] " + $msg) -ForegroundColor Cyan }
function Write-Ok   { param($msg) Write-Host ("[OK  ] " + $msg) -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host ("[WARN] " + $msg) -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host ("[FAIL] " + $msg) -ForegroundColor Red; exit 1 }

# ---- 0. Kiem tra file dau vao ----
Write-Step "Kiem tra baseline: $BaselineXlsm"
if (-not (Test-Path $BaselineXlsm)) { Write-Fail "Khong tim thay baseline XLSM: $BaselineXlsm" }

foreach ($mod in $VbaModules) {
    $p = Join-Path $PatchDir "$mod.bas"
    if (-not (Test-Path $p)) { Write-Fail "Thieu file patch: $p" }
}
Write-Ok "Tat ca file dau vao da co mat."

# ---- 1. Tao thu muc output ----
Write-Step "Tao thu muc output: $OutputDir"
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }
$outputXlsm = Join-Path $OutputDir $OutputName
Copy-Item -Path $BaselineXlsm -Destination $outputXlsm -Force
Write-Ok "Da copy baseline -> $outputXlsm"

# ---- 2. Kiem tra kich thuoc vbaProject truoc ----
Write-Step "Kiem tra kich thuoc vbaProject.bin baseline"
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($outputXlsm)
$vbaEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/vbaProject.bin" }
if ($null -eq $vbaEntry) { $zip.Dispose(); Write-Fail "Khong tim thay vbaProject.bin trong XLSM" }
$vbaSize = $vbaEntry.Length
$zip.Dispose()
Write-Ok "vbaProject.bin baseline: $vbaSize bytes (~$([math]::Round($vbaSize/1MB,2)) MB)"
if ($vbaSize -lt $MinVbaSize) { Write-Fail "vbaProject.bin qua nho ($vbaSize bytes). File baseline co the da bi loi truoc." }

# ---- 3. Mo Excel va import module ----
Write-Step "Mo Excel COM de import module"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($outputXlsm)
    
    # Kiem tra so sheet
    $sheetCount = $wb.Worksheets.Count
    Write-Ok "So sheet: $sheetCount"
    if ($sheetCount -ne $RequiredSheets) { Write-Warn "So sheet khac so ky vong ($RequiredSheets vs $sheetCount)" }
    
    # Kiem tra VBA Project
    try {
        $vbaProj = $wb.VBProject
        $compCount = $vbaProj.VBComponents.Count
        Write-Ok "So VBA component truoc khi patch: $compCount"
    } catch {
        Write-Fail "Khong the truy cap VBA Project. Hay bat 'Trust access to VBA project object model' trong Excel Trust Center."
    }
    
    # Xoa module cu, import module moi
    foreach ($modName in $VbaModules) {
        $patchFile = Join-Path $PatchDir "$modName.bas"
        
        # Xoa component cu neu ton tai
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
        
        # Import module moi
        $vbaProj.VBComponents.Import($patchFile) | Out-Null
        Write-Ok "Da import module moi: $modName"
        Start-Sleep -Milliseconds 300
    }
    
    # Kiem tra so component sau khi import
    $compCountAfter = $vbaProj.VBComponents.Count
    Write-Ok "So VBA component sau khi patch: $compCountAfter"
    if ($compCountAfter -ne $compCount) {
        Write-Warn "So component thay doi: $compCount -> $compCountAfter. Kiem tra lai module."
    }
    
    # Compile VBA de kiem tra loi cu phap
    Write-Step "Compile VBA Project"
    try {
        $vbaProj.Compile() | Out-Null
        Write-Ok "Compile thanh cong - khong co loi cu phap."
    } catch {
        Write-Warn "Compile bao loi: $_"
    }
    
    # Kiem tra cac macro bat buoc
    Write-Step "Kiem tra cac macro bat buoc"
    foreach ($macro in $RequiredMacros) {
        $found = $false
        foreach ($comp in $vbaProj.VBComponents) {
            $code = $comp.CodeModule.Lines(1, $comp.CodeModule.CountOfLines)
            if ($code -match "(?i)Sub\s+$macro\s*\(") {
                $found = $true
                Write-Ok "  Macro ton tai: $macro (trong $($comp.Name))"
                break
            }
        }
        if (-not $found) { Write-Fail "  THIEU MACRO: $macro" }
    }
    
    # Luu file
    Write-Step "Luu workbook V1.7"
    $wb.Save()
    $wb.Close($false)
    Write-Ok "Da luu: $outputXlsm"
    
} catch {
    Write-Fail "Loi khi xu ly Excel: $_"
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# ---- 4. Kiem tra kich thuoc vbaProject sau khi build ----
Write-Step "Kiem tra kich thuoc vbaProject.bin sau build"
$zip2 = [System.IO.Compression.ZipFile]::OpenRead($outputXlsm)
$vbaEntry2 = $zip2.Entries | Where-Object { $_.FullName -eq "xl/vbaProject.bin" }
$vbaSizeAfter = $vbaEntry2.Length
$zip2.Dispose()
Write-Ok "vbaProject.bin sau build: $vbaSizeAfter bytes (~$([math]::Round($vbaSizeAfter/1MB,2)) MB)"

if ($vbaSizeAfter -lt $MinVbaSize) {
    Write-Fail "HARD GATE FAIL: vbaProject.bin qua nho sau build ($vbaSizeAfter bytes < $MinVbaSize bytes). Build bi loi - KHONG BÀNTRAI."
}
if ($vbaSizeAfter -gt $MaxVbaSize) {
    Write-Warn "vbaProject.bin qua lon sau build ($vbaSizeAfter bytes > $MaxVbaSize bytes). Kiem tra lai."
}

# ---- 5. Kiem tra nut bam (button macro assignments) ----
Write-Step "Kiem tra macro duoc gan cho nut bam"
$excel2 = New-Object -ComObject Excel.Application
$excel2.Visible = $false
$excel2.DisplayAlerts = $false
try {
    $wb2 = $excel2.Workbooks.Open($outputXlsm)
    $expectedButtons = @{
        "btnChooseFolder" = "SelectSourceFolder_NoOCR"
        "btnCheck"        = "CheckFiles_NoOCR"
        "btnHeadText"     = "SAP_FillHeadTextOnly"
        "btnAttach"       = "SAP_AttachCheckedFiles"
    }
    $buttonResults = @{}
    foreach ($ws in $wb2.Worksheets) {
        foreach ($shape in $ws.Shapes) {
            if ($expectedButtons.ContainsKey($shape.Name)) {
                $buttonResults[$shape.Name] = $shape.OnAction
            }
        }
    }
    foreach ($btnName in $expectedButtons.Keys) {
        $expectedMacro = $expectedButtons[$btnName]
        if ($buttonResults.ContainsKey($btnName)) {
            $actualMacro = $buttonResults[$btnName]
            if ($actualMacro -eq $expectedMacro) {
                Write-Ok "  $btnName -> $actualMacro [OK]"
            } else {
                Write-Warn "  $btnName -> $actualMacro [KHAC NHAU! Expected: $expectedMacro]"
            }
        } else {
            Write-Warn "  $btnName - Khong tim thay tren bat ky sheet nao"
        }
    }
    $wb2.Close($false)
} catch {
    Write-Warn "Loi kiem tra button: $_"
} finally {
    $excel2.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel2) | Out-Null
}

# ---- 6. Tao file ZIP ----
Write-Step "Dong goi ZIP"
$zipPath = Join-Path $OutputDir ($OutputName -replace "\.xlsm$", ".zip")
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path $outputXlsm -DestinationPath $zipPath -Force
Write-Ok "Da tao ZIP: $zipPath"

# ---- 7. Tao BUILD_MANIFEST.txt ----
Write-Step "Tao BUILD_MANIFEST.txt"
$manifestPath = Join-Path $OutputDir "BUILD_MANIFEST.txt"
$xlsmInfo = Get-Item $outputXlsm
$zipInfo = Get-Item $zipPath
$manifestContent = @"
BUILD MANIFEST - V1.7 SMART GOS ATT
=====================================
Build Date    : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Baseline      : $BaselineXlsm
Output XLSM   : $outputXlsm
Output ZIP    : $zipPath

File Sizes:
  XLSM   : $($xlsmInfo.Length) bytes
  ZIP    : $($zipInfo.Length) bytes
  vbaProject.bin baseline : $vbaSize bytes
  vbaProject.bin V1.7     : $vbaSizeAfter bytes

VBA Modules Patched:
  - modWinFileDialog (Unicode Win32 API, ho tro tieng Viet, UNC)
  - MIGO311_RunPost_V24_ErrorGuard (Smart GOS 3-step retry, SAP Security Preflight)

Required Macros Verified:
  - SelectSourceFolder_NoOCR
  - CheckFiles_NoOCR
  - SAP_FillHeadTextOnly
  - SAP_AttachCheckedFiles

V1.7 Changes:
  [modWinFileDialog]
  - Nang cap tu ANSI (FindWindowA) len Unicode (FindWindowW / SendMessageW)
  - Ho tro ten file tieng Viet co dau (UTF-16 via StrPtr)
  - Ho tro UNC path, Network Drive, Local Drive
  - Them WinBringSecurityDialogToFront(), WinIsSapSecurityVisible()
  - Them ComboBoxEx32 fallback khi tim Edit control

  [MIGO311_RunPost_V24_ErrorGuard]
  - ATT_ResolveSourceFolder(): khong khoa duong dan, tu dong bat user chon
  - ATT_GosSecurityPreflight(): kiem tra SAP Security truoc vong lap chinh
  - ATT_GosOpenWithRetry(): 3-step retry (Technical ID / Refresh / Menu Fallback)
  - ATT_AttachOneSmartGOS(): thay the ATT_AttachOne, goi Smart GOS + Unicode dialog
  - ATT_SetSapFileDialogSmart(): nhan dien DY_PATH+DY_FILENAME / RLGRAP-FILENAME
  - SelectSourceFolder_NoOCR(): luu vao CONFIG!B4, ho tro Local/Network/UNC
  - Log stage bao gom: SAP_SECURITY_WAIT, ACTION_SENT, RESULT_READY, SAP_NOT_IDLE,
    TRY_FAIL, GOS_OPEN_FAILED_AFTER_RETRY, GOS_NO_DIALOG_POSSIBLE_SECURITY_RULE_OR_NOT_READY,
    ATT_VERIFIED, ATT_ALREADY_EXISTS, SOURCE_FOLDER_EMPTY, SOURCE_FOLDER_NOT_FOUND
"@
Set-Content -Path $manifestPath -Value $manifestContent -Encoding UTF8
Write-Ok "Da tao: $manifestPath"

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  BUILD HOAN THANH - V1.7 SMART GOS ATT   " -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host "  XLSM : $outputXlsm" -ForegroundColor White
Write-Host "  ZIP  : $zipPath" -ForegroundColor White
Write-Host "  vbaProject.bin: $vbaSizeAfter bytes" -ForegroundColor White
Write-Host ""
Write-Host "TRUOC KHI BAN GIAO, hay kiem tra thu cong:" -ForegroundColor Yellow
Write-Host "  1. Bam nut '1. CHON THU MUC' -> chon duong dan tieng Viet" -ForegroundColor Yellow
Write-Host "  2. Bam nut '2. SO TEN FILE' -> kiem tra sheet FILE_SCAN" -ForegroundColor Yellow
Write-Host "  3. Mo SAP -> bam '3. ATTACH SAP' -> kiem tra SAP Security" -ForegroundColor Yellow
