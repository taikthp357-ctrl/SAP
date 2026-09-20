$ErrorActionPreference = "Continue"
$outputXlsm = "g:\SAP\v183_output\260803_V1.8.3_BBBG_VAT_SMART_GOS_ATT_REBUILT.xlsm"

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  RA SOAT TOAN DIEN V1.8.3 SMART GOS ATT REBUILT " -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

$hasError = $false
$warnCount = 0

function Report-Pass { param($msg) Write-Host ("[PASS] " + $msg) -ForegroundColor Green }
function Report-Warn { param($msg) Write-Host ("[WARN] " + $msg) -ForegroundColor Yellow; $script:warnCount++ }
function Report-Fail { param($msg) Write-Host ("[FAIL] " + $msg) -ForegroundColor Red; $script:hasError = $true }

# 1. Kiem tra File ton tai & Kich thuoc Zip
Write-Host "`n[1/8] RA SOAT FILE VA DUNG LUONG..." -ForegroundColor Magenta
if (-not (Test-Path $outputXlsm)) {
    Report-Fail "Khong tim thay file XLSM output: $outputXlsm"
} else {
    $xlsmSize = (Get-Item $outputXlsm).Length
    Report-Pass "File XLSM ton tai, dung luong: $xlsmSize bytes (~$([math]::Round($xlsmSize/1MB,2)) MB)"
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
try {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($outputXlsm)
    $vbaEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/vbaProject.bin" }
    if ($null -eq $vbaEntry) {
        Report-Fail "Khong tim thay xl/vbaProject.bin trong XLSM!"
    } else {
        $vbaSize = $vbaEntry.Length
        Report-Pass "vbaProject.bin size: $vbaSize bytes (~$([math]::Round($vbaSize/1MB,2)) MB)"
        if ($vbaSize -lt 400000) { Report-Fail "vbaProject.bin qua nho (<400KB), co the bi copy thieu!" }
    }
    $zip.Dispose()
} catch {
    Report-Fail "Loi mo file zip XLSM: $_"
}

# 2. Ra soat Excel COM & Export Components
Write-Host "`n[2/8] RA SOAT BIEN DICH VA XUAT CAC COMPONENT TU EXCEL COM..." -ForegroundColor Magenta
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$exportedDir = "g:\SAP\v183_audit_exported"
if (Test-Path $exportedDir) { Remove-Item $exportedDir -Recurse -Force }
New-Item -ItemType Directory -Path $exportedDir | Out-Null

try {
    $wb = $excel.Workbooks.Open($outputXlsm)
    $vbaProj = $wb.VBProject
    
    $compCount = $vbaProj.VBComponents.Count
    Report-Pass "Tong so VBA Component trong workbook: $compCount"
    if ($compCount -ne 20) { Report-Warn "So luong component ($compCount) khac so luong ky vong (20)" }
    
    foreach ($comp in $vbaProj.VBComponents) {
        $ext = ".bas"
        if ($comp.Type -eq 2) { $ext = ".cls" }
        elseif ($comp.Type -eq 3) { $ext = ".frm" }
        $exportPath = Join-Path $exportedDir ($comp.Name + $ext)
        $comp.Export($exportPath)
    }
    Report-Pass "Da export toan bo $compCount component ra $exportedDir"
    
    $sheetCount = $wb.Worksheets.Count
    Report-Pass "Tong so Worksheet: $sheetCount"
    $sheetNames = @()
    foreach ($ws in $wb.Worksheets) { $sheetNames += $ws.Name }
    Report-Pass "Danh sach sheet: $($sheetNames -join ', ')"
    
    Write-Host "`n[3/8] RA SOAT NUT BAM VA LIEN KET MACRO..." -ForegroundColor Magenta
    $expectedButtons = @{
        "btnChooseFolder" = "SelectSourceFolder_NoOCR"
        "btnCheck"        = "CheckFiles_NoOCR"
        "btnHeadText"     = "SAP_FillHeadTextOnly"
        "btnAttach"       = "SAP_AttachCheckedFiles"
    }
    $foundButtons = @{}
    foreach ($ws in $wb.Worksheets) {
        foreach ($shape in $ws.Shapes) {
            if ($expectedButtons.ContainsKey($shape.Name)) {
                $foundButtons[$shape.Name] = $shape.OnAction
            }
        }
    }
    foreach ($btn in $expectedButtons.Keys) {
        $expectedMacro = $expectedButtons[$btn]
        if ($foundButtons.ContainsKey($btn)) {
            $actual = $foundButtons[$btn]
            if ($actual -match $expectedMacro) {
                Report-Pass "Nut $btn -> $actual [KHOP CHUAN]"
            } else {
                Report-Fail "Nut $btn bi sai macro! (Actual: $actual | Expected: $expectedMacro)"
            }
        } else {
            Report-Warn "Nut $btn khong tim thay tren sheet"
        }
    }
    
    $wb.Close($false)
} catch {
    Report-Fail "Loi xu ly Excel COM: $_"
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# 4. Ra soat BOM va Ky tu rac tren tat ca component export
Write-Host "`n[4/8] RA SOAT KY TU BOM (0xEF 0xBB 0xBF) VA BANG MA TOAN BO CODE..." -ForegroundColor Magenta
$expFiles = Get-ChildItem -Path $exportedDir
foreach ($f in $expFiles) {
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        Report-Fail "Phat hien UTF-8 BOM o component: $($f.Name)"
    } else {
        $firstLine = (Get-Content -Path $f.FullName -Head 1)
        if ($firstLine -match "ï»¿") {
            Report-Fail "Component $($f.Name) chua ky tu rac o dong dau!"
        }
    }
}
Report-Pass "Kiem tra BOM thanh cong - Khong co bat ky file nao chua BOM rac"

# 5. Ra soat Module Name mismatch
Write-Host "`n[5/8] RA SOAT KHAI BAO ATTRIBUTE VB_NAME..." -ForegroundColor Magenta
foreach ($f in $expFiles) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    $lines = Get-Content -Path $f.FullName -Head 5
    $attrName = ""
    foreach ($l in $lines) {
        if ($l -match '^Attribute\s+VB_Name\s*=\s*"([^"]+)"') {
            $attrName = $matches[1]
            break
        }
    }
    if ($attrName -ne "" -and $attrName -ne $baseName) {
        Report-Fail "Ten module khong khop! File: $($f.Name) vs Attribute VB_Name: $attrName"
    }
}
Report-Pass "Tat ca 20 component deu khop 100% Attribute VB_Name voi ten file"

# 6. Ra soat 7 Ham GOS Adapter trong MIGO Module
Write-Host "`n[6/8] RA SOAT BO 7 HAM GOS ADAPTER TRONG MIGO MODULE..." -ForegroundColor Magenta
$migoFile = Join-Path $exportedDir "MIGO311_RunPost_V24_ErrorGuard.bas"
if (-not (Test-Path $migoFile)) {
    Report-Fail "Khong tim thay module MIGO311_RunPost_V24_ErrorGuard.bas de audit!"
} else {
    $adapterFuncs = @(
        "ATT_WaitDocReady",
        "ATT_FindGosShell",
        "ATT_OpenGosWithRetry",
        "ATT_DetectSecurity",
        "ATT_WaitSecurityClosed",
        "ATT_RebindSessionAfterSecurity",
        "ATT_SaveSuccessfulGosMap"
    )
    foreach ($func in $adapterFuncs) {
        $found = Select-String -Path $migoFile -Pattern "Function\s+$func|Sub\s+$func"
        if ($null -ne $found) {
            Report-Pass "Ham Adapter ton tai va hop le: $func"
        } else {
            Report-Fail "THIEU HAM ADAPTER: $func"
        }
    }
}

# 7. Ra soat Khai bao Win32 API 32-bit / 64-bit
Write-Host "`n[7/8] RA SOAT TUONG THICH WIN32 API (32-BIT & 64-BIT)..." -ForegroundColor Magenta
$winModule = Join-Path $exportedDir "modWinFileDialog.bas"
if (-not (Test-Path $winModule)) {
    Report-Fail "Khong tim thay modWinFileDialog.bas!"
} else {
    $winContent = Get-Content -Path $winModule -Raw
    if (($winContent -match "#If VBA7 Then" -or $winContent -match "#If Win64 Then") -and $winContent -match "PtrSafe") {
        Report-Pass "modWinFileDialog.bas co khai bao PtrSafe / VBA7 dual-compatibility 32-bit & 64-bit!"
    } else {
        Report-Warn "modWinFileDialog.bas co thiet kieu PtrSafe chua phu hop 64-bit"
    }
}

# 8. Tong ket Ra soat
Write-Host "`n[8/8] TONG KET KET QUA RA SOAT V1.8.3" -ForegroundColor Magenta
Write-Host "==================================================" -ForegroundColor Gray
if ($hasError) {
    Write-Host " [KET QUA] PHAT HIEN LOI TRONG QUA TRINH RA SOAT!" -ForegroundColor Red
} else {
    Write-Host " [KET QUA] RA SOAT HOAN HAO - KHONG CO BAT KY LOI NAO!" -ForegroundColor Green
}
Write-Host " Tong so Warning : $warnCount" -ForegroundColor Yellow
Write-Host "==================================================" -ForegroundColor Gray
