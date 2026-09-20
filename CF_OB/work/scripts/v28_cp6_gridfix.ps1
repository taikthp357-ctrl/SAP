# CP6: Grid lazy-load fix + Vietnamese messages
$ErrorActionPreference = 'Stop'
$cp5a = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP5a_CHECKBAND.xlsm'
$wip = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP6_WIP.xlsm'
$vbaBase = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$logPath = 'g:\SAP\CF_OB\work\REPORTS\v28_cp6_gridfix.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Host $m }

if (-not (Test-Path $cp5a)) { throw "NO_CP5A=$cp5a" }
Copy-Item $cp5a $wip -Force
$wipHash = (Get-FileHash $wip -Algorithm SHA256).Hash
Log "WIP_CREATED=$wipHash"

# Ground truth VBA paths
$readOnlyPath = Join-Path $vbaBase 'modCodexSAPReadOnly.bas'
$vnNewPath = Join-Path $vbaBase 'modCF_VietnameseMessages.bas'

# 1. Modify modCodexSAPReadOnly.bas - add helper + insert calls at 11 sites
$roContent = [System.IO.File]::ReadAllText($readOnlyPath, [Text.Encoding]::GetEncoding(1252))

# Helper function to insert after Option Explicit
$helper = @'

' Grid lazy-load fix: force SAP GUI to render all rows before reading.
' Issue #6/#2: weak machine + many rows → unrendered cells return blank → bỏ sót mã.
Private Sub Codex_EnsureGridRendered(ByVal grid As Object, ByVal ses As Object)
    Dim rowCount As Long
    On Error Resume Next
    rowCount = CLng(grid.rowCount)
    On Error GoTo 0
    If rowCount > 20 Then
        grid.SetCurrentCell rowCount - 1, 0
        CWR_WaitIdle ses, 0.5
        grid.SetCurrentCell 0, 0
    End If
End Sub
'@

# Insert helper after Option Explicit (with proper line ending preservation)
if ($roContent -match '(?m)^(Option Explicit)\r?\n') {
    $roContent = $roContent -replace '(?m)^(Option Explicit)\r?\n', "`$1`r`n$helper`r`n"
    Log "INSERTED_HELPER_AFTER_OPTION_EXPLICIT"
} else {
    throw "CANNOT_FIND_OPTION_EXPLICIT"
}

# 2. Insert Codex_EnsureGridRendered calls before affected loops
# Find all "For r = 0 To CLng(grid.rowCount) - 1" loops and insert before them
$loopsFixed = 0
$lines = $roContent -split "`r`n"
$newLines = [System.Collections.ArrayList]::new()
for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    # Check if this line is a grid loop start
    if ($line -match '^\s+For r = 0 To CLng\(grid\.rowCount\) - 1\s*$') {
        # Check if previous line already has Codex_EnsureGridRendered
        if ($i -eq 0 -or $newLines[-1] -notmatch 'Codex_EnsureGridRendered') {
            # Insert the call with same indentation as the For loop
            $indent = $line -replace '^(\s+).*', '$1'
            [void]$newLines.Add("${indent}Codex_EnsureGridRendered grid, ses")
            $loopsFixed++
        }
    }
    [void]$newLines.Add($line)
}
$roContent = $newLines -join "`r`n"
Log "INSERTED_CODEX_ENSUREGRIDRENDERED_AT_${loopsFixed}_SITES"

[System.IO.File]::WriteAllText($readOnlyPath, $roContent, [Text.Encoding]::GetEncoding(1252))
Log "MODIFIED_modCodexSAPReadOnly.bas"

# 3. Create modCF_VietnameseMessages.bas
$vnContent = @'
Attribute VB_Name = "modCF_VietnameseMessages"
Option Explicit

' Vietnamese message mapper for user-facing feedback.
' Maps Result enum (English technical) + Detail context → Vietnamese string.
' Used by Phase 1b worklist to display CHECK feedback in Vietnamese.

Public Function CF_VN_GetMessage(ByVal result As String, ByVal detail As String) As String
    Dim msg As String
    result = UCase$(Trim$(result))
    detail = Trim$(detail)

    Select Case result
        Case "PASS"
            msg = "Sẵn sàng xử lý"

        Case "BLOCKED_SOURCE_PLAN_INCOMPLETE"
            If InStr(1, detail, "Multi-batch", vbTextCompare) > 0 Then
                msg = "Vị trí có nhiều lô, cần chọn lô cụ thể"
            ElseIf InStr(1, detail, "PICK_BIN_CANDIDATE_COUNT", vbTextCompare) > 0 Then
                msg = "Có nhiều vị trí nguồn, cần chọn vị trí cụ thể"
            Else
                msg = "Dữ liệu nguồn chưa đầy đủ"
            End If

        Case "BLOCKED_QTY_GROUP_BASELINE"
            msg = "Nhóm số lượng chưa xác định được baseline"

        Case "BLOCKED_INCOMPLETE_EXECUTION"
            msg = "Hành động trước chưa hoàn tất"

        Case "BLOCKED_CREATE_OPEN"
            msg = "Chứng từ PRDO đang mở, cần đóng trước khi tiếp tục"

        Case "STOP"
            If InStr(1, detail, "SOURCE_PLAN_INCOMPLETE", vbTextCompare) > 0 Then
                msg = "Dừng: dữ liệu nguồn chưa đủ"
            ElseIf InStr(1, detail, "REGRESSION_GATE", vbTextCompare) > 0 Then
                msg = "Dừng: không thỏa điều kiện an toàn regression"
            Else
                msg = "Dừng: " & detail
            End If

        Case "READY_FOR_CF"
            msg = "Sẵn sàng để user xác nhận WT"

        Case "DONE_BY_USER_CF"
            msg = "Đã hoàn tất bởi user CF"

        Case "DONE"
            msg = "Hoàn tất"

        Case "PLAN_NEEDS_TAI_POSTING"
            msg = "Cần TAI POSTING để tiếp tục"

        Case Else
            msg = result
            If detail <> "" Then msg = msg & " (" & Left$(detail, 100) & ")"
    End Select

    CF_VN_GetMessage = msg
End Function

Public Function CF_VN_GetStageLabel(ByVal stage As String) As String
    stage = UCase$(Trim$(stage))
    Select Case stage
        Case "CHECK PREPARE": CF_VN_GetStageLabel = "Chuẩn bị kiểm tra"
        Case "CHECK DONE": CF_VN_GetStageLabel = "Kiểm tra hoàn tất"
        Case "EXECUTION_CONTEXT_REUSE": CF_VN_GetStageLabel = "Tái sử dụng context"
        Case "CANCEL_QUEUE": CF_VN_GetStageLabel = "Hàng đợi cancel"
        Case "CANCEL_GROUP": CF_VN_GetStageLabel = "Cancel nhóm"
        Case "CANCEL_BATCH_READBACK": CF_VN_GetStageLabel = "Đọc lại sau cancel"
        Case "LOCATION_GROUP": CF_VN_GetStageLabel = "Thay đổi vị trí"
        Case "CREATE_GROUP": CF_VN_GetStageLabel = "Tạo WT mới"
        Case "FINAL_VERIFY": CF_VN_GetStageLabel = "Xác minh cuối"
        Case "WAIT_USER_CF": CF_VN_GetStageLabel = "Chờ user CF"
        Case Else: CF_VN_GetStageLabel = stage
    End Select
End Function
'@
[System.IO.File]::WriteAllText($vnNewPath, $vnContent, [Text.Encoding]::GetEncoding(1252))
Log "CREATED_modCF_VietnameseMessages.bas"

# 4. Import into Excel COM
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false
$wb = $null
try {
    $wb = $xl.Workbooks.Open($wip, 0, $false)
    Log "OPENED_WIP"

    # Remove old modCodexSAPReadOnly
    $found = $false
    foreach ($comp in $wb.VBProject.VBComponents) {
        if ($comp.Name -eq 'modCodexSAPReadOnly') {
            $wb.VBProject.VBComponents.Remove($comp)
            Log "REMOVED_OLD_modCodexSAPReadOnly"
            $found = $true
            break
        }
    }
    if (-not $found) { throw "OLD_MODULE_NOT_FOUND=modCodexSAPReadOnly" }

    # Check if modCF_VietnameseMessages already exists (shouldn't on CP5a baseline)
    foreach ($comp in $wb.VBProject.VBComponents) {
        if ($comp.Name -eq 'modCF_VietnameseMessages') {
            $wb.VBProject.VBComponents.Remove($comp)
            Log "REMOVED_OLD_modCF_VietnameseMessages"
            break
        }
    }

    # Import modified modCodexSAPReadOnly
    $wb.VBProject.VBComponents.Import($readOnlyPath)
    Log "IMPORTED_MODIFIED_modCodexSAPReadOnly"

    # Import new modCF_VietnameseMessages
    $wb.VBProject.VBComponents.Import($vnNewPath)
    Log "IMPORTED_modCF_VietnameseMessages"

    $wb.Save()
    Log "SAVED_CP6_WIP"
    $wb.Close($false)
    $wb = $null
} catch {
    Log ("FATAL=" + $_.Exception.Message)
    if ($wb -ne $null) { try { $wb.Close($false) } catch {} }
    throw
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    # Kill any Excel processes started by this script
    $postExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $newExcel = $postExcel | Where-Object { $_ -notin $preExcel }
    if ($newExcel) {
        foreach ($excelPid in $newExcel) {
            if ((Get-Date) - $scriptStart -gt [TimeSpan]::FromMinutes(5)) {
                Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
                Log "KILLED_HUNG_EXCEL=$excelPid"
            }
        }
    }
}

Set-Content -Path $logPath -Value ($L -join "`r`n") -Encoding UTF8
Log "DONE"
