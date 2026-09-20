$ErrorActionPreference = "Stop"

$PatchDir = "g:\SAP\v184_patch"
$V17Xlsm = "g:\SAP\BBBG_VAT\260803_V1.7_BBBG_VAT_SMART_GOS_ATT.xlsm"
$V17Zip = "g:\SAP\BBBG_VAT\260803_V1.7_BBBG_VAT_SMART_GOS_ATT.zip"
$V165Xlsm = "g:\SAP\CF_OB\260803_V1.6.5_BBBG_VAT.xlsm"
$RepoDir = "g:\SAP\repo_push"

$migoPath = Join-Path $PatchDir "MIGO311_RunPost_V24_ErrorGuard.bas"
$lines = [System.IO.File]::ReadAllLines($migoPath)

$newLines = @()
for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    if ($line -match "Private Sub ATT_ListFilesNoOcr") {
        $newLines += "Private Sub ATT_ListFilesNoOcr(ByVal files As Collection)"
        $newLines += "    Dim wsS As Worksheet, p As Variant, typ As String, total As Long"
        $newLines += "    Dim dataArr() As Variant, i As Long"
        $newLines += "    Set wsS = ThisWorkbook.Worksheets(SH_SCAN)"
        $newLines += "    wsS.Rows("2:" & wsS.Rows.Count).ClearContents"
        $newLines += "    total = files.Count"
        $newLines += "    If total = 0 Then Exit Sub"
        $newLines += "    ReDim dataArr(1 To total, 1 To 14)"
        $newLines += "    i = 0"
        $newLines += "    For Each p In files"
        $newLines += "        i = i + 1"
        $newLines += "        typ = ATT_Classify(ATT_FileName(CStr(p)), "")"
        $newLines += "        dataArr(i, 1) = i"
        $newLines += "        dataArr(i, 2) = ATT_FileName(CStr(p))"
        $newLines += "        dataArr(i, 3) = CStr(p)"
        $newLines += "        dataArr(i, 4) = LCase"
        $newLines += "        dataArr(i, 5) = typ"
        $newLines += "        dataArr(i, 6) = "FILENAME_ONLY""
        $newLines += "        dataArr(i, 13) = "SCANNED""
        $newLines += "    Next p"
        $newLines += "    wsS.Range(wsS.Cells(2, 1), wsS.Cells(total + 1, 14)).Value = dataArr"
        $newLines += "End Sub"
        # skip old sub
        while ($lines[$i] -ne "End Sub") { $i++ }
    }
    elseif ($line -eq "    ses.findById("wnd[0]").sendVKey 0") {
        $newLines += "    ses.findById("wnd[0]").sendVKey 0"
        $newLines += "    ATT_Wait 250"
        $newLines += "    On Error Resume Next"
        $newLines += "    Dim wndPopup As Object"
        $newLines += "    Set wndPopup = ses.findById("wnd[1]", False)"
        $newLines += "    If Not wndPopup Is Nothing Then"
        $newLines += "        Dim popText As String"
        $newLines += "        popText = CStr(wndPopup.Text) & " " & CStr(wndPopup.findById("usr/txtMESSTXT", False).Text)"
        $newLines += "        If InStr(1, popText, "lock", vbTextCompare) > 0 Or InStr(1, popText, "process", vbTextCompare) > 0 Or InStr(1, popText, "khoa", vbTextCompare) > 0 Or InStr(1, popText, "user", vbTextCompare) > 0 Then"
        $newLines += "            errText = "DOC_LOCKED_BY_USER: Chung tu dang bi khoa/su dung boi user khac tren SAP""
        $newLines += "            wndPopup.sendVKey 0"
        $newLines += "            ATT_Wait 200"
        $newLines += "            Exit Function"
        $newLines += "        End If"
        $newLines += "    End If"
        $newLines += "    On Error GoTo EH"
    }
    else {
        $newLines += $line
    }
}

# Append utility subs for Reset and Stop
$newLines += ""
$newLines += "Public Sub ResetStatusColumns()"
$newLines += "    Dim ws As Worksheet, lr As Long, ans As VbMsgBoxResult"
$newLines += "    Set ws = ThisWorkbook.Worksheets(SH_INPUT)"
$newLines += "    lr = Application.Max(ATT_LastRow(ws, C_MAT), ATT_LastRow(ws, C_DOC), ATT_LastRow(ws, C_PO), ATT_LastRow(ws, C_TICKET))"
$newLines += "    If lr < FIRST_ROW Then lr = FIRST_ROW + 10"
$newLines += "    ans = MsgBox("Ban co chac chan muon XOA KET QUA & TRANG THAI CU (cot 5 den 16) de quet moi khong?" & vbCrLf & vbCrLf & _"
$newLines += "                 "- Du lieu goc tai Cot 1-4 (Material, DOC, PO, So phieu) van duoc GIU NGUYEN." & vbCrLf & _"
$newLines += "                 "- Toan bo ket qua file match, trang thai HeadText va ket qua ATT se duoc xoa sach.", _"
$newLines += "                 vbQuestion + vbYesNo + vbDefaultButton2, "XAC NHAN XOA KET QUA CU")"
$newLines += "    If ans <> vbYes Then Exit Sub"
$newLines += "    Application.ScreenUpdating = False"
$newLines += "    Application.Calculation = xlCalculationManual"
$newLines += "    Application.EnableEvents = False"
$newLines += "    ws.Range(ws.Cells(FIRST_ROW, C_BBBG), ws.Cells(lr + 50, C_NOTE)).ClearContents"
$newLines += "    ws.Range(ws.Cells(FIRST_ROW, C_CHECK), ws.Cells(lr + 50, C_NOTE)).Interior.ColorIndex = xlNone"
$newLines += "    Application.Calculation = xlCalculationAutomatic"
$newLines += "    Application.EnableEvents = True"
$newLines += "    Application.ScreenUpdating = True"
$newLines += "    MsgBox "Da xoa sach ket qua cu. Bang INPUT da san sang cho luot quet moi!", vbInformation, "XOA TRANG THAI THANH CONG""
$newLines += "End Sub"
$newLines += ""
$newLines += "Public Sub StopCurrentRun()"
$newLines += "    If Not mToolBusy Then"
$newLines += "        MsgBox "Hien tai khong co tien trinh nao dang chay.", vbInformation, "THONG BAO""
$newLines += "        Exit Sub"
$newLines += "    End If"
$newLines += "    mStop = True"
$newLines += "    Application.StatusBar = "DANG YEU CAU DUNG... Tool se dung an toan sau khi hoan tat DOC hien tai.""
$newLines += "    MsgBox "Da gui lenh DUNG. Tool se hoan tat chung tu dang xu ly va dung lai an toan.", vbExclamation, "DUNG TIEN TRINH""
$newLines += "End Sub"

[System.IO.File]::WriteAllLines($migoPath, $newLines, (New-Object System.Text.UTF8Encoding $false))
Write-Host "Updated MIGO311 module with batch array writing, doc lock check, reset & stop subs." -ForegroundColor Green

# Refurbish UI via Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Opening $V17Xlsm for UI refurbishment..."
    $wb = $excel.Workbooks.Open($V17Xlsm)
    $wsIn = $wb.Sheets.Item("INPUT")
    
    # Delete existing Reset or Stop shapes
    foreach ($sp in $wsIn.Shapes) {
        if ($sp.Name -eq "btnReset" -or $sp.Name -eq "btnStop") {
            $sp.Delete()
        }
    }
    
    # Add btnReset
    $btnReset = $wsIn.Shapes.AddShape(1, 600, 48, 140, 30)
    $btnReset.Name = "btnReset"
    $btnReset.TextFrame2.TextRange.Text = "XOA KET QUA CU"
    $btnReset.OnAction = "ResetStatusColumns"
    $btnReset.Fill.Solid()
    $btnReset.Fill.ForeColor.RGB = 0x595959
    $btnReset.TextFrame2.TextRange.Font.Bold = $true
    $btnReset.TextFrame2.TextRange.Font.Size = 10
    $btnReset.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
    
    # Add btnStop
    $btnStop = $wsIn.Shapes.AddShape(1, 600, 84, 140, 30)
    $btnStop.Name = "btnStop"
    $btnStop.TextFrame2.TextRange.Text = "DUNG TOOL"
    $btnStop.OnAction = "StopCurrentRun"
    $btnStop.Fill.Solid()
    $btnStop.Fill.ForeColor.RGB = 0x2626C0
    $btnStop.TextFrame2.TextRange.Font.Bold = $true
    $btnStop.TextFrame2.TextRange.Font.Size = 10
    $btnStop.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
    
    # Format Material DOC Column in ATT_HISTORY as Text
    $wsHist = $wb.Sheets.Item("ATT_HISTORY")
    $wsHist.Columns.Item(2).NumberFormat = "@"
    
    # Clean CONFIG default path
    $wsCfg = $wb.Sheets.Item("CONFIG")
    $wsCfg.Cells.Item(4, 2).Value = ""
    
    # Re-import VBA modules
    $vbaModules = @("modSupports", "ModuleSAP", "modWinFileDialog", "MIGO311_RunPost_V24_ErrorGuard")
    $vbaProj = $wb.VBProject
    foreach ($mName in $vbaModules) {
        $patchFile = Join-Path $PatchDir "$mName.bas"
        $comp = $null
        foreach ($c in $vbaProj.VBComponents) {
            if ($c.Name -eq $mName) { $comp = $c; break }
        }
        if ($null -ne $comp) { $vbaProj.VBComponents.Remove($comp) }
        $vbaProj.VBComponents.Import($patchFile) | Out-Null
        Write-Host "Imported upgraded module: $mName" -ForegroundColor Green
    }
    
    $wb.Save()
    $wb.Close($false)
    Write-Host "Saved $V17Xlsm successfully" -ForegroundColor Green
    
    # Also patch V1.6.5
    if (Test-Path $V165Xlsm) {
        Write-Host "Patching $V165Xlsm..."
        $wb2 = $excel.Workbooks.Open($V165Xlsm)
        $vbaProj2 = $wb2.VBProject
        foreach ($mName in $vbaModules) {
            $patchFile = Join-Path $PatchDir "$mName.bas"
            $comp2 = $null
            foreach ($c in $vbaProj2.VBComponents) {
                if ($c.Name -eq $mName) { $comp2 = $c; break }
            }
            if ($null -ne $comp2) { $vbaProj2.VBComponents.Remove($comp2) }
            $vbaProj2.VBComponents.Import($patchFile) | Out-Null
        }
        $wb2.Save()
        $wb2.Close($false)
        Write-Host "Saved $V165Xlsm successfully" -ForegroundColor Green
    }
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# Pack ZIP & Copy
if (Test-Path $V17Zip) { Remove-Item $V17Zip -Force }
Compress-Archive -Path $V17Xlsm -DestinationPath $V17Zip -Force
Write-Host "Created ZIP: $V17Zip" -ForegroundColor Green

Copy-Item -Path $V17Xlsm -Destination (Join-Path $RepoDir "260803_V1.7_BBBG_VAT_SMART_GOS_ATT.xlsm") -Force
Copy-Item -Path $V17Zip -Destination (Join-Path $RepoDir "260803_V1.7_BBBG_VAT_SMART_GOS_ATT.zip") -Force
Write-Host "Copied updated release files to $RepoDir" -ForegroundColor Green
