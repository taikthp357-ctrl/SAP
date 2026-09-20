# =============================================================================
#  apply_perfect_v17_release.ps1
#  Complete UI Refurbishment, Enhanced HeadText & History Logging, and Packaging
# =============================================================================

$ErrorActionPreference = "Stop"

$PatchDir = "g:\SAP\v184_patch"
$V17Xlsm = "g:\SAP\BBBG_VAT\260803_V1.7_BBBG_VAT_SMART_GOS_ATT.xlsm"
$V17Zip = "g:\SAP\BBBG_VAT\260803_V1.7_BBBG_VAT_SMART_GOS_ATT.zip"
$V165Xlsm = "g:\SAP\CF_OB\260803_V1.6.5_BBBG_VAT.xlsm"
$RepoDir = "g:\SAP\repo_push"

# 1. Update MIGO311_RunPost_V24_ErrorGuard.bas with complete logging for SAP_AttachCheckedFiles
$migoPath = Join-Path $PatchDir "MIGO311_RunPost_V24_ErrorGuard.bas"
$migoContent = [System.IO.File]::ReadAllText($migoPath)

# Add sapUser to SAP_AttachCheckedFiles and enhanced BBBG/VAT history
$oldAttachLoop = @"
            ' -- ATT BBBG --
            Dim bParts As Variant, bIdx As Long, bPath As String, bAllOk As Boolean
            If Trim$(CStr(g("BBBG"))) = "" Then
                ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "NOT_FOUND"
            ElseIf ATT_TitleExists(titles, titleB) Or ATT_WasPreviouslyAttached(CStr(docKey), "BBBG", CStr(g("BBBG"))) Then
                ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "ALREADY_ATTACHED"
                ATT_Log "ATT", "ATT_ALREADY_EXISTS", CStr(docKey), "BBBG: " & titleB
            Else
                bParts = Split(CStr(g("BBBG")), "|"): bAllOk = True
                For bIdx = LBound(bParts) To UBound(bParts)
                    bPath = Trim$(CStr(bParts(bIdx)))
                    If bPath <> "" Then
                        If Not ATT_AttachOneSmartGOS(ses, bPath, titleB, errText) Then
                            bAllOk = False: ATT_AddNoteRows ws, g("ROWS"), "BBBG: " & errText
                            ATT_Log "ATT", "ATT_FAIL_BBBG", CStr(docKey), bPath & " | " & errText
                        Else
                            ATT_Log "ATT", "ATT_VERIFIED", CStr(docKey), "BBBG: " & titleB
                        End If
                    End If
                Next bIdx
                If bAllOk Then
                    ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "ATT_SENT"
                Else
                    ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "ATT_PARTIAL_FAIL"
                    docAttachFailed = True
                    Exit For
                End If
            End If

            ' -- ATT VAT --
            If Trim$(CStr(g("VAT"))) = "" Then
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "NOT_FOUND"
            ElseIf ATT_TitleExists(titles, titleV) Or ATT_WasPreviouslyAttached(CStr(docKey), "VAT", CStr(g("VAT"))) Then
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "ALREADY_ATTACHED"
                ATT_Log "ATT", "ATT_ALREADY_EXISTS", CStr(docKey), "VAT: " & titleV
            ElseIf ATT_AttachOneSmartGOS(ses, CStr(g("VAT")), titleV, errText) Then
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "ATT_SENT"
                ATT_Log "ATT", "ATT_VERIFIED", CStr(docKey), "VAT: " & titleV
            Else
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "ATT_FAIL"
                ATT_AddNoteRows ws, g("ROWS"), "VAT: " & errText
                ATT_Log "ATT", "ATT_FAIL_VAT", CStr(docKey), CStr(g("VAT")) & " | " & errText
                docAttachFailed = True
                Exit For
            End If
"@

$newAttachLoop = @"
            ' -- ATT BBBG --
            Dim bParts As Variant, bIdx As Long, bPath As String, bAllOk As Boolean
            If Trim$(CStr(g("BBBG"))) = "" Then
                ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "NOT_FOUND"
                ATT_History CStr(docKey), CStr(g("TICKET")), "BBBG", "", "NOT_FOUND", sapUser, "Khong tim thay file BBBG trong thu muc nguon."
            ElseIf ATT_TitleExists(titles, titleB) Or ATT_WasPreviouslyAttached(CStr(docKey), "BBBG", CStr(g("BBBG"))) Then
                ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "ALREADY_ATTACHED"
                ATT_History CStr(docKey), CStr(g("TICKET")), "BBBG", titleB, "ALREADY_EXISTS", sapUser, "File BBBG da co tren SAP (khong can attach lai)."
                ATT_Log "ATT", "ATT_ALREADY_EXISTS", CStr(docKey), "BBBG: " & titleB, "", runId
            Else
                bParts = Split(CStr(g("BBBG")), "|"): bAllOk = True
                For bIdx = LBound(bParts) To UBound(bParts)
                    bPath = Trim$(CStr(bParts(bIdx)))
                    If bPath <> "" Then
                        If Not ATT_AttachOneSmartGOS(ses, bPath, titleB, errText) Then
                            bAllOk = False: ATT_AddNoteRows ws, g("ROWS"), "BBBG: " & errText
                            ATT_History CStr(docKey), CStr(g("TICKET")), "BBBG", bPath, "FAIL", sapUser, "Loi attach BBBG: " & errText & ". Kiem tra SAP Security hoac file bi lock."
                            ATT_Log "ATT", "ATT_FAIL_BBBG", CStr(docKey), bPath & " | " & errText, bPath, runId
                        Else
                            ATT_History CStr(docKey), CStr(g("TICKET")), "BBBG", bPath, "DONE", sapUser, "Attach file BBBG thanh cong qua SAP GOS."
                            ATT_Log "ATT", "ATT_VERIFIED", CStr(docKey), "BBBG: " & titleB, bPath, runId
                        End If
                    End If
                Next bIdx
                If bAllOk Then
                    ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "ATT_SENT"
                Else
                    ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "ATT_PARTIAL_FAIL"
                    docAttachFailed = True
                    Exit For
                End If
            End If

            ' -- ATT VAT --
            If Trim$(CStr(g("VAT"))) = "" Then
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "NOT_FOUND"
                ATT_History CStr(docKey), CStr(g("TICKET")), "VAT", "", "NOT_FOUND", sapUser, "Khong tim thay file VAT trong thu muc nguon."
            ElseIf ATT_TitleExists(titles, titleV) Or ATT_WasPreviouslyAttached(CStr(docKey), "VAT", CStr(g("VAT"))) Then
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "ALREADY_ATTACHED"
                ATT_History CStr(docKey), CStr(g("TICKET")), "VAT", titleV, "ALREADY_EXISTS", sapUser, "File VAT da co tren SAP (khong can attach lai)."
                ATT_Log "ATT", "ATT_ALREADY_EXISTS", CStr(docKey), "VAT: " & titleV, "", runId
            ElseIf ATT_AttachOneSmartGOS(ses, CStr(g("VAT")), titleV, errText) Then
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "ATT_SENT"
                ATT_History CStr(docKey), CStr(g("TICKET")), "VAT", CStr(g("VAT")), "DONE", sapUser, "Attach file VAT thanh cong qua SAP GOS."
                ATT_Log "ATT", "ATT_VERIFIED", CStr(docKey), "VAT: " & titleV, CStr(g("VAT")), runId
            Else
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "ATT_FAIL"
                ATT_AddNoteRows ws, g("ROWS"), "VAT: " & errText
                ATT_History CStr(docKey), CStr(g("TICKET")), "VAT", CStr(g("VAT")), "FAIL", sapUser, "Loi attach VAT: " & errText & ". Kiem tra SAP Security hoac file bi lock."
                ATT_Log "ATT", "ATT_FAIL_VAT", CStr(docKey), CStr(g("VAT")) & " | " & errText, CStr(g("VAT")), runId
                docAttachFailed = True
                Exit For
            End If
"@

$migoContentNorm = $migoContent -replace "`r`n", "`n"
$oldAttachLoopNorm = $oldAttachLoop -replace "`r`n", "`n"
$newAttachLoopNorm = $newAttachLoop -replace "`r`n", "`n"

if ($migoContentNorm.Contains($oldAttachLoopNorm)) {
    $migoContentNorm = $migoContentNorm.Replace($oldAttachLoopNorm, $newAttachLoopNorm)
    Write-Host "Updated SAP_AttachCheckedFiles loop history logging successfully." -ForegroundColor Green
}

# Also ensure sapUser is declared in SAP_AttachCheckedFiles
$migoContentNorm = $migoContentNorm -replace "Dim preflightDone As Boolean\n", "Dim preflightDone As Boolean, sapUser As String, runId As String`n"
$migoContentNorm = $migoContentNorm -replace "ATT_ActivateMigoSession ses\n", "ATT_ActivateMigoSession ses`n    runId = ""ATT_"" & Format$(Now, ""yyyymmdd_hhnnss"")`n    On Error Resume Next: sapUser = CStr(ses.Info.User): If sapUser = """" Then sapUser = Environ$(""USERNAME""): On Error GoTo EH`n"
$migoContentNorm = $migoContentNorm -replace "ATT_MarkVerified ws, docInfo\n", "ATT_MarkVerified ws, docInfo, sapUser`n"

$migoFinal = $migoContentNorm -replace "`n", "`r`n"
[System.IO.File]::WriteAllText($migoPath, $migoFinal, (New-Object System.Text.UTF8Encoding $false))
Write-Host "MIGO module updated successfully without BOM." -ForegroundColor Green

# 2. Open Excel COM and refurbish UI + Import VBA Modules
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Opening $V17Xlsm for UI refurbishment and VBA patching..."
    $wb = $excel.Workbooks.Open($V17Xlsm)
    
    # ------------------ SHEET INPUT UI ------------------
    $wsIn = $wb.Sheets("INPUT")
    
    # Banner
    $wsIn.Range("A1").Value = "TOOL ATT BBBG / VAT - CHECK HEADTEXT + FILENAME - V1.7 SMART GOS ATT (HARDENED)"
    $wsIn.Range("A1:P1").Merge()
    $wsIn.Range("A1:P1").Font.Name = "Calibri"
    $wsIn.Range("A1:P1").Font.Size = 14
    $wsIn.Range("A1:P1").Font.Bold = $true
    $wsIn.Range("A1:P1").Font.Color = 0xFFFFFF # White
    $wsIn.Range("A1:P1").Interior.Color = 0x7D491F # Navy Blue (RGB: 31, 73, 125)
    $wsIn.Range("A1:P1").HorizontalAlignment = -4108 # Center
    $wsIn.Range("A1:P1").VerticalAlignment = -4108 # Center
    $wsIn.Rows.Item(1).RowHeight = 28
    
    # Subtitle
    $wsIn.Range("A2").Value = "QUY TRINH 4 BUOC: 1. Chon thu muc | 2. So ten file & HeadText | 3. Dien HeadText (SAP) | 4. Attach SAP (GOS) - Tra cuu lich su tai ATT_HISTORY & LOG"
    $wsIn.Range("A2:P2").Merge()
    $wsIn.Range("A2:P2").Font.Name = "Calibri"
    $wsIn.Range("A2:P2").Font.Size = 10
    $wsIn.Range("A2:P2").Font.Italic = $true
    $wsIn.Range("A2:P2").Font.Color = 0x1F497D # Dark Blue
    $wsIn.Range("A2:P2").Interior.Color = 0xF2E6DC # Light Ice Blue
    $wsIn.Range("A2:P2").HorizontalAlignment = -4108
    $wsIn.Range("A2:P2").VerticalAlignment = -4108
    $wsIn.Rows.Item(2).RowHeight = 18
    
    # Clean rows 3-5 height
    $wsIn.Rows.Item(3).RowHeight = 12
    $wsIn.Rows.Item(4).RowHeight = 35
    $wsIn.Rows.Item(5).RowHeight = 35
    
    # Format and align 4 Buttons
    foreach ($sp in $wsIn.Shapes) {
        if ($sp.Name -eq "btnChooseFolder") {
            $sp.TextFrame2.TextRange.Text = "1. CHON THU MUC"
            $sp.OnAction = "SelectSourceFolder_NoOCR"
            $sp.Top = 50
            $sp.Left = 10
            $sp.Width = 280
            $sp.Height = 30
            $sp.Fill.Solid()
            $sp.Fill.ForeColor.RGB = 0x804000 # Blue/Navy
            $sp.TextFrame2.TextRange.Font.Bold = $true
            $sp.TextFrame2.TextRange.Font.Size = 11
            $sp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
        }
        elseif ($sp.Name -eq "btnCheck") {
            $sp.TextFrame2.TextRange.Text = "2. SO TEN FILE & HEADTEXT"
            $sp.OnAction = "CheckFiles_NoOCR"
            $sp.Top = 50
            $sp.Left = 305
            $sp.Width = 280
            $sp.Height = 30
            $sp.Fill.Solid()
            $sp.Fill.ForeColor.RGB = 0x7A5B00 # Teal / Slate Blue
            $sp.TextFrame2.TextRange.Font.Bold = $true
            $sp.TextFrame2.TextRange.Font.Size = 11
            $sp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
        }
        elseif ($sp.Name -eq "btnHeadText") {
            $sp.TextFrame2.TextRange.Text = "3. DIEN HEADTEXT (SAP)"
            $sp.OnAction = "SAP_FillHeadTextOnly"
            $sp.Top = 88
            $sp.Left = 10
            $sp.Width = 280
            $sp.Height = 30
            $sp.Fill.Solid()
            $sp.Fill.ForeColor.RGB = 0x0066CC # Amber / Orange
            $sp.TextFrame2.TextRange.Font.Bold = $true
            $sp.TextFrame2.TextRange.Font.Size = 11
            $sp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
        }
        elseif ($sp.Name -eq "btnAttach") {
            $sp.TextFrame2.TextRange.Text = "4. ATTACH SAP (GOS)"
            $sp.OnAction = "SAP_AttachCheckedFiles"
            $sp.Top = 88
            $sp.Left = 305
            $sp.Width = 280
            $sp.Height = 30
            $sp.Fill.Solid()
            $sp.Fill.ForeColor.RGB = 0x33802A # Forest Green
            $sp.TextFrame2.TextRange.Font.Bold = $true
            $sp.TextFrame2.TextRange.Font.Size = 11
            $sp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
        }
    }
    
    # Headers on Row 6
    $headers = @(
        "Material", "Material DOC", "PO/SA", "So phieu (INPUT)", 
        "File BBBG", "File VAT", "Match BBBG", "Match VAT", 
        "HeadText SAP", "Match Note", "CHECK Status", "HeadText Status", 
        "ATT BBBG", "ATT VAT", "ATT Status", "Ghi chu tong hop"
    )
    for ($c = 1; $c -le $headers.Count; $c++) {
        $wsIn.Cells.Item(6, $c).Value = $headers[$c-1]
    }
    $wsIn.Range("A6:P6").Font.Bold = $true
    $wsIn.Range("A6:P6").Font.Color = 0xFFFFFF
    $wsIn.Range("A6:P6").Interior.Color = 0x7D491F
    $wsIn.Range("A6:P6").HorizontalAlignment = -4108
    $wsIn.Rows.Item(6).RowHeight = 24
    
    # ------------------ SHEET ATT_HISTORY UI ------------------
    $wsHist = $wb.Sheets("ATT_HISTORY")
    $histHeaders = @(
        "Thoi gian (Time)", "Material DOC", "So phieu (Ticket)", 
        "Loai (Type)", "Doi tuong / File / HeadText", "Ket qua (Result)", 
        "Nguoi dung SAP (User)", "Chi tiet & Huong dan sua loi (Diagnostic Note)"
    )
    for ($c = 1; $c -le $histHeaders.Count; $c++) {
        $wsHist.Cells.Item(1, $c).Value = $histHeaders[$c-1]
    }
    $wsHist.Range("A1:H1").Font.Bold = $true
    $wsHist.Range("A1:H1").Font.Color = 0xFFFFFF
    $wsHist.Range("A1:H1").Interior.Color = 0x926036 # Dark Steel Blue
    $wsHist.Range("A1:H1").HorizontalAlignment = -4108
    $wsHist.Rows.Item(1).RowHeight = 24
    $wsHist.Columns.Item(1).ColumnWidth = 20
    $wsHist.Columns.Item(2).ColumnWidth = 15
    $wsHist.Columns.Item(3).ColumnWidth = 22
    $wsHist.Columns.Item(4).ColumnWidth = 14
    $wsHist.Columns.Item(5).ColumnWidth = 42
    $wsHist.Columns.Item(6).ColumnWidth = 16
    $wsHist.Columns.Item(7).ColumnWidth = 16
    $wsHist.Columns.Item(8).ColumnWidth = 60
    
    # ------------------ SHEET LOG UI ------------------
    $wsLog = $wb.Sheets("LOG")
    $logHeaders = @(
        "Thoi gian (Time)", "Stage", "Status", 
        "Material DOC", "Message / Chi tiet", "File / Duong dan", "Run ID"
    )
    for ($c = 1; $c -le $logHeaders.Count; $c++) {
        $wsLog.Cells.Item(1, $c).Value = $logHeaders[$c-1]
    }
    $wsLog.Range("A1:G1").Font.Bold = $true
    $wsLog.Range("A1:G1").Font.Color = 0xFFFFFF
    $wsLog.Range("A1:G1").Interior.Color = 0x3B3B3B # Dark Slate
    $wsLog.Range("A1:G1").HorizontalAlignment = -4108
    $wsLog.Rows.Item(1).RowHeight = 24
    $wsLog.Columns.Item(1).ColumnWidth = 20
    $wsLog.Columns.Item(2).ColumnWidth = 22
    $wsLog.Columns.Item(3).ColumnWidth = 16
    $wsLog.Columns.Item(4).ColumnWidth = 15
    $wsLog.Columns.Item(5).ColumnWidth = 50
    $wsLog.Columns.Item(6).ColumnWidth = 42
    $wsLog.Columns.Item(7).ColumnWidth = 24

    # ------------------ SHEET GUIDE UI ------------------
    $wsGuide = $wb.Sheets("GUIDE")
    $wsGuide.Range("A1").Value = "HUONG DAN SU DUNG TOOL ATT BBBG/VAT (V1.7 SMART GOS ATT HARDENED)"
    $wsGuide.Range("A1:H1").Merge()
    $wsGuide.Range("A1:H1").Font.Bold = $true
    $wsGuide.Range("A1:H1").Font.Size = 13
    $wsGuide.Range("A1:H1").Font.Color = 0xFFFFFF
    $wsGuide.Range("A1:H1").Interior.Color = 0x7D491F
    $wsGuide.Range("A1:H1").HorizontalAlignment = -4108
    $wsGuide.Rows.Item(1).RowHeight = 26
    
    # ------------------ IMPORT VBA MODULES ------------------
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
    Write-Host "Successfully refurbished and saved $V17Xlsm" -ForegroundColor Green

    # Also apply to V1.6.5
    if (Test-Path $V165Xlsm) {
        Write-Host "Patching and formatting $V165Xlsm..."
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
        Write-Host "Successfully patched and saved $V165Xlsm" -ForegroundColor Green
    }
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# 3. Verify VBA binary size and package ZIP
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($V17Xlsm)
$vbaEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/vbaProject.bin" }
$vbaSize = $vbaEntry.Length
$zip.Dispose()
Write-Host "vbaProject.bin size: $vbaSize bytes" -ForegroundColor Green

if (Test-Path $V17Zip) { Remove-Item $V17Zip -Force }
Compress-Archive -Path $V17Xlsm -DestinationPath $V17Zip -Force
Write-Host "Created release ZIP: $V17Zip" -ForegroundColor Green

# 4. Copy to repo_push
Copy-Item -Path $V17Xlsm -Destination (Join-Path $RepoDir "260803_V1.7_BBBG_VAT_SMART_GOS_ATT.xlsm") -Force
Copy-Item -Path $V17Zip -Destination (Join-Path $RepoDir "260803_V1.7_BBBG_VAT_SMART_GOS_ATT.zip") -Force
Write-Host "Copied updated deliverables to $RepoDir" -ForegroundColor Green

Write-Host "==================================================" -ForegroundColor Green
Write-Host "  PERFECT V1.7 BUILD & UI REFURBISHMENT COMPLETE! " -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Green
