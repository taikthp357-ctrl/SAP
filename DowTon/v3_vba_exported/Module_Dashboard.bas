Attribute VB_Name = "Module_Dashboard"
Option Explicit

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Then
        SafeStr = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        SafeStr = ""
    Else
        SafeStr = Trim(CStr(v))
    End If
End Function

Private Function SafeVal(ByVal v As Variant) As Double
    If IsError(v) Then
        SafeVal = 0
    ElseIf IsEmpty(v) Or IsNull(v) Then
        SafeVal = 0
    ElseIf IsNumeric(v) Then
        SafeVal = CDbl(v)
    Else
        SafeVal = Val(v)
    End If
End Function

Public Function GetCheckSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If InStr(1, ws.Name, "Check", vbTextCompare) > 0 Or ws.CodeName = "Sheet4" Then
            Set GetCheckSheet = ws
            Exit Function
        End If
    Next ws
    On Error Resume Next
    Set GetCheckSheet = wb.Sheets("Check " & ChrW(7891) & "n")
    If GetCheckSheet Is Nothing Then Set GetCheckSheet = wb.Sheets(3)
    On Error GoTo 0
End Function

Public Function UniText(ByVal key As String) As String
    Select Case key
        Case "Rat_cao": UniText = "R" & ChrW(7845) & "t cao"
        Case "Cao": UniText = "Cao"
        Case "Trung_binh": UniText = "Trung b" & ChrW(236) & "nh"
        Case "Thap": UniText = "Th" & ChrW(7845) & "p"
        Case "Het_ton": UniText = "H" & ChrW(7871) & "t t" & ChrW(7891) & "n"
        Case "Duoi_min": UniText = "D" & ChrW(432) & ChrW(7899) & "i m" & ChrW(7913) & "c min"
        Case "Keo_line": UniText = ChrW(431) & "u ti" & ChrW(234) & "n k" & ChrW(233) & "o line"
        Case "UT_xuat": UniText = "UT xu" & ChrW(7845) & "t"
        Case "Du_ton": UniText = ChrW(272) & ChrW(7911) & " t" & ChrW(7891) & "n"
        Case "Theo_doi": UniText = "Theo d" & ChrW(245) & "i, l" & ChrW(234) & "n k" & ChrW(7871) & " ho" & ChrW(7841) & "ch"
        Case "Xuat_kho": UniText = ChrW(431) & "u ti" & ChrW(234) & "n xu" & ChrW(7845) & "t kho"
        Case "Binh_thuong": UniText = "B" & ChrW(236) & "nh th" & ChrW(434) & ChrW(7890) & "ng"
        Case "Tat_ca": UniText = "T" & ChrW(7845) & "t c" & ChrW(7843)
        Case Else: UniText = key
    End Select
End Function

Public Sub Dashboard_Refresh_MB52()
    ' Nút Refresh MB52: k?t n?i th?c SAP GUI MB52
    Call Export_MB52
End Sub

Public Sub Show_Select_Models()
    On Error Resume Next
    frmSelectModel.Show
End Sub

Public Sub Copy_CanhBao(Optional ByVal ShowMessage As Boolean = True)
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("DASHBOARD")
    
    Dim lastRow As Long
    lastRow = wsDash.Cells(wsDash.Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 10 Then
        If ShowMessage Then
            MsgBox "Kh" & ChrW(244) & "ng c" & ChrW(243) & " m" & ChrW(227) & " n" & ChrW(224) & "o c" & ChrW(7847) & "n c" & ChrW(7843) & "nh b" & ChrW(225) & "o!", vbInformation, "Th" & ChrW(244) & "ng B" & ChrW(225) & "o"
        End If
        Exit Sub
    End If
    
    wsDash.Range("B9:N" & lastRow).Copy
    
    If ShowMessage Then
        MsgBox ChrW(272) & ChrW(227) & " copy danh s" & ChrW(225) & "ch c" & ChrW(7843) & "nh b" & ChrW(225) & "o (" & (lastRow - 9) & " m" & ChrW(227) & ") v" & ChrW(224) & "o Clipboard!" & vbCrLf & _
               "B" & ChrW(7841) & "n c" & ChrW(243) & " th" & ChrW(7875) & " Paste tr" & ChrW(7921) & "c ti" & ChrW(7871) & "p v" & ChrW(224) & "o Teams, Zalo, Chat ho" & ChrW(7863) & "c Excel.", vbInformation, "Copy C" & ChrW(7843) & "nh B" & ChrW(225) & "o"
    End If
End Sub

Public Sub Sync_Pending_From_Model(ByVal strModels As String)
    Dim wb As Workbook
    Dim wsCheck As Worksheet, wsPending As Worksheet
    Dim lastRowCheck As Long, lastRowPending As Long
    Dim dicSelectedModels As Object, dicUniqueMat As Object
    Dim isAllModels As Boolean
    Dim parts() As String, p As Variant, mName As String
    Dim arrCheck() As Variant, r As Long
    Dim mat As String, cModel As String
    Dim k As Variant, outArr() As Variant, idx As Long
    
    Set wb = ThisWorkbook
    Set wsCheck = GetCheckSheet(wb)
    Set wsPending = wb.Sheets("Pending")
    
    If wsCheck.FilterMode Then
        On Error Resume Next
        wsCheck.ShowAllData
        On Error GoTo 0
    End If
    
    If wsPending.AutoFilterMode Then
        On Error Resume Next
        wsPending.ShowAllData
        wsPending.AutoFilterMode = False
        On Error GoTo 0
    End If
    wsPending.Rows.Hidden = False
    
    If Len(Trim(strModels)) = 0 Then strModels = "ALL"
    isAllModels = (UCase(Trim(strModels)) = "ALL" Or strModels = UniText("Tat_ca") Or strModels = "T" & ChrW(7845) & "t c" & ChrW(7843) Or strModels = "*")
    
    Set dicSelectedModels = CreateObject("Scripting.Dictionary")
    If Not isAllModels Then
        parts = Split(strModels, ",")
        For Each p In parts
            mName = Trim(CStr(p))
            If Len(mName) > 0 Then dicSelectedModels(UCase(mName)) = True
        Next p
    End If
    
    lastRowCheck = wsCheck.Cells(wsCheck.Rows.Count, "A").End(xlUp).Row
    If lastRowCheck < 2 Then Exit Sub
    
    arrCheck = wsCheck.Range("A2:J" & lastRowCheck).Value
    Set dicUniqueMat = CreateObject("Scripting.Dictionary")
    
    For r = 1 To UBound(arrCheck, 1)
        mat = SafeStr(arrCheck(r, 1))
        If Len(mat) > 0 Then
            cModel = SafeStr(arrCheck(r, 10))
            If isAllModels Or dicSelectedModels.Exists(UCase(cModel)) Then
                If Not dicUniqueMat.Exists(mat) Then
                    dicUniqueMat(mat) = True
                End If
            End If
        End If
    Next r
    
    ' Clear old Material in Pending Column A
    lastRowPending = wsPending.Cells(wsPending.Rows.Count, "A").End(xlUp).Row
    If lastRowPending >= 2 Then
        wsPending.Range("A2:A" & lastRowPending).ClearContents
    End If
    
    ' Write new unique materials to Pending starting at A2
    If dicUniqueMat.Count > 0 Then
        ReDim outArr(1 To dicUniqueMat.Count, 1 To 1)
        idx = 0
        For Each k In dicUniqueMat.Keys
            idx = idx + 1
            outArr(idx, 1) = k
        Next k
        wsPending.Range("A2").Resize(dicUniqueMat.Count, 1).Value = outArr
    End If
    
    ' Ensure Plant 3001 and 3093 in Column B (B2=3001, B3=3093)
    wsPending.Range("B2").Value = 3001
    wsPending.Range("B3").Value = 3093
    ' Clear any leftover plants below B3
    Dim lastPlantRow As Long
    lastPlantRow = wsPending.Cells(wsPending.Rows.Count, "B").End(xlUp).Row
    If lastPlantRow > 3 Then
        wsPending.Range("B4:B" & lastPlantRow).ClearContents
    End If
    
    ' Ensure Header and Control params in Pending
    wsPending.Range("A1").Value = "M" & ChrW(227)
    wsPending.Range("B1").Value = "Plant"
    wsPending.Range("D1").Value = "VL06O"
    wsPending.Range("D2").Value = "Shipping Point"
    wsPending.Range("D3").Value = 3093
    wsPending.Range("E2").Value = "Shipto"
    wsPending.Range("E3").Value = 3001
    wsPending.Range("F1").Value = "ME2N"
    wsPending.Range("F2").Value = "Plant"
    wsPending.Range("F3").Value = 3001
    wsPending.Range("G2").Value = "Supplying Plant"
    wsPending.Range("G3").Value = 3093
End Sub

Public Sub Calculate_All_Stock(Optional ByVal IsInteractive As Boolean = True, Optional ByVal IsSAPRefresh As Boolean = False)
    Dim wsDash As Worksheet, wsFull As Worksheet, wsCheck As Worksheet, wsMB As Worksheet
    Dim wb As Workbook
    Dim lastRowMB As Long
    Dim dicKhoBody As Object, dicLineBody As Object
    Dim dicTransit As Object, dicKhoMA As Object, dicLineMA As Object
    Dim dicDesc As Object
    Dim arrMB() As Variant
    Dim i As Long
    Dim mat As String, plant As String, sloc As String, mDesc As String
    Dim qty As Double, transit As Double
    Dim strModels As String, minLevel As Double
    Dim dicSelectedModels As Object
    Dim isAllModels As Boolean
    Dim parts() As String, p As Variant, mName As String
    Dim lastRowCheck As Long
    Dim arrCheck() As Variant
    Dim totalMaster As Long
    Dim arrFull() As Variant, arrDash() As Variant
    Dim idxFull As Long, idxDash As Long
    Dim cntDuoiMin As Long, cntKeoLine As Long, cntUTXuat As Long, cntHetTon As Long
    Dim r As Long, colIdx As Long
    Dim cModel As String, cUsage As Double
    Dim tKhoBody As Double, tLineBody As Double, cCanBuild As Double
    Dim inTrans As Double, tKhoMA As Double, tLineMA As Double, tongMA As Double
    Dim cRack As String, cPack As String, cThuTu As String, cTram As String
    Dim mucDo As String, canhBao As String, hanhDong As String
    Dim isLowStock As Boolean, totalStockBody As Double
    Dim matchModel As Boolean
    Dim oldLastFull As Long, subArrFull() As Variant
    Dim oldLastDash As Long, subArrDash() As Variant
    Dim modelTitle As String
    
    Set wb = ThisWorkbook
    Set wsDash = wb.Sheets("DASHBOARD")
    Set wsFull = wb.Sheets("Ton_Full")
    Set wsCheck = GetCheckSheet(wb)
    Set wsMB = wb.Sheets("MB52")
    
    application.ScreenUpdating = False
    application.EnableEvents = False
    application.Calculation = xlCalculationManual
    
    On Error GoTo ErrHandler
    
    ' 1. Read MB52 into Memory & Dictionaries
    lastRowMB = wsMB.Cells(wsMB.Rows.Count, "A").End(xlUp).Row
    
    Set dicKhoBody = CreateObject("Scripting.Dictionary")
    Set dicLineBody = CreateObject("Scripting.Dictionary")
    Set dicTransit = CreateObject("Scripting.Dictionary")
    Set dicKhoMA = CreateObject("Scripting.Dictionary")
    Set dicLineMA = CreateObject("Scripting.Dictionary")
    Set dicDesc = CreateObject("Scripting.Dictionary")
    
    If lastRowMB >= 2 Then
        arrMB = wsMB.Range("A2:I" & lastRowMB).Value
        
        For i = 1 To UBound(arrMB, 1)
            mat = SafeStr(arrMB(i, 1))
            If Len(mat) > 0 Then
                qty = SafeVal(arrMB(i, 4))
                plant = SafeStr(arrMB(i, 6))
                If InStr(plant, ".") > 0 Then plant = Left(plant, InStr(plant, ".") - 1)
                sloc = SafeStr(arrMB(i, 7))
                If InStr(sloc, ".") > 0 Then sloc = Left(sloc, InStr(sloc, ".") - 1)
                transit = SafeVal(arrMB(i, 8))
                mDesc = SafeStr(arrMB(i, 9))
                
                If Not dicDesc.Exists(mat) And Len(mDesc) > 0 Then dicDesc(mat) = mDesc
                
                If plant = "3001" And sloc = "1003" Then
                    dicKhoBody(mat) = dicKhoBody(mat) + qty
                ElseIf plant = "3001" And sloc = "3003" Then
                    dicLineBody(mat) = dicLineBody(mat) + qty
                End If
                
                dicTransit(mat) = dicTransit(mat) + transit
                
                If plant = "3093" And (sloc = "3093" Or sloc = "3003") Then
                    dicKhoMA(mat) = dicKhoMA(mat) + qty
                ElseIf plant = "3093" And (sloc = "3000" Or sloc = "3001") Then
                    dicLineMA(mat) = dicLineMA(mat) + qty
                End If
            End If
        Next i
    End If
    
    ' 2. Parse Selected Models and Min Level
    strModels = SafeStr(wsDash.Range("C3").Value)
    If Len(strModels) = 0 Then strModels = "ALL"
    
    minLevel = SafeVal(wsDash.Range("C4").Value)
    If minLevel <= 0 Then minLevel = 200: wsDash.Range("C4").Value = 200
    
    Set dicSelectedModels = CreateObject("Scripting.Dictionary")
    
    isAllModels = (UCase(strModels) = "ALL" Or strModels = UniText("Tat_ca") Or strModels = "T" & ChrW(7845) & "t c" & ChrW(7843) Or strModels = "*")
    
    If Not isAllModels Then
        parts = Split(strModels, ",")
        For Each p In parts
            mName = Trim(CStr(p))
            If Len(mName) > 0 Then dicSelectedModels(UCase(mName)) = True
        Next p
    End If
    
    ' 3. Process Master Data from Check ton
    If wsCheck.FilterMode Then
        On Error Resume Next
        wsCheck.ShowAllData
        On Error GoTo 0
    End If
    lastRowCheck = wsCheck.Cells(wsCheck.Rows.Count, "A").End(xlUp).Row
    If lastRowCheck < 2 Then GoTo FinishExit
    
    arrCheck = wsCheck.Range("A2:N" & lastRowCheck).Value
    totalMaster = UBound(arrCheck, 1)
    
    ReDim arrFull(1 To totalMaster, 1 To 19)
    ReDim arrDash(1 To totalMaster, 1 To 13)
    
    idxFull = 0: idxDash = 0
    cntDuoiMin = 0: cntKeoLine = 0: cntUTXuat = 0: cntHetTon = 0
    
    For r = 1 To totalMaster
        mat = SafeStr(arrCheck(r, 1))
        If Len(mat) > 0 Then
            cModel = SafeStr(arrCheck(r, 10))
            
            If isAllModels Then
                matchModel = True
            Else
                matchModel = dicSelectedModels.Exists(UCase(cModel))
            End If
            
            If matchModel Then
                cUsage = SafeVal(arrCheck(r, 2))
                If cUsage <= 0 Then cUsage = 1
                cRack = SafeStr(arrCheck(r, 11))
                cPack = SafeStr(arrCheck(r, 12))
                cThuTu = SafeStr(arrCheck(r, 13))
                cTram = SafeStr(arrCheck(r, 14))
                
                tKhoBody = 0: If dicKhoBody.Exists(mat) Then tKhoBody = dicKhoBody(mat)
                tLineBody = 0: If dicLineBody.Exists(mat) Then tLineBody = dicLineBody(mat)
                inTrans = 0: If dicTransit.Exists(mat) Then inTrans = dicTransit(mat)
                tKhoMA = 0: If dicKhoMA.Exists(mat) Then tKhoMA = dicKhoMA(mat)
                tLineMA = 0: If dicLineMA.Exists(mat) Then tLineMA = dicLineMA(mat)
                tongMA = tKhoMA + tLineMA
                cCanBuild = Round((tKhoBody + tLineBody) / cUsage, 1)
                
                mDesc = ""
                If dicDesc.Exists(mat) Then mDesc = dicDesc(mat)
                
                totalStockBody = tKhoBody + tLineBody
                
                ' Classification
                isLowStock = False
                If totalStockBody = 0 Then
                    mucDo = UniText("Rat_cao")
                    canhBao = UniText("Het_ton")
                    If inTrans > 0 Then
                        hanhDong = UniText("Xuat_kho")
                        cntUTXuat = cntUTXuat + 1
                    Else
                        hanhDong = UniText("Keo_line")
                        cntKeoLine = cntKeoLine + 1
                    End If
                    cntHetTon = cntHetTon + 1
                    cntDuoiMin = cntDuoiMin + 1
                    isLowStock = True
                ElseIf tLineBody = 0 Then
                    mucDo = UniText("Rat_cao")
                    canhBao = UniText("Keo_line")
                    hanhDong = UniText("Keo_line")
                    cntKeoLine = cntKeoLine + 1
                    cntDuoiMin = cntDuoiMin + 1
                    isLowStock = True
                ElseIf totalStockBody < minLevel Or tLineBody < minLevel Then
                    isLowStock = True
                    cntDuoiMin = cntDuoiMin + 1
                    If totalStockBody < minLevel * 0.3 Then
                        mucDo = UniText("Cao")
                        canhBao = UniText("Duoi_min")
                        If tKhoBody > 0 And tLineBody < 50 Then
                            hanhDong = UniText("Keo_line")
                            cntKeoLine = cntKeoLine + 1
                        Else
                            hanhDong = UniText("Theo_doi")
                        End If
                    Else
                        mucDo = UniText("Trung_binh")
                        canhBao = UniText("Duoi_min")
                        hanhDong = UniText("Theo_doi")
                    End If
                ElseIf inTrans > 0 And tKhoBody < minLevel Then
                    isLowStock = True
                    mucDo = UniText("Cao")
                    canhBao = UniText("UT_xuat")
                    hanhDong = UniText("Xuat_kho")
                    cntUTXuat = cntUTXuat + 1
                Else
                    mucDo = UniText("Thap")
                    canhBao = UniText("Du_ton")
                    hanhDong = UniText("Binh_thuong")
                End If
                
                ' Add to Ton_Full
                idxFull = idxFull + 1
                arrFull(idxFull, 1) = idxFull
                arrFull(idxFull, 2) = mat
                arrFull(idxFull, 3) = mDesc
                arrFull(idxFull, 4) = cModel
                arrFull(idxFull, 5) = cUsage
                arrFull(idxFull, 6) = tKhoBody
                arrFull(idxFull, 7) = tLineBody
                arrFull(idxFull, 8) = cCanBuild
                arrFull(idxFull, 9) = inTrans
                arrFull(idxFull, 10) = tKhoMA
                arrFull(idxFull, 11) = tLineMA
                arrFull(idxFull, 12) = tongMA
                arrFull(idxFull, 13) = cRack
                arrFull(idxFull, 14) = cPack
                arrFull(idxFull, 15) = cThuTu
                arrFull(idxFull, 16) = cTram
                arrFull(idxFull, 17) = mucDo
                arrFull(idxFull, 18) = canhBao
                arrFull(idxFull, 19) = hanhDong
                
                ' Add to DASHBOARD (Low stock only)
                If isLowStock Then
                    idxDash = idxDash + 1
                    arrDash(idxDash, 1) = idxDash
                    arrDash(idxDash, 2) = mat
                    arrDash(idxDash, 3) = mDesc
                    arrDash(idxDash, 4) = cModel
                    arrDash(idxDash, 5) = tLineBody
                    arrDash(idxDash, 6) = tKhoBody
                    arrDash(idxDash, 7) = tongMA
                    arrDash(idxDash, 8) = inTrans
                    arrDash(idxDash, 9) = cCanBuild
                    arrDash(idxDash, 10) = minLevel
                    arrDash(idxDash, 11) = mucDo
                    arrDash(idxDash, 12) = canhBao
                    arrDash(idxDash, 13) = hanhDong
                End If
            End If
        End If
    Next r
    
    ' 4. Write Ton_Full
    oldLastFull = wsFull.Cells(wsFull.Rows.Count, "A").End(xlUp).Row
    If oldLastFull >= 3 Then wsFull.Range("A3:S" & oldLastFull).Clear
    
    If idxFull > 0 Then
        ReDim subArrFull(1 To idxFull, 1 To 19)
        For r = 1 To idxFull
            For colIdx = 1 To 19: subArrFull(r, colIdx) = arrFull(r, colIdx): Next colIdx
        Next r
        wsFull.Range("A3").Resize(idxFull, 19).Value = subArrFull
        Call Format_Ton_Full_Rows(wsFull, idxFull)
    End If
    
    ' 5. Write DASHBOARD Table
    oldLastDash = wsDash.Cells(wsDash.Rows.Count, "B").End(xlUp).Row
    If oldLastDash >= 10 Then wsDash.Range("B10:N" & oldLastDash).Clear
    
    If idxDash > 0 Then
        ReDim subArrDash(1 To idxDash, 1 To 13)
        For r = 1 To idxDash
            For colIdx = 1 To 13: subArrDash(r, colIdx) = arrDash(r, colIdx): Next colIdx
        Next r
        wsDash.Range("B10").Resize(idxDash, 13).Value = subArrDash
        Call Format_Dashboard_Rows(wsDash, idxDash)
    End If
    
    ' 6. Update Control Info & KPI Summary
    ' Ch? c?p nh?t C5 và C7 khi th?c s? ch?y qua SAP Refresh!
    If IsSAPRefresh Then
        wsDash.Range("C5").Value = Format(Now, "dd/mm/yyyy hh:nn:ss")
        wsDash.Range("C7").Value = ChrW(272) & ChrW(227) & " t" & ChrW(7843) & "i MB52"
    End If
    
    wsDash.Range("C6").Value = idxDash
    
    If isAllModels Then modelTitle = UniText("Tat_ca") Else modelTitle = strModels
    wsDash.Range("L2").Value = "T" & ChrW(7892) & "NG H" & ChrW(7906) & "P C" & ChrW(7842) & "NH B" & ChrW(193) & "O (Model: " & modelTitle & ")"
    wsDash.Range("N3").Value = cntDuoiMin
    wsDash.Range("N4").Value = cntKeoLine
    wsDash.Range("N5").Value = cntUTXuat
    wsDash.Range("N6").Value = cntHetTon

FinishExit:
    application.ScreenUpdating = True
    application.EnableEvents = True
    application.Calculation = xlCalculationAutomatic
    
    If IsInteractive Then
        MsgBox ChrW(272) & ChrW(227) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t d" & ChrW(7919) & " li" & ChrW(7879) & "u th" & ChrW(224) & "nh c" & ChrW(244) & "ng!" & vbCrLf & _
               "- Model: " & strModels & vbCrLf & _
               "- T" & ChrW(7893) & "ng m" & ChrW(227) & " t" & ChrW(7843) & "i t" & ChrW(7891) & "n (Ton_Full): " & idxFull & vbCrLf & _
               "- T" & ChrW(7893) & "ng m" & ChrW(227) & " c" & ChrW(7843) & "nh b" & ChrW(225) & "o (DASHBOARD): " & idxDash, vbInformation, "DowTon Dashboard"
    End If
    Exit Sub

ErrHandler:
    application.ScreenUpdating = True
    application.EnableEvents = True
    application.Calculation = xlCalculationAutomatic
    MsgBox "L" & ChrW(7895) & "i khi c" & ChrW(7853) & "p nh" & ChrW(7853) & "t Dashboard: " & Err.Description, vbCritical, "L" & ChrW(7895) & "i"
End Sub

Private Sub Format_Dashboard_Rows(ByVal ws As Worksheet, ByVal numRows As Long)
    Dim dataRng As Range
    Set dataRng = ws.Range("B10:N" & (9 + numRows))
    
    dataRng.Font.Name = "Calibri"
    dataRng.Font.Size = 10
    ws.Range("B10:N" & (9 + numRows)).RowHeight = 20
    
    ' Alignments
    ws.Range("B10:B" & (9 + numRows)).HorizontalAlignment = -4108
    ws.Range("C10:C" & (9 + numRows)).HorizontalAlignment = -4108
    ws.Range("C10:C" & (9 + numRows)).Font.Bold = True
    ws.Range("D10:D" & (9 + numRows)).HorizontalAlignment = -4131
    ws.Range("E10:E" & (9 + numRows)).HorizontalAlignment = -4108
    
    ' Numbers
    ws.Range("F10:I" & (9 + numRows)).HorizontalAlignment = -4152
    ws.Range("F10:I" & (9 + numRows)).NumberFormat = "#,##0"
    ws.Range("J10:J" & (9 + numRows)).HorizontalAlignment = -4152
    ws.Range("J10:J" & (9 + numRows)).NumberFormat = "#,##0.0"
    ws.Range("K10:K" & (9 + numRows)).HorizontalAlignment = -4152
    ws.Range("K10:K" & (9 + numRows)).NumberFormat = "#,##0"
    
    ws.Range("L10:N" & (9 + numRows)).HorizontalAlignment = -4108
    
    ' Borders
    On Error Resume Next
    dataRng.Borders.LineStyle = 1
    dataRng.Borders.Color = RGB(217, 217, 217)
    
    Dim r As Long, cellMucDo As Range, cellCanhBao As Range, cellHanhDong As Range
    Dim valMucDo As String, valCanhBao As String, valHanhDong As String
    
    For r = 10 To 9 + numRows
        Set cellMucDo = ws.Cells(r, "L")
        Set cellCanhBao = ws.Cells(r, "M")
        Set cellHanhDong = ws.Cells(r, "N")
        
        valMucDo = Trim(CStr(cellMucDo.Value))
        valCanhBao = Trim(CStr(cellCanhBao.Value))
        valHanhDong = Trim(CStr(cellHanhDong.Value))
        
        ' Muc do
        If InStr(valMucDo, "cao") > 0 And InStr(valMucDo, "R") > 0 Then
            cellMucDo.Interior.Color = RGB(192, 0, 0)
            cellMucDo.Font.Color = RGB(255, 255, 255)
            cellMucDo.Font.Bold = True
        ElseIf valMucDo = "Cao" Then
            cellMucDo.Interior.Color = RGB(237, 125, 49)
            cellMucDo.Font.Color = RGB(255, 255, 255)
            cellMucDo.Font.Bold = True
        ElseIf InStr(valMucDo, "Trung") > 0 Then
            cellMucDo.Interior.Color = RGB(255, 230, 153)
            cellMucDo.Font.Color = RGB(128, 96, 0)
            cellMucDo.Font.Bold = True
        ElseIf InStr(valMucDo, "Th") > 0 Then
            cellMucDo.Interior.Color = RGB(169, 208, 142)
            cellMucDo.Font.Color = RGB(0, 97, 0)
            cellMucDo.Font.Bold = True
        End If
        
        ' Canh bao
        If InStr(valCanhBao, "ton") > 0 Or InStr(valCanhBao, "t" & ChrW(7891) & "n") > 0 Then
            If InStr(valCanhBao, "H") > 0 Then
                cellCanhBao.Interior.Color = RGB(252, 228, 214)
                cellCanhBao.Font.Color = RGB(192, 0, 0)
                cellCanhBao.Font.Bold = True
            ElseIf InStr(valCanhBao, ChrW(272)) > 0 Or InStr(valCanhBao, "D") > 0 Then
                cellCanhBao.Interior.Color = RGB(198, 239, 206)
                cellCanhBao.Font.Color = RGB(0, 97, 0)
                cellCanhBao.Font.Bold = True
            End If
        ElseIf InStr(valCanhBao, "line") > 0 Then
            cellCanhBao.Interior.Color = RGB(252, 228, 214)
            cellCanhBao.Font.Color = RGB(198, 89, 17)
            cellCanhBao.Font.Bold = True
        ElseIf InStr(valCanhBao, "min") > 0 Then
            cellCanhBao.Interior.Color = RGB(255, 242, 204)
            cellCanhBao.Font.Color = RGB(128, 96, 0)
            cellCanhBao.Font.Bold = True
        ElseIf InStr(valCanhBao, "UT") > 0 Then
            cellCanhBao.Interior.Color = RGB(221, 235, 247)
            cellCanhBao.Font.Color = RGB(31, 78, 121)
            cellCanhBao.Font.Bold = True
        End If
        
        ' Hanh dong uu tien
        If InStr(valHanhDong, "line") > 0 Then
            cellHanhDong.Font.Color = RGB(192, 0, 0)
            cellHanhDong.Font.Bold = True
        ElseIf InStr(valHanhDong, "xuat") > 0 Or InStr(valHanhDong, "xu" & ChrW(7845) & "t") > 0 Then
            cellHanhDong.Font.Color = RGB(31, 78, 121)
            cellHanhDong.Font.Bold = True
        End If
    Next r
    On Error GoTo 0
End Sub

Private Sub Format_Ton_Full_Rows(ByVal ws As Worksheet, ByVal numRows As Long)
    Dim dataRng As Range
    Set dataRng = ws.Range("A3:S" & (2 + numRows))
    
    dataRng.Font.Name = "Calibri"
    dataRng.Font.Size = 10
    ws.Range("A3:S" & (2 + numRows)).RowHeight = 19
    
    ws.Range("A3:A" & (2 + numRows)).HorizontalAlignment = -4108
    ws.Range("B3:B" & (2 + numRows)).HorizontalAlignment = -4108
    ws.Range("B3:B" & (2 + numRows)).Font.Bold = True
    ws.Range("C3:C" & (2 + numRows)).HorizontalAlignment = -4131
    ws.Range("D3:E" & (2 + numRows)).HorizontalAlignment = -4108
    
    ws.Range("F3:L" & (2 + numRows)).HorizontalAlignment = -4152
    ws.Range("F3:L" & (2 + numRows)).NumberFormat = "#,##0"
    ws.Range("H3:H" & (2 + numRows)).NumberFormat = "#,##0.0"
    
    ws.Range("M3:P" & (2 + numRows)).HorizontalAlignment = -4108
    ws.Range("Q3:S" & (2 + numRows)).HorizontalAlignment = -4108
    
    On Error Resume Next
    dataRng.Borders.LineStyle = 1
    dataRng.Borders.Color = RGB(217, 217, 217)
    On Error GoTo 0
End Sub


