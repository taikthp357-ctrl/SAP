Attribute VB_Name = "modCF_Posting"
Option Explicit

Private Const CF_UI_POSTING As String = "POSTING"
Private Const CF_RAW_POSTING As String = "POSTING_RAW"
Private Const CF_QUERY_ORDER As String = "ORDER_QUERY"

Public Sub CF_LoadPostingFromWH()
    Dim ob As String, msg As String, started As Double
    Dim products As Object, wsWH As Worksheet, wsOrder As Worksheet
    Dim wsUI As Worksheet, wsRaw As Worksheet
    Dim lastRow As Long, r As Long, outRow As Long, mat As String, obState As String
    Dim stateRows As Long
    Dim oldSU As Boolean, oldEv As Boolean, oldCalc As XlCalculation
    Dim ok As Boolean, stage As String, k As Variant
    Dim mappedRows As Long
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Failed

    gMPOST_LastLoadOK = False
    gMPOST_LastLoadError = ""
    gMPOST_LastRows = 0
    gMPOST_LastProductCount = 0

    If Not RT_GetSingleInputOB(ob, msg) Then Err.Raise vbObjectError + 9501, , msg
    started = RT_StartTimer()
    RT_SetStatus "STAGE", "LOAD POSTING"
    RT_Log "LOAD_POSTING", ob, "", "START", "Bat dau tai ton tu /SCWM/POST."

    Set wsUI = ThisWorkbook.Worksheets(CF_UI_POSTING)
    Set products = CreateObject("Scripting.Dictionary")
    products.CompareMode = vbTextCompare

    If gMPOST_UseOrderQuery Then
        ' Phase 2 source: ORDER_QUERY was built only from PASS rows in OB_WH_CACHE.
        Set wsOrder = CF_EnsureSheet(CF_QUERY_ORDER, True)
        Set products = MPOST_ReadOrderProducts(wsOrder)
        If products.count = 0 Then Err.Raise vbObjectError + 9502, , _
            "ORDER_QUERY khong co Material hop le tu OB_WH_CACHE cho Phase 2."
    Else
        ' Single/current-OB compatibility source: WH_ACTUAL remains the working view.
        Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
        lastRow = wsWH.Cells(wsWH.rows.count, 6).End(xlUp).row
        For r = 5 To lastRow
            mat = Trim$(CStr(wsWH.Cells(r, 6).Value2 & ""))
            obState = UCase$(Trim$(CStr(wsWH.Cells(r, 25).Value2 & "")))
            If mat <> "" Then
                stateRows = stateRows + 1
                If obState = "OB_OPEN_OR_PARTIAL" Then products(mat) = True
            End If
        Next r
        If products.count = 0 Then
            If stateRows > 0 Then Err.Raise vbObjectError + 9502, , _
                "Khong co Material thuoc OB_OPEN_OR_PARTIAL de truy van POSTING; OB_DONE/NOT_PROVEN duoc STOP."
            RT_Log "LOAD_POSTING", ob, "", "INFO", _
                   "WH_ACTUAL chua co Material; tu dong chay LOAD WH truoc."
            RT_LoadMON
            lastRow = wsWH.Cells(wsWH.rows.count, 6).End(xlUp).row
            For r = 5 To lastRow
                mat = Trim$(CStr(wsWH.Cells(r, 6).Value2 & ""))
                obState = UCase$(Trim$(CStr(wsWH.Cells(r, 25).Value2 & "")))
                If mat <> "" And obState = "OB_OPEN_OR_PARTIAL" Then products(mat) = True
            Next r
        End If
        If products.count = 0 Then Err.Raise vbObjectError + 9502, , _
            "Khong co Material trong WH_ACTUAL de truy van POSTING."

        Set wsOrder = CF_EnsureSheet(CF_QUERY_ORDER, True)
        wsOrder.Cells.ClearContents
        wsOrder.Range("A1:C1").Value = Array("Product", "Qty", "Tram")
        outRow = 2
        For Each k In products.keys
            wsOrder.Cells(outRow, 1).Value2 = CStr(k)
            wsOrder.Cells(outRow, 2).Value2 = 1
            wsOrder.Cells(outRow, 3).Value2 = "CF_OB_SINGLE"
            outRow = outRow + 1
        Next k
    End If

    Set wsRaw = CF_EnsureSheet(CF_RAW_POSTING, True)
    wsRaw.Cells.ClearContents
    wsRaw.Range("A1:Q1").Value = Array("Product", "A.Qty AUoM", "Batch", "Source Bin", "TST", _
        "Total Stk", "Dest. Bin", "QtyPostChg", "Dest.Batch", "ST", "GR Date", "GR Time", _
        "Proc. Type", "AvlQtyBUoM", "", "", "Allocated Qty")

    oldSU = Application.ScreenUpdating
    oldEv = Application.EnableEvents
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    stage = "OPEN_POST_AND_DOWNLOAD"
    CF_SwapPostingSheets wsUI, wsRaw, True
    ok = ModuleSAP.DownloadPostingForOrder(False)
    If Not ok Then Err.Raise vbObjectError + 9503, , _
        "ModuleSAP khong tai duoc du lieu POSTING. " & ModuleSAP.LastPostingError
    RT_Log "LOAD_POSTING", ob, "", "INFO", _
           "SAP raw after download: lastRow=" & CStr(CF_LastDataRow(wsRaw, 1, 2)) & _
           "; A2=" & Trim$(CStr(wsRaw.Cells(2, 1).Value2 & ""))

    stage = "MAP_POSTING"
    CF_MapRawPosting wsRaw, wsUI, ob
    RT_Log "LOAD_POSTING", ob, "", "INFO", _
           "SAP read valid rows=" & CStr(ModuleSAP.LastPostingRowsRead) & _
           "; POSTING map preview: A2=" & Trim$(CStr(wsUI.Cells(2, 1).Value2 & "")) & _
           "; B2=" & Trim$(CStr(wsUI.Cells(2, 2).Value2 & ""))
    CF_SwapPostingSheets wsUI, wsRaw, False
    Application.ScreenUpdating = oldSU
    Application.EnableEvents = oldEv
    Application.Calculation = oldCalc

    RT_SetStatus "STAGE", "POSTING DONE"
    RT_SetStatus "TIME", Format$(RT_Elapsed(started), "0.000") & " s"
    mappedRows = CF_CountPostingRows(wsUI)
    If mappedRows > 0 Then
        RT_Log "LOAD_POSTING", ob, "", "PASS", _
               "Da tai va map " & CStr(mappedRows) & " dong ton vao POSTING."
    Else
        RT_Log "LOAD_POSTING", ob, "", "INFO", _
               "SAP khong tra dong ton hop le cho Material; POSTING duoc de trong."
    End If
    gMPOST_LastLoadOK = True
    gMPOST_LastRows = ModuleSAP.LastPostingRowsRead
    gMPOST_LastProductCount = products.count
    If Not gRT_Silent Then MsgBox "LOAD POSTING PASS" & vbCrLf & _
        "Material: " & CStr(products.count), vbInformation, "SAP CF OB"
    RT_SetStatus "STAGE", "IDLE"
    Exit Sub

Failed:
    ' Capture the original error before cleanup.  CF_SwapPostingSheets and
    ' Excel state restoration run under Resume Next and can otherwise clear
    ' Err, leaving the user with the useless "Err 0" diagnostic.
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    CF_SwapPostingSheets wsUI, wsRaw, False
    Application.ScreenUpdating = oldSU
    Application.EnableEvents = oldEv
    Application.Calculation = oldCalc
    RT_SetStatus "STAGE", "POSTING FAIL"
    RT_SetStatus "TIME", Format$(RT_Elapsed(started), "0.000") & " s"
    gMPOST_LastLoadError = "Stage=" & stage & "; Err " & CStr(failureNumber) & ": " & failureDescription
    RT_Log "LOAD_POSTING", ob, "", "FAIL", gMPOST_LastLoadError
    If Not gRT_Silent Then MsgBox "LOAD POSTING failed" & vbCrLf & _
        Err.Description, vbExclamation, "SAP CF OB"
    RT_SetStatus "STAGE", "IDLE"
End Sub

Private Function MPOST_ReadOrderProducts(ByVal wsOrder As Worksheet) As Object
    Dim d As Object, lastRow As Long, r As Long, material As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    If wsOrder Is Nothing Then Set MPOST_ReadOrderProducts = d: Exit Function
    lastRow = wsOrder.Cells(wsOrder.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        material = Trim$(CStr(wsOrder.Cells(r, 1).Value2 & ""))
        If material <> "" Then d(material) = True
    Next r
    Set MPOST_ReadOrderProducts = d
End Function

Private Function CF_EnsureSheet(ByVal sheetName As String, ByVal makeVeryHidden As Boolean) As Worksheet
    On Error Resume Next
    Set CF_EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If CF_EnsureSheet Is Nothing Then
        Set CF_EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        CF_EnsureSheet.Name = sheetName
    End If
    If makeVeryHidden Then CF_EnsureSheet.Visible = xlSheetVeryHidden
End Function

Private Sub CF_SwapPostingSheets(ByVal wsUI As Worksheet, ByVal wsRaw As Worksheet, ByVal toRaw As Boolean)
    If toRaw Then
        wsUI.Name = "POSTING_UI"
        wsRaw.Name = "POSTING"
    Else
        wsRaw.Name = "POSTING_RAW"
        wsUI.Name = "POSTING"
    End If
End Sub
Private Sub CF_MapRawPosting(ByVal wsRaw As Worksheet, ByVal wsUI As Worksheet, ByVal ob As String)
    Dim lastRaw As Long, r As Long, c As Long, outRow As Long, mat As String
    Dim formatLastRow As Long

    ' Visible POSTING keeps SAP read fields A:F and action payload G:I.
    ' J:W is the CHECK analysis area; no SAP write is performed here.
    wsUI.Range("A2:W2000").ClearContents
    lastRaw = CF_LastDataRow(wsRaw, 1, 2)
    outRow = 2

    For r = 2 To lastRaw
        mat = Trim$(CStr(wsRaw.Cells(r, 1).Value2 & ""))
        If mat <> "" Then
            For c = 1 To 6
                wsUI.Cells(outRow, c).Value2 = wsRaw.Cells(r, c).Value2
            Next c
            outRow = outRow + 1
            If outRow > 2000 Then Exit For
        End If
    Next r

    formatLastRow = Application.Max(2, outRow - 1)
    wsUI.Range("B2:B" & CStr(formatLastRow)).NumberFormat = "#,##0.###"
    wsUI.Range("C2:C" & CStr(formatLastRow)).NumberFormat = "@"
    wsUI.Range("G2:G" & CStr(formatLastRow)).NumberFormat = "@"
    wsUI.Range("H2:H" & CStr(formatLastRow)).NumberFormat = "#,##0.###"
    wsUI.Range("I2:I" & CStr(formatLastRow)).NumberFormat = "@"
    wsUI.Range("L2:M" & CStr(formatLastRow)).NumberFormat = "#,##0.###"
    wsUI.Range("R2:S" & CStr(formatLastRow)).NumberFormat = "#,##0.###"
    wsUI.Range("J2:K" & CStr(formatLastRow)).NumberFormat = "@"
    wsUI.Range("N2:Q" & CStr(formatLastRow)).NumberFormat = "@"
    wsUI.Range("T2:W" & CStr(formatLastRow)).NumberFormat = "@"
End Sub

Private Function CF_LastDataRow(ByVal ws As Worksheet, ByVal colNo As Long, Optional ByVal firstRow As Long = 2) As Long
    Dim r As Long
    r = ws.Cells(ws.rows.count, colNo).End(xlUp).row
    If r < firstRow Then CF_LastDataRow = firstRow - 1 Else CF_LastDataRow = r
End Function

Private Function CF_CountPostingRows(ByVal ws As Worksheet) As Long
    Dim r As Long
    For r = 2 To 2000
        If Trim$(CStr(ws.Cells(r, 1).Value2 & "")) <> "" Then CF_CountPostingRows = CF_CountPostingRows + 1
    Next r
End Function

' Tao ke hoach MIGO tu dung cac gia tri da xac nhan trong WH_ACTUAL.
' Khong dung Allocation lam so luong: Allocation chi la so dong/phan bo cua WH.
Public Sub CF_RunPostingFromWH()
    ' Compatibility entry point: POSTING chi tai ton tu WH_ACTUAL, khong goi MIGO.
    CF_LoadPostingFromWH
End Sub

Private Sub CF_RunPostingFromWH_MIGO()
    Dim ob As String, msg As String, started As Double, stage As String
    Dim wsWH As Worksheet, wsPlan As Worksheet, wsGroup As Worksheet
    Dim lastWH As Long, r As Long, n As Long, planRow As Long, groupRow As Long
    Dim material As String, batch As String, actualBin As String, whNo As String, wtNo As String
    Dim qty As Double, statusText As String, wtStateText As String, stateText As String, groupStatus As String, groupMsg As String
    Dim groupDelivery As String, allPosted As Boolean, okRows As Long
    Dim plant As String, sloc As String, slocTo As String, tramText As String
    On Error GoTo Failed

    If Not RT_GetSingleInputOB(ob, msg) Then Err.Raise vbObjectError + 9601, , msg
    started = RT_StartTimer()
    RT_SetStatus "STAGE", "PREPARE MIGO 311"
    RT_Log "POSTING", ob, "", "START", "Bat dau tao du lieu MIGO tu WH_ACTUAL; khong doc lai OB tu man hinh."

    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
    Set wsPlan = ThisWorkbook.Worksheets(CF_UI_POSTING)
    Set wsGroup = CF_EnsureSheet("OB_GROUPS", True)
    wsGroup.Cells.ClearContents
    wsGroup.Range("A1:Q1").Value = Array("OB", "Tram", "Product", "Qty", "Plant", "SLoc", "SLoc Transfer", "Batch", _
        "Status", "SAP Message", "Delivery", "Post Time", "Run ID", "Product Count", "Row Count", "Group Key", "Warning")

    plant = Trim$(ModuleHelper.ToolConfig("DEFAULT_PLANT", "3093"))
    sloc = Trim$(ModuleHelper.ToolConfig("DEFAULT_SLOC", "1005"))
    slocTo = Trim$(ModuleHelper.ToolConfig("DEFAULT_SLOC_TRANSFER", "3001"))
    If plant = "" Or sloc = "" Or slocTo = "" Then _
        Err.Raise vbObjectError + 9602, , "Thieu DEFAULT_PLANT/DEFAULT_SLOC/DEFAULT_SLOC_TRANSFER trong CONFIG."
    If sloc = slocTo Then Err.Raise vbObjectError + 9603, , "DEFAULT_SLOC va DEFAULT_SLOC_TRANSFER dang giong nhau."

    wsPlan.Range("A5:M504").ClearContents
    lastWH = wsWH.Cells(wsWH.rows.count, 1).End(xlUp).row
    planRow = 5: groupRow = 2
    For r = 5 To lastWH
        If Trim$(CStr(wsWH.Cells(r, 1).Value2 & "")) = ob Then
            material = Trim$(CStr(wsWH.Cells(r, 6).Value2 & ""))
            batch = Trim$(CStr(wsWH.Cells(r, 9).Value2 & ""))
            actualBin = Trim$(CStr(wsWH.Cells(r, 13).Value2 & ""))
            whNo = Trim$(CStr(wsWH.Cells(r, 4).Value2 & ""))
            wtNo = Trim$(CStr(wsWH.Cells(r, 5).Value2 & ""))
            qty = RT_Number(wsWH.Cells(r, 12).Value2)
            stateText = UCase$(Trim$(CStr(wsWH.Cells(r, 17).Value2 & "")))
            statusText = UCase$(Trim$(CStr(wsWH.Cells(r, 10).Value2 & "")))
            wtStateText = UCase$(Trim$(CStr(wsWH.Cells(r, 26).Value2 & "")))
            If material <> "" And qty > 0 And stateText <> "POSTED" _
               And (wtStateText = "WT_ACTIVE" Or wtStateText = "WT_CANCELLED_REMAINING") Then
                tramText = "CF OB " & ob & " WH " & whNo & " WT " & wtNo & " BIN " & actualBin

                wsPlan.Cells(planRow, 1).Value2 = material
                wsPlan.Cells(planRow, 2).Value2 = actualBin
                wsPlan.Cells(planRow, 3).Value2 = batch
                wsPlan.Cells(planRow, 4).Value2 = ""
                wsPlan.Cells(planRow, 5).Value2 = batch
                wsPlan.Cells(planRow, 6).Value2 = qty
                wsPlan.Cells(planRow, 7).Value2 = "WH=" & whNo & "; WT=" & wtNo
                wsPlan.Cells(planRow, 8).Value2 = "READY"
                wsPlan.Cells(planRow, 9).Value2 = "PENDING"
                wsPlan.Cells(planRow, 11).Value2 = r
                wsPlan.Cells(planRow, 12).Value2 = CStr(wsWH.Cells(r, 20).Value2 & "")
                wsPlan.Cells(planRow, 13).Value2 = "WH_ACTUAL -> MIGO 311"

                wsGroup.Cells(groupRow, 1).Value2 = ob
                wsGroup.Cells(groupRow, 2).Value2 = tramText
                wsGroup.Cells(groupRow, 3).Value2 = material
                wsGroup.Cells(groupRow, 4).Value2 = qty
                wsGroup.Cells(groupRow, 5).Value2 = plant
                wsGroup.Cells(groupRow, 6).Value2 = sloc
                wsGroup.Cells(groupRow, 7).Value2 = slocTo
                wsGroup.Cells(groupRow, 8).NumberFormat = "@"
                wsGroup.Cells(groupRow, 8).Value2 = batch
                wsGroup.Cells(groupRow, 9).Value2 = "READY"
                wsGroup.Cells(groupRow, 13).Value2 = gRT_RunID
                wsGroup.Cells(groupRow, 14).Value2 = 1
                wsGroup.Cells(groupRow, 15).Value2 = 1
                wsGroup.Cells(groupRow, 16).Value2 = ob & "|" & material & "|" & batch & "|" & whNo & "|" & wtNo
                planRow = planRow + 1
                groupRow = groupRow + 1
                n = n + 1
            End If
        End If
    Next r
    If n = 0 Then Err.Raise vbObjectError + 9604, , "WH_ACTUAL khong co dong OPEN co Actual Qty > 0 cho OB " & ob & "."

    RT_Log "POSTING", ob, "", "INFO", "Da map " & n & " dong WH_ACTUAL -> OB_GROUPS; Plant=" & plant & "; SLoc=" & sloc & " -> " & slocTo & "."
    RT_SetStatus "STAGE", "RUN MIGO 311"
    Run_MIGO311_OneClick

    allPosted = True
    For groupRow = 2 To n + 1
        groupStatus = Trim$(CStr(wsGroup.Cells(groupRow, 9).Value2 & ""))
        groupMsg = Trim$(CStr(wsGroup.Cells(groupRow, 10).Value2 & ""))
        groupDelivery = Trim$(CStr(wsGroup.Cells(groupRow, 11).Value2 & ""))
        r = CF_FindWHRowByGroupKey(wsWH, ob, CStr(wsGroup.Cells(groupRow, 16).Value2 & ""))
        If UCase$(groupStatus) = "POSTED" Then
            okRows = okRows + 1
            CF_UpdatePostingResult wsPlan, groupRow + 3, "POSTED", groupDelivery, groupMsg
            If r > 0 Then CF_UpdateWHResult wsWH, r, "POSTED", groupDelivery, groupMsg
        Else
            allPosted = False
            CF_UpdatePostingResult wsPlan, groupRow + 3, groupStatus, groupDelivery, groupMsg
            If r > 0 Then CF_UpdateWHResult wsWH, r, groupStatus, groupDelivery, groupMsg
        End If
    Next groupRow

    If allPosted Then
        RT_Log "POSTING", ob, "", "PASS", "MIGO 311 da POSTED " & okRows & " dong tu WH_ACTUAL."
    Else
        RT_Log "POSTING", ob, "", "FAIL", "MIGO 311 chua POSTED het; OK=" & okRows & "/" & n & ". Xem OB_GROUPS va POSTING."
    End If
    RT_SetStatus "TIME", Format$(RT_Elapsed(started), "0.000") & " s"
    RT_SetStatus "STAGE", IIf(allPosted, "POSTING DONE", "POSTING ERROR")
    Exit Sub

Failed:
    RT_SetStatus "STAGE", "POSTING ERROR"
    RT_SetStatus "TIME", Format$(RT_Elapsed(started), "0.000") & " s"
    RT_Log "POSTING", ob, "", "FAIL", "Stage=" & stage & "; Err " & CStr(Err.Number) & ": " & Err.Description
End Sub

Public Sub CF_DownloadPostingStock()
    CF_LoadPostingFromWH
End Sub

Public Sub CF_RebuildPostingFromRaw()
    Dim wsRaw As Worksheet, wsUI As Worksheet
    On Error GoTo Failed
    Set wsRaw = ThisWorkbook.Worksheets(CF_RAW_POSTING)
    Set wsUI = ThisWorkbook.Worksheets(CF_UI_POSTING)
    CF_MapRawPosting wsRaw, wsUI, Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("A7").Value2 & ""))
    RT_Log "LOAD_POSTING", Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("A7").Value2 & "")), "", "PASS", _
        "Da map lai POSTING tu POSTING_RAW, cho phep A.Qty AUoM = 0; so dong=" & CStr(CF_CountPostingRows(wsUI)) & "."
    Exit Sub
Failed:
    RT_Log "LOAD_POSTING", "", "", "FAIL", "REMAP_RAW: Err " & CStr(Err.Number) & ": " & Err.Description
End Sub

Private Function CF_FindWHRowByGroupKey(ByVal ws As Worksheet, ByVal ob As String, ByVal groupKey As String) As Long
    Dim lastRow As Long, r As Long, keyText As String
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 5 To lastRow
        keyText = Trim$(CStr(ws.Cells(r, 1).Value2 & "")) & "|" & Trim$(CStr(ws.Cells(r, 6).Value2 & "")) & "|" & _
            Trim$(CStr(ws.Cells(r, 9).Value2 & "")) & "|" & Trim$(CStr(ws.Cells(r, 4).Value2 & "")) & "|" & Trim$(CStr(ws.Cells(r, 5).Value2 & ""))
        If Trim$(CStr(ws.Cells(r, 1).Value2 & "")) = ob And keyText = groupKey Then CF_FindWHRowByGroupKey = r: Exit Function
    Next r
End Function
Private Sub CF_UpdatePostingResult(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal statusText As String, ByVal deliveryNo As String, ByVal messageText As String)
    Dim noteText As String
    noteText = messageText

    If deliveryNo <> "" Then
        If noteText <> "" Then
            noteText = "Delivery=" & deliveryNo & "; " & noteText
        Else
            noteText = "Delivery=" & deliveryNo
        End If
    End If

    ' CHECK area uses U=Action, V=Check Result, W=Note.
    ws.Cells(rowNo, 22).Value2 = statusText
    ws.Cells(rowNo, 23).Value2 = noteText
End Sub

Private Sub CF_UpdateWHResult(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal statusText As String, ByVal deliveryNo As String, ByVal messageText As String)
    ws.Cells(rowNo, 16).Value2 = IIf(deliveryNo <> "", "MIGO_311=" & deliveryNo, "MIGO_311")
    ws.Cells(rowNo, 17).Value2 = statusText
    ws.Cells(rowNo, 18).Value2 = IIf(deliveryNo <> "", deliveryNo, statusText)
    ws.Cells(rowNo, 19).Value2 = messageText
End Sub
Public Sub CF_CheckActualPosting()
    Dim wsWH As Worksheet, wsPost As Worksheet
    Dim lastWH As Long, lastPost As Long
    Dim whRow As Long, postRow As Long
    Dim material As String, sysBatch As String, sysBin As String
    Dim actualBin As String, areaText As String, actualBatchInput As String
    Dim actualBinResolved As String, actualBatch As String, matchBatch As String
    Dim actualQtyText As String, sysQtyText As String
    Dim postMaterial As String, postBatch As String, postBin As String
    Dim availableText As String, totalText As String
    Dim actualQty As Double, sysQty As Double, available As Double, totalStock As Double
    Dim needQtyText As String, whNo As String, wtNo As String
    Dim matCount As Long, batchCount As Long, candidateCount As Long
    Dim firstMaterialRow As Long, firstBatchRow As Long, firstCandidateRow As Long
    Dim resultText As String, actionText As String, noteText As String, holdText As String
    Dim checked As Long, readyCount As Long, needMonCount As Long, reviewCount As Long, failCount As Long
    Dim started As Double, tol As Double, ob As String, doc As String
    Dim hasUserInput As Boolean, postNeeded As Boolean, binDiff As Boolean, batchDiff As Boolean, qtyDiff As Boolean
    Dim stockAvailable As Boolean, stockTotal As Boolean, areaPrefix As String, postStep As String
    Dim statusText As String, obStateText As String, wtStateText As String, needsPostingData As Boolean

    On Error GoTo Failed

    started = RT_StartTimer()
    tol = 0.000001
    RT_SetStatus "STAGE", "CHECK ACTUAL"

    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
    Set wsPost = modSourceStockRepo.SSR_BackingSheet()   ' V2.8 Phase2 redirect (SHEET mode = POSTING). Analysis cols G:W relocation deferred to Phase3.

    lastWH = wsWH.Cells(wsWH.rows.count, 6).End(xlUp).row
    lastPost = CF_LastDataRow(wsPost, 1, 2)
    For whRow = 5 To lastWH
        statusText = UCase$(Trim$(CStr(wsWH.Cells(whRow, 10).Value2 & "")))
        obStateText = UCase$(Trim$(CStr(wsWH.Cells(whRow, 25).Value2 & "")))
        wtStateText = UCase$(Trim$(CStr(wsWH.Cells(whRow, 26).Value2 & "")))
        actualQtyText = Trim$(CStr(wsWH.Cells(whRow, 12).Value2 & ""))
        If Trim$(CStr(wsWH.Cells(whRow, 6).Value2 & "")) <> "" And _
           obStateText = "OB_OPEN_OR_PARTIAL" _
           And (wtStateText = "WT_ACTIVE" Or wtStateText = "WT_CANCELLED_REMAINING") _
           And actualQtyText <> "" And IsNumeric(actualQtyText) Then
            If RT_Number(actualQtyText) > 0# Then needsPostingData = True
        End If
    Next whRow
    If lastPost < 2 And needsPostingData Then Err.Raise vbObjectError + 9701, , _
        "POSTING chua co du lieu A:F. Hay TAI POSTING truoc."

    ' CHECK only analyzes workbook data; it does not call SAP.
    wsPost.Range("G2:W2000").ClearContents
    wsWH.Range("O5:S504").ClearContents

    For whRow = 5 To lastWH
        material = Trim$(CStr(wsWH.Cells(whRow, 6).Value2 & ""))
        actualQtyText = Trim$(CStr(wsWH.Cells(whRow, 12).Value2 & ""))
        actualBin = Trim$(CStr(wsWH.Cells(whRow, 13).Value2 & ""))
        areaText = Trim$(CStr(wsWH.Cells(whRow, 21).Value2 & ""))
        actualBatchInput = Trim$(CStr(wsWH.Cells(whRow, 22).Value2 & ""))
        statusText = UCase$(Trim$(CStr(wsWH.Cells(whRow, 10).Value2 & "")))
        obStateText = UCase$(Trim$(CStr(wsWH.Cells(whRow, 25).Value2 & "")))
        wtStateText = UCase$(Trim$(CStr(wsWH.Cells(whRow, 26).Value2 & "")))

        If obStateText = "OB_DONE" Then
            resultText = "DONE"
            actionText = "DONE"
            noteText = "OB_STATE=OB_DONE: all four OB header statuses are Completed; no WT/Posting/CF action."
            GoTo WriteWHResult
        End If

        If obStateText <> "OB_OPEN_OR_PARTIAL" Then
            resultText = "STOP_OB_STATE_NOT_PROVEN"
            actionText = "STOP"
            noteText = "OB header state is not proven; no Posting/CF action is allowed."
            GoTo WriteWHResult
        End If

        If wtStateText = "WT_DONE" Then
            resultText = "DONE"
            actionText = "DONE"
            noteText = "WT_STATE=WT_DONE: WT already confirmed; no Posting/CF action required."
            GoTo WriteWHResult
        End If

        If wtStateText <> "WT_ACTIVE" And wtStateText <> "WT_CANCELLED_REMAINING" Then
            resultText = "STOP_WT_STATE"
            actionText = "STOP"
            noteText = "WT state is not eligible for Posting in V1.2.1; Remaining Qty reader is not complete."
            GoTo WriteWHResult
        End If

        hasUserInput = (actualQtyText <> "" Or actualBin <> "" Or _
                        areaText <> "" Or actualBatchInput <> "")
        If material = "" Or Not hasUserInput Then GoTo NextWHRow

        checked = checked + 1
        sysBatch = Trim$(CStr(wsWH.Cells(whRow, 9).Value2 & ""))
        sysBin = Trim$(CStr(wsWH.Cells(whRow, 8).Value2 & ""))
        sysQtyText = Trim$(CStr(wsWH.Cells(whRow, 7).Value2 & ""))
        ob = Trim$(CStr(wsWH.Cells(whRow, 1).Value2 & ""))
        doc = Trim$(CStr(wsWH.Cells(whRow, 2).Value2 & ""))
        whNo = Trim$(CStr(wsWH.Cells(whRow, 4).Value2 & ""))
        wtNo = Trim$(CStr(wsWH.Cells(whRow, 5).Value2 & ""))

        resultText = "REVIEW"
        actionText = "REVIEW"
        noteText = ""
        holdText = ""
        postNeeded = False
        actualBatch = actualBatchInput
        If actualBatch = "" Then actualBatch = sysBatch
        matchBatch = actualBatch
        actualBinResolved = actualBin

        If actualQtyText = "" Then
            resultText = "MISSING_ACTUAL_QTY"
            noteText = "Nhap Actual Qty truoc khi CHECK."
        ElseIf Not IsNumeric(actualQtyText) Then
            resultText = "INVALID_ACTUAL_QTY"
            noteText = "Actual Qty khong phai so: " & actualQtyText
        Else
            actualQty = RT_Number(actualQtyText)
            If actualQty <= 0# Then
                resultText = "USER_SKIP"
                actionText = "USER_SKIP"
                noteText = "Actual Qty=0: user explicitly skips this WT; no Posting/CF action."
                GoTo WriteWHResult
            ElseIf actualBin = "" And areaText = "" Then
                resultText = "MISSING_ACTUAL_BIN"
                noteText = "Nhap Actual Bin hoac Area (P/ZF/HR)."
            ElseIf matchBatch = "" Then
                resultText = "MISSING_ACTUAL_BATCH"
                noteText = "Khong co Actual Batch va Sys Batch de match."
            Else
                matCount = 0
                batchCount = 0
                candidateCount = 0
                firstMaterialRow = 0
                firstBatchRow = 0
                firstCandidateRow = 0

                For postRow = 2 To lastPost
                    postMaterial = Trim$(CStr(wsPost.Cells(postRow, 1).Value2 & ""))
                    postBatch = Trim$(CStr(wsPost.Cells(postRow, 3).Value2 & ""))
                    postBin = Trim$(CStr(wsPost.Cells(postRow, 4).Value2 & ""))

                    If StrComp(postMaterial, material, vbTextCompare) = 0 Then
                        matCount = matCount + 1
                        If firstMaterialRow = 0 Then firstMaterialRow = postRow

                        If StrComp(postBatch, matchBatch, vbTextCompare) = 0 Then
                            batchCount = batchCount + 1
                            If firstBatchRow = 0 Then firstBatchRow = postRow

                            If actualBin <> "" Then
                                If StrComp(postBin, actualBin, vbTextCompare) = 0 Then
                                    candidateCount = candidateCount + 1
                                    If firstCandidateRow = 0 Then firstCandidateRow = postRow
                                End If
                            ElseIf CF_BinInArea(postBin, areaText) Then
                                candidateCount = candidateCount + 1
                                If firstCandidateRow = 0 Then firstCandidateRow = postRow
                            End If
                        End If
                    End If
                Next postRow

                If batchCount = 0 Then
                    resultText = "BATCH_NOT_FOUND"
                    noteText = "Khong thay Material/Actual Batch trong POSTING."
                    If matCount > 0 Then
                        CF_WritePostingCheck wsPost, firstMaterialRow, ob, whNo, _
                            sysQtyText, actualQtyText, sysBin, actualBin, _
                            sysBatch, actualBatch, actualQtyText, "", "", _
                            actionText, resultText, noteText, False
                    End If

                ElseIf candidateCount = 0 Then
                    If actualBin = "" And areaText <> "" Then
                        resultText = "NO_BIN_IN_AREA"
                        noteText = "Khong thay bin trong Area=" & areaText & _
                            " voi Material/Actual Batch."
                    Else
                        resultText = "BIN_NOT_IN_POSTING"
                        noteText = "Khong thay Actual Bin voi Material/Actual Batch."
                    End If
                    CF_WritePostingCheck wsPost, firstBatchRow, ob, whNo, _
                        sysQtyText, actualQtyText, sysBin, actualBin, _
                        sysBatch, actualBatch, actualQtyText, "", "", _
                        actionText, resultText, noteText, False

                ElseIf candidateCount > 1 Then
                    resultText = "MULTI_BIN"
                    actionText = "REVIEW"
                    noteText = "Co " & CStr(candidateCount) & _
                        " dong cung Material/Batch; khong tu chon bin."
                    For postRow = 2 To lastPost
                        postMaterial = Trim$(CStr(wsPost.Cells(postRow, 1).Value2 & ""))
                        postBatch = Trim$(CStr(wsPost.Cells(postRow, 3).Value2 & ""))
                        postBin = Trim$(CStr(wsPost.Cells(postRow, 4).Value2 & ""))
                        If StrComp(postMaterial, material, vbTextCompare) = 0 And _
                           StrComp(postBatch, matchBatch, vbTextCompare) = 0 Then
                            If (actualBin <> "" And StrComp(postBin, actualBin, vbTextCompare) = 0) Or _
                               (actualBin = "" And CF_BinInArea(postBin, areaText)) Then
                                CF_WritePostingCheck wsPost, postRow, ob, whNo, _
                                    sysQtyText, actualQtyText, sysBin, postBin, _
                                    sysBatch, actualBatch, actualQtyText, _
                                    Trim$(CStr(wsPost.Cells(postRow, 2).Value2 & "")), _
                                    "YES", actionText, resultText, noteText, False
                            End If
                        End If
                    Next postRow

                Else
                    postRow = firstCandidateRow
                    postBin = Trim$(CStr(wsPost.Cells(postRow, 4).Value2 & ""))
                    If actualBin = "" Then
                        actualBinResolved = postBin
                        noteText = "Area=" & areaText & " auto-select bin=" & postBin & "; "
                    Else
                        actualBinResolved = actualBin
                    End If

                    availableText = Trim$(CStr(wsPost.Cells(postRow, 2).Value2 & ""))
                    totalText = Trim$(CStr(wsPost.Cells(postRow, 6).Value2 & ""))
                    needQtyText = actualQtyText
                    noteText = noteText & "Match=Material+Batch+Bin; " & _
                        "SysQty=" & sysQtyText & "; ActualQty=" & actualQtyText & _
                        "; SysBin=" & sysBin & "; ActualBin=" & actualBinResolved & _
                        "; SysBatch=" & sysBatch & "; ActualBatch=" & actualBatch

                    If sysQtyText = "" Or Not IsNumeric(sysQtyText) Then
                        resultText = "SYS_QTY_MISSING"
                        actionText = "REVIEW"
                        noteText = noteText & "; Sys Qty khong hop le."
                    ElseIf sysBin = "" Then
                        resultText = "SYS_BIN_MISSING"
                        actionText = "REVIEW"
                        noteText = noteText & "; Sys Bin khong hop le."
                    ElseIf Not IsNumeric(availableText) Or Not IsNumeric(totalText) Then
                        resultText = "POSTING_QTY_MISSING"
                        actionText = "REVIEW"
                        noteText = noteText & "; B Available/F Total Stk khong hop le."
                    Else
                        sysQty = RT_Number(sysQtyText)
                        available = RT_Number(availableText)
                        totalStock = RT_Number(totalText)
                        binDiff = (StrComp(actualBinResolved, sysBin, vbTextCompare) <> 0)
                        batchDiff = (StrComp(actualBatch, sysBatch, vbTextCompare) <> 0)
                        qtyDiff = (Abs(actualQty - sysQty) > tol)
                        stockAvailable = (actualQty <= available + tol)
                        stockTotal = (actualQty <= totalStock + tol)

                        noteText = noteText & "; Available=" & availableText & _
                            "; TotalStk=" & totalText
                        If qtyDiff Then
                            noteText = noteText & _
                                "; Qty diff: PRDO dung OB Qty, POST dung Need Qty."
                        End If

                        If Not stockTotal Then
                            resultText = "NOT_ENOUGH_STOCK"
                            actionText = "REVIEW"
                            holdText = "NO"
                        Else
                            If stockAvailable Then
                                resultText = "READY"
                                holdText = "NO"
                            Else
                                resultText = "NEED_MON"
                                holdText = "YES"
                            End If

                            If binDiff Or batchDiff Then
                                postNeeded = True
                                If qtyDiff Then
                                    If batchDiff Then
                                        postStep = "POST_CHANGE_BATCH_BIN_QTY"
                                    Else
                                        postStep = "POST_CHANGE_BIN_QTY"
                                    End If
                                ElseIf batchDiff Then
                                    postStep = "POST_CHANGE_BATCH_BIN"
                                Else
                                    postStep = "CHANGE_BIN"
                                End If
                            End If

                            If qtyDiff Then
                                If postNeeded Then
                                    actionText = postStep & " -> PRDO_CHANGE_QTY -> WT_PROCESS -> CF"
                                Else
                                    actionText = "CHANGE_QTY"
                                End If
                            ElseIf postNeeded Then
                                actionText = postStep
                            Else
                                actionText = "CF_ONLY"
                            End If

                            If Not stockAvailable Then
                                actionText = "RELEASE_OTHER_WT -> " & actionText
                            End If
                        End If

                        If postNeeded Then
                            noteText = noteText & "; G:I prefilled for /SCWM/POST."
                        Else
                            noteText = noteText & "; G:I left blank because no POST step."
                        End If
                    End If

                    CF_WritePostingCheck wsPost, postRow, ob, whNo, _
                        sysQtyText, actualQtyText, sysBin, actualBinResolved, _
                        sysBatch, actualBatch, needQtyText, availableText, holdText, _
                        actionText, resultText, noteText, postNeeded
                End If
            End If
        End If

WriteWHResult:
        wsWH.Cells(whRow, 15).Value2 = resultText
        If actualQtyText <> "" And IsNumeric(actualQtyText) Then
            wsWH.Cells(whRow, 16).Value2 = RT_Number(actualQtyText)
        Else
            wsWH.Cells(whRow, 16).Value2 = ""
        End If
        wsWH.Cells(whRow, 17).Value2 = "CHECKED"
        wsWH.Cells(whRow, 18).Value2 = actionText
        If resultText = "READY" Then
            wsWH.Cells(whRow, 19).Value2 = ""
            readyCount = readyCount + 1
        ElseIf resultText = "NEED_MON" Then
            wsWH.Cells(whRow, 19).Value2 = noteText
            needMonCount = needMonCount + 1
        ElseIf resultText = "DONE" Or resultText = "USER_SKIP" Then
            wsWH.Cells(whRow, 19).Value2 = ""
        ElseIf InStr(1, resultText, "REVIEW", vbTextCompare) > 0 Or _
               resultText = "MULTI_BIN" Or resultText = "NO_BIN_IN_AREA" Then
            wsWH.Cells(whRow, 19).Value2 = noteText
            reviewCount = reviewCount + 1
        Else
            wsWH.Cells(whRow, 19).Value2 = noteText
            failCount = failCount + 1
        End If

        RT_Log "CHECK_ACTUAL", ob, doc, resultText, _
            "WH=" & whNo & "; WT=" & wtNo & "; Action=" & actionText & _
            "; " & noteText

NextWHRow:
    Next whRow

    RT_SetStatus "TIME", Format$(RT_Elapsed(started), "0.000") & " s"
    RT_SetStatus "STAGE", "CHECK DONE"

    If checked = 0 Then
        MsgBox "Khong co dong nao co du lieu Actual de CHECK.", _
            vbInformation, "CHECK"
    Else
        MsgBox "CHECK hoan tat." & vbCrLf & _
            "Rows: " & CStr(checked) & _
            " | READY: " & CStr(readyCount) & _
            " | NEED_MON: " & CStr(needMonCount) & _
            " | REVIEW: " & CStr(reviewCount) & _
            " | FAIL: " & CStr(failCount) & vbCrLf & _
            "Xem Action/Check Result/Note trong POSTING.", _
            vbInformation, "CHECK"
    End If
    Exit Sub

Failed:
    RT_SetStatus "STAGE", "CHECK FAIL"
    RT_SetStatus "TIME", Format$(RT_Elapsed(started), "0.000") & " s"
    RT_Log "CHECK_ACTUAL", ob, doc, "FAIL", _
        "Err " & CStr(Err.Number) & ": " & Err.Description
    MsgBox "CHECK failed" & vbCrLf & _
        "Err " & CStr(Err.Number) & ": " & Err.Description, _
        vbExclamation, "CHECK"
End Sub
Private Sub CF_WritePostingCheck(ByVal ws As Worksheet, ByVal rowNo As Long, _
                                 ByVal ob As String, ByVal whNo As String, _
                                 ByVal sysQtyText As String, ByVal actualQtyText As String, _
                                 ByVal sysBin As String, ByVal actualBin As String, _
                                 ByVal sysBatch As String, ByVal actualBatch As String, _
                                 ByVal needQtyText As String, ByVal availableText As String, _
                                 ByVal holdText As String, ByVal actionText As String, _
                                 ByVal resultText As String, ByVal noteText As String, _
                                 ByVal postNeeded As Boolean)
    If ws Is Nothing Then Exit Sub
    If rowNo < 2 Then Exit Sub

    If postNeeded Then
        ws.Cells(rowNo, 7).Value2 = actualBin
        If IsNumeric(actualQtyText) Then
            ws.Cells(rowNo, 8).Value2 = RT_Number(actualQtyText)
        Else
            ws.Cells(rowNo, 8).Value2 = actualQtyText
        End If
        ws.Cells(rowNo, 9).Value2 = actualBatch
    Else
        ws.Cells(rowNo, 7).Value2 = ""
        ws.Cells(rowNo, 8).Value2 = ""
        ws.Cells(rowNo, 9).Value2 = ""
    End If

    ws.Cells(rowNo, 10).Value2 = ob
    ws.Cells(rowNo, 11).Value2 = whNo

    If IsNumeric(sysQtyText) Then
        ws.Cells(rowNo, 12).Value2 = RT_Number(sysQtyText)
    Else
        ws.Cells(rowNo, 12).Value2 = sysQtyText
    End If
    If IsNumeric(actualQtyText) Then
        ws.Cells(rowNo, 13).Value2 = RT_Number(actualQtyText)
    Else
        ws.Cells(rowNo, 13).Value2 = actualQtyText
    End If

    ws.Cells(rowNo, 14).Value2 = sysBin
    ws.Cells(rowNo, 15).Value2 = actualBin
    ws.Cells(rowNo, 16).Value2 = sysBatch
    ws.Cells(rowNo, 17).Value2 = actualBatch

    If IsNumeric(needQtyText) Then
        ws.Cells(rowNo, 18).Value2 = RT_Number(needQtyText)
    Else
        ws.Cells(rowNo, 18).Value2 = needQtyText
    End If
    If IsNumeric(availableText) Then
        ws.Cells(rowNo, 19).Value2 = RT_Number(availableText)
    Else
        ws.Cells(rowNo, 19).Value2 = availableText
    End If

    ws.Cells(rowNo, 20).Value2 = holdText
    ws.Cells(rowNo, 21).Value2 = actionText
    ws.Cells(rowNo, 22).Value2 = resultText
    ws.Cells(rowNo, 23).Value2 = noteText
End Sub

Private Function CF_NormalizeAreaPrefix(ByVal areaText As String) As String
    Dim areaKey As String
    areaKey = UCase$(Trim$(areaText))
    If areaKey = "" Then Exit Function

    If Left$(areaKey, 5) = "APFL-" Then
        CF_NormalizeAreaPrefix = areaKey
    ElseIf areaKey = "P" Then
        CF_NormalizeAreaPrefix = "APFL-P"
    ElseIf areaKey = "ZF" Then
        CF_NormalizeAreaPrefix = "APFL-ZF"
    ElseIf areaKey = "HR" Then
        CF_NormalizeAreaPrefix = "APFL-HR"
    Else
        CF_NormalizeAreaPrefix = areaKey
    End If
End Function

Private Function CF_BinInArea(ByVal binText As String, ByVal areaText As String) As Boolean
    Dim prefixText As String, binKey As String
    prefixText = CF_NormalizeAreaPrefix(areaText)
    binKey = UCase$(Trim$(binText))
    If prefixText = "" Or binKey = "" Then Exit Function
    CF_BinInArea = (Left$(binKey, Len(prefixText)) = prefixText)
End Function


