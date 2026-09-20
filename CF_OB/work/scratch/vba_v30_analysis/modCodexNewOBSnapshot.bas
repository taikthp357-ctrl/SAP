Attribute VB_Name = "modCodexNewOBSnapshot"
Option Explicit

' Bounded SAP read + local snapshot repair for the currently entered OB.
' It performs only MON query/navigation.  It never creates, changes, saves,
' cancels, posts, confirms, or triggers CF in SAP.

Private Const CNS_NODE_KEY As String = "N000000010"
Private Const CNS_UPPER_ID As String = "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell"
Private Const CNS_WT_ID As String = "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell"
Private Const CNS_ERP_ID As String = "wnd[1]/usr/txtS_ERP_H-LOW"
Private Const CNS_DOC_ID As String = "wnd[1]/usr/txtS_DOCNO-LOW"
Private Const CNS_EXECUTE_ID As String = "wnd[1]/tbar[0]/btn[8]"
Private Const CNS_WT_ACTION As String = "N000000096"

Private Const CNS_TARGET_QTY As Double = 128#
Private Const CNS_TARGET_AREA As String = "ZF"
Private Const CNS_TARGET_BIN_NO As String = "03-01-13"
Private Const CNS_TARGET_BIN As String = "APFL-ZF03-01-13"
Private Const CNS_TARGET_BATCH As String = "260713-07"

Public Function Codex_LoadNewOBSnapshotAndSetTarget() As String
    Dim wsInput As Worksheet, wsActual As Worksheet, wsQueue As Worksheet, wsCache As Worksheet
    Dim ob As String, doc As String, who As String, wt As String, itemNo As String, material As String
    Dim sysBin As String, sysBatch As String, wtStatus As String, sourceKey As String
    Dim waStatus As String, pickStatus As String, giStatus As String, pickPIStatus As String, obState As String
    Dim sysQty As Double, runID As String, rowError As String, snapshotMessage As String
    Dim oldActual As Variant, oldActualRun As Variant, oldActualInfo As Variant, oldQueue As Variant, oldCache As Variant
    Dim backupTaken As Boolean

    On Error GoTo Failed
    Set wsInput = CF_GetInputSheet()
    Set wsActual = ThisWorkbook.Worksheets("DATA_CORE")
    Set wsQueue = ThisWorkbook.Worksheets("OB_QUEUE")
    Set wsCache = ThisWorkbook.Worksheets("OB_WH_CACHE")
    ob = CNS_Text(wsInput.Range("A7").Value2)
    If Not CNS_IsDigits(ob) Then Err.Raise vbObjectError + 30101, , "STOP_INPUT_OB_INVALID: INPUT!A7 must contain one numeric OB."

    If Not CNS_ReadExactLiveMON(ob, doc, who, wt, itemNo, material, sysQty, sysBin, sysBatch, wtStatus, _
                                waStatus, pickStatus, giStatus, pickPIStatus, obState, snapshotMessage) Then
        Err.Raise vbObjectError + 30102, , snapshotMessage
    End If
    If sysQty + 0.000001 < CNS_TARGET_QTY Then _
        Err.Raise vbObjectError + 30103, , "STOP_TARGET_QTY_EXCEEDS_LIVE_WT: live=" & CStr(sysQty) & "; target=" & CStr(CNS_TARGET_QTY)

    ' Capture local state before any range is cleared.  The caller saves only
    ' after the complete post-verification below passes.
    oldActual = wsActual.Range("A5:AD504").Value2
    oldActualRun = wsActual.Range("W1:X2").Value2
    oldActualInfo = wsActual.Range("AF1:AF2").Value2
    oldQueue = wsQueue.Range("A1:X504").Value2
    oldCache = wsCache.Range("A1:Z504").Value2
    backupTaken = True

    runID = Format$(Now, "yyyymmdd_hhnnss") & "_EXACTMON"
    sourceKey = ob & "|" & doc & "|" & itemNo & "|" & wt

    CNS_WriteHeaders wsQueue, wsCache
    wsQueue.Range("A2:X504").ClearContents
    wsCache.Range("A2:Z504").ClearContents
    wsActual.Range("A5:V504").ClearContents
    wsActual.Range("Y5:AD504").ClearContents
    wsActual.Range("Y4:AD4").value = Array("OB State", "WT State", "Remaining Qty", "Process Allowed", "Decision", "Reason")

    CNS_WriteActual wsActual, ob, doc, itemNo, who, wt, material, sysQty, sysBin, sysBatch, wtStatus, sourceKey, obState
    AP_PrepareActualPickingSheet

    ' User input is deliberately separated from live SAP facts: system source
    ' remains in G:I, requested quantity/bin/batch are L/N/V.
    wsActual.Cells(5, 11).Value2 = CNS_TARGET_AREA
    wsActual.Cells(5, 12).Value2 = CNS_TARGET_QTY
    wsActual.Cells(5, 13).Value2 = CNS_TARGET_BIN_NO
    wsActual.Cells(5, 14).ClearContents
    wsActual.Cells(5, 22).NumberFormat = "@"
    wsActual.Cells(5, 22).Value2 = CNS_TARGET_BATCH
    rowError = ""
    AP_NormalizeRow wsActual, 5, True, rowError
    If rowError <> "" Then Err.Raise vbObjectError + 30104, , "STOP_TARGET_BIN_INVALID: " & rowError
    If StrComp(CNS_Text(wsActual.Cells(5, 14).Value2), CNS_TARGET_BIN, vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 30105, , "STOP_TARGET_BIN_READBACK: " & CNS_Text(wsActual.Cells(5, 14).Value2)
    If Abs(CNS_Number(wsActual.Cells(5, 12).Value2) - CNS_TARGET_QTY) > 0.000001 Then _
        Err.Raise vbObjectError + 30106, , "STOP_TARGET_QTY_READBACK"
    If StrComp(CNS_Text(wsActual.Cells(5, 22).Value2), CNS_TARGET_BATCH, vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 30107, , "STOP_TARGET_BATCH_READBACK"

    CNS_WriteQueue wsQueue, ob, runID, waStatus, pickStatus, giStatus, pickPIStatus, obState
    CNS_WriteCache wsCache, ob, doc, itemNo, who, wt, material, sysQty, sysBin, sysBatch, wtStatus, sourceKey, runID, obState
    wsActual.Range("W1").Value2 = "WH Run ID"
    wsActual.Range("W2").Value2 = runID
    wsActual.Range("X1").Value2 = "WH Actual Sync Rows"
    wsActual.Range("X2").Value2 = 1
    wsActual.Range("AF1").Value2 = "Exact Load Result"
    wsActual.Range("AF2").Value2 = "PASS: verified live MON snapshot; target prepared; CF remains user handoff."

    snapshotMessage = ""
    If Not CF_CheckPlanner_ValidateSnapshot(snapshotMessage) Then _
        Err.Raise vbObjectError + 30108, , "STOP_SNAPSHOT_VALIDATE: " & snapshotMessage

    If StrComp(CNS_Text(wsQueue.Cells(2, 3).Value2), ob, vbTextCompare) <> 0 Or _
       StrComp(CNS_Text(wsQueue.Cells(2, 4).Value2), ob, vbTextCompare) <> 0 Or _
       StrComp(UCase$(CNS_Text(wsQueue.Cells(2, 7).Value2)), "PASS", vbTextCompare) <> 0 Or _
       StrComp(CNS_Text(wsCache.Cells(2, 18).Value2), sourceKey, vbTextCompare) <> 0 Or _
       StrComp(CNS_Text(wsActual.Cells(5, 20).Value2), sourceKey, vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 30109, , "STOP_LOCAL_SNAPSHOT_POSTVERIFY"

    Codex_LoadNewOBSnapshotAndSetTarget = "PASS|OB=" & ob & "|DOC=" & doc & "|WT=" & wt & _
        "|SYS_QTY=" & CStr(sysQty) & "|SYS_BIN=" & sysBin & "|SYS_BATCH=" & sysBatch & _
        "|TARGET_QTY=" & CStr(CNS_TARGET_QTY) & "|TARGET_BIN=" & CNS_TARGET_BIN & _
        "|TARGET_BATCH=" & CNS_TARGET_BATCH & "|RUN=" & runID & "|CF=USER_HANDOFF"
    Exit Function

Failed:
    If backupTaken Then CNS_RestoreLocalState wsActual, wsQueue, wsCache, oldActual, oldActualRun, oldActualInfo, oldQueue, oldCache
    Codex_LoadNewOBSnapshotAndSetTarget = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function CNS_ReadExactLiveMON(ByVal expectedOB As String, _
                                      ByRef doc As String, ByRef who As String, ByRef wt As String, _
                                      ByRef itemNo As String, ByRef material As String, ByRef sysQty As Double, _
                                      ByRef sysBin As String, ByRef sysBatch As String, ByRef wtStatus As String, _
                                      ByRef waStatus As String, ByRef pickStatus As String, ByRef giStatus As String, _
                                      ByRef pickPIStatus As String, ByRef obState As String, ByRef messageText As String) As Boolean
    Dim gui As Object, sapApp As Object, con As Object, ses As Object, tree As Object, modal As Object
    Dim erpCtl As Object, docCtl As Object, executeBtn As Object, upperGrid As Object, wtGrid As Object
    Dim i As Long, r As Long, n As Long, docHits As Long, upperRow As Long, wtHits As Long
    Dim tcode As String, rowDoc As String, currentDoc As String, currentWho As String
    Dim rowItem As String, rowWT As String, rowMat As String, rowStatus As String, rowDest As String, rowProc As String
    Dim rowQty As Double

    On Error GoTo Failed
    messageText = ""
    Set gui = GetObject("SAPGUI"): Set sapApp = gui.GetScriptingEngine: Set con = sapApp.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = ""
        On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(Trim$(tcode), "/", "")) = "SCWMMON" Then Exit For
        Set ses = Nothing
    Next i
    If ses Is Nothing Then messageText = "STOP_MON_SESSION_NOT_FOUND": Exit Function
    If CBool(ses.busy) Then messageText = "STOP_MON_SESSION_BUSY": Exit Function

    Set erpCtl = ses.FindById(CNS_ERP_ID, False)
    If erpCtl Is Nothing Then
        Set modal = ses.FindById("wnd[1]", False)
        Set docCtl = ses.FindById(CNS_DOC_ID, False)
        If Not modal Is Nothing Then
            If docCtl Is Nothing Then
                messageText = "STOP_UNEXPECTED_MON_MODAL"
                Exit Function
            End If
            modal.sendVKey 12
            If Not CNS_WaitIdle(ses, 8#) Then messageText = "STOP_MON_CLOSE_STALE_DIALOG_TIMEOUT": Exit Function
        End If
        Set tree = ses.FindById("wnd[0]/usr/shell/shellcont[0]/shell", False)
        If tree Is Nothing Then messageText = "STOP_MON_TREE_NOT_FOUND": Exit Function
        CallByName tree, "DoubleClickNode", VbMethod, CNS_NODE_KEY
        If Not CNS_WaitIdle(ses, 12#) Then messageText = "STOP_MON_NODE_TIMEOUT": Exit Function
        Set erpCtl = ses.FindById(CNS_ERP_ID, False)
    End If
    If erpCtl Is Nothing Then messageText = "STOP_MON_ERP_FIELD_NOT_FOUND": Exit Function
    Set docCtl = ses.FindById(CNS_DOC_ID, False)
    If docCtl Is Nothing Then messageText = "STOP_MON_DOC_FIELD_NOT_FOUND": Exit Function

    ' The document field must be empty so this is proved by the OB / LE
    ' Delivery criterion alone, not by a leftover ODO filter.
    docCtl.text = ""
    If CNS_Text(docCtl.text) <> "" Then messageText = "STOP_MON_DOC_CLEAR_READBACK": Exit Function
    erpCtl.text = expectedOB
    If StrComp(CNS_Text(erpCtl.text), expectedOB, vbTextCompare) <> 0 Then
        messageText = "STOP_MON_ERP_READBACK: expected=" & expectedOB & "; actual=" & CNS_Text(erpCtl.text)
        Exit Function
    End If
    Set executeBtn = ses.FindById(CNS_EXECUTE_ID, False)
    If executeBtn Is Nothing Then messageText = "STOP_MON_EXECUTE_NOT_FOUND": Exit Function
    executeBtn.Press
    If Not CNS_WaitIdle(ses, 30#) Then messageText = "STOP_MON_EXECUTE_TIMEOUT": Exit Function
    If Not CNS_WaitForGrid(ses, CNS_UPPER_ID, 20#) Then messageText = "STOP_MON_UPPER_GRID_TIMEOUT": Exit Function
    Set upperGrid = ses.FindById(CNS_UPPER_ID, False)
    If upperGrid Is Nothing Then messageText = "STOP_MON_UPPER_GRID_NOT_FOUND": Exit Function

    n = CLng(upperGrid.rowCount)
    upperRow = -1
    For r = 0 To n - 1
        rowDoc = CNS_Cell(upperGrid, r, "DOCNO_H")
        If rowDoc <> "" Then
            docHits = docHits + 1
            upperRow = r
            doc = rowDoc
        End If
    Next r
    If docHits <> 1 Or doc = "" Then
        messageText = "STOP_MON_FILTERED_DOC_COUNT=" & CStr(docHits)
        Exit Function
    End If
    waStatus = CNS_Cell(upperGrid, upperRow, "STATUS_WM_ACTIVITY")
    pickStatus = CNS_Cell(upperGrid, upperRow, "STATUS_PICKING")
    giStatus = CNS_Cell(upperGrid, upperRow, "STATUS_GI")
    pickPIStatus = CNS_Cell(upperGrid, upperRow, "STATUS_PICKING_PLAN")
    If waStatus = "" Or pickStatus = "" Or giStatus = "" Or pickPIStatus = "" Then
        messageText = "STOP_OB_STATE_NOT_PROVEN: blank MON header status"
        Exit Function
    End If
    If UCase$(waStatus) = "COMPLETED" And UCase$(pickStatus) = "COMPLETED" And _
       UCase$(giStatus) = "COMPLETED" And UCase$(pickPIStatus) = "COMPLETED" Then
        messageText = "STOP_OB_DONE: all MON header statuses are Completed"
        Exit Function
    End If
    obState = "OB_OPEN_OR_PARTIAL"

    upperGrid.selectedRows = CStr(upperRow)
    upperGrid.SetCurrentCell upperRow, "DOCNO_H"
    upperGrid.PressToolbarButton CNS_WT_ACTION
    If Not CNS_WaitIdle(ses, 30#) Then messageText = "STOP_MON_WT_ACTION_TIMEOUT": Exit Function
    If Not CNS_WaitForGrid(ses, CNS_WT_ID, 20#) Then messageText = "STOP_MON_WT_GRID_TIMEOUT": Exit Function
    Set wtGrid = ses.FindById(CNS_WT_ID, False)
    If wtGrid Is Nothing Then messageText = "STOP_MON_WT_GRID_NOT_FOUND": Exit Function

    n = CLng(wtGrid.rowCount)
    currentDoc = "": currentWho = ""
    For r = 0 To n - 1
        rowDoc = CNS_Cell(wtGrid, r, "DOCNO")
        If rowDoc <> "" Then
            If StrComp(rowDoc, currentDoc, vbTextCompare) <> 0 Then currentWho = ""
            currentDoc = rowDoc
        End If
        If CNS_Cell(wtGrid, r, "WHO") <> "" Then currentWho = CNS_Cell(wtGrid, r, "WHO")
        rowItem = CNS_Cell(wtGrid, r, "ITEMNO")
        rowWT = CNS_Cell(wtGrid, r, "TANUM")
        rowMat = CNS_Cell(wtGrid, r, "MATNR")
        If StrComp(currentDoc, doc, vbTextCompare) = 0 And (rowItem <> "" Or rowWT <> "" Or rowMat <> "") Then
            wtHits = wtHits + 1
            If wtHits > 1 Then messageText = "STOP_UNEXPECTED_WT_COUNT=" & CStr(wtHits): Exit Function
            who = currentWho: wt = rowWT: itemNo = rowItem: material = rowMat
            sysBin = CNS_Cell(wtGrid, r, "VLPLA")
            sysBatch = CNS_Cell(wtGrid, r, "CHARG")
            wtStatus = CNS_Cell(wtGrid, r, "TOSTAT")
            rowDest = CNS_Cell(wtGrid, r, "NLPLA")
            rowProc = CNS_Cell(wtGrid, r, "PROCTY")
            rowQty = CNS_Number(CNS_Cell(wtGrid, r, "VSOLM"))
            If who = "" Or wt = "" Or itemNo = "" Or material = "" Or sysBin = "" Or sysBatch = "" Or rowQty <= 0# Then
                messageText = "STOP_WT_IDENTITY_INCOMPLETE"
                Exit Function
            End If
            If wtStatus <> "" Then messageText = "STOP_WT_NOT_ACTIVE: status=" & wtStatus: Exit Function
            If StrComp(rowDest, "GI-ZONE", vbTextCompare) <> 0 Then messageText = "STOP_WT_DESTINATION=" & rowDest: Exit Function
            If StrComp(rowProc, "2010", vbTextCompare) <> 0 Then messageText = "STOP_WT_PROCTY=" & rowProc: Exit Function
            sysQty = rowQty
        End If
    Next r
    If wtHits <> 1 Then messageText = "STOP_WT_COUNT=" & CStr(wtHits): Exit Function

    CNS_ReadExactLiveMON = True
    Exit Function
Failed:
    messageText = "STOP_MON_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub CNS_WriteActual(ByVal ws As Worksheet, ByVal ob As String, ByVal doc As String, _
                            ByVal itemNo As String, ByVal who As String, ByVal wt As String, _
                            ByVal material As String, ByVal sysQty As Double, ByVal sysBin As String, _
                            ByVal sysBatch As String, ByVal wtStatus As String, ByVal sourceKey As String, _
                            ByVal obState As String)
    ws.Cells(5, 1).Value2 = ob
    ws.Cells(5, 2).Value2 = doc
    ws.Cells(5, 3).Value2 = itemNo
    ws.Cells(5, 4).Value2 = who
    ws.Cells(5, 5).Value2 = wt
    ws.Cells(5, 6).Value2 = material
    ws.Cells(5, 7).Value2 = sysQty
    ws.Cells(5, 8).Value2 = sysBin
    ws.Cells(5, 9).NumberFormat = "@"
    ws.Cells(5, 9).Value2 = sysBatch
    ws.Cells(5, 10).Value2 = wtStatus
    ws.Cells(5, 12).Value2 = sysQty
    ws.Cells(5, 14).Value2 = sysBin
    ws.Cells(5, 17).Value2 = "LOADED"
    ws.Cells(5, 18).Value2 = "MON_LOADED_EXACT"
    ws.Cells(5, 20).Value2 = sourceKey
    ws.Cells(5, 25).Value2 = obState
    ws.Cells(5, 26).Value2 = "WT_ACTIVE"
    ws.Cells(5, 28).Value2 = "TRUE"
    ws.Cells(5, 30).Value2 = "Exact MON snapshot; active WT with blank status."
End Sub

Private Sub CNS_WriteQueue(ByVal ws As Worksheet, ByVal ob As String, ByVal runID As String, _
                           ByVal waStatus As String, ByVal pickStatus As String, ByVal giStatus As String, _
                           ByVal pickPIStatus As String, ByVal obState As String)
    ws.Cells(2, 1).Value2 = 1
    ws.Cells(2, 2).Value2 = 7
    ws.Cells(2, 3).Value2 = ob
    ws.Cells(2, 4).Value2 = ob
    ws.Cells(2, 6).Value2 = 1
    ws.Cells(2, 7).Value2 = "PASS"
    ws.Cells(2, 8).Value2 = "EXACT_CURRENT_MON"
    ws.Cells(2, 9).Value2 = Now
    ws.Cells(2, 10).Value2 = Now
    ws.Cells(2, 11).Value2 = "TRUE"
    ws.Cells(2, 12).Value2 = "YES"
    ws.Cells(2, 13).Value2 = 0
    ws.Cells(2, 14).Value2 = 0
    ws.Cells(2, 15).Value2 = 0
    ws.Cells(2, 16).Value2 = 0
    ws.Cells(2, 17).Value2 = "RequestedOB=ReturnedOB; exact verified MON filter."
    ws.Cells(2, 18).Value2 = runID
    ws.Cells(2, 19).Value2 = waStatus
    ws.Cells(2, 20).Value2 = pickStatus
    ws.Cells(2, 21).Value2 = giStatus
    ws.Cells(2, 22).Value2 = pickPIStatus
    ws.Cells(2, 23).Value2 = obState
    ws.Cells(2, 24).Value2 = "Header completion is not 4/4."
End Sub

Private Sub CNS_WriteCache(ByVal ws As Worksheet, ByVal ob As String, ByVal doc As String, _
                           ByVal itemNo As String, ByVal who As String, ByVal wt As String, _
                           ByVal material As String, ByVal sysQty As Double, ByVal sysBin As String, _
                           ByVal sysBatch As String, ByVal wtStatus As String, ByVal sourceKey As String, _
                           ByVal runID As String, ByVal obState As String)
    ws.Cells(2, 1).Value2 = 1
    ws.Cells(2, 2).Value2 = ob
    ws.Cells(2, 3).Value2 = ob
    ws.Cells(2, 4).Value2 = doc
    ws.Cells(2, 5).Value2 = itemNo
    ws.Cells(2, 6).Value2 = who
    ws.Cells(2, 7).Value2 = wt
    ws.Cells(2, 8).Value2 = material
    ws.Cells(2, 9).Value2 = sysQty
    ws.Cells(2, 10).Value2 = sysBin
    ws.Cells(2, 11).NumberFormat = "@"
    ws.Cells(2, 11).Value2 = sysBatch
    ws.Cells(2, 12).Value2 = wtStatus
    ws.Cells(2, 13).Value2 = sysQty
    ws.Cells(2, 14).Value2 = sysBin
    ws.Cells(2, 16).Value2 = "PASS"
    ws.Cells(2, 17).Value2 = "RequestedOB=ReturnedOB; exact verified MON filter."
    ws.Cells(2, 18).Value2 = sourceKey
    ws.Cells(2, 19).Value2 = Now
    ws.Cells(2, 20).Value2 = runID
    ws.Cells(2, 21).Value2 = obState
    ws.Cells(2, 22).Value2 = "WT_ACTIVE"
    ws.Cells(2, 24).Value2 = "TRUE"
    ws.Cells(2, 26).Value2 = "Exact live MON snapshot."
End Sub

Private Sub CNS_WriteHeaders(ByVal wsQueue As Worksheet, ByVal wsCache As Worksheet)
    wsQueue.Range("A1:X1").value = Array("Sequence", "Input Row", "Requested OB", "Returned OB", "Note", _
        "WH Rows", "Status", "Load Result", "Started", "Finished", "Process Allowed", _
        "Read Only", "SAP Write", "Cancel", "Save", "CF", "Reason", "Run ID", _
        "Warehouse Activity Status", "Picking Status", "Goods Issue Status", "Picking PI Status", _
        "OB State", "OB State Reason")
    wsCache.Range("A1:Z1").value = Array("Sequence", "Requested OB", "Returned OB", "Document", _
        "Item", "WH", "WT", "Material", "Sys Qty", "Sys Bin", "Sys Batch", "WT Status", _
        "Actual Qty", "Actual Bin", "Destination Bin", "Load Result", "Reason", "Source Key", _
        "Snapshot Time", "Load Run ID", "OB State", "WT State", "Remaining Qty", _
        "Process Allowed", "Decision", "Reason")
End Sub

Private Sub CNS_RestoreLocalState(ByVal wsActual As Worksheet, ByVal wsQueue As Worksheet, ByVal wsCache As Worksheet, _
                                  ByVal oldActual As Variant, ByVal oldActualRun As Variant, ByVal oldActualInfo As Variant, _
                                  ByVal oldQueue As Variant, ByVal oldCache As Variant)
    On Error Resume Next
    wsActual.Range("A5:AD504").Value2 = oldActual
    wsActual.Range("W1:X2").Value2 = oldActualRun
    wsActual.Range("AF1:AF2").Value2 = oldActualInfo
    wsQueue.Range("A1:X504").Value2 = oldQueue
    wsCache.Range("A1:Z504").Value2 = oldCache
    On Error GoTo 0
End Sub

Private Function CNS_WaitIdle(ByVal ses As Object, ByVal timeoutSeconds As Double) As Boolean
    Dim started As Double
    started = Timer
    Do
        DoEvents
        On Error Resume Next
        If Not CBool(ses.busy) Then CNS_WaitIdle = True: Exit Function
        On Error GoTo 0
        If CNS_Elapsed(started) >= timeoutSeconds Then Exit Do
    Loop
End Function

Private Function CNS_WaitForGrid(ByVal ses As Object, ByVal controlID As String, ByVal timeoutSeconds As Double) As Boolean
    Dim started As Double, ctl As Object
    started = Timer
    Do
        DoEvents
        Set ctl = ses.FindById(controlID, False)
        If Not ctl Is Nothing Then
            On Error Resume Next
            If CLng(ctl.rowCount) >= 0 Then CNS_WaitForGrid = True: Exit Function
            On Error GoTo 0
        End If
        If CNS_Elapsed(started) >= timeoutSeconds Then Exit Do
    Loop
End Function

Private Function CNS_Elapsed(ByVal started As Double) As Double
    CNS_Elapsed = Timer - started
    If CNS_Elapsed < 0# Then CNS_Elapsed = CNS_Elapsed + 86400#
End Function

Private Function CNS_Cell(ByVal grid As Object, ByVal rowNo As Long, ByVal columnName As String) As String
    On Error Resume Next
    CNS_Cell = Trim$(CStr(grid.GetCellValue(rowNo, columnName) & ""))
End Function

Private Function CNS_Text(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    CNS_Text = Trim$(CStr(valueIn))
End Function

Private Function CNS_Number(ByVal valueIn As Variant) As Double
    Dim textIn As String
    textIn = Replace(CNS_Text(valueIn), ",", "")
    If IsNumeric(textIn) Then CNS_Number = CDbl(textIn)
End Function

Private Function CNS_IsDigits(ByVal textIn As String) As Boolean
    Dim i As Long, ch As String
    textIn = Trim$(textIn)
    If Len(textIn) < 5 Then Exit Function
    For i = 1 To Len(textIn)
        ch = Mid$(textIn, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    CNS_IsDigits = True
End Function

