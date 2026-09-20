Attribute VB_Name = "modMultiPostingLoader"
Option Explicit

' PHASE 2 ONLY - MULTI-OB POSTING LOAD PROOF.
' Input comes only from the validated Phase 1 WH snapshot.

Public gMPOST_UseOrderQuery As Boolean
Public gMPOST_LastLoadOK As Boolean
Public gMPOST_LastLoadError As String
Public gMPOST_LastRows As Long
Public gMPOST_ExpectedRows As Long
Public gMPOST_SAPExpectedRows As Long
Public gMPOST_CapturedRawRows As Long
Public gMPOST_ValidStockRows As Long
Public gMPOST_CaptureComplete As Boolean
Public gMPOST_GridRowCount As Long
Public gMPOST_ClassifiedDataRows As Long
Public gMPOST_AuxRows As Long
Public gMPOST_UnknownRows As Long
Public gMPOST_CaptureProof As String
Public gMPOST_LastProductCount As Long
Public gMPOST_CheckMaterialFilter As Object
Public gMPOST_LastRunOK As Boolean
Public gMPOST_LastRunMessage As String

Private Const MPOST_QUEUE As String = "OB_QUEUE"
Private Const MPOST_WH_CACHE As String = "OB_WH_CACHE"
Private Const MPOST_ORDER As String = "ORDER_QUERY"
Private Const MPOST_RAW As String = "POSTING_RAW"
Private Const MPOST_POSTING As String = "POSTING"
Private Const MPOST_SUMMARY As String = "POSTING_SUMMARY"
Private Const MPOST_CACHE As String = "POSTING_CACHE"
Private Const MPOST_INPUT_FIRST As Long = 7

Public Sub MPOST_Setup()
    Dim wsSummary As Worksheet, wsCache As Worksheet
    Set wsSummary = MPOST_EnsureSheet(MPOST_SUMMARY, False)
    Set wsCache = MPOST_EnsureSheet(MPOST_CACHE, False)
    MPOST_WriteSummaryHeader wsSummary
    If MPOST_Text(wsCache.Cells(1, 1).Value2) = "" Then
        wsCache.Range("A1:S1").value = Array("Load Run ID", "Raw Row", "Material", _
            "Available Qty", "Batch", "Source Bin", "TST", "Total Stock", "Dest Bin", _
            "QtyPostChg", "Dest.Batch", "ST", "GR Date", "GR Time", "Proc. Type", _
            "AvlQtyBUoM", "Allocated Qty", "Source", "Snapshot Result")
        wsCache.rows(1).Font.Bold = True
    End If
End Sub

Public Sub MPOST_LoadPosting()
    Dim runID As String, messageText As String
    gMPOST_LastRunOK = False
    gMPOST_LastRunMessage = ""
    MPOST_ResetCaptureProof
    If Not SAPRun_RequireRuntime("POSTING_LOAD", _
        "POST_GRID_ID,POST_PRODUCT_COL,POST_AVAILABLE_QTY_COL,POST_BATCH_COL,POST_SOURCE_BIN_COL,POST_TST_COL,POST_TOTAL_STOCK_COL,POST_DEST_BIN_COL,POST_QTY_CHG_COL,POST_DEST_BATCH_COL,POST_RESULT_COUNT_ID,POST_RESULT_COUNT_PATTERN,POST_ADVANCED_TOGGLE_ID,POST_ADVANCED_RESET_ID,POST_ADVANCED_START_ID,POST_PRODUCT_MULTI_ID,POST_OWNER_ID,POST_PRODUCT_POPUP_CLOSE_ID,POST_PRODUCT_POPUP_CLEAR_ID,POST_PRODUCT_PASTE_ID,POST_PRODUCT_POPUP_OK_ID", messageText) Then
        gMPOST_LastRunMessage = messageText
        MPOST_WriteStop messageText, runID
        If Not gRT_Silent Then MsgBox messageText, vbExclamation, "PHASE 2"
        Exit Sub
    End If
    If Not MPOST_HasExistingPostSession(messageText) Then
        gMPOST_LastRunMessage = messageText
        MPOST_WriteStop messageText, runID
        If Not gRT_Silent Then MsgBox messageText, vbExclamation, "PHASE 2"
        Exit Sub
    End If
    If Not MPOST_BuildMaterialList(runID, messageText) Then
        gMPOST_LastRunMessage = messageText
        MPOST_WriteStop messageText, runID
        If Not gRT_Silent Then MsgBox messageText, vbExclamation, "PHASE 2"
        Exit Sub
    End If
    If Not MPOST_ValidateWHActualRunID(runID, messageText) Then
        gMPOST_LastRunMessage = messageText
        MPOST_WriteStop messageText, runID
        If Not gRT_Silent Then MsgBox messageText, vbExclamation, "PHASE 2"
        Exit Sub
    End If
    MPOST_ResetOutputs
    If Not MPOST_LoadPostingOnce(runID, messageText) Then
        gMPOST_LastRunMessage = messageText
        MPOST_WriteStop messageText, runID
        If Not gRT_Silent Then MsgBox messageText, vbExclamation, "PHASE 2"
        Exit Sub
    End If
    If Not MPOST_NormalizeRows(runID, messageText) Then
        gMPOST_LastRunMessage = messageText
        MPOST_WriteStop messageText, runID
        If Not gRT_Silent Then MsgBox messageText, vbExclamation, "PHASE 2"
        Exit Sub
    End If
    If Not MPOST_ValidateCoverage(runID, messageText) Then
        gMPOST_LastRunMessage = messageText
        MPOST_WriteStop messageText, runID
        If Not gRT_Silent Then MsgBox messageText, vbExclamation, "PHASE 2"
        Exit Sub
    End If
    MPOST_FinalizeSnapshot runID, messageText
    gMPOST_LastRunOK = True
    gMPOST_LastRunMessage = messageText
    If Not gRT_Silent Then MsgBox messageText, vbInformation, "PHASE 2"
End Sub

' Read-only session discovery for CHECK / technical Posting load.
' It must never create a new SAP tab: CHECK is allowed to reuse /SCWM/POST
' only when the user already has that transaction open.
Public Function MPOST_HasExistingPostSession(ByRef messageText As String) As Boolean
    Dim guiAuto As Object, app As Object, connections As Object, sessions As Object
    Dim conn As Object, ses As Object
    Dim ci As Long, si As Long, connCount As Long, sesCount As Long
    Dim tcode As String

    On Error GoTo Failed
    messageText = ""
    On Error Resume Next
    Set guiAuto = GetObject("SAPGUI")
    If Not guiAuto Is Nothing Then Set app = guiAuto.GetScriptingEngine
    On Error GoTo Failed
    If app Is Nothing Then
        messageText = "STOP_POST_SESSION_REQUIRED: SAP GUI Scripting is not available; open an idle /SCWM/POST session."
        Exit Function
    End If

    On Error Resume Next
    Set connections = app.children
    connCount = CLng(connections.count)
    On Error GoTo Failed
    For ci = 0 To connCount - 1
        Set conn = Nothing
        On Error Resume Next
        Set conn = connections.ElementAt(ci)
        On Error GoTo Failed
        If Not conn Is Nothing Then
            Set sessions = Nothing
            On Error Resume Next
            Set sessions = conn.children
            sesCount = CLng(sessions.count)
            On Error GoTo Failed
            For si = 0 To sesCount - 1
                Set ses = Nothing
                On Error Resume Next
                Set ses = sessions.ElementAt(si)
                If Not ses Is Nothing Then
                    If Not ses.busy Then tcode = UCase$(Trim$(CStr(ses.Info.Transaction)))
                End If
                On Error GoTo Failed
                If Replace(tcode, "/", "") = "SCWMPOST" Then
                    MPOST_HasExistingPostSession = True
                    messageText = "POST_SESSION_REUSE_READY"
                    Exit Function
                End If
                tcode = ""
            Next si
        End If
    Next ci

    messageText = "STOP_POST_SESSION_REQUIRED: Open an idle /SCWM/POST session, then run CHECK again."
    Exit Function

Failed:
    messageText = "STOP_POST_SESSION_REQUIRED: Existing /SCWM/POST session check failed. Err " & _
                  CStr(Err.Number) & ": " & Err.Description
End Function

Private Function MPOST_ValidateWHActualRunID(ByVal runID As String, _
                                             ByRef messageText As String) As Boolean
    Dim ws As Worksheet, actualRunID As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    actualRunID = MPOST_Text(ws.Range("W2").Value2)
    If actualRunID = "" Then
        messageText = "STOP_WH_ACTUAL_NOT_SYNCED: run TAI WH and complete cache-to-actual sync before TAI POSTING."
        Exit Function
    End If
    If StrComp(actualRunID, runID, vbTextCompare) <> 0 Then
        messageText = "POSTING_STALE: WH_ACTUAL RunID=" & actualRunID & _
                      "; Phase1 RunID=" & runID & ". TAI WH again before TAI POSTING."
        Exit Function
    End If
    MPOST_ValidateWHActualRunID = True
    Exit Function
Failed:
    messageText = "STOP_WH_ACTUAL_RUNID_CHECK_FAIL: " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Sub CF_Phase2_LoadPostingProof()
    MPOST_LoadPosting
End Sub

Public Function MPOST_BuildMaterialList(ByRef runID As String, _
                                        ByRef messageText As String) As Boolean
    Dim wsQ As Worksheet, wsWH As Worksheet, wsOrder As Worksheet
    Dim passSeq As Object, passOB As Object, materialOBs As Object
    Dim lastQ As Long, lastWH As Long, r As Long, seq As String
    Dim ob As String, returnedOB As String, material As String, obState As String, wtState As String
    Dim obSet As Object
    Dim rowRunID As String, queueRunID As String, sourceText As String
    Dim outRow As Long, k As Variant, removeKeys As Collection, removeKey As Variant

    On Error GoTo Failed
    runID = ""
    messageText = ""
    Set wsQ = ThisWorkbook.Worksheets(MPOST_QUEUE)
    Set wsWH = ThisWorkbook.Worksheets(MPOST_WH_CACHE)
    Set wsOrder = MPOST_EnsureSheet(MPOST_ORDER, True)
    wsOrder.Cells.ClearContents
    wsOrder.Range("A1:F1").value = Array("Product", "Qty", "Tram", _
        "Phase1 Seq", "Phase1 OB Count", "Source OBs")

    Set passSeq = CreateObject("Scripting.Dictionary")
    passSeq.CompareMode = vbTextCompare
    Set passOB = CreateObject("Scripting.Dictionary")
    passOB.CompareMode = vbTextCompare
    Set materialOBs = CreateObject("Scripting.Dictionary")
    materialOBs.CompareMode = vbTextCompare

    lastQ = MPOST_LastRow(wsQ, 1, 2)
    For r = 2 To lastQ
        If UCase$(MPOST_Text(wsQ.Cells(r, 7).Value2)) = "PASS" Then
            seq = MPOST_Text(wsQ.Cells(r, 1).Value2)
            ob = MPOST_Text(wsQ.Cells(r, 3).Value2)
            queueRunID = MPOST_Text(wsQ.Cells(r, 18).Value2)
            If seq <> "" And ob <> "" And queueRunID <> "" Then
                If runID = "" Then runID = queueRunID
                If StrComp(runID, queueRunID, vbTextCompare) <> 0 Then
                    messageText = "STOP_PHASE1_RUNID_MISMATCH: OB_QUEUE contains multiple Phase 1 Run IDs."
                    Exit Function
                End If
                passSeq(seq) = True
                passOB(ob) = True
            End If
        End If
    Next r
    If passSeq.count = 0 Or runID = "" Then
        messageText = "STOP_PHASE1_NOT_READY: no Phase 1 PASS row with LoadRunID was found."
        Exit Function
    End If

    lastWH = MPOST_LastRow(wsWH, 1, 2)
    For r = 2 To lastWH
        seq = MPOST_Text(wsWH.Cells(r, 1).Value2)
        ob = MPOST_Text(wsWH.Cells(r, 2).Value2)
        returnedOB = MPOST_Text(wsWH.Cells(r, 3).Value2)
        material = MPOST_Text(wsWH.Cells(r, 8).Value2)
        rowRunID = MPOST_Text(wsWH.Cells(r, 20).Value2)
        obState = UCase$(MPOST_Text(wsWH.Cells(r, 21).Value2))
        wtState = UCase$(MPOST_Text(wsWH.Cells(r, 22).Value2))
        If passSeq.Exists(seq) And StrComp(ob, returnedOB, vbTextCompare) = 0 _
           And StrComp(rowRunID, runID, vbTextCompare) = 0 And material <> "" _
           And obState = "OB_OPEN_OR_PARTIAL" _
           And (wtState = "WT_ACTIVE" Or wtState = "WT_CANCELLED_REMAINING") Then
            If Not materialOBs.Exists(material) Then
                Set obSet = CreateObject("Scripting.Dictionary")
                obSet.CompareMode = vbTextCompare
                materialOBs.Add material, obSet
            End If
            Set obSet = materialOBs(material)
            obSet(ob) = True
        End If
    Next r
    ' CHECK may provide an exception-only Material filter. Keep the normal
    ' Phase 2 path unchanged when the filter is Nothing; otherwise remove
    ' materials that do not participate in a Posting-required plan row.
    If Not gMPOST_CheckMaterialFilter Is Nothing Then
        Set removeKeys = New Collection
        For Each k In materialOBs.keys
            If Not gMPOST_CheckMaterialFilter.Exists(CStr(k)) Then removeKeys.Add CStr(k)
        Next k
        For Each removeKey In removeKeys
            materialOBs.Remove CStr(removeKey)
        Next removeKey
    End If

    If materialOBs.count = 0 Then
        messageText = "STOP_NO_PHASE1_MATERIAL: Phase 1 PASS snapshot has no valid Material rows."
        Exit Function
    End If

    outRow = 2
    For Each k In materialOBs.keys
        sourceText = MPOST_DictionaryKeys(materialOBs(k))
        wsOrder.Cells(outRow, 1).Value2 = CStr(k)
        wsOrder.Cells(outRow, 2).Value2 = 1
        wsOrder.Cells(outRow, 3).Value2 = "CF_OB_PHASE2"
        wsOrder.Cells(outRow, 4).Value2 = ""
        Set obSet = materialOBs(k)
        wsOrder.Cells(outRow, 5).Value2 = obSet.count
        wsOrder.Cells(outRow, 6).Value2 = sourceText
        outRow = outRow + 1
    Next k
    wsOrder.Range("A1:F1").Font.Bold = True
    wsOrder.Visible = xlSheetVeryHidden

    messageText = "Material list ready; requested=" & CStr(materialOBs.count) & _
                  "; unique=" & CStr(materialOBs.count) & "; LoadRunID=" & runID
    MPOST_BuildMaterialList = True
    Exit Function

Failed:
    messageText = "PHASE2_BUILD_MATERIAL_FAIL: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function MPOST_LoadPostingOnce(ByVal runID As String, _
                                       ByRef messageText As String) As Boolean
    Dim wsIn As Worksheet, wsRaw As Worksheet, oldA7 As Variant
    Dim oldSilent As Boolean, oldBatchRunID As String

    On Error GoTo Failed
    Set wsIn = CF_GetInputSheet()
    Set wsRaw = ThisWorkbook.Worksheets(MPOST_RAW)
    oldA7 = wsIn.Range("A7").Value2
    oldSilent = gRT_Silent
    oldBatchRunID = gRT_ActiveBatchRunID
    gRT_ActiveBatchRunID = runID
    wsRaw.Range("A2:Q20000").ClearContents
    wsIn.Range("A7").Value2 = "PHASE2_MULTI_MATERIAL"
    gMPOST_UseOrderQuery = True
    gRT_Silent = True
    CF_LoadPostingFromWH
    gRT_Silent = oldSilent
    gMPOST_UseOrderQuery = False
    gRT_ActiveBatchRunID = oldBatchRunID
    wsIn.Range("A7").Value2 = oldA7

    If Not gMPOST_LastLoadOK Then
        messageText = "POST_READ_ERROR: " & gMPOST_LastLoadError
        Exit Function
    End If
    messageText = "Posting loaded; rows=" & CStr(gMPOST_LastRows) & _
                  "; LoadRunID=" & runID
    MPOST_LoadPostingOnce = True
    Exit Function

Failed:
    On Error Resume Next
    gRT_Silent = oldSilent
    gMPOST_UseOrderQuery = False
    gRT_ActiveBatchRunID = oldBatchRunID
    wsIn.Range("A7").Value2 = oldA7
    messageText = "POST_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function MPOST_NormalizeRows(ByVal runID As String, _
                                     ByRef messageText As String) As Boolean
    Dim wsRaw As Worksheet, wsCache As Worksheet
    Dim lastRow As Long, r As Long, outRow As Long, material As String

    On Error GoTo Failed
    Set wsRaw = ThisWorkbook.Worksheets(MPOST_RAW)
    Set wsCache = MPOST_EnsureSheet(MPOST_CACHE, False)
    wsCache.Cells.ClearContents
    wsCache.Range("A1:S1").value = Array("Load Run ID", "Raw Row", "Material", _
        "Available Qty", "Batch", "Source Bin", "TST", "Total Stock", "Dest Bin", _
        "QtyPostChg", "Dest.Batch", "ST", "GR Date", "GR Time", "Proc. Type", _
        "AvlQtyBUoM", "Allocated Qty", "Source", "Snapshot Result")
    wsCache.rows(1).Font.Bold = True

    lastRow = MPOST_LastRow(wsRaw, 1, 2)
    outRow = 2
    For r = 2 To lastRow
        material = MPOST_Text(wsRaw.Cells(r, 1).Value2)
        If material <> "" And UCase$(MPOST_Text(wsRaw.Cells(r, 15).Value2)) = "VALID_STOCK" Then
            wsCache.Cells(outRow, 1).Value2 = runID
            wsCache.Cells(outRow, 2).Value2 = r
            wsCache.Cells(outRow, 3).Value2 = material
            wsCache.Cells(outRow, 4).Value2 = wsRaw.Cells(r, 2).Value2
            wsCache.Cells(outRow, 5).Value2 = MPOST_Text(wsRaw.Cells(r, 3).Value2)
            wsCache.Cells(outRow, 6).Value2 = MPOST_Text(wsRaw.Cells(r, 4).Value2)
            wsCache.Cells(outRow, 7).Value2 = MPOST_Text(wsRaw.Cells(r, 5).Value2)
            wsCache.Cells(outRow, 8).Value2 = wsRaw.Cells(r, 6).Value2
            wsCache.Cells(outRow, 9).Value2 = MPOST_Text(wsRaw.Cells(r, 7).Value2)
            wsCache.Cells(outRow, 10).Value2 = wsRaw.Cells(r, 8).Value2
            wsCache.Cells(outRow, 11).Value2 = MPOST_Text(wsRaw.Cells(r, 9).Value2)
            wsCache.Cells(outRow, 12).Value2 = MPOST_Text(wsRaw.Cells(r, 10).Value2)
            wsCache.Cells(outRow, 13).Value2 = MPOST_Text(wsRaw.Cells(r, 11).Value2)
            wsCache.Cells(outRow, 14).Value2 = MPOST_Text(wsRaw.Cells(r, 12).Value2)
            wsCache.Cells(outRow, 15).Value2 = MPOST_Text(wsRaw.Cells(r, 13).Value2)
            wsCache.Cells(outRow, 16).Value2 = wsRaw.Cells(r, 14).Value2
            wsCache.Cells(outRow, 17).Value2 = wsRaw.Cells(r, 17).Value2
            wsCache.Cells(outRow, 18).Value2 = "POSTING_RAW"
            wsCache.Cells(outRow, 19).Value2 = "VALID_STOCK"
            outRow = outRow + 1
        End If
    Next r
    wsCache.Range("A1:S1").EntireColumn.AutoFit
    messageText = "Posting snapshot normalized; rows=" & CStr(outRow - 2)
    MPOST_NormalizeRows = True
    Exit Function

Failed:
    messageText = "POST_NORMALIZE_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function MPOST_ValidateCoverage(ByVal runID As String, _
                                        ByRef messageText As String) As Boolean
    Dim wsOrder As Worksheet, wsRaw As Worksheet, wsSummary As Worksheet
    Dim requested As Object, rowCounts As Object, sourceOBs As Object
    Dim lastOrder As Long, lastRaw As Long, r As Long, outRow As Long
    Dim material As String, returnedMaterial As String, reasonText As String
    Dim k As Variant, countRows As Long, staleCount As Long, noStockCount As Long, errorText As String
    Dim dataReady As Boolean

    On Error GoTo Failed
    Set wsOrder = ThisWorkbook.Worksheets(MPOST_ORDER)
    Set wsRaw = ThisWorkbook.Worksheets(MPOST_RAW)
    Set wsSummary = MPOST_EnsureSheet(MPOST_SUMMARY, False)
    wsSummary.Cells.ClearContents
    MPOST_WriteSummaryHeader wsSummary

    Set requested = CreateObject("Scripting.Dictionary")
    requested.CompareMode = vbTextCompare
    Set rowCounts = CreateObject("Scripting.Dictionary")
    rowCounts.CompareMode = vbTextCompare
    Set sourceOBs = CreateObject("Scripting.Dictionary")
    sourceOBs.CompareMode = vbTextCompare

    lastOrder = MPOST_LastRow(wsOrder, 1, 2)
    For r = 2 To lastOrder
        material = MPOST_Text(wsOrder.Cells(r, 1).Value2)
        If material <> "" Then
            requested(material) = True
            rowCounts(material) = 0
            sourceOBs(material) = MPOST_Text(wsOrder.Cells(r, 6).Value2)
        End If
    Next r
    If requested.count = 0 Then
        messageText = "POST_COVERAGE_ERROR: requested Material set is empty."
        Exit Function
    End If

    lastRaw = MPOST_LastRow(wsRaw, 1, 2)
    For r = 2 To lastRaw
        returnedMaterial = MPOST_Text(wsRaw.Cells(r, 1).Value2)
        If UCase$(MPOST_Text(wsRaw.Cells(r, 15).Value2)) = "VALID_STOCK" And returnedMaterial <> "" Then
            If requested.Exists(returnedMaterial) Then
                rowCounts(returnedMaterial) = CLng(rowCounts(returnedMaterial)) + 1
            Else
                staleCount = staleCount + 1
            End If
        End If
    Next r
    If gMPOST_CapturedRawRows <> MPOST_CountRawRows(wsRaw) Then
        errorText = "STOP_POST_CAPTURE_INCOMPLETE: raw row count differs from captured SAP row count."
    ElseIf staleCount > 0 Then
        errorText = "POST_STALE_DATA: returned Material outside requested set; rows=" & CStr(staleCount)
    End If

    dataReady = (errorText = "" And gMPOST_LastLoadOK And _
                  gMPOST_LastRows >= 0 And gMPOST_CaptureComplete And _
                  gMPOST_SAPExpectedRows >= 0 And gMPOST_CapturedRawRows >= 0 And _
                  gMPOST_SAPExpectedRows = gMPOST_CapturedRawRows And _
                  gMPOST_CapturedRawRows = MPOST_CountRawRows(wsRaw) And runID <> "")
    outRow = 2
    For Each k In requested.keys
        countRows = CLng(rowCounts(k))
        If errorText <> "" Then
            reasonText = errorText
            material = "POST_STALE_DATA"
            dataReady = False
        ElseIf Not gMPOST_LastLoadOK Then
            reasonText = gMPOST_LastLoadError
            material = "POST_READ_ERROR"
            dataReady = False
        ElseIf countRows = 0 Then
            If Not gMPOST_CaptureComplete Or gMPOST_CapturedRawRows <> gMPOST_SAPExpectedRows Then
                reasonText = "POST_CAPTURE_INCOMPLETE: SAP result was not fully captured; no-stock classification is blocked."
                material = "POST_CAPTURE_INCOMPLETE"
                noStockCount = noStockCount + 1
            Else
                reasonText = "SAP_RETURNED_NONE_AFTER_FULL_CAPTURE"
                material = "POST_NO_STOCK_ROW"
            End If
        ElseIf countRows > 1 Then
            reasonText = "MULTI_SOURCE"
            material = "POST_FOUND"
        Else
            reasonText = "OK"
            material = "POST_FOUND"
        End If
        wsSummary.Cells(outRow, 1).Value2 = CStr(k)
        wsSummary.Cells(outRow, 2).Value2 = "YES"
        wsSummary.Cells(outRow, 3).Value2 = countRows
        wsSummary.Cells(outRow, 4).Value2 = material
        wsSummary.Cells(outRow, 5).Value2 = reasonText
        wsSummary.Cells(outRow, 6).Value2 = runID
        wsSummary.Cells(outRow, 7).Value2 = MPOST_CountCsv(sourceOBs(k))
        wsSummary.Cells(outRow, 8).Value2 = sourceOBs(k)
        wsSummary.Cells(outRow, 9).Value2 = IIf(dataReady, "TRUE", "FALSE")
        wsSummary.Cells(outRow, 10).Value2 = 0
        wsSummary.Cells(outRow, 11).Value2 = 0
        wsSummary.Cells(outRow, 12).Value2 = 0
        wsSummary.Cells(outRow, 13).Value2 = gMPOST_SAPExpectedRows
        wsSummary.Cells(outRow, 14).Value2 = gMPOST_CapturedRawRows
        wsSummary.Cells(outRow, 15).Value2 = gMPOST_ValidStockRows
        wsSummary.Cells(outRow, 16).Value2 = gMPOST_GridRowCount
        wsSummary.Cells(outRow, 17).Value2 = gMPOST_ClassifiedDataRows
        wsSummary.Cells(outRow, 18).Value2 = gMPOST_AuxRows
        wsSummary.Cells(outRow, 19).Value2 = gMPOST_UnknownRows
        wsSummary.Cells(outRow, 20).Value2 = gMPOST_CaptureProof
        outRow = outRow + 1
    Next k

    wsSummary.Range("A1:T1").EntireColumn.AutoFit
    If dataReady Then
        messageText = "POSTING_DATA_READY=TRUE; RequestedMaterialCount=" & _
                      CStr(requested.count) & "; ReturnedRows=" & CStr(gMPOST_LastRows) & _
                      "; SAPExpectedRows=" & CStr(gMPOST_SAPExpectedRows) & _
                      "; CapturedRawRows=" & CStr(gMPOST_CapturedRawRows) & _
                      "; ValidStockRows=" & CStr(gMPOST_ValidStockRows) & _
                      "; NoStockMaterials=" & CStr(noStockCount) & "; LoadRunID=" & runID
        MPOST_ValidateCoverage = True
    Else
        messageText = "POSTING_DATA_READY=FALSE; " & IIf(errorText <> "", errorText, _
                      "STOP_POST_CAPTURE_INCOMPLETE: independent expected/captured/raw validation failed")
    End If
    Exit Function

Failed:
    messageText = "POST_COVERAGE_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub MPOST_FinalizeSnapshot(ByVal runID As String, ByRef messageText As String)
    RT_Log "PHASE2_POSTING", "", "", "PASS", messageText & _
           "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
End Sub

Private Sub MPOST_ResetOutputs()
    Dim ws As Worksheet
    Set ws = MPOST_EnsureSheet(MPOST_POSTING, False)
    ws.Range("A2:AE2000").ClearContents
    Set ws = MPOST_EnsureSheet(MPOST_RAW, True)
    ws.Range("A2:Q20000").ClearContents
    Set ws = MPOST_EnsureSheet(MPOST_SUMMARY, False)
    ws.Cells.ClearContents
    Set ws = MPOST_EnsureSheet(MPOST_CACHE, False)
    ws.Cells.ClearContents
End Sub

Private Sub MPOST_WriteStop(ByVal messageText As String, ByVal runID As String)
    Dim ws As Worksheet
    Set ws = MPOST_EnsureSheet(MPOST_SUMMARY, False)
    MPOST_WriteSummaryHeader ws
    ws.Cells(2, 1).Value2 = "PHASE2_STOP"
    ws.Cells(2, 4).Value2 = "STOP"
    ws.Cells(2, 5).Value2 = messageText
    ws.Cells(2, 6).Value2 = runID
    ws.Cells(2, 9).Value2 = "FALSE"
    ws.Cells(2, 10).Value2 = 0
    ws.Cells(2, 11).Value2 = 0
    ws.Cells(2, 12).Value2 = 0
    ws.Cells(2, 13).Value2 = gMPOST_SAPExpectedRows
    ws.Cells(2, 14).Value2 = gMPOST_CapturedRawRows
    ws.Cells(2, 15).Value2 = gMPOST_ValidStockRows
    ws.Cells(2, 16).Value2 = gMPOST_GridRowCount
    ws.Cells(2, 17).Value2 = gMPOST_ClassifiedDataRows
    ws.Cells(2, 18).Value2 = gMPOST_AuxRows
    ws.Cells(2, 19).Value2 = gMPOST_UnknownRows
    ws.Cells(2, 20).Value2 = gMPOST_CaptureProof
    RT_Log "PHASE2_POSTING", "", "", "FAIL", messageText & _
           "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
End Sub

Private Sub MPOST_WriteSummaryHeader(ByVal ws As Worksheet)
    ws.Range("A1:T1").value = Array("Material", "Requested", "Rows Found", "Result", _
        "Reason", "Load Run ID", "Phase1 OB Count", "Source OBs", "Data Ready", _
        "SAP Write", "Cancel", "Save", "SAP Expected Rows", "Captured Raw Rows", _
        "Valid Stock Rows", "Grid RowCount", "Classified Data Rows", "Aux Rows", _
        "Unknown Rows", "Capture Proof")
    ws.rows(1).Font.Bold = True
End Sub

Private Sub MPOST_ResetCaptureProof()
    gMPOST_SAPExpectedRows = -1
    gMPOST_CapturedRawRows = 0
    gMPOST_ValidStockRows = 0
    gMPOST_CaptureComplete = False
    gMPOST_GridRowCount = -1
    gMPOST_ClassifiedDataRows = 0
    gMPOST_AuxRows = 0
    gMPOST_UnknownRows = 0
    gMPOST_CaptureProof = "NOT_RUN"
End Sub

Public Sub MPOST_CompileProbe()
    If False Then MPOST_LoadPosting
End Sub

Private Function MPOST_CountRawRows(ByVal wsRaw As Worksheet) As Long
    Dim lastRow As Long, r As Long
    lastRow = MPOST_LastRow(wsRaw, 1, 2)
    For r = 2 To lastRow
        If MPOST_Text(wsRaw.Cells(r, 15).Value2) <> "" Then MPOST_CountRawRows = MPOST_CountRawRows + 1
    Next r
End Function

Private Function MPOST_EnsureSheet(ByVal sheetName As String, _
                                   ByVal veryHidden As Boolean) As Worksheet
    On Error Resume Next
    Set MPOST_EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If MPOST_EnsureSheet Is Nothing Then
        Set MPOST_EnsureSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        MPOST_EnsureSheet.Name = sheetName
    End If
    MPOST_EnsureSheet.Visible = IIf(veryHidden, xlSheetVeryHidden, xlSheetVisible)
End Function

Private Function MPOST_LastRow(ByVal ws As Worksheet, ByVal col As Long, _
                               ByVal firstRow As Long) As Long
    Dim lastRow As Long
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.rows.count, col).End(xlUp).row
    If lastRow < firstRow Then MPOST_LastRow = firstRow - 1 Else MPOST_LastRow = lastRow
End Function

Private Function MPOST_Text(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then
        MPOST_Text = ""
    Else
        MPOST_Text = Trim$(CStr(valueIn))
    End If
End Function

Private Function MPOST_DictionaryKeys(ByVal d As Object) As String
    Dim k As Variant, s As String
    If d Is Nothing Then Exit Function
    For Each k In d.keys
        If s <> "" Then s = s & ","
        s = s & CStr(k)
    Next k
    MPOST_DictionaryKeys = s
End Function

Private Function MPOST_CountCsv(ByVal valueText As String) As Long
    If Trim$(valueText) = "" Then Exit Function
    MPOST_CountCsv = UBound(Split(valueText, ",")) + 1
End Function







