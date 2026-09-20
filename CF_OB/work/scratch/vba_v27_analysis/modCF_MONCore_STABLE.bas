Attribute VB_Name = "modCF_MONCore_STABLE"
Option Explicit

' Stable MON flow: OB -> Document -> Warehouse Task -> WT/WH rows.
' Dependencies in the target workbook: modCF_UI and modCF_WarehouseTask.

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const RT_VERSION As String = "V0.5 MON ONE OB - STABLE CONNECTION (COM HARDENED)"
Private Const CF_UPPER_ID As String = "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell"
Private Const CF_WT_ID As String = "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell"
Private Const CF_OB_WA_COL As String = "STATUS_WM_ACTIVITY"
Private Const CF_OB_PICKING_COL As String = "STATUS_PICKING"
Private Const CF_OB_GI_COL As String = "STATUS_GI"
Private Const CF_OB_PICKING_PI_COL As String = "STATUS_PICKING_PLAN"
Private Const CF_MON_TCODE As String = "/SCWM/MON"
Private Const CF_POLL_MS As Long = 100
Private Const CF_TIMEOUT_SEC As Long = 12
Private Const CF_MAX_WT_ROWS As Long = 500

Public gRT_StopRequested As Boolean
Public gRT_Running As Boolean
Public gRT_ActiveBatchRunID As String
Public gRT_RunID As String
Public gRT_LastSessionDiagnostic As String
Public gRT_Silent As Boolean
Public gRT_LastLoadError As String
Public gRT_LastLoadSystemFailure As Boolean
Public gRT_LastOBWAStatus As String
Public gRT_LastOBPickingStatus As String
Public gRT_LastOBGIStatus As String
Public gRT_LastOBPickingPIStatus As String
Public gRT_LastOBState As String
Public gRT_LastOBStateReason As String
Private mRT_StageStarted As Double
Private mRT_StageName As String
Private gRT_DOCUMENT_OPEN_COUNT As Long

'============================================================
' PUBLIC ENTRY POINTS - match the existing CF OB buttons
'============================================================
Public Sub RT_InitializeUI()
    On Error Resume Next
    ' Keep the title and visual language supplied by the UI-only template.
    RT_SetStatus "SAP", "CHUA CHAY"
    RT_SetStatus "MON", "CHUA CHAY"
    RT_SetStatus "LOAD", "CHUA CHAY"
    RT_SetStatus "PLAN", "TAI WH -> CHECK"
    RT_SetStatus "STAGE", "IDLE"
    RT_SetStatus "TIME", "0.000 s"
    AP_PrepareActualPickingSheet
    CF_ApplyInputLayout
    On Error GoTo 0
End Sub

Public Sub RT_Preflight()
    Dim ses As Object, upperGrid As Object, wtGrid As Object
    Dim ob As String, msg As String
    On Error GoTo Failed

    If Not SAPRun_RequireRuntime("PREFLIGHT", _
        "MON_RESULT_GRID_ID,MON_WT_GRID_ID,MON_UPPER_OB_COL,MON_UPPER_DOC_COL,MON_WT_DOC_COL,MON_WT_WHO_COL,MON_WT_COL,MON_WT_ITEM_COL,MON_WT_MATERIAL_COL,MON_WT_BATCH_COL,MON_WT_BIN_COL,MON_WT_QTY_COL,MON_WT_STATUS_COL,MON_OB_INPUT_ID,MON_EXECUTE_ID,MON_WT_ACTION_ID", msg) Then
        Err.Raise vbObjectError + 9400, , msg
    End If

    If Not RT_GetSingleInputOB(ob, msg) Then Err.Raise vbObjectError + 9401, , msg
    RT_SetStatus "STAGE", "PREFLIGHT"

    Set ses = RT_FindMONSession()
    If ses Is Nothing Then Err.Raise vbObjectError + 9402, , "MON_SESSION_NOT_FOUND. Mo san /SCWM/MON tren SAP."
    If Not RT_WaitReady(ses, CF_TIMEOUT_SEC) Then Err.Raise vbObjectError + 9403, , "MON_SESSION_BUSY_TIMEOUT"

    Set upperGrid = RT_FindById(ses, RT_ExactID("MON_RESULT_GRID_ID", CF_UPPER_ID))
    Set wtGrid = RT_FindById(ses, RT_ExactID("MON_WT_GRID_ID", CF_WT_ID))
    If upperGrid Is Nothing Then Err.Raise vbObjectError + 9404, , "MON_UPPER_GRID_NOT_FOUND"
    If wtGrid Is Nothing Then Err.Raise vbObjectError + 9405, , "MON_WT_GRID_NOT_FOUND"

    RT_RequireColumn upperGrid, RT_ColumnID("MON_UPPER_OB_COL", "ERP_NUMBER_H")
    RT_RequireColumn upperGrid, RT_ColumnID("MON_UPPER_DOC_COL", "DOCNO_H")
    RT_RequireColumn upperGrid, RT_ColumnID("MON_UPPER_WA_COL", CF_OB_WA_COL)
    RT_RequireColumn upperGrid, RT_ColumnID("MON_UPPER_PICKING_COL", CF_OB_PICKING_COL)
    RT_RequireColumn upperGrid, RT_ColumnID("MON_UPPER_GI_COL", CF_OB_GI_COL)
    RT_RequireColumn upperGrid, RT_ColumnID("MON_UPPER_PICKING_PLAN_COL", CF_OB_PICKING_PI_COL)
    RT_RequireColumn wtGrid, RT_ColumnID("MON_WT_DOC_COL", "DOCNO")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_WT_WHO_COL", "WHO")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_WT_COL", "TANUM")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_WT_ITEM_COL", "ITEMNO")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_MATERIAL_COL", "MATNR")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_BATCH_COL", "CHARG")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_BIN_COL", "VLPLA")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_QTY_COL", "VSOLM")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_STATUS_COL", "TOSTAT")

    RT_SetStatus "SAP", "READY"
    RT_SetStatus "MON", "READY"
    RT_SetStatus "LOAD", "1 OB"
    RT_SetStatus "STAGE", "PREFLIGHT PASS"
    RT_Log "PREFLIGHT", ob, "", "PASS", "MON session + exact technical columns OK"
    MsgBox "PREFLIGHT PASS" & vbCrLf & "OB: " & ob & vbCrLf & "MON va cac technical key deu san sang.", vbInformation, "SAP CF OB"
    RT_SetStatus "STAGE", "IDLE"
    GoTo CleanExit

Failed:
    RT_SetStatus "MON", "ERROR"
    RT_SetStatus "STAGE", "PREFLIGHT FAIL"
    RT_Log "PREFLIGHT", ob, "", "FAIL", "Err " & CStr(Err.Number) & ": " & Err.Description
    MsgBox "PREFLIGHT failed" & vbCrLf & "Err " & CStr(Err.Number) & ": " & Err.Description, vbExclamation, "SAP CF OB"
    RT_SetStatus "STAGE", "IDLE"

CleanExit:
    ' Always release SAP COM references so a stale/dead pointer never survives
    ' into the next call (this is what previously left later calls, even
    ' unrelated ones like CF_SelfTest, throwing Catastrophic failure).
    Set upperGrid = Nothing
    Set wtGrid = Nothing
    Set ses = Nothing
End Sub

Public Sub RT_LoadMON()
    Dim t0 As Double
    Dim ses As Object, upperGrid As Object, wtGrid As Object
    Dim ws As Worksheet
    Dim ob As String, doc As String, msg As String
    Dim upperRow As Long, matchCount As Long
    Dim rowCount As Long, i As Long, used As Long
    Dim rawDoc As String, currentDoc As String, rawWH As String, currentWH As String
    Dim itemText As String, wtText As String, matText As String, binText As String, batchText As String
    Dim statusText As String, caseText As String, sourceKey As String
    Dim loadStage As String, errorNo As Long, errorDesc As String, errorDetail As String
    Dim errorTCode As String, errorScreen As String
    Dim monMapMessage As String, runGuardMessage As String
    Dim obWAStatus As String, obPickingStatus As String
    Dim obGIStatus As String, obPickingPIStatus As String
    Dim obState As String, obStateReason As String
    Dim normalizedWTState As String, normalizedReason As String, normalizedDecision As String
    Dim normalizedProcessAllowed As Boolean, normalizedPostingAllowed As String
    Dim normalizedNeedRemaining As Boolean
    Dim qty As Double
    Dim output(1 To CF_MAX_WT_ROWS, 1 To 20) As Variant
    Dim wtStateOutput(1 To CF_MAX_WT_ROWS) As String
    Dim processOutput(1 To CF_MAX_WT_ROWS) As String
    Dim decisionOutput(1 To CF_MAX_WT_ROWS) As String
    Dim reasonOutput(1 To CF_MAX_WT_ROWS) As String
    Dim preservedActual As Object, preservedVals As Variant

    On Error GoTo Failed
    gRT_DOCUMENT_OPEN_COUNT = 0
    CFWT_ResetNavigationCounters
    If Not SAPRun_RequireRuntime("PHASE1_WH", _
        "MON_RESULT_GRID_ID,MON_WT_GRID_ID,MON_UPPER_OB_COL,MON_UPPER_DOC_COL,MON_WT_DOC_COL,MON_WT_WHO_COL,MON_WT_COL,MON_WT_ITEM_COL,MON_WT_MATERIAL_COL,MON_WT_BATCH_COL,MON_WT_BIN_COL,MON_WT_QTY_COL,MON_WT_STATUS_COL,MON_OB_INPUT_ID,MON_EXECUTE_ID,MON_WT_ACTION_ID", runGuardMessage) Then
        Err.Raise vbObjectError + 9400, , runGuardMessage
    End If
    If gRT_Running Then Err.Raise vbObjectError + 9410, , "TOOL_ALREADY_RUNNING"
    gRT_Running = True
    gRT_StopRequested = False
    gRT_LastLoadError = ""
    gRT_LastLoadSystemFailure = False
    gRT_LastOBWAStatus = ""
    gRT_LastOBPickingStatus = ""
    gRT_LastOBGIStatus = ""
    gRT_LastOBPickingPIStatus = ""
    gRT_LastOBState = "OB_STATE_NOT_PROVEN"
    gRT_LastOBStateReason = "OB_STATE_NOT_PROVEN: header status was not read."
    If Trim$(gRT_ActiveBatchRunID) <> "" Then
        gRT_RunID = gRT_ActiveBatchRunID
    Else
        gRT_RunID = Format$(Now, "yyyymmdd_hhnnss")
    End If
    t0 = RT_StartTimer()
    RT_SetStatus "STAGE", "LOAD MON (180s MAX)"
    loadStage = "INIT"
    RT_SetStatus "STAGE", "LOAD MON"

    loadStage = "READ_INPUT_OB"
    If Not RT_GetSingleInputOB(ob, msg) Then Err.Raise vbObjectError + 9411, , msg

    loadStage = "FIND_MON_SESSION"
    RT_StageStart "FIND_MON_SESSION"
    Set ses = RT_FindMONSession()
    RT_StageRequire ob, doc, "FIND_MON_SESSION", 15#
    RT_StageLog ob, doc, "FIND_MON_SESSION"
    If ses Is Nothing Then
        If gRT_LastSessionDiagnostic = "" Then gRT_LastSessionDiagnostic = "Khong tim thay SAP session co the truy cap."
        Err.Raise vbObjectError + 9412, , "MON_SESSION_NOT_FOUND. " & gRT_LastSessionDiagnostic
    End If
    loadStage = "WAIT_MON_READY"
    RT_StageStart "WAIT_MON_IDLE"
    If Not RT_WaitReady(ses, 15#) Then Err.Raise vbObjectError + 9413, , "STOP_STAGE_TIMEOUT: WAIT_MON_IDLE"
    RT_StageLog ob, doc, "WAIT_MON_IDLE"
    RT_SetStatus "SAP", "READY"

    loadStage = "FIND_MON_GRIDS"
    RT_StageStart "FIND_UPPER_GRID"
    Set upperGrid = RT_FindById(ses, RT_ExactID("MON_RESULT_GRID_ID", CF_UPPER_ID))
    Set wtGrid = RT_FindById(ses, RT_ExactID("MON_WT_GRID_ID", CF_WT_ID))
    RT_StageRequire ob, doc, "FIND_UPPER_GRID", 15#
    RT_StageLog ob, doc, "FIND_UPPER_GRID"

    ' A newly opened /SCWM/MON session has no result grid yet.  An existing
    ' grid can also belong to a prior node/layout and is not evidence for the
    ' OB about to be loaded.  Do not let that stale surface block a fresh
    ' read-only MON filter; the exact contract is enforced after the filter
    ' returns its result grid below.
    If Not upperGrid Is Nothing Then
        loadStage = "DISCARD_EXISTING_MON_CONTEXT"
        RT_Log "LOAD_MON", ob, "", "INFO", "EXISTING_MON_CONTEXT_IGNORED; fresh OB filter required before a WH snapshot is accepted."
        Set upperGrid = Nothing
        Set wtGrid = Nothing
    End If

    loadStage = "FIND_OB_IN_MON"
    RT_StageStart "FIND_OB"
    If Not upperGrid Is Nothing Then upperRow = RT_FindUpperOBRow(upperGrid, ob, matchCount)
    If matchCount = 0 Then
        loadStage = "MAP_MON_LE_DELIVERY"
        monMapMessage = ""
        If Not ModuleSAP.CF_MON_SetLEDelivery(ses, ob, monMapMessage) Then
            Err.Raise vbObjectError + 9415, , "OB_NOT_FOUND_IN_MON: " & ob & "; MON_MAP=" & monMapMessage & _
                      "; GridRows=" & CStr(RT_SafeLong(upperGrid, "RowCount")) & "; Samples=" & RT_UpperGridSample(upperGrid, 8)
        End If
        loadStage = "REFIND_OB_AFTER_MON_FILTER"
        Set upperGrid = Nothing
        Set upperGrid = RT_FindById(ses, RT_ExactID("MON_RESULT_GRID_ID", CF_UPPER_ID))
        If upperGrid Is Nothing Then Err.Raise vbObjectError + 9414, , "MON_UPPER_GRID_NOT_FOUND_AFTER_FILTER"
        loadStage = "CHECK_MON_COLUMNS_AFTER_FILTER"
        RT_RequireUpperLoadColumns upperGrid
        matchCount = 0
        upperRow = RT_FindUpperOBRow(upperGrid, ob, matchCount)
        ' In this MON layout the LE Delivery column can be rendered as
        ' "Not Determined" even though the filtered result row is present.
        ' The selection screen has already filtered by S_ERP_H-LOW, so accept
        ' one non-empty document row as the filtered ODO result and continue
        ' with the Warehouse Task toolbar button.
        If matchCount = 0 Then
            upperRow = RT_FindSingleFilteredResultRow(upperGrid, matchCount)
        End If
    End If
    If matchCount = 0 Then Err.Raise vbObjectError + 9415, , "OB_NOT_FOUND_IN_MON: " & ob & "; GridRows=" & CStr(RT_SafeLong(upperGrid, "RowCount")) & "; Samples=" & RT_UpperGridSample(upperGrid, 8)
    If matchCount > 1 Then Err.Raise vbObjectError + 9416, , "DUPLICATE_OB_IN_MON: " & ob
    RT_StageRequire ob, doc, "FIND_OB", 20#
    RT_StageLog ob, doc, "FIND_OB"

    doc = RT_GridCell(upperGrid, upperRow, RT_ColumnID("MON_UPPER_DOC_COL", "DOCNO_H"))
    If doc = "" Then Err.Raise vbObjectError + 9417, , "DOCUMENT_EMPTY_FOR_OB: " & ob
    gRT_DOCUMENT_OPEN_COUNT = 1
    RT_Log "WT_NAV", ob, doc, "PASS", "DOCUMENT_OPEN_COUNT=1; REOPEN_DOCUMENT=0; DOCUMENT_CONTEXT_LOCKED"
    If Not RT_ReadOBHeaderState(upperGrid, upperRow, obWAStatus, obPickingStatus, _
                                obGIStatus, obPickingPIStatus, obState, obStateReason) Then
        Err.Raise vbObjectError + 9430, , obStateReason
    End If
    gRT_LastOBWAStatus = obWAStatus
    gRT_LastOBPickingStatus = obPickingStatus
    gRT_LastOBGIStatus = obGIStatus
    gRT_LastOBPickingPIStatus = obPickingPIStatus
    gRT_LastOBState = obState
    gRT_LastOBStateReason = obStateReason
    RT_SetStatus "MON", ob
    RT_SetStatus "LOAD", doc

    ' Selecting the OB row and switching to the Warehouse Task tab both trigger
    ' a real SAP screen transition. A previous retry loop was accidentally
    ' configured as 1 To 1 even though its diagnostics claimed two attempts.
    ' Keep one bounded recovery retry, with no new SAP session/tab, so a transient
    ' control-tree rebuild cannot look like an Excel freeze.
    Dim attemptNo As Long, navOK As Boolean, lastErrNum As Long, lastErrDesc As String
    loadStage = "SELECT_OB_AND_WT"
    RT_StageStart "OPEN_WT"
    navOK = False
    For attemptNo = 1 To 2
        On Error GoTo NavAttemptFailed

        If attemptNo > 1 Then
            RT_SetStatus "STAGE", "RETRY NAV (" & attemptNo & "/2)"
            RT_WaitReady ses, 2#
            Set upperGrid = Nothing
        Set upperGrid = RT_FindById(ses, RT_ExactID("MON_RESULT_GRID_ID", CF_UPPER_ID))
            If upperGrid Is Nothing Then Err.Raise vbObjectError + 9414, , "MON_UPPER_GRID_NOT_FOUND (retry)"
        End If

        RT_SelectUpperRow ses, upperGrid, upperRow, "ERP_NUMBER_H"
        If Not RT_WaitReady(ses, CF_TIMEOUT_SEC) Then Err.Raise vbObjectError + 9418, , "MON_BUSY_AFTER_SELECT_OB"

        If Not CFWT_SelectWarehouseTask(ses, CF_TIMEOUT_SEC, True) Then
            Err.Raise vbObjectError + 9419, , "WAREHOUSE_TASK_GRID_NOT_READY. " & CFWT_LastDiagnostic
        End If
        If Not RT_WaitReady(ses, CF_TIMEOUT_SEC) Then Err.Raise vbObjectError + 9419, , "MON_BUSY_AFTER_WAREHOUSE_TASK"

        ' SAP can leave the first Warehouse Task activation on a partially
        ' rendered grid.  Press the same verified WT action once more in the
        ' same OB/document context so the complete row set is materialized.
        ' This is a bounded read/navigation action; it never reopens the OB.
        If Not CFWT_SelectWarehouseTask(ses, CF_TIMEOUT_SEC, True) Then
            Err.Raise vbObjectError + 9419, , "WT_SECOND_CLICK_CONTROL_NOT_FOUND. " & CFWT_LastDiagnostic
        End If
        If Not RT_WaitReady(ses, CF_TIMEOUT_SEC) Then Err.Raise vbObjectError + 9419, , "MON_BUSY_AFTER_WAREHOUSE_TASK_SECOND_CLICK"
        RT_Log "WT_NAV", ob, doc, "PASS", "DOCUMENT_OPEN_COUNT=" & CStr(gRT_DOCUMENT_OPEN_COUNT) & _
               "; WT_CLICK_COUNT=" & CStr(CFWT_WT_ACTION_CLICK_COUNT) & "; REOPEN_DOCUMENT=0"

        ' Important: do not keep the old WT grid object after changing the upper OB.
        ' Holding a reference to a control that SAP has already torn down after a
        ' screen transition is one of the classic causes of Catastrophic failure.
        Set wtGrid = Nothing
        If Not RT_WaitWTContext(ses, doc, wtGrid, CF_TIMEOUT_SEC) Then
            Err.Raise vbObjectError + 9419, , "WT_CONTEXT_NOT_READY. Expected Document=" & doc
        End If
        RT_StageRequire ob, doc, "OPEN_WT", 30#
        RT_StageLog ob, doc, "OPEN_WT"
        navOK = True
        Exit For

NavAttemptFailed:
        lastErrNum = Err.Number
        lastErrDesc = Err.Description
        RT_Log "LOAD_MON_NAV", ob, doc, "RETRY", "Attempt " & attemptNo & " Err " & lastErrNum & ": " & lastErrDesc
        Err.Clear
        On Error GoTo 0
        Set wtGrid = Nothing
        Set upperGrid = Nothing
        If attemptNo < 2 Then Sleep 500
    Next attemptNo

    On Error GoTo Failed
    If Not navOK Then Err.Raise lastErrNum, , lastErrDesc

    loadStage = "CHECK_WT_COLUMNS_AFTER_NAV"
    RT_RequireColumn wtGrid, RT_ColumnID("MON_WT_DOC_COL", "DOCNO")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_WT_WHO_COL", "WHO")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_WT_COL", "TANUM")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_WT_ITEM_COL", "ITEMNO")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_MATERIAL_COL", "MATNR")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_BATCH_COL", "CHARG")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_BIN_COL", "VLPLA")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_QTY_COL", "VSOLM")
    RT_RequireColumn wtGrid, RT_ColumnID("MON_STATUS_COL", "TOSTAT")

    loadStage = "CLEAR_WH_ACTUAL"
    Set ws = ThisWorkbook.Worksheets("WH_ACTUAL")
    Set preservedActual = RT_CaptureActualInput(ws)
    ws.Range("A5:V504").ClearContents
    AP_PrepareActualPickingSheet

    loadStage = "READ_WT_ROWS"
    RT_StageStart "READ_WT_ROWS"
    rowCount = RT_SafeLong(wtGrid, "RowCount")
    currentDoc = ""
    currentWH = ""

    For i = 0 To rowCount - 1
        If RT_Elapsed(t0) >= 180# Then Err.Raise vbObjectError + 9470, , "TIMEOUT_STOP: LOAD_MON exceeded 180 seconds."
        If gRT_StopRequested Then Exit For

        rawDoc = RT_GridCell(wtGrid, i, RT_ColumnID("MON_WT_DOC_COL", "DOCNO"))
        If rawDoc <> "" Then
            If StrComp(rawDoc, currentDoc, vbTextCompare) <> 0 Then
                currentDoc = rawDoc
                currentWH = ""   ' hard reset WH when Document changes
            End If
        End If

        rawWH = RT_GridCell(wtGrid, i, RT_ColumnID("MON_WT_WHO_COL", "WHO"))
        If rawWH <> "" Then currentWH = rawWH

        If StrComp(currentDoc, doc, vbTextCompare) = 0 Then
            itemText = RT_GridCell(wtGrid, i, RT_ColumnID("MON_WT_ITEM_COL", "ITEMNO"))
            wtText = RT_GridCell(wtGrid, i, RT_ColumnID("MON_WT_COL", "TANUM"))
            matText = RT_GridCell(wtGrid, i, RT_ColumnID("MON_MATERIAL_COL", "MATNR"))

            If itemText <> "" Or wtText <> "" Or matText <> "" Then
                If currentWH = "" Then
                    Err.Raise vbObjectError + 9420, , "WH_NOT_RESOLVED at WT row " & CStr(i + 1) & " for Document=" & doc
                End If

                used = used + 1
                If used > CF_MAX_WT_ROWS Then Err.Raise vbObjectError + 9421, , "WT_ROWS_OVER_LIMIT"

                qty = RT_Number(RT_GridCell(wtGrid, i, RT_ColumnID("MON_QTY_COL", "VSOLM")))
                binText = RT_GridCell(wtGrid, i, RT_ColumnID("MON_BIN_COL", "VLPLA"))
                batchText = RT_GridCell(wtGrid, i, RT_ColumnID("MON_BATCH_COL", "CHARG"))
                statusText = UCase$(RT_GridCell(wtGrid, i, RT_ColumnID("MON_STATUS_COL", "TOSTAT")))
                WT_NormalizeState obState, statusText, False, 0#, normalizedWTState, _
                                  normalizedProcessAllowed, normalizedPostingAllowed, _
                                  normalizedNeedRemaining, normalizedDecision, normalizedReason
                If normalizedWTState = WT_STATE_ACTIVE Or normalizedWTState = WT_STATE_CANCELLED_REMAINING Then
                    caseText = "DIRECT_CONFIRM"
                ElseIf normalizedWTState = WT_STATE_DONE Then
                    caseText = "ALREADY_CONFIRMED"
                Else
                    caseText = "REVIEW_STOP"
                End If

                sourceKey = ob & "|" & doc & "|" & itemText & "|" & wtText
                If RT_SourceKeyExists(output, used - 1, sourceKey) Then
                    Err.Raise vbObjectError + 9422, , "DUPLICATE_WT_SOURCE_KEY: " & sourceKey
                End If

                output(used, 1) = ob
                output(used, 2) = doc
                output(used, 3) = itemText
                output(used, 4) = currentWH
                output(used, 5) = wtText
                output(used, 6) = matText
                output(used, 7) = qty
                output(used, 8) = binText
                output(used, 9) = batchText
                output(used, 10) = statusText
                ' Actual Picking contract: K=Area, L=Actual Qty, M=Bin No., N=Actual Bin.
                output(used, 11) = ""
                output(used, 12) = qty
                output(used, 13) = ""
                output(used, 14) = binText
                output(used, 15) = ""
                output(used, 16) = ""
                output(used, 17) = normalizedWTState
                output(used, 18) = "MON_LOADED"
                output(used, 19) = ""
                output(used, 20) = sourceKey
                wtStateOutput(used) = normalizedWTState
                processOutput(used) = IIf(normalizedProcessAllowed, "TRUE", "FALSE")
                decisionOutput(used) = normalizedDecision
                reasonOutput(used) = normalizedReason
            End If
        End If
    Next i

    If used = 0 Then Err.Raise vbObjectError + 9423, , "NO_WT_ROWS_FOR_DOCUMENT: " & doc
    RT_StageRequire ob, doc, "READ_WT_ROWS", 30#
    RT_StageLog ob, doc, "READ_WT_ROWS"

    loadStage = "WRITE_WH_ACTUAL"
    ws.Range("A5:T504").Value2 = output
    ws.Range("Y4:AD504").ClearContents
    ws.Range("Y4:AD4").Value = Array("OB State", "WT State", "Remaining Qty", _
        "Process Allowed", "Decision", "Reason")
    For i = 0 To used - 1
        ws.Cells(i + 5, 25).Value2 = obState
        ws.Cells(i + 5, 26).Value2 = wtStateOutput(i + 1)
        ws.Cells(i + 5, 27).Value2 = ""
        ws.Cells(i + 5, 28).Value2 = processOutput(i + 1)
        ws.Cells(i + 5, 29).Value2 = decisionOutput(i + 1)
        ws.Cells(i + 5, 30).Value2 = reasonOutput(i + 1)
        sourceKey = CStr(output(i + 1, 20) & "")
        If Not preservedActual Is Nothing Then
            If preservedActual.Exists(sourceKey) And obState <> "OB_DONE" _
               And (wtStateOutput(i + 1) = WT_STATE_ACTIVE Or _
                    wtStateOutput(i + 1) = WT_STATE_CANCELLED_PENDING Or _
                    wtStateOutput(i + 1) = WT_STATE_CANCELLED_REMAINING) Then
                preservedVals = preservedActual(sourceKey)
                ws.Cells(i + 5, 11).Value2 = preservedVals(0)
                ws.Cells(i + 5, 12).Value2 = preservedVals(1)
                ws.Cells(i + 5, 13).Value2 = preservedVals(2)
                ws.Cells(i + 5, 14).Value2 = preservedVals(3)
                ws.Cells(i + 5, 21).Value2 = preservedVals(4)
                ws.Cells(i + 5, 22).Value2 = preservedVals(5)
            End If
        End If
        AP_NormalizeRow ws, i + 5, False
    Next i
    RT_SetStatus "LOAD", doc
    RT_SetStatus "PLAN", CStr(used) & " WT"
    RT_SetStatus "TIME", Format$(RT_Elapsed(t0), "0.000") & " s"
    RT_SetStatus "STAGE", "DONE"
    RT_Log "LOAD_MON", ob, doc, "PASS", CStr(used) & " WT rows; OB_STATE=" & obState & _
           "; WA=" & obWAStatus & "; PICKING=" & obPickingStatus & "; GI=" & obGIStatus & _
           "; PICKING_PI=" & obPickingPIStatus & "; WH carry locked by Document; TOTAL_OB_MS=" & _
           Format$(RT_Elapsed(t0) * 1000#, "0") & "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
    If Not gRT_Silent Then
        MsgBox "LOAD MON PASS" & vbCrLf & _
               "OB: " & ob & vbCrLf & _
               "Document: " & doc & vbCrLf & _
               "WT rows: " & CStr(used), vbInformation, "SAP CF OB"
    End If

SafeExit:
    gRT_Running = False
    RT_SetStatus "STAGE", "IDLE"
    ' Release SAP COM references before leaving, always - success or failure.
    Set upperGrid = Nothing
    Set wtGrid = Nothing
    Set ses = Nothing
    Exit Sub

Failed:
    ' Capture the original error before status/log helpers can clear Err.
    errorNo = Err.Number
    errorDesc = Err.Description
    If errorNo = 0 Then errorNo = vbObjectError + 9499
    If Trim$(errorDesc) = "" Then errorDesc = "Unknown VBA/SAP error; Err object was empty."
    On Error Resume Next
    If Not ses Is Nothing Then
        errorTCode = RT_SessionTransaction(ses)
        errorScreen = RT_SafeText(RT_FindById(ses, "wnd[0]"), "Text")
    End If
    On Error GoTo 0
    errorDetail = "Stage=" & loadStage & "; Err=" & CStr(errorNo) & "; Message=" & errorDesc & _
                  "; TOTAL_OB_MS=" & Format$(RT_Elapsed(t0) * 1000#, "0") & _
                  "; OB=" & ob & "; Document=" & doc & _
                  "; TCode=" & errorTCode & "; Screen=" & errorScreen & _
                  "; UpperRow=" & CStr(upperRow) & "; OBMatches=" & CStr(matchCount) & _
                  "; WTRows=" & CStr(rowCount) & "; Used=" & CStr(used) & _
                  "; CurrentDoc=" & currentDoc & "; CurrentWH=" & currentWH & _
                  "; Attempt=" & CStr(attemptNo)
    If i > 0 Then errorDetail = errorDetail & "; WTRow=" & CStr(i + 1)
    If lastErrNum <> 0 Or Trim$(lastErrDesc) <> "" Then
        errorDetail = errorDetail & "; LastNavErr=" & CStr(lastErrNum) & ": " & lastErrDesc
    End If
    RT_StageLog ob, doc, mRT_StageName
    gRT_LastLoadError = errorDetail
    gRT_LastLoadSystemFailure = RT_IsSystemFailureDetail(errorDetail)
    On Error Resume Next
    RT_SetStatus "MON", "ERROR"
    RT_SetStatus "STAGE", "LOAD MON FAIL"
    RT_SetStatus "TIME", Format$(RT_Elapsed(t0), "0.000") & " s"
    RT_Log "LOAD_MON", ob, doc, "FAIL", errorDetail
    If Not gRT_Silent Then
        MsgBox "LOAD MON failed" & vbCrLf & errorDetail & vbCrLf & _
               "Chi tiet da duoc ghi vao RUN_LOG.", vbExclamation, "SAP CF OB"
    End If
    Resume SafeExit
End Sub

Public Function RT_MultiOBStartupGate(ByRef messageText As String) As Boolean
    Dim ses As Object

    On Error GoTo Failed
    messageText = ""

    ' Bootstrap only: find an idle MON session or open /SCWM/MON automatically.
    ' No OB is selected here. WT and upper result grids may appear later
    ' because a fresh MON session starts before Execute has run.
    Set ses = RT_FindMONSession()
    If ses Is Nothing Then
        messageText = "STOP_MON_SESSION_REQUIRED: Could not find or automatically open /SCWM/MON. " & gRT_LastSessionDiagnostic
        GoTo CleanExit
    End If

    If Not RT_WaitReady(ses, 15#) Then
        messageText = "STOP_STAGE_TIMEOUT: WAIT_MON_IDLE exceeded 15 seconds."
        GoTo CleanExit
    End If

    ' Startup establishes only a usable idle MON session.  The currently
    ' visible result grid may be a stale layout from a previous activity, so
    ' it must never veto the next load. RT_LoadMON filters the requested OB
    ' and then verifies every exact technical column on that fresh result.
    RT_MultiOBStartupGate = True

CleanExit:
    Set ses = Nothing
    Exit Function

Failed:
    messageText = "STOP_MON_STARTUP_GATE: " & CStr(Err.Number) & ": " & Err.Description
    Resume CleanExit
End Function

Public Sub RT_LoadMON_Silent()
    Dim oldSilent As Boolean
    On Error GoTo CleanExit
    oldSilent = gRT_Silent
    gRT_Silent = True
    RT_LoadMON
CleanExit:
    gRT_Silent = oldSilent
End Sub

' Compile-only probe for the full read-only loader. It intentionally does not
' connect to SAP or execute the loader.
Public Sub RT_MONCompileProbe()
    If False Then RT_LoadMON
End Sub

Public Sub RT_LoadPosting()
    CF_LoadPosting
End Sub

Public Sub RT_BuildPlan()
    MsgBox "V0.5 chua bat PHAN LUONG. Can PASS MON 1 OB truoc.", vbInformation, "SAP CF OB"
End Sub

Public Sub RT_RunPlan()
    MsgBox "V0.5 READ ONLY. Khong ghi thay doi len SAP.", vbInformation, "SAP CF OB"
End Sub

Public Sub RT_RetryErrors()
    RT_LoadMON
End Sub

Public Sub RT_Stop()
    gRT_StopRequested = True
    RT_SetStatus "STAGE", "STOP REQUESTED"
End Sub

Public Sub RT_OpenLog()
    On Error GoTo Failed
    ThisWorkbook.Worksheets("RUN_LOG").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("RUN_LOG").Activate
    Exit Sub
Failed:
    MsgBox "Khong mo duoc RUN_LOG: " & Err.Description, vbExclamation, "SAP CF OB"
End Sub

' Compatibility entry points retained for older buttons / workbook code.
Public Sub RT_FreezePlan()
    RT_BuildPlan
End Sub

Public Sub RT_RunPlanReadOnly()
    RT_RunPlan
End Sub

Public Sub RT_RetryErrorRows_ReadOnly()
    RT_RetryErrors
End Sub

Public Sub RT_ListSessions()
    RT_Preflight
End Sub

Public Sub RT_MapCurrentScreen()
    RT_Preflight
End Sub

Public Sub RT_SetupButtons()
End Sub

'============================================================
' STABLE SAP BOOTSTRAP - hardened against Catastrophic failure
'============================================================
Private Function RT_IsSystemFailureDetail(ByVal detailText As String) As Boolean
    Dim u As String
    u = UCase$(detailText)
    RT_IsSystemFailureDetail = (InStr(u, "MON_SESSION") > 0 Or _
                                 InStr(u, "SAP_SESSION") > 0 Or _
                                 InStr(u, "STOP_STAGE_TIMEOUT") > 0 Or _
                                 InStr(u, "TIMEOUT_STOP") > 0 Or _
                                 InStr(u, "AUTOMATION") > 0 Or _
                                 InStr(u, "CATASTROPHIC") > 0 Or _
                                 InStr(u, "UNEXPECTED_POPUP") > 0 Or _
                                 InStr(u, "COM_FAILURE") > 0 Or _
                                 InStr(u, "MON_UPPER_GRID_NOT_FOUND") > 0 Or _
                                 InStr(u, "WAREHOUSE_TASK_GRID_NOT_READY") > 0 Or _
                                 InStr(u, "WT_CONTEXT_NOT_READY") > 0)
End Function
Private Sub RT_StageStart(ByVal stageName As String)
    mRT_StageName = stageName
    mRT_StageStarted = RT_StartTimer()
End Sub

Private Sub RT_StageLog(ByVal ob As String, ByVal doc As String, ByVal stageName As String)
    If mRT_StageName <> stageName Then Exit Sub
    RT_Log "GATE0_STAGE", ob, doc, "INFO", stageName & "_MS=" & _
           Format$(RT_Elapsed(mRT_StageStarted) * 1000#, "0") & _
           "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
    mRT_StageName = ""
    mRT_StageStarted = 0#
End Sub

Private Sub RT_StageRequire(ByVal ob As String, ByVal doc As String, _
                            ByVal stageName As String, ByVal seconds As Double)
    If mRT_StageName <> stageName Then Exit Sub
    If RT_Elapsed(mRT_StageStarted) >= seconds Then
        Err.Raise vbObjectError + 9471, , "STOP_STAGE_TIMEOUT: " & stageName & _
                  " exceeded " & CStr(seconds) & " seconds."
    End If
End Sub
Private Function RT_GetSAPApplicationStable(Optional ByRef seedSession As Object) As Object
    Dim sapGuiAuto As Object, sapApp As Object, conn As Object, ses As Object
    Dim connColl As Object, sesColl As Object
    Dim ci As Long, si As Long, retryNo As Long
    Dim connCount As Long, sesCount As Long

    For retryNo = 0 To 0
        Set sapGuiAuto = Nothing
        Set sapApp = Nothing
        On Error Resume Next
        Set sapGuiAuto = GetObject("SAPGUI")
        If Not sapGuiAuto Is Nothing Then Set sapApp = sapGuiAuto.GetScriptingEngine
        Err.Clear
        On Error GoTo 0

        If Not sapApp Is Nothing Then
            Set connColl = RT_SafeChildren(sapApp)
            connCount = RT_CollectionCount(connColl)
            For ci = 0 To connCount - 1
                Set conn = RT_CollectionObjectAt(connColl, ci)
                If Not conn Is Nothing Then
                    Set sesColl = RT_SafeChildren(conn)
                    sesCount = RT_CollectionCount(sesColl)
                    For si = 0 To sesCount - 1
                        Set ses = RT_CollectionObjectAt(sesColl, si)
                        If Not ses Is Nothing Then
                            If Not RT_FindById(ses, "wnd[0]") Is Nothing Then
                                Set seedSession = ses
                                Set RT_GetSAPApplicationStable = sapApp
                                Exit Function
                            End If
                        End If
                        Set ses = Nothing
                    Next si
                    Set sesColl = Nothing
                End If
                Set conn = Nothing
            Next ci
            Set connColl = Nothing
        End If

        DoEvents
        Sleep 500
    Next retryNo
End Function

Private Function RT_FindMONSession(Optional ByVal requireUpperGrid As Boolean = True, _
                                    Optional ByVal allowCreate As Boolean = True, _
                                    Optional ByVal allowCreateInRun As Boolean = False) As Object
    Dim app As Object, seed As Object, conn As Object, ses As Object, candidate As Object
    Dim connColl As Object, sesColl As Object
    Dim ci As Long, si As Long, tcode As String, titleText As String
    Dim upperGrid As Object, wtGrid As Object
    Dim connCount As Long, sesCount As Long, openMsg As String
    Dim busyMonFound As Boolean, monFound As Boolean

    gRT_LastSessionDiagnostic = ""

    ' First reuse an idle MON session. If a MON session already exists but is
    ' busy, fail closed instead of opening another MON and multiplying tabs.
    Set app = RT_GetSAPApplicationStable(seed)
    If Not app Is Nothing Then
        Set connColl = RT_SafeChildren(app)
        connCount = RT_CollectionCount(connColl)
        For ci = 0 To connCount - 1
            Set conn = RT_CollectionObjectAt(connColl, ci)
            If Not conn Is Nothing Then
                Set sesColl = RT_SafeChildren(conn)
                sesCount = RT_CollectionCount(sesColl)
                For si = 0 To sesCount - 1
                    Set ses = RT_CollectionObjectAt(sesColl, si)
                    If Not ses Is Nothing Then
                        tcode = UCase$(Trim$(RT_SessionTransaction(ses)))
                        titleText = UCase$(RT_SafeText(RT_FindById(ses, "wnd[0]"), "Text"))
                        If RT_TcodeEquals(tcode, CF_MON_TCODE) Or InStr(1, titleText, "WAREHOUSE MANAGEMENT MONITOR", vbTextCompare) > 0 Then
                            monFound = True
                            ' Never probe a session that is mid-repaint - this
                            ' is the common trigger for SAP GUI catastrophic
                            ' failures. More importantly, do not create a
                            ' replacement MON while this one is still busy.
                            If RT_SafeBool(ses, "Busy") Then
                                busyMonFound = True
                            Else
                                If Not requireUpperGrid Then
                                    Set RT_FindMONSession = ses
                                    gRT_LastSessionDiagnostic = "MON_EXISTING_IDLE_SESSION_FOUND"
                                    Exit Function
                                ElseIf Not RT_FindById(ses, RT_ExactID("MON_RESULT_GRID_ID", CF_UPPER_ID)) Is Nothing Then
                                    Set RT_FindMONSession = ses
                                    gRT_LastSessionDiagnostic = "MON_EXISTING_IDLE_SESSION_FOUND"
                                    Set ses = Nothing
                                    Set sesColl = Nothing
                                    Set conn = Nothing
                                    Set connColl = Nothing
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                    Set ses = Nothing
                Next si
                Set sesColl = Nothing
            End If
            Set conn = Nothing
        Next ci
        Set connColl = Nothing
    End If

    If busyMonFound Then
        gRT_LastSessionDiagnostic = "MON_EXISTING_SESSION_BUSY: no replacement MON tab will be opened."
        Exit Function
    End If
    If monFound Then
        gRT_LastSessionDiagnostic = "MON_EXISTING_SESSION_NOT_READY: no replacement MON tab will be opened."
        Exit Function
    End If
    If Not allowCreate Then
        gRT_LastSessionDiagnostic = "MON_EXISTING_IDLE_SESSION_NOT_FOUND"
        Exit Function
    End If

    If SAPRun_IsRunMode() And Not allowCreateInRun Then
        gRT_LastSessionDiagnostic = "MON_EXISTING_IDLE_SESSION_NOT_FOUND: RUN does not open a new SAP tab."
        Exit Function
    End If

    ' In RUN, the only permitted creation is one bounded dedicated reader tab.
    ' It is never the user's executor session and is reused on later calls.
    If SAPRun_IsRunMode() And allowCreateInRun Then
        Set candidate = RT_OpenDedicatedMONSession(seed, openMsg)
        If candidate Is Nothing Then
            gRT_LastSessionDiagnostic = "MON_OPEN_FAILED: " & openMsg
            RT_Log "MON_SESSION_RECOVERY", "", "", "FAIL", _
                   gRT_LastSessionDiagnostic & "; SECOND_TAB=TRUE; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
            Exit Function
        End If
        Set RT_FindMONSession = candidate
        gRT_LastSessionDiagnostic = "DEDICATED_MON_SESSION_READY"
        RT_Log "MON_SESSION_RECOVERY", "", "", "PASS", _
               gRT_LastSessionDiagnostic & "; SECOND_TAB=TRUE; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
        Exit Function
    End If

    ' Outside RUN, the generic helper may reuse or create a MON session.
    openMsg = ""
    On Error Resume Next
    Set candidate = ModuleSAP.CF_OpenTCodeSession(CF_MON_TCODE, openMsg)
    If Err.Number <> 0 Then
        openMsg = "Err " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    If Not candidate Is Nothing Then
        If RT_WaitReady(candidate, 15#) Then
            Set RT_FindMONSession = candidate
            gRT_LastSessionDiagnostic = "MON_SESSION_AUTO_OPENED_OR_REUSED"
            RT_Log "MON_SESSION_RECOVERY", "", "", "PASS", _
                   gRT_LastSessionDiagnostic & "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
            Exit Function
        End If
        openMsg = openMsg & IIf(openMsg = "", "", "; ") & "MON_SESSION_BUSY_TIMEOUT"
    End If

    If openMsg = "" Then openMsg = "SAP GUI Scripting did not return a MON session."
    gRT_LastSessionDiagnostic = "AUTO_OPEN_MON_FAILED: " & openMsg
    RT_Log "MON_SESSION_RECOVERY", "", "", "FAIL", _
           gRT_LastSessionDiagnostic & "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
End Function

' Opens exactly one new SAP GUI session and navigates it to the read-only
' Warehouse Management Monitor. The executor session is never repurposed.
' The caller has already proved that no reusable MON session exists, and the
' caller bounds creation to the dedicated RUN reader contract.
Private Function RT_OpenDedicatedMONSession(ByVal seed As Object, _
                                             ByRef messageText As String) As Object
    Dim conn As Object, sessions As Object, candidate As Object
    Dim beforeCount As Long, currentCount As Long, t0 As Double
    Dim errNo As Long, errText As String

    On Error GoTo Failed
    messageText = ""
    If seed Is Nothing Then
        messageText = "NO_LOGGED_IN_SAP_SESSION"
        Exit Function
    End If
    Set conn = seed.parent
    Set sessions = RT_SafeChildren(conn)
    beforeCount = RT_CollectionCount(sessions)
    If beforeCount < 1 Then
        messageText = "SAP_CONNECTION_HAS_NO_SESSION"
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    seed.CreateSession
    errNo = Err.Number
    errText = Err.Description
    Err.Clear
    On Error GoTo Failed
    If errNo <> 0 Then
        messageText = "CREATE_SESSION_ERROR: " & CStr(errNo) & ": " & errText
        Exit Function
    End If

    t0 = Timer
    Do
        Set sessions = RT_SafeChildren(conn)
        currentCount = RT_CollectionCount(sessions)
        If currentCount > beforeCount Then
            Set candidate = RT_CollectionObjectAt(sessions, currentCount - 1)
            If Not candidate Is Nothing Then
                If Not RT_SafeBool(candidate, "Busy") Then
                    If Not RT_FindById(candidate, "wnd[0]") Is Nothing Then Exit Do
                End If
            End If
        End If
        DoEvents
        Sleep 250
    Loop While RT_Elapsed(t0) < 15#

    If candidate Is Nothing Then
        messageText = "CREATE_SESSION_TIMEOUT"
        Exit Function
    End If
    If RT_SafeBool(candidate, "Busy") Then
        messageText = "CREATE_SESSION_BUSY_TIMEOUT"
        Exit Function
    End If
    If RT_FindById(candidate, "wnd[0]") Is Nothing Then
        messageText = "CREATE_SESSION_WINDOW_TIMEOUT"
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    candidate.StartTransaction CF_MON_TCODE
    errNo = Err.Number
    errText = Err.Description
    Err.Clear
    On Error GoTo Failed
    If errNo <> 0 Then
        messageText = "START_MON_ERROR: " & CStr(errNo) & ": " & errText
        Exit Function
    End If
    If Not RT_WaitReady(candidate, 20#) Then
        messageText = "START_MON_TIMEOUT"
        Exit Function
    End If
    Set RT_OpenDedicatedMONSession = candidate
    Exit Function

Failed:
    messageText = "OPEN_MON_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_MON_GetExistingReaderSession(ByRef messageText As String) As Object
    On Error GoTo Failed
    messageText = ""

    If Not SAPRun_RequireRuntime("PHASE1_WH", _
        "MON_RESULT_GRID_ID,MON_WT_GRID_ID,MON_UPPER_OB_COL,MON_UPPER_DOC_COL,MON_WT_DOC_COL,MON_WT_WHO_COL,MON_WT_COL,MON_WT_ITEM_COL,MON_WT_MATERIAL_COL,MON_WT_BATCH_COL,MON_WT_BIN_COL,MON_WT_QTY_COL,MON_WT_STATUS_COL,MON_OB_INPUT_ID,MON_EXECUTE_ID,MON_WT_ACTION_ID", messageText) Then Exit Function
    Set CF_MON_GetExistingReaderSession = RT_FindMONSession(False, False)
    If CF_MON_GetExistingReaderSession Is Nothing Then
        messageText = "STOP_MON_SESSION_REQUIRED: open an idle /SCWM/MON session before loading release holders."
    End If
    Exit Function
Failed:
    messageText = "STOP_MON_SESSION_REQUIRED: existing MON session check failed. Err " & _
                  CStr(Err.Number) & ": " & Err.Description
    Set CF_MON_GetExistingReaderSession = Nothing
End Function

' The executor owns same-tab PRDO-to-MON recovery. This helper never creates
' a new SAP tab in RUN; a missing reader is a safe stop for other callers.
Public Function CF_MON_EnsureReaderSession(ByRef messageText As String) As Object
    On Error GoTo Failed
    messageText = ""

    If Not SAPRun_RequireRuntime("PHASE1_WH", _
        "MON_RESULT_GRID_ID,MON_WT_GRID_ID,MON_UPPER_OB_COL,MON_UPPER_DOC_COL,MON_WT_DOC_COL,MON_WT_WHO_COL,MON_WT_COL,MON_WT_ITEM_COL,MON_WT_MATERIAL_COL,MON_WT_BATCH_COL,MON_WT_BIN_COL,MON_WT_QTY_COL,MON_WT_STATUS_COL,MON_OB_INPUT_ID,MON_EXECUTE_ID,MON_WT_ACTION_ID", messageText) Then GoTo CleanExit

    Set CF_MON_EnsureReaderSession = RT_FindMONSession(False, False)
    If CF_MON_EnsureReaderSession Is Nothing Then
        messageText = "STOP_MON_SESSION_REQUIRED: could not establish an idle /SCWM/MON reader session. " & gRT_LastSessionDiagnostic
    End If
    Exit Function

Failed:
    messageText = "STOP_MON_SESSION_REQUIRED: MON recovery failed. Err " & _
                  CStr(Err.Number) & ": " & Err.Description
    Set CF_MON_EnsureReaderSession = Nothing
CleanExit:
End Function

' The outbound-WT executor may create exactly one dedicated MON reader tab in
' RUN when no idle MON exists. Subsequent calls discover and reuse that tab;
' RT_FindMONSession refuses to create a duplicate while one is busy.
Public Function CF_MON_EnsureDedicatedReaderSession(ByRef messageText As String) As Object
    On Error GoTo Failed
    messageText = ""
    If Not SAPRun_RequireRuntime("PHASE1_WH", _
        "MON_RESULT_GRID_ID,MON_WT_GRID_ID,MON_UPPER_OB_COL,MON_UPPER_DOC_COL,MON_WT_DOC_COL,MON_WT_WHO_COL,MON_WT_COL,MON_WT_ITEM_COL,MON_WT_MATERIAL_COL,MON_WT_BATCH_COL,MON_WT_BIN_COL,MON_WT_QTY_COL,MON_WT_STATUS_COL,MON_OB_INPUT_ID,MON_EXECUTE_ID,MON_WT_ACTION_ID", messageText) Then Exit Function
    Set CF_MON_EnsureDedicatedReaderSession = RT_FindMONSession(False, True, True)
    If CF_MON_EnsureDedicatedReaderSession Is Nothing Then
        messageText = "STOP_MON_SESSION_REQUIRED: could not establish the bounded second-tab /SCWM/MON reader session. " & gRT_LastSessionDiagnostic
    End If
    Exit Function
Failed:
    messageText = "STOP_MON_SESSION_REQUIRED: dedicated MON recovery failed. Err " & _
                  CStr(Err.Number) & ": " & Err.Description
    Set CF_MON_EnsureDedicatedReaderSession = Nothing
End Function

Public Function CF_MON_WaitExistingReaderIdle(ByVal ses As Object, ByVal timeoutSeconds As Double) As Boolean
    CF_MON_WaitExistingReaderIdle = RT_WaitReady(ses, timeoutSeconds)
End Function

' The .children property itself (not just enumerating it) is a genuine COM
' call into the out-of-process SAP GUI Scripting engine. Reading it on a
' connection/session object that is stale, mid-teardown, or disconnected is
' one of the most common real-world triggers of an unwrapped, raw
' "-2147418113 Automation error: Catastrophic failure" - and it was NOT
' guarded before (only the enumeration helpers were), which is why the raw
' error kept escaping past every retry/try-catch layer around it.
Private Function RT_SafeChildren(ByVal obj As Object) As Object
    If obj Is Nothing Then Exit Function
    On Error Resume Next
    Set RT_SafeChildren = obj.children
    Err.Clear
    On Error GoTo 0
End Function

Private Function RT_CollectionCount(ByVal items As Object) As Long
    If items Is Nothing Then Exit Function
    On Error Resume Next
    RT_CollectionCount = CLng(items.count)
    On Error GoTo 0
End Function

' Single, safe access path (no more speculative multi-method fallback chain -
' repeatedly probing an out-of-process COM object with several different call
' styles back to back is what was poisoning the SAP GUI Scripting proxy and
' causing the untrappable "Catastrophic failure").
Private Function RT_CollectionObjectAt(ByVal items As Object, ByVal indexNo As Long) As Object
    If items Is Nothing Then Exit Function
    On Error Resume Next
    Set RT_CollectionObjectAt = items.ElementAt(indexNo)
    On Error GoTo 0
End Function

Private Function RT_CollectionValueAt(ByVal items As Object, ByVal indexNo As Long) As Variant
    If items Is Nothing Then Exit Function
    On Error Resume Next
    RT_CollectionValueAt = items.ElementAt(indexNo)
    On Error GoTo 0
End Function

Private Function RT_TcodeEquals(ByVal actualTcode As String, ByVal expectedTcode As String) As Boolean
    Dim a As String, b As String
    a = UCase$(Trim$(actualTcode))
    b = UCase$(Trim$(expectedTcode))
    If Left$(a, 1) = "/" Then a = Mid$(a, 2)
    If Left$(b, 1) = "/" Then b = Mid$(b, 2)
    RT_TcodeEquals = (a = b)
End Function

Private Function RT_SessionTransaction(ByVal ses As Object) As String
    If ses Is Nothing Then Exit Function
    On Error Resume Next
    RT_SessionTransaction = CStr(ses.Info.Transaction)
    On Error GoTo 0
End Function

Private Function RT_FindById(ByVal ses As Object, ByVal controlID As String) As Object
    If ses Is Nothing Then Exit Function
    On Error Resume Next
    Set RT_FindById = ses.FindById(controlID, False)
    On Error GoTo 0
End Function

Private Function RT_ExactID(ByVal configKey As String, ByVal mapDefault As String) As String
    If SAPRun_IsRunMode() Then
        RT_ExactID = Trim$(ModuleHelper.ToolConfig(configKey, ""))
    Else
        RT_ExactID = Trim$(ModuleHelper.ToolConfig(configKey, mapDefault))
    End If
End Function

Private Function RT_ColumnID(ByVal configKey As String, ByVal mapDefault As String) As String
    If SAPRun_IsRunMode() Then
        RT_ColumnID = Trim$(ModuleHelper.ToolConfig(configKey, ""))
    Else
        RT_ColumnID = Trim$(ModuleHelper.ToolConfig(configKey, mapDefault))
    End If
End Function

Private Function RT_SafeText(ByVal obj As Object, ByVal propertyName As String) As String
    If obj Is Nothing Then Exit Function
    On Error Resume Next
    RT_SafeText = CStr(CallByName(obj, propertyName, VbGet))
    On Error GoTo 0
End Function

Private Function RT_SafeLong(ByVal obj As Object, ByVal propertyName As String) As Long
    If obj Is Nothing Then Exit Function
    On Error Resume Next
    RT_SafeLong = CLng(CallByName(obj, propertyName, VbGet))
    On Error GoTo 0
End Function

Private Function RT_SafeBool(ByVal obj As Object, ByVal propertyName As String) As Boolean
    If obj Is Nothing Then Exit Function
    On Error Resume Next
    RT_SafeBool = CBool(CallByName(obj, propertyName, VbGet))
    On Error GoTo 0
End Function

Private Function RT_WaitReady(ByVal ses As Object, ByVal timeoutSeconds As Double) As Boolean
    Dim t0 As Double
    If ses Is Nothing Then Exit Function
    t0 = RT_StartTimer()
    Do
        DoEvents
        If gRT_StopRequested Then Exit Function
        If Not RT_SafeBool(ses, "Busy") Then
            RT_WaitReady = True
            Exit Function
        End If
        If RT_Elapsed(t0) >= timeoutSeconds Then Exit Do
        Sleep CF_POLL_MS
    Loop
End Function

'============================================================
' MON EXACT-CONTEXT ENGINE
'============================================================
Private Sub RT_RequireColumn(ByVal grid As Object, ByVal columnKey As String)
    If Trim$(columnKey) = "" Then Err.Raise vbObjectError + 9438, , "STOP_CONTROL_NOT_FOUND: blank technical column ID."
    If Not RT_HasColumn(grid, columnKey) Then
        Err.Raise vbObjectError + 9440, , "STOP_CONTROL_NOT_FOUND: MON column=" & columnKey
    End If
End Sub

Private Sub RT_RequireUpperLoadColumns(ByVal upperGrid As Object)
    RT_RequireColumn upperGrid, RT_ColumnID("MON_UPPER_OB_COL", "ERP_NUMBER_H")
    RT_RequireColumn upperGrid, RT_ColumnID("MON_UPPER_DOC_COL", "DOCNO_H")
    RT_RequireColumn upperGrid, RT_ColumnID("MON_UPPER_WA_COL", CF_OB_WA_COL)
    RT_RequireColumn upperGrid, RT_ColumnID("MON_UPPER_PICKING_COL", CF_OB_PICKING_COL)
    RT_RequireColumn upperGrid, RT_ColumnID("MON_UPPER_GI_COL", CF_OB_GI_COL)
    RT_RequireColumn upperGrid, RT_ColumnID("MON_UPPER_PICKING_PLAN_COL", CF_OB_PICKING_PI_COL)
End Sub

Private Function RT_HasColumn(ByVal grid As Object, ByVal columnKey As String) As Boolean
    Dim order As Object, i As Long, v As Variant, colCount As Long
    If grid Is Nothing Then Exit Function
    On Error Resume Next
    Set order = grid.columnOrder
    On Error GoTo 0
    If order Is Nothing Then Exit Function
    colCount = RT_CollectionCount(order)
    For i = 0 To colCount - 1
        v = RT_CollectionValueAt(order, i)
        If StrComp(Trim$(CStr(v & "")), columnKey, vbTextCompare) = 0 Then
            RT_HasColumn = True
            Exit Function
        End If
    Next i
    Set order = Nothing
End Function

Private Function RT_FindUpperOBRow(ByVal upperGrid As Object, ByVal ob As String, ByRef matchCount As Long) As Long
    Dim r As Long, rows As Long, v As String
    RT_FindUpperOBRow = -1
    rows = RT_SafeLong(upperGrid, "RowCount")
    For r = 0 To rows - 1
        v = RT_GridCell(upperGrid, r, RT_ColumnID("MON_UPPER_OB_COL", "ERP_NUMBER_H"))
        If RT_NormalizeOB(v) = RT_NormalizeOB(ob) Then
            matchCount = matchCount + 1
            If RT_FindUpperOBRow < 0 Then RT_FindUpperOBRow = r
        End If
    Next r
End Function

Private Function RT_FindSingleFilteredResultRow(ByVal upperGrid As Object, ByRef resultCount As Long) As Long
    Dim r As Long, rows As Long, docText As String
    RT_FindSingleFilteredResultRow = -1
    resultCount = 0
    rows = RT_SafeLong(upperGrid, "RowCount")
    For r = 0 To rows - 1
        docText = RT_GridCell(upperGrid, r, RT_ColumnID("MON_UPPER_DOC_COL", "DOCNO_H"))
        If docText <> "" Then
            resultCount = resultCount + 1
            If RT_FindSingleFilteredResultRow < 0 Then RT_FindSingleFilteredResultRow = r
        End If
    Next r
    If resultCount <> 1 Then RT_FindSingleFilteredResultRow = -1
End Function

Private Function RT_NormalizeOB(ByVal rawOB As String) As String
    Dim s As String
    s = Replace(Trim$(rawOB), " ", "")
    Do While Len(s) > 1 And Left$(s, 1) = "0"
        s = Mid$(s, 2)
    Loop
    RT_NormalizeOB = UCase$(s)
End Function

Private Function RT_UpperGridSample(ByVal upperGrid As Object, ByVal sampleLimit As Long) As String
    Dim rows As Long, r As Long, found As Long, obText As String, docText As String
    If upperGrid Is Nothing Then Exit Function
    rows = RT_SafeLong(upperGrid, "RowCount")
    For r = 0 To rows - 1
        obText = RT_GridCell(upperGrid, r, RT_ColumnID("MON_UPPER_OB_COL", "ERP_NUMBER_H"))
        docText = RT_GridCell(upperGrid, r, RT_ColumnID("MON_UPPER_DOC_COL", "DOCNO_H"))
        If obText <> "" Or docText <> "" Then
            RT_UpperGridSample = RT_UpperGridSample & "[r" & CStr(r + 1) & ": OB=" & obText & ", DOC=" & docText & "] "
            found = found + 1
            If found >= sampleLimit Then Exit For
        End If
    Next r
    If RT_UpperGridSample = "" Then RT_UpperGridSample = "<upper grid khong co dong co OB/DOC>"
End Function

Private Sub RT_SelectUpperRow(ByVal ses As Object, ByVal grid As Object, ByVal rowNo As Long, ByVal columnKey As String)
    Dim stepName As String

    stepName = "SelectedRows"
    On Error GoTo StepFailed
    grid.selectedRows = CStr(rowNo)

    ' The MON result row only needs to be selected before pressing the
    ' Warehouse Task toolbar button.  Clicking the hidden ERP_NUMBER_H cell
    ' causes a server transition/RPC error in this layout and is unnecessary.
    stepName = "SetCurrentCell"
    On Error Resume Next
    grid.SetCurrentCell rowNo, RT_ColumnID("MON_UPPER_DOC_COL", "DOCNO_H")
    If Err.Number <> 0 Then Err.Clear: grid.SetCurrentCell rowNo, columnKey
    Err.Clear
    On Error GoTo 0

    ' ClickCurrentCell can trigger a full SAP screen transition (control tree
    ' torn down and rebuilt). Give SAP a moment to settle before touching the
    ' grid object again - this is the single most common source of a raw,
    ' untrappable Catastrophic failure in this exact spot.
    RT_WaitReady ses, 3#

    ' Scrolling the row into view is purely cosmetic - never abort the whole
    ' load just because this one call fails.
    On Error Resume Next
    grid.FirstVisibleRow = rowNo
    Err.Clear
    On Error GoTo 0
    Exit Sub

StepFailed:
    Err.Raise vbObjectError + 9441, , "MON_SELECT_OB_FAILED at " & stepName & ": " & Err.Description
End Sub

Private Function RT_WaitWTContext(ByVal ses As Object, ByVal expectedDoc As String, ByRef wtGrid As Object, ByVal timeoutSeconds As Double) As Boolean
    Dim t0 As Double, grid As Object
    Dim rows As Long, r As Long, currentDoc As String, rawDoc As String

    t0 = RT_StartTimer()
    Do
        DoEvents
        If gRT_StopRequested Then Exit Function
        RT_WaitReady ses, 1#

        Set grid = Nothing
        Set grid = RT_FindById(ses, RT_ExactID("MON_WT_GRID_ID", CF_WT_ID))
        If Not grid Is Nothing Then
            rows = RT_SafeLong(grid, "RowCount")
            currentDoc = ""
            For r = 0 To rows - 1
                rawDoc = RT_GridCell(grid, r, RT_ColumnID("MON_WT_DOC_COL", "DOCNO"))
                If rawDoc <> "" Then currentDoc = rawDoc
                If StrComp(currentDoc, expectedDoc, vbTextCompare) = 0 Then
                    Set wtGrid = grid
                    RT_WaitWTContext = True
                    Exit Function
                End If
            Next r
        End If

        If RT_Elapsed(t0) >= timeoutSeconds Then Exit Do
        Sleep CF_POLL_MS
    Loop
    Set grid = Nothing
End Function

Private Function RT_GridCell(ByVal grid As Object, ByVal rowNo As Long, ByVal columnKey As String) As String
    If grid Is Nothing Then Exit Function
    On Error Resume Next
    RT_GridCell = Trim$(CStr(grid.GetCellValue(rowNo, columnKey) & ""))
    On Error GoTo 0
End Function

Private Function RT_ReadOBHeaderState(ByVal upperGrid As Object, ByVal rowNo As Long, _
                                      ByRef waStatus As String, ByRef pickingStatus As String, _
                                      ByRef giStatus As String, ByRef pickingPIStatus As String, _
                                      ByRef obState As String, ByRef reasonText As String) As Boolean
    Dim missing As String
    waStatus = RT_GridCell(upperGrid, rowNo, RT_ColumnID("MON_UPPER_WA_COL", CF_OB_WA_COL))
    pickingStatus = RT_GridCell(upperGrid, rowNo, RT_ColumnID("MON_UPPER_PICKING_COL", CF_OB_PICKING_COL))
    giStatus = RT_GridCell(upperGrid, rowNo, RT_ColumnID("MON_UPPER_GI_COL", CF_OB_GI_COL))
    pickingPIStatus = RT_GridCell(upperGrid, rowNo, RT_ColumnID("MON_UPPER_PICKING_PLAN_COL", CF_OB_PICKING_PI_COL))
    If waStatus = "" Then missing = missing & RT_ColumnID("MON_UPPER_WA_COL", CF_OB_WA_COL) & ","
    If pickingStatus = "" Then missing = missing & RT_ColumnID("MON_UPPER_PICKING_COL", CF_OB_PICKING_COL) & ","
    If giStatus = "" Then missing = missing & RT_ColumnID("MON_UPPER_GI_COL", CF_OB_GI_COL) & ","
    If pickingPIStatus = "" Then missing = missing & RT_ColumnID("MON_UPPER_PICKING_PLAN_COL", CF_OB_PICKING_PI_COL) & ","
    If missing <> "" Then
        obState = "OB_STATE_NOT_PROVEN"
        reasonText = "OB_STATE_NOT_PROVEN: missing header value(s)=" & Left$(missing, Len(missing) - 1)
        Exit Function
    End If
    If UCase$(Trim$(waStatus)) = "COMPLETED" And _
       UCase$(Trim$(pickingStatus)) = "COMPLETED" And _
       UCase$(Trim$(giStatus)) = "COMPLETED" And _
       UCase$(Trim$(pickingPIStatus)) = "COMPLETED" Then
        obState = "OB_DONE"
        reasonText = "OB_DONE: all four header statuses are Completed."
    Else
        obState = "OB_OPEN_OR_PARTIAL"
        reasonText = "OB_OPEN_OR_PARTIAL: header completion is not 4/4."
    End If
    RT_ReadOBHeaderState = True
End Function

Private Function RT_SourceKeyExists(ByRef output() As Variant, ByVal usedRows As Long, ByVal sourceKey As String) As Boolean
    Dim i As Long
    For i = 1 To usedRows
        If StrComp(CStr(output(i, 20) & ""), sourceKey, vbTextCompare) = 0 Then
            RT_SourceKeyExists = True
            Exit Function
        End If
    Next i
End Function

Private Function RT_CaptureActualInput(ByVal ws As Worksheet) As Object
    Dim result As Object, lastRow As Long, r As Long, keyText As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    If ws Is Nothing Then Set RT_CaptureActualInput = result: Exit Function
    lastRow = ws.Cells(ws.rows.count, 20).End(xlUp).row
    If lastRow < 5 Then Set RT_CaptureActualInput = result: Exit Function
    For r = 5 To lastRow
        keyText = Trim$(CStr(ws.Cells(r, 20).Value2 & ""))
        If keyText <> "" Then
            If Not result.Exists(keyText) Then
                result.Add keyText, Array(CStr(ws.Cells(r, 11).Value2 & ""), ws.Cells(r, 12).Value2, _
                    CStr(ws.Cells(r, 13).Value2 & ""), CStr(ws.Cells(r, 14).Value2 & ""), _
                    CStr(ws.Cells(r, 21).Value2 & ""), CStr(ws.Cells(r, 22).Value2 & ""))
            End If
        End If
    Next r
    Set RT_CaptureActualInput = result
End Function

' NOTE: RT_StartTimer / RT_Elapsed / RT_SetStatus / RT_Log / RT_GetSingleInputOB /
' RT_Number are defined in modCF_UI.bas (unchanged) - do not redefine them here,
' VBA would raise "Ambiguous name" at compile time if the same Public name
' exists in two standard modules.









