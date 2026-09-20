Attribute VB_Name = "modCF_MultiOB"
Option Explicit

' TAI WH MULTI-OB - READ-ONLY LOAD PROOF.
' This module deliberately does not load POSTING and does not call any SAP
' write, cancel, save, or CF action. It reuses RT_LoadMON through CF_LoadWH.

Private Const MO_QUEUE_SHEET As String = "OB_QUEUE"
Private Const MO_WH_CACHE_SHEET As String = "OB_WH_CACHE"
Private Const MO_FIRST_INPUT_ROW As Long = 7
Private Const MO_LAST_INPUT_ROW As Long = 506
Private Const MO_WH_FIRST_ROW As Long = 5
Private Const MO_WH_LAST_ROW As Long = 504

Private mMO_Running As Boolean
Private mMO_PreservedActualState As Object

Public Sub CF_MultiOB_Setup()
    Dim wsQ As Worksheet, wsW As Worksheet, wsIn As Worksheet

    On Error GoTo Failed
    Set wsQ = MO_EnsureSheet(MO_QUEUE_SHEET)
    Set wsW = MO_EnsureSheet(MO_WH_CACHE_SHEET)
    Set wsIn = ThisWorkbook.Worksheets("CF_CONTROL")

    MO_WriteQueueHeaders wsQ
    MO_WriteWHHeaders wsW
    CF_ApplyInputLayout

    If Not gRT_Silent Then
        MsgBox "TAI WH MULTI-OB SAN SANG." & vbCrLf & _
               "Nhap OB tai o vang INPUT!A7:C7, sau do chay TAI WH.", _
               vbInformation, "TAI WH"
    End If
    Exit Sub

Failed:
    If Not gRT_Silent Then MsgBox "TAI WH SETUP LOI: " & Err.Description, _
                                  vbExclamation, "TAI WH"
End Sub

Public Sub CF_MultiOB_CompileProbe()
    Dim wsQ As Worksheet, wsW As Worksheet
    Set wsQ = MO_EnsureSheet(MO_QUEUE_SHEET)
    Set wsW = MO_EnsureSheet(MO_WH_CACHE_SHEET)
    MO_WriteQueueHeaders wsQ
    MO_WriteWHHeaders wsW
End Sub

Public Sub CF_MultiOB_MigrateLegacyStateColumns()
    Dim wsCache As Worksheet, wsQ As Worksheet, wsActual As Worksheet
    Dim lastRow As Long, lastQ As Long, r As Long, q As Long
    Dim obState As String, rawWT As String, rowState As String, decision As String, reasonText As String
    Dim requestedOB As String, queueReason As String
    Dim processAllowed As Boolean, postingAllowed As String, needRemaining As Boolean
    Dim batchRunID As String, syncMessage As String

    On Error GoTo Failed
    Set wsCache = ThisWorkbook.Worksheets(MO_WH_CACHE_SHEET)
    Set wsQ = ThisWorkbook.Worksheets(MO_QUEUE_SHEET)
    Set wsActual = ThisWorkbook.Worksheets("DATA_CORE")

    lastRow = MO_LastDataRow(wsCache, 1, 2, wsCache.rows.count)
    For r = 2 To lastRow
        obState = UCase$(MO_Text(wsCache.Cells(r, 21).Value2))
        rawWT = UCase$(MO_Text(wsCache.Cells(r, 12).Value2))
        If obState = "" Then obState = "OB_STATE_NOT_PROVEN"
        WT_NormalizeState obState, rawWT, False, 0#, rowState, processAllowed, postingAllowed, _
                          needRemaining, decision, reasonText
        wsCache.Cells(r, 21).Value2 = obState
        wsCache.Cells(r, 22).Value2 = rowState
        wsCache.Cells(r, 23).Value2 = ""
        wsCache.Cells(r, 24).Value2 = IIf(processAllowed, "TRUE", "FALSE")
        wsCache.Cells(r, 25).Value2 = decision
        wsCache.Cells(r, 26).Value2 = reasonText
        ' Legacy cache had a different user-input layout. Keep only the proven
        ' System Qty/Bin defaults; user input is rebuilt in WH_ACTUAL.
        wsCache.Cells(r, 13).Value2 = wsCache.Cells(r, 9).Value2
        wsCache.Cells(r, 14).Value2 = MO_Text(wsCache.Cells(r, 10).Value2)
    Next r

    lastQ = MO_LastDataRow(wsQ, 1, 2, wsQ.rows.count)
    For q = 2 To lastQ
        requestedOB = MO_Text(wsQ.Cells(q, 3).Value2)
        obState = ""
        queueReason = ""
        For r = 2 To lastRow
            If StrComp(MO_Text(wsCache.Cells(r, 2).Value2), requestedOB, vbTextCompare) = 0 Then
                obState = UCase$(MO_Text(wsCache.Cells(r, 21).Value2))
                queueReason = MO_Text(wsCache.Cells(r, 26).Value2)
                Exit For
            End If
        Next r
        If obState = "" Then
            obState = "OB_STATE_NOT_PROVEN"
            queueReason = "OB_STATE_NOT_PROVEN: legacy snapshot had no proven header state."
        End If
        wsQ.Cells(q, 19).Value2 = ""
        wsQ.Cells(q, 20).Value2 = ""
        wsQ.Cells(q, 21).Value2 = ""
        wsQ.Cells(q, 22).Value2 = ""
        wsQ.Cells(q, 23).Value2 = obState
        wsQ.Cells(q, 24).Value2 = queueReason
        wsQ.Cells(q, 11).Value2 = IIf(obState = "OB_OPEN_OR_PARTIAL", "TRUE", "FALSE")
    Next q

    batchRunID = MO_Text(wsActual.Range("W2").Value2)
    If batchRunID <> "" Then
        If Not MO_SyncCacheToActual(wsCache, wsActual, batchRunID, syncMessage) Then
            Err.Raise vbObjectError + 9861, , syncMessage
        End If
    End If
    Exit Sub

Failed:
    Err.Raise vbObjectError + 9862, , "LEGACY_STATE_MIGRATION_FAIL: " & Err.Description
End Sub

Public Sub CF_RunAllMultiOB()
    Dim wsIn As Worksheet, wsQ As Worksheet, wsCache As Worksheet, wsActual As Worksheet
    Dim inputRows As Collection
    Dim oldA7 As Variant
    Dim seq As Long, qRow As Long, inputCount As Long, inputRow As Long
    Dim passCount As Long, failCount As Long
    Dim attemptedCount As Long, expectedSnapshotRows As Long, actualSnapshotRows As Long
    Dim batchRunID As String
    Dim ob As String, noteText As String, returnedOB As String
    Dim reasonText As String, whCount As Long
    Dim started As Date, finished As Date
    Dim ok As Boolean
    Dim batchPass As Boolean
    Dim gateMessage As String
    Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation
    Dim syncOK As Boolean, syncMessage As String, obProcessAllowed As Boolean
    Dim systemStop As Boolean
    Dim failureNo As Long, failureSource As String, failureDescription As String

    On Error GoTo BatchFailed
    If mMO_Running Then
        MsgBox "TAI WH DANG CHAY: vui long cho luong hien tai ket thuc.", _
               vbExclamation, "TAI WH"
        Exit Sub
    End If
    mMO_Running = True
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsIn = ThisWorkbook.Worksheets("CF_CONTROL")
    oldA7 = wsIn.Range("A7").Value2
    Set wsQ = MO_EnsureSheet(MO_QUEUE_SHEET)
    Set wsCache = MO_EnsureSheet(MO_WH_CACHE_SHEET)
    Set wsActual = ThisWorkbook.Worksheets("DATA_CORE")
    Set mMO_PreservedActualState = MO_CaptureActualUserState(wsActual)
    MO_WriteQueueHeaders wsQ
    MO_WriteWHHeaders wsCache
    Set inputRows = MO_ReadInputOBs(wsIn)
    inputCount = inputRows.count
    If inputCount = 0 Then Err.Raise vbObjectError + 9851, , _
        "INPUT_OB_LIST_EMPTY: no numeric OB was found in INPUT!A7:A506."

    ' Preserve the prior run unless SAP MON has passed this read-only gate.
    If Not RT_MultiOBStartupGate(gateMessage) Then
        MO_MarkSnapshotStale wsActual, "STOP_STALE_WH_SNAPSHOT: " & gateMessage
        RT_Log "MULTI_WH_START", "", "", "STOP", gateMessage & _
               "; SAP_DIAGNOSTIC=" & gRT_LastSessionDiagnostic & _
               "; PRIOR_SNAPSHOT_PRESERVED=TRUE; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
        mMO_Running = False
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        Application.Calculation = oldCalculation
        Set mMO_PreservedActualState = Nothing
        If Not gRT_Silent Then MsgBox "TAI WH CHUA THUC HIEN." & vbCrLf & _
            gateMessage & vbCrLf & _
            "OB_QUEUE va OB_WH_CACHE duoc giu nguyen de bao toan snapshot cu.", _
            vbExclamation, "TAI WH"
        Exit Sub
    End If

    MO_ResetPhase1Sheets wsQ, wsCache
    gRT_RunID = Format$(Now, "yyyymmdd_hhnnss")
    batchRunID = gRT_RunID
    gRT_ActiveBatchRunID = batchRunID
    RT_Log "MULTI_WH_START", "", "", "START", _
           "OB_COUNT=" & CStr(inputCount) & "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"

    For seq = 1 To inputCount
        ob = CStr(inputRows(seq)(0))
        noteText = CStr(inputRows(seq)(1))
        inputRow = CLng(inputRows(seq)(2))
        qRow = seq + 1
        started = Now
        returnedOB = ""
        reasonText = ""
        whCount = 0

        wsQ.Cells(qRow, 1).Value2 = seq
        wsQ.Cells(qRow, 2).Value2 = inputRow
        wsQ.Cells(qRow, 3).Value2 = ob
        wsQ.Cells(qRow, 5).Value2 = noteText
        wsQ.Cells(qRow, 7).Value2 = "RUNNING"
        wsQ.Cells(qRow, 9).Value2 = started
        wsQ.Cells(qRow, 11).Value2 = "FALSE"
        wsQ.Cells(qRow, 12).Value2 = "YES"
        wsQ.Cells(qRow, 13).Value2 = 0
        wsQ.Cells(qRow, 14).Value2 = 0
        wsQ.Cells(qRow, 15).Value2 = 0
        wsQ.Cells(qRow, 16).Value2 = 0
        wsQ.Cells(qRow, 18).Value2 = batchRunID

        RT_Log "MULTI_WH_OB", ob, "", "START", _
               "SEQ=" & CStr(seq) & "; NOTE=" & noteText

        ' RT_LoadMON reads INPUT!A7. The working WH sheet is cleared before
        ' every call so a loader failure can never leave a prior OB as data.
        wsIn.Range("A7").Value2 = ob
        wsActual.Range("A5:V504").ClearContents
        wsActual.Range("Y5:AD504").ClearContents

        attemptedCount = attemptedCount + 1
        ok = MO_LoadAndSnapshot(seq, ob, batchRunID, returnedOB, whCount, reasonText)
        finished = Now

        wsQ.Cells(qRow, 4).Value2 = returnedOB
        wsQ.Cells(qRow, 6).Value2 = whCount
        wsQ.Cells(qRow, 7).Value2 = IIf(ok, "PASS", "FAIL")
        wsQ.Cells(qRow, 8).Value2 = IIf(ok, "WH_SNAPSHOT_APPENDED", "WH_SNAPSHOT_NOT_SAVED")
        wsQ.Cells(qRow, 9).Value2 = started
        wsQ.Cells(qRow, 10).Value2 = finished
        ' The read-only loader can return no header state after a technical
        ' failure. Never apply Trim$ directly to that value: Null must become
        ' a logged non-processable state, not abort the batch.
        obProcessAllowed = (UCase$(MO_Text(gRT_LastOBState)) = "OB_OPEN_OR_PARTIAL")
        wsQ.Cells(qRow, 11).Value2 = IIf(ok And obProcessAllowed, "TRUE", "FALSE")
        wsQ.Cells(qRow, 12).Value2 = "YES"
        wsQ.Cells(qRow, 13).Value2 = 0
        wsQ.Cells(qRow, 14).Value2 = 0
        wsQ.Cells(qRow, 15).Value2 = 0
        wsQ.Cells(qRow, 16).Value2 = 0
        wsQ.Cells(qRow, 17).Value2 = reasonText
        wsQ.Cells(qRow, 18).Value2 = batchRunID
        wsQ.Cells(qRow, 19).Value2 = MO_Text(gRT_LastOBWAStatus)
        wsQ.Cells(qRow, 20).Value2 = MO_Text(gRT_LastOBPickingStatus)
        wsQ.Cells(qRow, 21).Value2 = MO_Text(gRT_LastOBGIStatus)
        wsQ.Cells(qRow, 22).Value2 = MO_Text(gRT_LastOBPickingPIStatus)
        wsQ.Cells(qRow, 23).Value2 = MO_Text(gRT_LastOBState)
        wsQ.Cells(qRow, 24).Value2 = MO_Text(gRT_LastOBStateReason)

        If ok Then
            passCount = passCount + 1
            expectedSnapshotRows = expectedSnapshotRows + whCount
            RT_Log "MULTI_WH_OB", ob, returnedOB, "PASS", _
                   "SEQ=" & CStr(seq) & "; ROWS=" & CStr(whCount) & _
                   "; RequestedOB=ReturnedOB; SNAPSHOT=APPENDED; " & _
                   "SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
        Else
            failCount = failCount + 1
            RT_Log "MULTI_WH_OB", ob, returnedOB, "FAIL", _
                   "SEQ=" & CStr(seq) & "; " & reasonText & _
                   "; ProcessAllowed=FALSE; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
            systemStop = MO_IsSystemFailure(reasonText)
            If systemStop Then
                RT_Log "MULTI_WH_STOP", ob, returnedOB, "STOP", _
                       "SYSTEM_STOP_AFTER_OB=" & ob & "; " & reasonText & _
                       "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
                wsIn.Range("A7").Value2 = oldA7
                Exit For
            End If
        End If

        wsIn.Range("A7").Value2 = oldA7
    Next seq

    wsIn.Range("A7").Value2 = oldA7
    syncOK = MO_SyncCacheToActual(wsCache, wsActual, batchRunID, syncMessage)
    actualSnapshotRows = MO_CountSnapshotRows(wsCache, batchRunID)
    batchPass = (attemptedCount = inputCount And passCount = inputCount And _
                 failCount = 0 And expectedSnapshotRows > 0 And _
                 actualSnapshotRows = expectedSnapshotRows And syncOK)

    RT_Log "MULTI_WH_END", "", "", IIf(batchPass, "PASS", "FAIL"), _
           "REQUESTED=" & CStr(inputCount) & "; ATTEMPTED=" & CStr(attemptedCount) & _
           "; PASS=" & CStr(passCount) & "; FAIL=" & CStr(failCount) & _
           "; EXPECTED_SNAPSHOT_ROWS=" & CStr(expectedSnapshotRows) & _
           "; ACTUAL_SNAPSHOT_ROWS=" & CStr(actualSnapshotRows) & _
           "; ACTUAL_SYNC=" & IIf(syncOK, "PASS", "FAIL") & "; " & syncMessage & _
           "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
    mMO_Running = False
    Set mMO_PreservedActualState = Nothing
    gRT_ActiveBatchRunID = ""
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalculation

    If Not batchPass Then
        MO_MarkSnapshotStale wsActual, "STOP_STALE_WH_SNAPSHOT: latest TAI WH batch did not complete."
        If Not gRT_Silent Then MsgBox "TAI WH CHUA HOAN TAT." & vbCrLf & _
            "Requested: " & CStr(inputCount) & " | Attempted: " & CStr(attemptedCount) & vbCrLf & _
            "PASS: " & CStr(passCount) & " | FAIL: " & CStr(failCount) & vbCrLf & _
            "Snapshot rows expected/actual: " & CStr(expectedSnapshotRows) & "/" & CStr(actualSnapshotRows) & vbCrLf & _
            "See OB_QUEUE and RUN_LOG. No success was recorded.", _
            vbExclamation, "TAI WH"
        Exit Sub
    End If

    If Not gRT_Silent Then MsgBox "TAI WH MULTI-OB HOAN TAT." & vbCrLf & _
        "Input OB: " & CStr(inputCount) & vbCrLf & _
        "Attempted: " & CStr(attemptedCount) & vbCrLf & _
        "PASS: " & CStr(passCount) & " | FAIL: " & CStr(failCount) & vbCrLf & _
        "Snapshot rows: " & CStr(actualSnapshotRows) & vbCrLf & _
        "See OB_QUEUE and OB_WH_CACHE. Read-only; no posting.", _
        vbInformation, "TAI WH"
    Exit Sub
BatchFailed:
    failureNo = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    If Not wsIn Is Nothing Then wsIn.Range("A7").Value2 = oldA7
    If Not wsActual Is Nothing Then MO_MarkSnapshotStale wsActual, "STOP_STALE_WH_SNAPSHOT: TAI WH batch aborted."
    RT_Log "MULTI_WH_END", ob, returnedOB, "FAIL", _
           "BATCH_ABORTED|ERR_NO=" & CStr(failureNo) & _
           "|ERR_SOURCE=" & failureSource & _
           "|ERR_DESCRIPTION=" & failureDescription & _
           "|SAP_DIAGNOSTIC=" & gRT_LastSessionDiagnostic & _
           "|RUN_ID=" & gRT_RunID & _
           "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
    mMO_Running = False
    Set mMO_PreservedActualState = Nothing
    gRT_ActiveBatchRunID = ""
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalculation
    If Not gRT_Silent Then MsgBox "TAI WH LOI: " & failureDescription, _
                                  vbExclamation, "TAI WH"
End Sub


Private Function MO_SyncCacheToActual(ByVal wsCache As Worksheet, _
                                      ByVal wsActual As Worksheet, _
                                      ByVal batchRunID As String, _
                                      ByRef messageText As String) As Boolean
    Dim preserved As Object, seen As Object, vals As Variant
    Dim lastCache As Long, r As Long, outRow As Long, rowCount As Long
    Dim sourceKey As String, preservedCount As Long, wtState As String
    On Error GoTo Failed
    messageText = ""
    If Not mMO_PreservedActualState Is Nothing Then
        Set preserved = mMO_PreservedActualState
    Else
        Set preserved = MO_CaptureActualUserState(wsActual)
    End If
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    lastCache = MO_LastDataRow(wsCache, 1, 2, wsCache.rows.count)
    For r = 2 To lastCache
        If StrComp(MO_Text(wsCache.Cells(r, 20).Value2), batchRunID, vbTextCompare) = 0 _
           And StrComp(MO_Text(wsCache.Cells(r, 16).Value2), "PASS", vbTextCompare) = 0 Then
            sourceKey = MO_Text(wsCache.Cells(r, 18).Value2)
            If sourceKey = "" Then messageText = "WH_ACTUAL_SYNC_KEY_EMPTY: cache row " & CStr(r): Exit Function
            If seen.Exists(sourceKey) Then messageText = "WH_ACTUAL_SYNC_DUPLICATE_KEY: " & sourceKey: Exit Function
            seen.Add sourceKey, True
        End If
    Next r
    If seen.count = 0 Then messageText = "WH_ACTUAL_SYNC_EMPTY: no PASS rows for RunID " & batchRunID: Exit Function
    wsActual.Range("A5:V504").ClearContents
    wsActual.Range("Y5:AD504").ClearContents
    AP_PrepareActualPickingSheet
    wsActual.Range("Y4:AD4").Value = Array("OB State", "WT State", "Remaining Qty", _
        "Process Allowed", "Decision", "Reason")
    outRow = 5
    For r = 2 To lastCache
        If StrComp(MO_Text(wsCache.Cells(r, 20).Value2), batchRunID, vbTextCompare) = 0 _
           And StrComp(MO_Text(wsCache.Cells(r, 16).Value2), "PASS", vbTextCompare) = 0 Then
            sourceKey = MO_Text(wsCache.Cells(r, 18).Value2)
            wsActual.Cells(outRow, 1).Value2 = MO_Text(wsCache.Cells(r, 3).Value2)
            wsActual.Cells(outRow, 2).Value2 = MO_Text(wsCache.Cells(r, 4).Value2)
            wsActual.Cells(outRow, 3).Value2 = MO_Text(wsCache.Cells(r, 5).Value2)
            wsActual.Cells(outRow, 4).Value2 = MO_Text(wsCache.Cells(r, 6).Value2)
            wsActual.Cells(outRow, 5).Value2 = MO_Text(wsCache.Cells(r, 7).Value2)
            wsActual.Cells(outRow, 6).Value2 = MO_Text(wsCache.Cells(r, 8).Value2)
            wsActual.Cells(outRow, 7).Value2 = wsCache.Cells(r, 9).Value2
            wsActual.Cells(outRow, 8).Value2 = MO_Text(wsCache.Cells(r, 10).Value2)
            wsActual.Cells(outRow, 9).Value2 = MO_Text(wsCache.Cells(r, 11).Value2)
            wsActual.Cells(outRow, 10).Value2 = MO_Text(wsCache.Cells(r, 12).Value2)
            ' Actual Picking contract: K=Area, L=Actual Qty, M=Bin No., N=Actual Bin.
            wsActual.Cells(outRow, 11).Value2 = ""
            wsActual.Cells(outRow, 12).Value2 = wsCache.Cells(r, 13).Value2
            wsActual.Cells(outRow, 13).Value2 = ""
            wsActual.Cells(outRow, 14).Value2 = MO_Text(wsCache.Cells(r, 14).Value2)
            wsActual.Cells(outRow, 17).Value2 = "LOADED"
            wsActual.Cells(outRow, 18).Value2 = "MON_LOADED"
            wsActual.Cells(outRow, 20).Value2 = sourceKey
            wsActual.Cells(outRow, 25).Value2 = MO_Text(wsCache.Cells(r, 21).Value2)
            wsActual.Cells(outRow, 26).Value2 = MO_Text(wsCache.Cells(r, 22).Value2)
            wsActual.Cells(outRow, 27).Value2 = MO_Text(wsCache.Cells(r, 23).Value2)
            wsActual.Cells(outRow, 28).Value2 = MO_Text(wsCache.Cells(r, 24).Value2)
            wsActual.Cells(outRow, 29).Value2 = MO_Text(wsCache.Cells(r, 25).Value2)
            wsActual.Cells(outRow, 30).Value2 = MO_Text(wsCache.Cells(r, 26).Value2)
            If Not preserved Is Nothing Then
                If preserved.Exists(sourceKey) And _
                   UCase$(MO_Text(wsCache.Cells(r, 21).Value2)) <> "OB_DONE" And _
                   (UCase$(MO_Text(wsCache.Cells(r, 22).Value2)) = "WT_ACTIVE" Or _
                    UCase$(MO_Text(wsCache.Cells(r, 22).Value2)) = "WT_CANCELLED_PENDING_REMAINING" Or _
                    UCase$(MO_Text(wsCache.Cells(r, 22).Value2)) = "WT_CANCELLED_REMAINING") Then
                    vals = preserved(sourceKey)
                    wsActual.Cells(outRow, 11).Value2 = vals(0)
                    wsActual.Cells(outRow, 12).Value2 = vals(1)
                    wsActual.Cells(outRow, 13).Value2 = vals(2)
                    wsActual.Cells(outRow, 14).Value2 = vals(3)
                    wsActual.Cells(outRow, 21).Value2 = vals(4)
                    wsActual.Cells(outRow, 22).Value2 = vals(5)
                    AP_NormalizeRow wsActual, outRow, False
                    preservedCount = preservedCount + 1
                End If
            End If
            AP_NormalizeRow wsActual, outRow, False
            outRow = outRow + 1
        End If
    Next r
    rowCount = outRow - 5
    wsActual.Range("W1").Value2 = "WH Run ID"
    wsActual.Range("W2").Value2 = batchRunID
    wsActual.Range("X1").Value2 = "WH Actual Sync Rows"
    wsActual.Range("X2").Value2 = rowCount
    AP_PrepareActualPickingSheet
    messageText = "WH_ACTUAL_SYNC=PASS; rows=" & CStr(rowCount) & "; preserved=" & CStr(preservedCount) & "; RunID=" & batchRunID
    MO_SyncCacheToActual = True
    Exit Function
Failed:
    messageText = "WH_ACTUAL_SYNC_FAIL: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function MO_CaptureActualUserState(ByVal wsActual As Worksheet) As Object
    Dim result As Object, lastRow As Long, r As Long, sourceKey As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    If wsActual Is Nothing Then Set MO_CaptureActualUserState = result: Exit Function
    lastRow = MO_LastDataRow(wsActual, 1, MO_WH_FIRST_ROW, MO_WH_LAST_ROW)
    For r = MO_WH_FIRST_ROW To lastRow
        sourceKey = MO_Text(wsActual.Cells(r, 20).Value2)
        If sourceKey <> "" Then
            If Not result.Exists(sourceKey) Then result.Add sourceKey, Array( _
                MO_Text(wsActual.Cells(r, 11).Value2), wsActual.Cells(r, 12).Value2, _
                MO_Text(wsActual.Cells(r, 13).Value2), MO_Text(wsActual.Cells(r, 14).Value2), _
                MO_Text(wsActual.Cells(r, 21).Value2), MO_Text(wsActual.Cells(r, 22).Value2))
        End If
    Next r
    Set MO_CaptureActualUserState = result
End Function

Private Function MO_CountSnapshotRows(ByVal wsCache As Worksheet, ByVal batchRunID As String) As Long
    Dim r As Long, lastRow As Long
    If wsCache Is Nothing Then Exit Function
    lastRow = MO_LastDataRow(wsCache, 1, 2, wsCache.rows.count)
    For r = 2 To lastRow
        If StrComp(MO_Text(wsCache.Cells(r, 20).Value2), batchRunID, vbTextCompare) = 0 And _
           StrComp(MO_Text(wsCache.Cells(r, 16).Value2), "PASS", vbTextCompare) = 0 And _
           MO_Text(wsCache.Cells(r, 18).Value2) <> "" Then
            MO_CountSnapshotRows = MO_CountSnapshotRows + 1
        End If
    Next r
End Function
Private Function MO_IsSystemFailure(ByVal reasonText As String) As Boolean
    Dim u As String
    u = UCase$(reasonText)
    MO_IsSystemFailure = (InStr(u, "SYSTEM_FAIL_LOADER_ERROR") > 0 Or _
                          InStr(u, "SAP_SESSION") > 0 Or _
                          InStr(u, "MON_SESSION") > 0 Or _
                          InStr(u, "STOP_STAGE_TIMEOUT") > 0 Or _
                          InStr(u, "TIMEOUT_STOP") > 0 Or _
                          InStr(u, "AUTOMATION") > 0 Or _
                          InStr(u, "CATASTROPHIC") > 0 Or _
                          InStr(u, "UNEXPECTED_POPUP") > 0 Or _
                          InStr(u, "COM_FAILURE") > 0)
End Function

Private Function MO_LoadAndSnapshot(ByVal seq As Long, ByVal requestedOB As String, _
                                    ByVal batchRunID As String, _
                                    ByRef returnedOB As String, ByRef whCount As Long, _
                                    ByRef reasonText As String) As Boolean
    Dim wsWH As Worksheet
    Dim r As Long, lastRow As Long
    Dim rowOB As String, firstOB As String
    Dim hasData As Boolean, wrongOB As Boolean

    On Error GoTo Failed
    reasonText = ""
    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")

    ' Reuse the existing SAP GUI scripting WH loader only.
    RT_LoadMON_Silent
    If gRT_LastLoadError <> "" Then
        reasonText = "FAIL_LOADER_ERROR: " & gRT_LastLoadError
        If gRT_LastLoadSystemFailure Then reasonText = "SYSTEM_" & reasonText
        Exit Function
    End If

    lastRow = MO_LastDataRow(wsWH, 1, MO_WH_FIRST_ROW, MO_WH_LAST_ROW)
    For r = MO_WH_FIRST_ROW To lastRow
        rowOB = MO_Text(wsWH.Cells(r, 1).Value2)
        If rowOB <> "" Then
            hasData = True
            If firstOB = "" Then firstOB = rowOB
            If StrComp(rowOB, requestedOB, vbTextCompare) <> 0 Then wrongOB = True
            whCount = whCount + 1
        End If
    Next r
    returnedOB = firstOB

    If Not hasData Then
        reasonText = "FAIL_NO_WH_ROWS: loader returned no WH/WT rows"
        Exit Function
    End If
    If wrongOB Or StrComp(firstOB, requestedOB, vbTextCompare) <> 0 Then
        reasonText = "FAIL_STALE_DATA: RequestedOB=" & requestedOB & _
                     "; ReturnedOB=" & firstOB
        Exit Function
    End If

    MO_AppendSnapshot seq, requestedOB, batchRunID, wsWH, whCount
    reasonText = "RequestedOB=ReturnedOB; rows=" & CStr(whCount) & _
                 "; OB_STATE=" & MO_Text(gRT_LastOBState)
    MO_LoadAndSnapshot = True
    Exit Function

Failed:
    reasonText = "FAIL_LOADER_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub MO_AppendSnapshot(ByVal seq As Long, ByVal requestedOB As String, _
                              ByVal batchRunID As String, _
                              ByVal wsWH As Worksheet, ByVal whCount As Long)
    Dim wsCache As Worksheet
    Dim r As Long, outRow As Long, lastRow As Long
    Dim returnedOB As String, sourceKey As String

    Set wsCache = MO_EnsureSheet(MO_WH_CACHE_SHEET)
    outRow = MO_LastDataRow(wsCache, 1, 2, wsCache.rows.count) + 1
    If outRow < 2 Then outRow = 2

    lastRow = MO_LastDataRow(wsWH, 1, MO_WH_FIRST_ROW, MO_WH_LAST_ROW)
    For r = MO_WH_FIRST_ROW To lastRow
        returnedOB = MO_Text(wsWH.Cells(r, 1).Value2)
        If returnedOB <> "" Then
            sourceKey = MO_Text(wsWH.Cells(r, 20).Value2)
            wsCache.Cells(outRow, 1).Value2 = seq
            wsCache.Cells(outRow, 2).Value2 = requestedOB
            wsCache.Cells(outRow, 3).Value2 = returnedOB
            wsCache.Cells(outRow, 4).Value2 = MO_Text(wsWH.Cells(r, 2).Value2)
            wsCache.Cells(outRow, 5).Value2 = MO_Text(wsWH.Cells(r, 3).Value2)
            wsCache.Cells(outRow, 6).Value2 = MO_Text(wsWH.Cells(r, 4).Value2)
            wsCache.Cells(outRow, 7).Value2 = MO_Text(wsWH.Cells(r, 5).Value2)
            wsCache.Cells(outRow, 8).Value2 = MO_Text(wsWH.Cells(r, 6).Value2)
            wsCache.Cells(outRow, 9).Value2 = wsWH.Cells(r, 7).Value2
            wsCache.Cells(outRow, 10).Value2 = MO_Text(wsWH.Cells(r, 8).Value2)
            wsCache.Cells(outRow, 11).Value2 = MO_Text(wsWH.Cells(r, 9).Value2)
            wsCache.Cells(outRow, 12).Value2 = MO_Text(wsWH.Cells(r, 10).Value2)
            wsCache.Cells(outRow, 13).Value2 = wsWH.Cells(r, 12).Value2
            wsCache.Cells(outRow, 14).Value2 = AP_ResolvedBin(wsWH, r, MO_Text(wsWH.Cells(r, 8).Value2))
            wsCache.Cells(outRow, 15).Value2 = ""
            wsCache.Cells(outRow, 16).Value2 = "PASS"
            wsCache.Cells(outRow, 17).Value2 = "RequestedOB=ReturnedOB"
            wsCache.Cells(outRow, 18).Value2 = sourceKey
            wsCache.Cells(outRow, 19).Value2 = Now
            wsCache.Cells(outRow, 20).Value2 = batchRunID
            wsCache.Cells(outRow, 21).Value2 = MO_Text(wsWH.Cells(r, 25).Value2)
            wsCache.Cells(outRow, 22).Value2 = MO_Text(wsWH.Cells(r, 26).Value2)
            wsCache.Cells(outRow, 23).Value2 = MO_Text(wsWH.Cells(r, 27).Value2)
            wsCache.Cells(outRow, 24).Value2 = MO_Text(wsWH.Cells(r, 28).Value2)
            wsCache.Cells(outRow, 25).Value2 = MO_Text(wsWH.Cells(r, 29).Value2)
            wsCache.Cells(outRow, 26).Value2 = MO_Text(wsWH.Cells(r, 30).Value2)
            outRow = outRow + 1
        End If
    Next r
End Sub

Private Function MO_ReadInputOBs(ByVal wsIn As Worksheet) As Collection
    Dim result As New Collection
    Dim r As Long, lastRow As Long
    Dim ob As String, noteText As String, misplacedOB As String

    If Not MO_AssertOBInputColumn(wsIn, misplacedOB) Then
        Err.Raise vbObjectError + 9850, , misplacedOB
    End If
    lastRow = MO_LastDataRow(wsIn, 1, MO_FIRST_INPUT_ROW, MO_LAST_INPUT_ROW)
    For r = MO_FIRST_INPUT_ROW To lastRow
        ob = MO_Text(wsIn.Cells(r, 1).Value2)
        If MO_IsOBCandidate(ob) Then
            noteText = MO_Text(wsIn.Cells(r, 2).Value2)
            ' Preserve duplicates: no dictionary and no de-duplication rule.
            result.Add Array(ob, noteText, r)
        End If
    Next r
    Set MO_ReadInputOBs = result
End Function

' INPUT!A is the sole executable OB list.  Catch a numeric OB accidentally
' entered in the reserved status column before the loader can use an older
' value from A.  This is deliberately a local UI gate: SAP has not been read
' or written when it returns FALSE.
Private Function MO_AssertOBInputColumn(ByVal wsIn As Worksheet, _
                                         ByRef messageText As String) As Boolean
    Dim r As Long, statusText As String

    messageText = ""
    For r = MO_FIRST_INPUT_ROW To MO_LAST_INPUT_ROW
        statusText = MO_Text(wsIn.Cells(r, 3).Value2)
        If MO_IsOBCandidate(statusText) Then
            messageText = "STOP_INPUT_OB_WRONG_COLUMN: OB " & statusText & _
                          " is in INPUT!C" & CStr(r) & ". Move it to INPUT!A" & _
                          CStr(r) & "; no SAP read/write was run."
            Exit Function
        End If
    Next r
    MO_AssertOBInputColumn = True
End Function

Private Function MO_IsOBCandidate(ByVal valueText As String) As Boolean
    Dim i As Long, ch As String
    valueText = Trim$(valueText)
    If Len(valueText) < 5 Then Exit Function
    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    MO_IsOBCandidate = True
End Function

Private Sub MO_ResetPhase1Sheets(ByVal wsQ As Worksheet, ByVal wsW As Worksheet)
    MO_ClearBelowHeader wsQ, 24
    MO_ClearBelowHeader wsW, 26
End Sub

Private Sub MO_MarkSnapshotStale(ByVal wsActual As Worksheet, ByVal reasonText As String)
    If wsActual Is Nothing Then Exit Sub
    On Error Resume Next
    wsActual.Range("W1").Value2 = "WH Run ID"
    wsActual.Range("W2").Value2 = ""
    wsActual.Range("X1").Value2 = "WH Snapshot Gate"
    wsActual.Range("X2").Value2 = reasonText
    On Error GoTo 0
End Sub

Private Sub MO_ClearBelowHeader(ByVal ws As Worksheet, ByVal lastCol As Long)
    Dim lastRow As Long
    If ws Is Nothing Then Exit Sub
    lastRow = MO_LastDataRow(ws, 1, 2, ws.rows.count)
    If lastRow >= 2 Then ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).ClearContents
End Sub

Private Sub MO_WriteQueueHeaders(ByVal ws As Worksheet)
    ws.Range("A1:X1").Value = Array("Sequence", "Input Row", "Requested OB", "Returned OB", "Note", _
        "WH Rows", "Status", "Load Result", "Started", "Finished", "Process Allowed", _
        "Read Only", "SAP Write", "Cancel", "Save", "CF", "Reason", "Run ID", _
        "Warehouse Activity Status", "Picking Status", "Goods Issue Status", "Picking PI Status", _
        "OB State", "OB State Reason")
    ws.rows(1).Font.Bold = True
    ws.Range("I:J").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Range("K:P").NumberFormat = "@"
End Sub

Private Sub MO_WriteWHHeaders(ByVal ws As Worksheet)
    ws.Range("A1:Z1").Value = Array("Sequence", "Requested OB", "Returned OB", "Document", _
        "Item", "WH", "WT", "Material", "Sys Qty", "Sys Bin", "Sys Batch", "WT Status", _
        "Actual Qty", "Actual Bin", "Destination Bin", "Load Result", "Reason", "Source Key", _
        "Snapshot Time", "Load Run ID", "OB State", "WT State", "Remaining Qty", _
        "Process Allowed", "Decision", "Reason")
    ws.rows(1).Font.Bold = True
    ws.Range("S:S").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Range("T:T").NumberFormat = "@"
End Sub

Private Function MO_EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set MO_EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If MO_EnsureSheet Is Nothing Then
        Set MO_EnsureSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        MO_EnsureSheet.Name = sheetName
    End If
    MO_EnsureSheet.Visible = xlSheetVisible
End Function

Private Function MO_LastDataRow(ByVal ws As Worksheet, ByVal col As Long, _
                                ByVal minRow As Long, ByVal maxRow As Long) As Long
    Dim r As Long, lastCandidate As Long
    If ws Is Nothing Then Exit Function
    If maxRow > ws.rows.count Then maxRow = ws.rows.count
    lastCandidate = ws.Cells(ws.rows.count, col).End(xlUp).row
    If lastCandidate < minRow Then
        MO_LastDataRow = minRow - 1
        Exit Function
    End If
    If lastCandidate > maxRow Then lastCandidate = maxRow
    For r = lastCandidate To minRow Step -1
        If MO_Text(ws.Cells(r, col).Value2) <> "" Then
            MO_LastDataRow = r
            Exit Function
        End If
    Next r
    MO_LastDataRow = minRow - 1
End Function

Private Function MO_Text(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then
        MO_Text = ""
    Else
        MO_Text = Trim$(CStr(valueIn))
    End If
End Function















