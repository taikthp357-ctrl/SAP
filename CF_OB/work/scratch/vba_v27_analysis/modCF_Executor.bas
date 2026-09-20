Attribute VB_Name = "modCF_Executor"
Option Explicit

' Single-pilot CF executor for the exact MON Warehouse Task flow.
' The executor is intentionally narrow: it selects one exact WT row, opens
' the row Methods menu, chooses the exact Background confirmation caption,
' then proves Status=C by a bounded read-back. No coordinates, SendKeys,
' caption search for navigation, posting, cancel, save, create, or retry.

#If VBA7 Then
Private Declare PtrSafe Sub CFEX_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub CFEX_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

Private Const CF_CONFIRM_TEXT As String = "Confirm WT in Background"
Private Const CF_NAV_TIMEOUT_SEC As Double = 30#
Private Const CF_ACTION_TIMEOUT_SEC As Double = 30#
Private Const CF_POSTVERIFY_TIMEOUT_SEC As Double = 30#
Private Const CF_EPS As Double = 0.000001

Public Sub CF_RunSinglePilot()
    Dim runID As String, messageText As String, target As Object
    Dim ok As Boolean, sent As Boolean, started As Double
    Dim ob As String, doc As String, wt As String, actionName As String

    On Error GoTo Failed
    started = CFEX_StartTimer()
    runID = Format$(Now, "yyyymmdd_hhnnss") & "_CFPILOT"
    actionName = "Confirm WT in Background"
    If CFEX_UserHandoffRequired(messageText) Then
        CFEX_DisarmLive
        RT_SetStatus "STAGE", "CF HANDOFF - USER"
        RT_Log "CF", "", "", "STOP", messageText & "; SAP_WRITE=0; no CF command sent"
        Exit Sub
    End If
    gRT_RunID = runID
    gRT_ActiveBatchRunID = runID
    RT_SetStatus "STAGE", "CF ACTION_START"
    CF_PreLiveLogActionStage runID, "", "ACTION_START", "", "", actionName, "", "", _
                             "Single exact WT pilot", "", "STARTED", "", CFEX_Elapsed(started) * 1000#

    If Not CF_PreLiveAudit_Require(messageText) Then
        CFEX_LogFailure runID, "", "", "", actionName, "PRELIVE_GUARD", messageText, started
        GoTo CleanExit
    End If

    Set target = Nothing
    If Not CF_PreVerifyTarget(runID, target, messageText) Then
        CFEX_LogFailure runID, CFEX_TargetText(target, "ob"), CFEX_TargetText(target, "wt"), _
                        CFEX_TargetText(target, "doc"), actionName, "PRE_VERIFY", messageText, started
        GoTo CleanExit
    End If
    ob = CFEX_TargetText(target, "ob")
    doc = CFEX_TargetText(target, "doc")
    wt = CFEX_TargetText(target, "wt")

    If Not CF_ExecuteConfirm(runID, target, messageText) Then
        CFEX_LogFailure runID, ob, wt, doc, actionName, "ACTION_SENT", messageText, started
        GoTo CleanExit
    End If
    sent = True

    If Not CF_PostVerifyTarget(runID, target, messageText) Then
        CFEX_LogFailure runID, ob, wt, doc, actionName, "POST_VERIFY", messageText, started
        GoTo CleanExit
    End If

    If Not CF_VerifyStatusC(runID, target, messageText) Then
        CFEX_LogFailure runID, ob, wt, doc, actionName, "POST_VERIFY", messageText, started
        GoTo CleanExit
    End If

    CF_PreLiveLogActionStage runID, "P0001", "DONE", ob, wt, actionName, "", _
                             CFEX_TargetText(target, "status"), "Status=C", _
                             CFEX_TargetText(target, "status"), "PASS", "", CFEX_Elapsed(started) * 1000#
    RT_SetStatus "STAGE", "CF DONE - Status=C"
    RT_SetStatus "MON", "CF OK"
    RT_Log "CF", ob, doc, "PASS", "Confirm WT in Background; Status=C; SAP_WRITE=1; POSTING=0; CANCEL=0; SAVE=0; CREATE=0"

CleanExit:
    CFEX_DisarmLive
    Set target = Nothing
    gRT_ActiveBatchRunID = ""
    Exit Sub

Failed:
    On Error Resume Next
    CFEX_LogFailure runID, ob, wt, doc, IIf(actionName = "", "Confirm WT in Background", actionName), _
                    IIf(sent, "POST_VERIFY", "PRE_VERIFY"), _
                    "CF failed: Err " & CStr(Err.Number) & ": " & Err.Description, started
    RT_SetStatus "STAGE", "CF ERROR - stopped"
    RT_Log "CF", ob, doc, "FAIL", "Err " & CStr(Err.Number) & ": " & Err.Description & "; stopped; SAP_WRITE=" & IIf(sent, "1", "0")
    Resume CleanExit
End Sub

Public Function CF_PreVerifyTarget(ByVal runID As String, ByRef target As Object, _
                                   ByRef messageText As String) As Boolean
    Dim wsPlan As Worksheet, wsWH As Worksheet, planRow As Long, whRow As Long
    Dim ob As String, doc As String, wt As String, wh As String, itemNo As String, material As String
    Dim sysQty As Double, targetQty As Double, actualQty As Double
    Dim sysBin As String, targetBin As String, actualBin As String
    Dim sysBatch As String, targetBatch As String, actualBatch As String
    Dim statusLocal As String, obState As String, wtState As String, planRun As String, whRun As String
    Dim ses As Object, grid As Object, rowNo As Long, matchCount As Long, errText As String
    Dim t0 As Double, detail As String

    On Error GoTo Failed
    messageText = ""
    Set target = CreateObject("Scripting.Dictionary")
    target.CompareMode = vbTextCompare
    t0 = CFEX_StartTimer()
    ModuleHelper.ResetToolConfigCache
    If Not SAPRun_RequireRunMode("CF PRE_VERIFY", messageText) Then Exit Function

    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsWH = ThisWorkbook.Worksheets("WH_ACTUAL")
    ob = Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("A7").Value2 & ""))
    planRun = Trim$(CStr(wsPlan.Range("AB2").Value2 & ""))
    whRun = Trim$(CStr(wsWH.Range("W2").Value2 & ""))
    If ob = "" Then messageText = "STOP_PRE_VERIFY_TARGET: INPUT!A7 is blank.": Exit Function
    If planRun = "" Or whRun = "" Or StrComp(planRun, whRun, vbTextCompare) <> 0 Then
        messageText = "STOP_PRE_VERIFY_TARGET: PLAN/WH Run ID mismatch. PLAN=" & planRun & "; WH=" & whRun
        Exit Function
    End If

    For planRow = 2 To wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
        If StrComp(Trim$(CStr(wsPlan.Cells(planRow, 2).Value2 & "")), ob, vbTextCompare) = 0 And _
           UCase$(Trim$(CStr(wsPlan.Cells(planRow, 20).Value2 & ""))) = "CF_ONLY" Then Exit For
    Next planRow
    If planRow > wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row Then
        messageText = "STOP_PRE_VERIFY_TARGET: no CF_ONLY plan row for OB=" & ob
        Exit Function
    End If

    doc = Trim$(CStr(wsPlan.Cells(planRow, 3).Value2 & ""))
    wh = Trim$(CStr(wsPlan.Cells(planRow, 4).Value2 & ""))
    wt = Trim$(CStr(wsPlan.Cells(planRow, 5).Value2 & ""))
    itemNo = Trim$(CStr(wsWH.Cells(CLng(val(wsPlan.Cells(planRow, 27).Value2)), 3).Value2 & ""))
    material = Trim$(CStr(wsPlan.Cells(planRow, 6).Value2 & ""))
    sysQty = CFEX_Number(wsPlan.Cells(planRow, 7).Value2)
    targetQty = CFEX_Number(wsPlan.Cells(planRow, 8).Value2)
    sysBin = Trim$(CStr(wsPlan.Cells(planRow, 9).Value2 & ""))
    targetBin = Trim$(CStr(wsPlan.Cells(planRow, 10).Value2 & ""))
    sysBatch = Trim$(CStr(wsPlan.Cells(planRow, 11).Value2 & ""))
    targetBatch = Trim$(CStr(wsPlan.Cells(planRow, 14).Value2 & ""))
    whRow = CLng(val(wsPlan.Cells(planRow, 27).Value2))
    If whRow < 5 Then messageText = "STOP_PRE_VERIFY_TARGET: WH Actual row is missing.": Exit Function

    If StrComp(Trim$(CStr(wsWH.Cells(whRow, 1).Value2 & "")), ob, vbTextCompare) <> 0 Or _
       StrComp(Trim$(CStr(wsWH.Cells(whRow, 2).Value2 & "")), doc, vbTextCompare) <> 0 Or _
       StrComp(Trim$(CStr(wsWH.Cells(whRow, 4).Value2 & "")), wh, vbTextCompare) <> 0 Or _
       StrComp(Trim$(CStr(wsWH.Cells(whRow, 5).Value2 & "")), wt, vbTextCompare) <> 0 Or _
       StrComp(Trim$(CStr(wsWH.Cells(whRow, 6).Value2 & "")), material, vbTextCompare) <> 0 Then
        messageText = "STOP_PRE_VERIFY_TARGET: WH_ACTUAL identity does not match PLAN."
        Exit Function
    End If
    statusLocal = Trim$(CStr(wsWH.Cells(whRow, 10).Value2 & ""))
    obState = UCase$(Trim$(CStr(wsWH.Cells(whRow, 25).Value2 & "")))
    wtState = UCase$(Trim$(CStr(wsWH.Cells(whRow, 26).Value2 & "")))
    actualQty = CFEX_Number(wsWH.Cells(whRow, 12).Value2)
    actualBin = Trim$(CStr(wsWH.Cells(whRow, 14).Value2 & ""))
    actualBatch = Trim$(CStr(wsWH.Cells(whRow, 22).Value2 & ""))
    If actualBatch = "" And StrComp(actualBin, targetBin, vbTextCompare) = 0 Then actualBatch = sysBatch
    If obState <> "OB_OPEN_OR_PARTIAL" Or wtState <> "WT_ACTIVE" Or statusLocal <> "" Then
        messageText = "STOP_PRE_VERIFY_TARGET: state/status not eligible. OB_STATE=" & obState & _
                      "; WT_STATE=" & wtState & "; Status=" & statusLocal
        Exit Function
    End If
    If Abs(sysQty - targetQty) > CF_EPS Or Abs(actualQty - targetQty) > CF_EPS Then
        messageText = "STOP_PRE_VERIFY_TARGET: qty mismatch. Sys=" & CStr(sysQty) & _
                      "; Target=" & CStr(targetQty) & "; Actual=" & CStr(actualQty)
        Exit Function
    End If
    If StrComp(sysBin, targetBin, vbTextCompare) <> 0 Or StrComp(actualBin, targetBin, vbTextCompare) <> 0 Then
        messageText = "STOP_PRE_VERIFY_TARGET: bin mismatch. Sys=" & sysBin & "; Target=" & targetBin & "; Actual=" & actualBin
        Exit Function
    End If
    If StrComp(sysBatch, targetBatch, vbTextCompare) <> 0 Or StrComp(actualBatch, targetBatch, vbTextCompare) <> 0 Then
        messageText = "STOP_PRE_VERIFY_TARGET: batch mismatch. Sys=" & sysBatch & "; Target=" & targetBatch & "; Actual=" & actualBatch
        Exit Function
    End If
    If UCase$(Trim$(CStr(wsPlan.Cells(planRow, 21).Value2 & ""))) = "TRUE" Or _
       UCase$(Trim$(CStr(wsPlan.Cells(planRow, 22).Value2 & ""))) = "TRUE" Or _
       UCase$(Trim$(CStr(wsPlan.Cells(planRow, 23).Value2 & ""))) = "TRUE" Or _
       Abs(CFEX_Number(wsPlan.Cells(planRow, 17).Value2)) > CF_EPS Then
        messageText = "STOP_PRE_VERIFY_TARGET: planner change/release flag is active."
        Exit Function
    End If

    Set ses = CF_OpenTCodeSession("/SCWM/MON", errText)
    If ses Is Nothing Then messageText = "STOP_PRE_VERIFY_TARGET: " & errText: Exit Function
    If Not CFEX_WaitIdle(ses, CF_NAV_TIMEOUT_SEC, messageText) Then Exit Function
    Set grid = Nothing
    On Error Resume Next
    Set grid = ses.FindById(ModuleHelper.ToolConfig("CF_GRID_ID", ""), False)
    If grid Is Nothing Then Set grid = ses.FindById(ModuleHelper.ToolConfig("MON_WT_GRID_ID", ""), False)
    On Error GoTo Failed
    If grid Is Nothing Then messageText = "STOP_CONTROL_NOT_FOUND: exact CF/WT grid is unavailable.": Exit Function

    rowNo = CFEX_FindExactRow(grid, doc, wh, wt, itemNo, material, targetBatch, targetBin, targetQty, matchCount)
    If matchCount <> 1 Then
        messageText = "STOP_PRE_VERIFY_TARGET: exact WT match count=" & CStr(matchCount) & _
                      "; expected exactly 1. DOC=" & doc & "; WH=" & wh & "; WT=" & wt
        Exit Function
    End If
    statusLocal = CFEX_GridCell(grid, rowNo, ModuleHelper.ToolConfig("MON_WT_STATUS_COL", "TOSTAT"))
    If statusLocal <> "" Then
        messageText = "STOP_PRE_VERIFY_TARGET: live WT is not open; Status=" & statusLocal
        Exit Function
    End If
    On Error Resume Next
    grid.selectedRows = CStr(rowNo)
    grid.SetCurrentCell rowNo, ModuleHelper.ToolConfig("MON_WT_COL", "TANUM")
    If Err.Number <> 0 Then messageText = "STOP_PRE_VERIFY_TARGET: WT row selection failed. Err " & CStr(Err.Number) & ": " & Err.Description: Err.Clear: Exit Function
    On Error GoTo Failed

    target.Add "runID", runID
    target.Add "planRow", planRow
    target.Add "whRow", whRow
    target.Add "ob", ob
    target.Add "doc", doc
    target.Add "wh", wh
    target.Add "wt", wt
    target.Add "item", itemNo
    target.Add "material", material
    target.Add "qty", targetQty
    target.Add "bin", targetBin
    target.Add "batch", targetBatch
    target.Add "status", statusLocal
    target.Add "row", rowNo
    target.Add "grid", grid
    target.Add "session", ses
    detail = "OB=" & ob & "; DOC=" & doc & "; WH=" & wh & "; WT=" & wt & "; ITEM=" & itemNo & _
             "; MAT=" & material & "; QTY=" & CStr(targetQty) & "; BIN=" & targetBin & "; BATCH=" & targetBatch
    CF_PreLiveLogActionStage runID, "P0001", "PRE_VERIFY", ob, wt, CF_CONFIRM_TEXT, "TOSTAT blank", _
                             "TOSTAT blank", detail, "TOSTAT blank", "PASS", "", CFEX_Elapsed(t0) * 1000#
    CF_PreVerifyTarget = True
    Exit Function

Failed:
    messageText = "STOP_PRE_VERIFY_TARGET: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_ExecuteConfirm(ByVal runID As String, ByVal target As Object, _
                                  ByRef messageText As String) As Boolean
    Dim grid As Object, ses As Object, ignored As Variant, started As Double
    Dim confirmText As String, popupText As String

    On Error GoTo Failed
    messageText = ""
    started = CFEX_StartTimer()
    If target Is Nothing Then messageText = "STOP_CF_EXECUTOR: target is missing.": Exit Function
    If CFEX_UserHandoffRequired(messageText) Then Exit Function
    If Not CF_PreLiveActionAllowed("CF", messageText) Then Exit Function
    Set grid = target("grid")
    Set ses = target("session")
    If grid Is Nothing Or ses Is Nothing Then messageText = "STOP_CF_EXECUTOR: exact grid/session is missing.": Exit Function
    confirmText = Trim$(ModuleHelper.ToolConfig("CF_CONFIRM_ID", CF_CONFIRM_TEXT))
    If Left$(UCase$(confirmText), 5) = "TEXT:" Then confirmText = Mid$(confirmText, 6)
    If StrComp(confirmText, CF_CONFIRM_TEXT, vbTextCompare) <> 0 Then
        messageText = "STOP_CF_EXECUTOR: CF_CONFIRM_ID is not the reviewed Background confirmation text.": Exit Function
    End If
    CF_PreLiveLogActionStage runID, "P0001", "ACTION_START", CFEX_TargetText(target, "ob"), _
                             CFEX_TargetText(target, "wt"), CF_CONFIRM_TEXT, "TOSTAT blank", _
                             "TOSTAT blank", "Methods -> " & confirmText, "", "STARTED", "", CFEX_Elapsed(started) * 1000#

    On Error Resume Next
    Err.Clear
    ' Close any stale menu left by a prior read-only probe before opening the
    ' exact lower-grid Methods context menu for this run.
    grid.SelectToolbarMenuItem "__CLOSE_MENU__"
    Err.Clear
    grid.PressToolbarContextButton "METHODS"
    If Err.Number <> 0 Then messageText = "STOP_CF_EXECUTOR: Methods toolbar open failed. Err " & CStr(Err.Number) & ": " & Err.Description: Err.Clear: Exit Function
    Err.Clear
    ignored = grid.SelectContextMenuItemByText(confirmText)
    If Err.Number <> 0 Then messageText = "STOP_CF_EXECUTOR: Background confirm command failed. Err " & CStr(Err.Number) & ": " & Err.Description: Err.Clear: Exit Function
    On Error GoTo Failed

    If Not CFEX_WaitIdle(ses, CF_ACTION_TIMEOUT_SEC, messageText) Then Exit Function
    popupText = CFEX_PopupText(ses)
    If popupText <> "" Then
        If InStr(1, UCase$(popupText), "DISPLAY LOGS", vbTextCompare) > 0 Then
            ' Background confirmation may open SAP's read-only Display logs
            ' result window. Close only this known informational window; do
            ' not accept or press any unknown popup action.
            On Error Resume Next
            ses.FindById("wnd[1]", False).Close
            Err.Clear
            On Error GoTo Failed
        Else
            messageText = "STOP_UNKNOWN_SAP_POPUP: " & popupText
            Exit Function
        End If
    End If
    CF_PreLiveLogActionStage runID, "P0001", "ACTION_SENT", CFEX_TargetText(target, "ob"), _
                             CFEX_TargetText(target, "wt"), CF_CONFIRM_TEXT, "TOSTAT blank", _
                             "ACTION_SENT", "Background confirm sent once", "", "SENT", "", CFEX_Elapsed(started) * 1000#
    CF_ExecuteConfirm = True
    Exit Function

Failed:
    messageText = "STOP_CF_EXECUTOR: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_PostVerifyTarget(ByVal runID As String, ByVal target As Object, _
                                    ByRef messageText As String) As Boolean
    Dim grid As Object, ses As Object, started As Double, rowNo As Long, matches As Long
    Dim statusText As String, detail As String
    On Error GoTo Failed
    messageText = ""
    started = CFEX_StartTimer()
    If target Is Nothing Then messageText = "STOP_POSTVERIFY: target is missing.": Exit Function
    Set grid = target("grid")
    Set ses = target("session")
    If grid Is Nothing Or ses Is Nothing Then messageText = "STOP_POSTVERIFY: exact grid/session is missing.": Exit Function
    Do
        DoEvents
        If Not CFEX_WaitIdle(ses, 1#, messageText) Then Exit Function
        rowNo = CFEX_FindExactRow(grid, CFEX_TargetText(target, "doc"), CFEX_TargetText(target, "wh"), _
                                  CFEX_TargetText(target, "wt"), CFEX_TargetText(target, "item"), _
                                  CFEX_TargetText(target, "material"), CFEX_TargetText(target, "batch"), _
                                  CFEX_TargetText(target, "bin"), CDbl(target("qty")), matches)
        If matches = 1 Then
            statusText = CFEX_GridCell(grid, rowNo, ModuleHelper.ToolConfig("MON_WT_STATUS_COL", "TOSTAT"))
            target("row") = rowNo
            target("status") = statusText
            If UCase$(Trim$(statusText)) = "C" Then
                detail = "Live read-back DOC=" & CFEX_TargetText(target, "doc") & "; WT=" & CFEX_TargetText(target, "wt") & "; TOSTAT=" & statusText
                CF_PreLiveLogActionStage runID, "P0001", "POST_VERIFY", CFEX_TargetText(target, "ob"), _
                                         CFEX_TargetText(target, "wt"), CF_CONFIRM_TEXT, "TOSTAT blank", _
                                         statusText, "Status=C", statusText, "PASS", "", CFEX_Elapsed(started) * 1000#
                CF_PostVerifyTarget = True
                Exit Function
            End If
        ElseIf matches > 1 Then
            messageText = "STOP_POSTVERIFY: exact WT match became ambiguous; count=" & CStr(matches)
            Exit Function
        End If
        If CFEX_Elapsed(started) >= CF_POSTVERIFY_TIMEOUT_SEC Then Exit Do
        CFEX_Sleep 150
    Loop
    messageText = "STOP_POSTVERIFY_TIMEOUT: exact WT did not read back Status=C within " & CStr(CF_POSTVERIFY_TIMEOUT_SEC) & " seconds."
    Exit Function
Failed:
    messageText = "STOP_POSTVERIFY: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_VerifyStatusC(ByVal runID As String, ByVal target As Object, _
                                 ByRef messageText As String) As Boolean
    Dim statusText As String
    On Error GoTo Failed
    messageText = ""
    If target Is Nothing Then messageText = "STOP_CF_STATE_UNPROVEN: target is missing.": Exit Function
    statusText = UCase$(Trim$(CFEX_TargetText(target, "status")))
    If statusText = "C" Then
        CF_VerifyStatusC = True
        Exit Function
    End If
    messageText = "STOP_CF_STATE_UNPROVEN: explicit Status=C proof failed; live status='" & statusText & "'."
    Exit Function
Failed:
    messageText = "STOP_CF_STATE_UNPROVEN: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

' Read-only reconciliation entry point used when SAP has already completed
' the one-shot action but the informational Display logs window interrupted
' the original executor after ACTION_SENT. It never opens Methods and never
' sends a second CF command.
Public Sub CF_PostVerifyExistingPilot()
    Dim messageText As String, runID As String
    runID = Format$(Now, "yyyymmdd_hhnnss") & "_CF_RECONCILE"
    If CF_PostVerifyExistingPilot_Run(runID, messageText) Then
        RT_SetStatus "STAGE", "CF DONE - Status=C"
    Else
        RT_SetStatus "STAGE", "CF POST-VERIFY ERROR"
        RT_Log "CF", "", "", "FAIL", messageText & "; no second CF sent"
    End If
End Sub

Public Function CF_PostVerifyExistingPilot_Run(ByVal runID As String, _
                                               ByRef messageText As String) As Boolean
    Dim wsPlan As Worksheet, wsWH As Worksheet, planRow As Long, whRow As Long
    Dim ob As String, doc As String, wh As String, wt As String, itemNo As String, material As String
    Dim batchText As String, binText As String, qty As Double, ses As Object, grid As Object
    Dim rowNo As Long, matches As Long, statusText As String, errText As String, started As Double

    On Error GoTo Failed
    messageText = "": started = CFEX_StartTimer()
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsWH = ThisWorkbook.Worksheets("WH_ACTUAL")
    ob = Trim$(CStr(ThisWorkbook.Worksheets("INPUT").Range("A7").Value2 & ""))
    For planRow = 2 To wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
        If StrComp(Trim$(CStr(wsPlan.Cells(planRow, 2).Value2 & "")), ob, vbTextCompare) = 0 And _
           UCase$(Trim$(CStr(wsPlan.Cells(planRow, 20).Value2 & ""))) = "CF_ONLY" Then Exit For
    Next planRow
    If planRow > wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row Then
        messageText = "STOP_POSTVERIFY: CF_ONLY plan row not found.": Exit Function
    End If
    whRow = CLng(val(wsPlan.Cells(planRow, 27).Value2))
    If whRow < 5 Then messageText = "STOP_POSTVERIFY: WH Actual row is missing.": Exit Function
    doc = Trim$(CStr(wsPlan.Cells(planRow, 3).Value2 & ""))
    wh = Trim$(CStr(wsPlan.Cells(planRow, 4).Value2 & ""))
    wt = Trim$(CStr(wsPlan.Cells(planRow, 5).Value2 & ""))
    itemNo = Trim$(CStr(wsWH.Cells(whRow, 3).Value2 & ""))
    material = Trim$(CStr(wsPlan.Cells(planRow, 6).Value2 & ""))
    qty = CFEX_Number(wsPlan.Cells(planRow, 8).Value2)
    binText = Trim$(CStr(wsPlan.Cells(planRow, 10).Value2 & ""))
    batchText = Trim$(CStr(wsPlan.Cells(planRow, 14).Value2 & ""))
    If Not SAPRun_RequireRunMode("CF POST_VERIFY", messageText) Then Exit Function
    Set ses = CF_OpenTCodeSession("/SCWM/MON", errText)
    If ses Is Nothing Then messageText = "STOP_POSTVERIFY: " & errText: Exit Function
    If Not CFEX_WaitIdle(ses, CF_POSTVERIFY_TIMEOUT_SEC, messageText) Then Exit Function
    On Error Resume Next
    Set grid = ses.FindById(ModuleHelper.ToolConfig("CF_GRID_ID", ""), False)
    If grid Is Nothing Then Set grid = ses.FindById(ModuleHelper.ToolConfig("MON_WT_GRID_ID", ""), False)
    On Error GoTo Failed
    If grid Is Nothing Then messageText = "STOP_POSTVERIFY: exact WT grid is unavailable.": Exit Function
    rowNo = CFEX_FindExactRow(grid, doc, wh, wt, itemNo, material, batchText, binText, qty, matches)
    If matches <> 1 Then messageText = "STOP_POSTVERIFY: exact WT match count=" & CStr(matches): Exit Function
    statusText = UCase$(Trim$(CFEX_GridCell(grid, rowNo, ModuleHelper.ToolConfig("MON_WT_STATUS_COL", "TOSTAT"))))
    If statusText = "C" Then
        CF_PreLiveLogActionStage runID, "P0001", "POST_VERIFY", ob, wt, CF_CONFIRM_TEXT, "TOSTAT blank", _
                                 statusText, "Status=C", statusText, "PASS", "", CFEX_Elapsed(started) * 1000#
        CF_PreLiveLogActionStage runID, "P0001", "DONE", ob, wt, CF_CONFIRM_TEXT, "TOSTAT blank", _
                                 "Status=C", "Status=C", statusText, "PASS", "", CFEX_Elapsed(started) * 1000#
        RT_Log "CF", ob, doc, "PASS", "POST_VERIFY confirmed Status=C; no second CF sent"
        CF_PostVerifyExistingPilot_Run = True
    Else
        messageText = "STOP_CF_STATE_UNPROVEN: explicit Status=C proof failed; live status='" & statusText & "'."
    End If
    Exit Function
Failed:
    messageText = "STOP_POSTVERIFY: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CFEX_FindExactRow(ByVal grid As Object, ByVal doc As String, ByVal wh As String, _
                                   ByVal wt As String, ByVal itemNo As String, ByVal material As String, _
                                   ByVal batchText As String, ByVal binText As String, ByVal qty As Double, _
                                   ByRef matchCount As Long) As Long
    Dim r As Long, rows As Long, rowDoc As String, rowWH As String, rowWT As String
    Dim rowItem As String, rowMat As String, rowBatch As String, rowBin As String, rowQty As Double
    matchCount = 0
    If grid Is Nothing Then Exit Function
    rows = CLng(grid.rowCount)
    For r = 0 To rows - 1
        rowDoc = CFEX_GridCell(grid, r, ModuleHelper.ToolConfig("MON_WT_DOC_COL", "DOCNO"))
        rowWH = CFEX_GridCell(grid, r, ModuleHelper.ToolConfig("MON_WT_WHO_COL", "WHO"))
        rowWT = CFEX_GridCell(grid, r, ModuleHelper.ToolConfig("MON_WT_COL", "TANUM"))
        rowItem = CFEX_GridCell(grid, r, ModuleHelper.ToolConfig("MON_WT_ITEM_COL", "ITEMNO"))
        rowMat = CFEX_GridCell(grid, r, ModuleHelper.ToolConfig("MON_WT_MATERIAL_COL", "MATNR"))
        rowBatch = CFEX_GridCell(grid, r, ModuleHelper.ToolConfig("MON_WT_BATCH_COL", "CHARG"))
        rowBin = CFEX_GridCell(grid, r, ModuleHelper.ToolConfig("MON_WT_BIN_COL", "VLPLA"))
        rowQty = CFEX_Number(CFEX_GridCell(grid, r, ModuleHelper.ToolConfig("MON_WT_QTY_COL", "VSOLM")))
        If StrComp(rowDoc, doc, vbTextCompare) = 0 And StrComp(rowWH, wh, vbTextCompare) = 0 And _
           StrComp(rowWT, wt, vbTextCompare) = 0 And StrComp(rowItem, itemNo, vbTextCompare) = 0 And _
           StrComp(rowMat, material, vbTextCompare) = 0 And StrComp(rowBatch, batchText, vbTextCompare) = 0 And _
           StrComp(rowBin, binText, vbTextCompare) = 0 And Abs(rowQty - qty) <= CF_EPS Then
            matchCount = matchCount + 1
            CFEX_FindExactRow = r
        End If
    Next r
End Function

Private Function CFEX_GridCell(ByVal grid As Object, ByVal rowNo As Long, ByVal columnKey As String) As String
    Dim v As Variant
    If grid Is Nothing Or Trim$(columnKey) = "" Then Exit Function
    On Error Resume Next
    Err.Clear
    v = grid.GetCellValue(rowNo, columnKey)
    If Err.Number = 0 Then
        If Not IsNull(v) And Not IsEmpty(v) Then CFEX_GridCell = Trim$(CStr(v))
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function CFEX_Number(ByVal valueIn As Variant) As Double
    If IsNumeric(valueIn) Then
        CFEX_Number = CDbl(valueIn)
    Else
        CFEX_Number = RT_Number(valueIn)
    End If
End Function

Private Function CFEX_TargetText(ByVal target As Object, ByVal keyName As String) As String
    If target Is Nothing Then Exit Function
    On Error Resume Next
    If target.Exists(keyName) Then CFEX_TargetText = Trim$(CStr(target(keyName) & ""))
    Err.Clear
    On Error GoTo 0
End Function

Private Function CFEX_WaitIdle(ByVal ses As Object, ByVal timeoutSeconds As Double, _
                               ByRef messageText As String) As Boolean
    Dim started As Double, busy As Boolean
    messageText = ""
    If ses Is Nothing Then messageText = "STOP_CF_SESSION: session is missing.": Exit Function
    started = CFEX_StartTimer()
    Do
        DoEvents
        On Error Resume Next
        Err.Clear
        busy = CBool(ses.busy)
        If Err.Number <> 0 Then messageText = "STOP_CF_SESSION: Busy read failed. Err " & CStr(Err.Number) & ": " & Err.Description: Err.Clear: Exit Function
        On Error GoTo 0
        If Not busy Then CFEX_WaitIdle = True: Exit Function
        If CFEX_Elapsed(started) >= timeoutSeconds Then Exit Do
        CFEX_Sleep 100
    Loop
    messageText = "STOP_CF_TIMEOUT: SAP session stayed busy beyond " & CStr(timeoutSeconds) & " seconds."
End Function

Private Function CFEX_PopupText(ByVal ses As Object) As String
    Dim popup As Object, v As Variant
    If ses Is Nothing Then Exit Function
    On Error Resume Next
    Set popup = ses.FindById("wnd[1]", False)
    If Not popup Is Nothing Then
        v = popup.text
        If Not IsNull(v) Then CFEX_PopupText = Trim$(CStr(v))
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub CFEX_LogFailure(ByVal runID As String, ByVal ob As String, ByVal wt As String, _
                            ByVal doc As String, ByVal actionName As String, ByVal stageName As String, _
                            ByVal messageText As String, ByVal started As Double)
    On Error Resume Next
    CF_PreLiveLogActionStage runID, "P0001", stageName, ob, wt, actionName, "", "", "", "", "FAIL", messageText, CFEX_Elapsed(started) * 1000#
    RT_Log "CF", ob, doc, "FAIL", messageText & "; stopped; SAP_WRITE=0/1 depending on ACTION_SENT"
    On Error GoTo 0
End Sub

Private Sub CFEX_DisarmLive()
    On Error Resume Next
    CFEX_SetConfig "CF_LIVE_ENABLED", "FALSE"
    CFEX_SetConfig "ALLOW_CF", "FALSE"
    CFEX_SetConfig "CF_RUN_LEVEL", "OFFLINE_AUDIT"
    ModuleHelper.ResetToolConfigCache
    On Error GoTo 0
End Sub

Private Function CFEX_UserHandoffRequired(ByRef messageText As String) As Boolean
    On Error GoTo FailClosed
    messageText = ""
    If ModuleHelper.ToolConfigBool("CF_USER_HANDOFF_REQUIRED", True) Then
        messageText = "CF_USER_HANDOFF: real CF is intentionally stopped; user must confirm WT in SAP."
        CFEX_UserHandoffRequired = True
    End If
    Exit Function
FailClosed:
    messageText = "CF_USER_HANDOFF: handoff flag could not be read; CF is stopped fail-closed."
    CFEX_UserHandoffRequired = True
End Function

Private Sub CFEX_SetConfig(ByVal keyName As String, ByVal valueText As String)
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("CONFIG")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value2 & "")), keyName, vbTextCompare) = 0 Then
            ws.Cells(r, 2).Value2 = valueText
            Exit Sub
        End If
    Next r
End Sub

Private Function CFEX_StartTimer() As Double
    CFEX_StartTimer = Timer
End Function

Private Function CFEX_Elapsed(ByVal started As Double) As Double
    If Timer >= started Then CFEX_Elapsed = Timer - started Else CFEX_Elapsed = 86400# - started + Timer
End Function
