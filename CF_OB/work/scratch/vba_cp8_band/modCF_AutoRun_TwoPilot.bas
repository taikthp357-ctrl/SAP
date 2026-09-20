Attribute VB_Name = "modCF_AutoRun_TwoPilot"
Option Explicit

Public Sub CF_RunTwoPilotExact()
    Const EXPECTED_OB As String = "1011764693"
    Const EXPECTED_DOC As String = "1012022347"
    Const EXPECTED_WT1 As String = "5042662681"
    Const EXPECTED_WT2 As String = "5042661772"
    Dim reportText As String, messageText As String, target As Object
    Dim runID As String, started As Double, stepNo As Long, completed As Long
    Dim swapped As Boolean

    On Error GoTo Failed
    started = Timer
    runID = Format$(Now, "yyyymmdd_hhnnss") & "_CF2PILOT"

    If ModuleHelper.ToolConfigBool("CF_USER_HANDOFF_REQUIRED", True) Then
        CF_DisarmPilot
        RT_SetStatus "STAGE", "CF HANDOFF - USER"
        RT_Log "CF_BATCH", EXPECTED_OB, EXPECTED_DOC, "STOP", _
                "CF_USER_HANDOFF_REQUIRED=TRUE; no CF command sent; user action required"
        Exit Sub
    End If

    ' Read-only audit first, with the workbook in MAP/offline mode.  The
    ' compile probe is deliberately fail-closed and accepts MAP or RUN, not
    ' PROBE; no SAP control is touched by this audit stage.
    CF_SetConfigValue "SAP_MODE", "MAP"
    CF_SetConfigValue "CF_RUN_LEVEL", "OFFLINE_AUDIT"
    CF_SetConfigValue "CF_LIVE_ENABLED", "FALSE"
    CF_SetConfigValue "ALLOW_CF", "FALSE"
    ModuleHelper.ResetToolConfigCache
    If Not CF_PreLiveAudit_Run(reportText) Then
        Err.Raise vbObjectError + 9961, , "STOP_PRELIVE_AUDIT: " & reportText
    End If

    ' Arm only the reviewed, single action.  The cleanup path below always
    ' returns the workbook to the fail-closed baseline.
    CF_SetConfigValue "SAP_MODE", "RUN"
    CF_SetConfigValue "CF_RUN_LEVEL", "PILOT_LIVE"
    CF_SetConfigValue "CF_LIVE_ENABLED", "TRUE"
    CF_SetConfigValue "ALLOW_CF", "TRUE"
    ModuleHelper.ResetToolConfigCache

    For stepNo = 1 To 2
        Set target = Nothing
        messageText = ""
        If Not CF_PreVerifyTarget(runID, target, messageText) Then
            Err.Raise vbObjectError + 9962, , "STOP_PRE_VERIFY_STEP=" & CStr(stepNo) & ": " & messageText
        End If
        If stepNo = 1 Then
            If CF_TwoPilotTargetText(target, "wt") <> EXPECTED_WT1 Then _
                Err.Raise vbObjectError + 9963, , "STOP_TARGET_ORDER: expected WT=" & EXPECTED_WT1 & _
                    "; actual=" & CF_TwoPilotTargetText(target, "wt")
        Else
            If CF_TwoPilotTargetText(target, "wt") <> EXPECTED_WT2 Then _
                Err.Raise vbObjectError + 9964, , "STOP_TARGET_ORDER: expected WT=" & EXPECTED_WT2 & _
                    "; actual=" & CF_TwoPilotTargetText(target, "wt")
        End If

        messageText = ""
        If Not CF_ExecuteConfirm(runID, target, messageText) Then _
            Err.Raise vbObjectError + 9965, , "STOP_ACTION_STEP=" & CStr(stepNo) & ": " & messageText
        messageText = ""
        If Not CF_PostVerifyTarget(runID, target, messageText) Then _
            Err.Raise vbObjectError + 9966, , "STOP_POST_VERIFY_STEP=" & CStr(stepNo) & ": " & messageText
        messageText = ""
        If Not CF_VerifyStatusC(runID, target, messageText) Then _
            Err.Raise vbObjectError + 9967, , "STOP_STATUS_C_STEP=" & CStr(stepNo) & ": " & messageText

        completed = completed + 1
        If stepNo = 1 Then
            CF_SwapPlanRows
            swapped = True
        End If
    Next stepNo

    If swapped Then
        CF_SwapPlanRows
        swapped = False
    End If
    CF_MarkLocalCFDone EXPECTED_OB, EXPECTED_DOC, EXPECTED_WT1, EXPECTED_WT2
    CF_DisarmPilot
    RT_Log "CF_BATCH", EXPECTED_OB, EXPECTED_DOC, "PASS", _
            "ROWS=2; WT_QTY=10+10; Status=C for both; POST_VERIFY=2; SAP_WRITE=1; RETRY=0"
    Exit Sub

Failed:
    Dim failNo As Long, failDesc As String
    failNo = Err.Number
    failDesc = Err.Description
    On Error Resume Next
    If swapped Then CF_SwapPlanRows
    CF_DisarmPilot
    RT_Log "CF_BATCH", EXPECTED_OB, EXPECTED_DOC, "FAIL", _
            "COMPLETED=" & CStr(completed) & "; STOPPED; SAP_WRITE=1 after first sent action; RETRY=0; " & _
            "Err=" & CStr(failNo) & ": " & failDesc
    ' Do not rethrow through Excel COM: preserve the exact failure in RUN_LOG
    ' so the caller can inspect it without leaving a hidden VBA break state.
    Exit Sub
End Sub

Private Sub CF_SwapPlanRows()
    Dim ws As Worksheet, a As Variant, b As Variant
    Set ws = ThisWorkbook.Worksheets("PLAN")
    a = ws.Range("A2:AA2").Value2
    b = ws.Range("A3:AA3").Value2
    ws.Range("A2:AA2").Value2 = b
    ws.Range("A3:AA3").Value2 = a
End Sub

Private Sub CF_MarkLocalCFDone(ByVal ob As String, ByVal doc As String, _
                               ByVal wt1 As String, ByVal wt2 As String)
    Dim ws As Worksheet, r As Long, wt As String
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    For r = 5 To 504
        If CStr(ws.Cells(r, 17).Value2) = ob And CStr(ws.Cells(r, 18).Value2) = doc Then
            wt = CStr(ws.Cells(r, 5).Value2)
            If wt = wt1 Or wt = wt2 Then
                ws.Cells(r, 10).Value2 = "C"
                ws.Cells(r, 15).Value2 = "DONE"
                ws.Cells(r, 16).Value2 = "POST_VERIFY"
                ws.Cells(r, 18).Value2 = "CF"
                ws.Cells(r, 19).Value2 = "Status=C; exact Background confirmation"
                ws.Cells(r, 26).Value2 = "WT_DONE"
                ws.Cells(r, 29).Value2 = "DONE"
                ws.Cells(r, 30).Value2 = "WT status C: WT is already confirmed."
            End If
        End If
    Next r
End Sub

Private Sub CF_DisarmPilot()
    CF_SetConfigValue "CF_LIVE_ENABLED", "FALSE"
    CF_SetConfigValue "ALLOW_CF", "FALSE"
    CF_SetConfigValue "CF_RUN_LEVEL", "OFFLINE_AUDIT"
    ModuleHelper.ResetToolConfigCache
End Sub

Private Sub CF_SetConfigValue(ByVal keyName As String, ByVal valueText As String)
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("SYSTEM")
    lastRow = ws.Cells(ws.rows.count, 17).End(xlUp).row
    For r = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 17).Value2 & "")), keyName, vbTextCompare) = 0 Then
            ws.Cells(r, 18).Value2 = valueText
            Exit Sub
        End If
    Next r
    Err.Raise vbObjectError + 9968, , "CONFIG_KEY_NOT_FOUND: " & keyName
End Sub

Private Function CF_TwoPilotTargetText(ByVal target As Object, ByVal keyName As String) As String
    If target Is Nothing Then Exit Function
    On Error Resume Next
    If target.Exists(keyName) Then CF_TwoPilotTargetText = Trim$(CStr(target(keyName) & ""))
    Err.Clear
    On Error GoTo 0
End Function
