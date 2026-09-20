Attribute VB_Name = "modCFResetUI"
' ============================================================================
' CF Reset UI  -  the "LAM MOI" (button #4) entry point (V2.7 CP8)
' ============================================================================
' The 4-button contract is: 1.TAI WH -> 2.CHECK -> 3.XU LY -> 4.LAM MOI.
' LAM MOI returns the tool to a clean pre-CHECK state WITHOUT touching SAP and
' WITHOUT destroying history: it clears the volatile plan-lock / execution
' context and the transient status cell, so the next CHECK starts fresh and no
' stale locked PlanID can be executed.
'
' Fail-safe: read-only against SAP. Never deletes RUN_HISTORY/RUN_LOG. Never
' clears INPUT OB list (user's typed OBs) - only the derived status cell.
' ============================================================================
Option Explicit

Public Sub CF_ResetUI()
    Dim msg As String
    If CF_ResetUICore(msg) Then
        On Error Resume Next
        MsgBox "LAM MOI xong: " & msg & vbCrLf & _
               "SAP_WRITE=0. Lich su duoc giu nguyen. Chay 1.TAI WH / 2.CHECK lai.", _
               vbInformation, "LAM MOI"
        On Error GoTo 0
        CF_RefreshInputWorklist
    Else
        On Error Resume Next
        MsgBox "LAM MOI gap loi: " & msg, vbExclamation, "LAM MOI"
        On Error GoTo 0
    End If
End Sub

' Core (no dialog) so it can be unit-tested offline.
Public Function CF_ResetUICore(ByRef messageText As String) As Boolean
    Dim clearedPlanLock As Boolean, clearedStatus As Boolean
    On Error GoTo Failed

    ' 1) Invalidate any locked plan / execution context so a stale PlanID can
    '    never be executed after LAM MOI. Reuse the scheduler's own reset if
    '    present; otherwise clear the PLAN lock cells directly.
    clearedPlanLock = CF_ResetUI_InvalidatePlanLock()

    ' 2) Clear the transient status/result cell on INPUT (derived, not user data).
    clearedStatus = CF_ResetUI_ClearStatusCell()

    messageText = "PLAN_LOCK_INVALIDATED=" & CStr(clearedPlanLock) & _
                  "|STATUS_CLEARED=" & CStr(clearedStatus) & _
                  "|HISTORY_PRESERVED=TRUE|SAP_WRITE=0"
    CF_ResetUICore = True
    Exit Function
Failed:
    messageText = "RESET_ERROR|" & CStr(Err.Number) & ":" & Err.Description
End Function

' Invalidate the CHECK-issued plan lock so EXECUTE refuses until a fresh CHECK.
' Sets PLAN!AJ2 away from LOCKED (the exact cell CWR_AssertQueuePlanLock reads).
Private Function CF_ResetUI_InvalidatePlanLock() As Boolean
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = CF_ResetUI_Sheet("PLAN")
    If ws Is Nothing Then Exit Function
    ' Do NOT wipe the plan rows (audit/debug); only break the lock state so the
    ' guard forces a re-CHECK. AJ2="RECHECK_REQUIRED_AFTER_RESET".
    ws.Range("AJ2").Value2 = "RECHECK_REQUIRED_AFTER_RESET"
    CF_ResetUI_InvalidatePlanLock = True
    Exit Function
Failed:
    CF_ResetUI_InvalidatePlanLock = False
End Function

Private Function CF_ResetUI_ClearStatusCell() As Boolean
    Dim ws As Worksheet, c As Range
    On Error GoTo Failed
    Set ws = CF_ResetUI_Sheet("INPUT")
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    ws.Range("H1:N3").ClearContents
    If Err.Number <> 0 Then
        Err.Clear
        For Each c In ws.Range("H1:N3")
            c.Value2 = Empty
        Next c
    End If
    On Error GoTo 0
    CF_ResetUI_ClearStatusCell = True
    Exit Function
Failed:
    CF_ResetUI_ClearStatusCell = False
End Function

Private Function CF_ResetUI_Sheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    Set CF_ResetUI_Sheet = ws
End Function

' ============================================================================
' Offline self-test
' ============================================================================
Public Function CFResetUI_SelfTest() As String
    Dim ws As Worksheet, wsInput As Worksheet, msg As String, origAJ2 As Variant
    Dim origStatus As Variant
    On Error GoTo Failed
    Set ws = CF_ResetUI_Sheet("PLAN")
    Set wsInput = CF_ResetUI_Sheet("INPUT")
    If ws Is Nothing Then CFResetUI_SelfTest = "FAIL|NO_PLAN_SHEET": Exit Function
    If wsInput Is Nothing Then CFResetUI_SelfTest = "FAIL|NO_INPUT_SHEET": Exit Function
    origAJ2 = ws.Range("AJ2").Value2
    origStatus = wsInput.Range("H1:N3").Value2

    ' Seed only in-memory state and restore it before return. This proves the
    ' status-clear contract instead of merely accepting a False status result.
    wsInput.Range("N3").Value2 = "CFRESETUI_SELFTEST_SENTINEL"
    If Not CF_ResetUICore(msg) Then GoTo CoreFailed
    If UCase$(CStr(ws.Range("AJ2").Value2 & "")) = "LOCKED" Then GoTo PlanStillLocked
    If CStr(wsInput.Range("N3").Value2 & "") <> "" Then GoTo StatusStillPresent

    ws.Range("AJ2").Value2 = origAJ2
    wsInput.Range("H1:N3").Value2 = origStatus
    CFResetUI_SelfTest = "PASS|" & msg
    Exit Function
CoreFailed:
    ws.Range("AJ2").Value2 = origAJ2
    wsInput.Range("H1:N3").Value2 = origStatus
    CFResetUI_SelfTest = "FAIL|CORE|" & msg
    Exit Function
PlanStillLocked:
    ws.Range("AJ2").Value2 = origAJ2
    wsInput.Range("H1:N3").Value2 = origStatus
    CFResetUI_SelfTest = "FAIL|PLAN_STILL_LOCKED_AFTER_RESET"
    Exit Function
StatusStillPresent:
    ws.Range("AJ2").Value2 = origAJ2
    wsInput.Range("H1:N3").Value2 = origStatus
    CFResetUI_SelfTest = "FAIL|STATUS_STILL_PRESENT_AFTER_RESET"
    Exit Function
Failed:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Range("AJ2").Value2 = origAJ2
    If Not wsInput Is Nothing Then wsInput.Range("H1:N3").Value2 = origStatus
    On Error GoTo 0
    CFResetUI_SelfTest = "FAIL|ERR=" & CStr(Err.Number) & ":" & Err.Description
End Function
