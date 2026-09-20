Attribute VB_Name = "modSAPActionRegistry"
Option Explicit

' The production resolver for SAP_ACTION_MAP. It never discovers controls and
' treats a non-locked action as a hard stop unless a separately armed,
' one-shot action is explicitly present in the map.

Private Const SHEET_ACTION As String = "SYSTEM"
Private Const COL_KEY As Long = 27
Private Const COL_CONTEXT As Long = 29
Private Const COL_COMMAND As Long = 31
Private Const COL_STATUS As Long = 32
Private Const COL_EVIDENCE As Long = 33
Private Const COL_REVIEW As Long = 35
Private Const COL_NOTES As Long = 36

Public Function SAPAction_Command(ByVal actionKey As String, _
                                  Optional ByVal requiredContext As String = "") As String
    SAPAction_Command = SAPAction_CommandCore(actionKey, requiredContext, False)
End Function

Public Function SAPAction_CommandOneShot(ByVal actionKey As String, _
                                         Optional ByVal requiredContext As String = "") As String
    SAPAction_CommandOneShot = SAPAction_CommandCore(actionKey, requiredContext, True)
End Function

Public Function SAPAction_IsLivePass(ByVal actionKey As String) As Boolean
    Dim commandText As String, contextText As String, statusText As String, reviewText As String, messageText As String
    If SAPAction_Get(actionKey, commandText, contextText, statusText, reviewText, messageText) Then
        SAPAction_IsLivePass = (statusText = "LIVE_PASS" And reviewText = "LOCKED")
    End If
End Function

' Validates a logical/composite production action without treating its
' Technical ID as a SAP toolbar command.  Direct component commands still go
' through SAPAction_Command before they are pressed.
Public Function SAPAction_AssertLivePass(ByVal actionKey As String, _
                                         ByVal requiredContext As String, _
                                         ByRef messageText As String) As Boolean
    Dim ignored As String
    On Error GoTo Failed
    messageText = vbNullString
    ignored = SAPAction_Command(actionKey, requiredContext)
    SAPAction_AssertLivePass = True
    Exit Function
Failed:
    messageText = Err.Description
End Function

Public Function SAPAction_Get(ByVal actionKey As String, ByRef commandText As String, _
                              ByRef contextText As String, ByRef statusText As String, _
                              ByRef reviewText As String, ByRef messageText As String) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long, matches As Long
    On Error GoTo Failed
    commandText = vbNullString: contextText = vbNullString
    statusText = vbNullString: reviewText = vbNullString: messageText = vbNullString
    Set ws = ThisWorkbook.Worksheets(SHEET_ACTION)
    If StrComp(Trim$(CStr(ws.Cells(1, COL_KEY).Value2)), "ActionKey", vbTextCompare) <> 0 Then
        messageText = "STOP_ACTION_MAP_HEADER"
        Exit Function
    End If
    lastRow = ws.Cells(ws.rows.count, COL_KEY).End(xlUp).row
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, COL_KEY).Value2)), Trim$(actionKey), vbTextCompare) = 0 Then
            matches = matches + 1
            commandText = Trim$(CStr(ws.Cells(r, COL_COMMAND).Value2))
            contextText = Trim$(CStr(ws.Cells(r, COL_CONTEXT).Value2))
            statusText = UCase$(Trim$(CStr(ws.Cells(r, COL_STATUS).Value2)))
            reviewText = UCase$(Trim$(CStr(ws.Cells(r, COL_REVIEW).Value2)))
        End If
    Next r
    If matches <> 1 Then
        messageText = "STOP_ACTION_KEY_MATCH_COUNT|" & Trim$(actionKey) & "|" & CStr(matches)
        Exit Function
    End If
    If commandText = vbNullString Then
        messageText = "STOP_ACTION_NOT_REGISTERED|" & Trim$(actionKey)
        Exit Function
    End If
    SAPAction_Get = True
    Exit Function
Failed:
    messageText = "STOP_ACTION_MAP_READ_ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function SAPAction_PromoteLivePass(ByVal actionKey As String, ByVal evidenceText As String, _
                                           ByRef messageText As String) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long, matches As Long
    Dim priorEvidence As String, priorNotes As String
    On Error GoTo Failed
    messageText = vbNullString
    Set ws = ThisWorkbook.Worksheets(SHEET_ACTION)
    lastRow = ws.Cells(ws.rows.count, COL_KEY).End(xlUp).row
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, COL_KEY).Value2)), Trim$(actionKey), vbTextCompare) = 0 Then
            matches = matches + 1
            priorEvidence = Trim$(CStr(ws.Cells(r, COL_EVIDENCE).Value2))
            priorNotes = Trim$(CStr(ws.Cells(r, COL_NOTES).Value2))
            ws.Cells(r, COL_STATUS).Value2 = "LIVE_PASS"
            If priorEvidence = vbNullString Then
                ws.Cells(r, COL_EVIDENCE).Value2 = evidenceText
            Else
                ws.Cells(r, COL_EVIDENCE).Value2 = priorEvidence & " | " & evidenceText
            End If
            ws.Cells(r, COL_REVIEW).Value2 = "LOCKED"
            If priorNotes = vbNullString Then ws.Cells(r, COL_NOTES).Value2 = "Promoted only after exact post-verify."
        End If
    Next r
    If matches <> 1 Then
        messageText = "STOP_ACTION_PROMOTE_MATCH_COUNT|" & Trim$(actionKey) & "|" & CStr(matches)
        Exit Function
    End If
    SAPAction_PromoteLivePass = True
    Exit Function
Failed:
    messageText = "STOP_ACTION_PROMOTE_ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Sub SAPAction_CompileProbe()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_ACTION)
    If StrComp(Trim$(CStr(ws.Cells(1, COL_KEY).Value2)), "ActionKey", vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 9721, , "ACTION_MAP_KEY_HEADER"
    If InStr(1, UCase$(Trim$(CStr(ws.Cells(1, COL_COMMAND).Value2))), "TECHNICAL", vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 9722, , "ACTION_MAP_COMMAND_HEADER"
    If StrComp(Trim$(CStr(ws.Cells(1, COL_STATUS).Value2)), "Status", vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 9723, , "ACTION_MAP_STATUS_HEADER"
End Sub

Private Function SAPAction_CommandCore(ByVal actionKey As String, ByVal requiredContext As String, _
                                       ByVal allowOneShot As Boolean) As String
    Dim commandText As String, contextText As String, statusText As String, reviewText As String, messageText As String
    If Not SAPAction_Get(actionKey, commandText, contextText, statusText, reviewText, messageText) Then _
        Err.Raise vbObjectError + 9731, "SAPAction_Command", messageText
    If Len(requiredContext) > 0 And StrComp(contextText, requiredContext, vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 9732, "SAPAction_Command", "STOP_ACTION_CONTEXT_MISMATCH|" & actionKey & "|" & contextText
    Select Case statusText
        Case "LIVE_PASS"
            If reviewText <> "LOCKED" Then Err.Raise vbObjectError + 9733, "SAPAction_Command", _
                "STOP_ACTION_LIVE_PASS_NOT_LOCKED|" & actionKey
        Case "VERIFIED"
            If reviewText <> "REVIEWED" And reviewText <> "LOCKED" Then Err.Raise vbObjectError + 9734, _
                "SAPAction_Command", "STOP_ACTION_NOT_REVIEWED|" & actionKey
        Case "DISCOVERED"
            If Not allowOneShot Or reviewText <> "ONE_SHOT_APPROVED" Or _
               Not CF_ConfigBool("FAST_WT_CREATE_SAVE_ONE_SHOT_ARMED", False) Then _
                Err.Raise vbObjectError + 9735, "SAPAction_Command", "STOP_ACTION_NOT_VERIFIED|" & actionKey & "|DISCOVERED"
        Case Else
            Err.Raise vbObjectError + 9736, "SAPAction_Command", "STOP_ACTION_NOT_VERIFIED|" & actionKey & "|" & statusText
    End Select
    SAPAction_CommandCore = commandText
End Function

