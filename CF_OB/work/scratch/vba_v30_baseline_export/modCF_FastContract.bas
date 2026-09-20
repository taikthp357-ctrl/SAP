Attribute VB_Name = "modCF_FastContract"
Option Explicit

' CF OB fast production contract. This is the only high-level source a
' production executor needs before it reaches SAP_ACTION_MAP.

Public Const CF_FAST_CONTRACT_VERSION As String = "2026.08.16.02"

Public Const ST_PLAN_READY As String = "PLAN_READY"
Public Const ST_SOURCE_FOUND As String = "SOURCE_ROW_FOUND"
Public Const ST_QTY_SET As String = "SOURCE_QTY_SET"
Public Const ST_SOURCE_SELECTED As String = "SOURCE_ROW_SELECTED"
Public Const ST_ADOPT_SENT As String = "ADOPT_QUANTITY_SENT"
Public Const ST_ADOPT_PASS As String = "ADOPT_MINIMAL_VERIFY_PASS"
Public Const ST_CREATE_SAVE_SENT As String = "CREATE_SAVE_SENT"
Public Const ST_WT_CREATED As String = "WT_CREATED"
Public Const ST_READY_CF As String = "READY_FOR_CF"

Public Const AK_MON_OPEN_WT As String = "MON_OPEN_WT"
Public Const AK_MON_CANCEL_WT As String = "MON_CANCEL_WT"
Public Const AK_POST_CREATE_SAVE As String = "POST_CHANGE_CREATE_SAVE"
Public Const AK_WT_CREATE As String = "WT_CREATE"
Public Const AK_WT_SAVE As String = "WT_SAVE"
Public Const AK_WT_CREATE_SAVE As String = "WT_CREATE_SAVE"
Public Const AK_WT_CREATE_FOLLOWON As String = "WT_CREATE_FOLLOWON"
Public Const AK_WT_CREATE_QUERY As String = "WT_CREATE_QUERY"
Public Const AK_WT_CREATE_COMMAND As String = "WT_CREATE_COMMAND_ID"
Public Const AK_WT_SAVE_COMMAND As String = "WT_SAVE_COMMAND_ID"
Public Const AK_WT_STOCK_TAB As String = "WT_TAB_STK_CAN_BE_REMOVED"
Public Const AK_WT_ADOPT_QTY As String = "WT_ADOPT_QUANTITY"

Public Function CF_FastEnabled() As Boolean
    CF_FastEnabled = CF_ConfigBool("FAST_FLOW_ENABLED", False)
End Function

Public Function CF_FastProductionMode() As Boolean
    CF_FastProductionMode = (UCase$(CF_ConfigValue("SAP_MODE")) = "RUN")
End Function

Public Function CF_RealCFBlocked() As Boolean
    ' Confirmation remains a user-owned action in every fast-flow state.
    CF_RealCFBlocked = CF_ConfigBool("FAST_CF_HANDOFF", False)
End Function

Public Function CF_AllowDiagnosticDiscovery() As Boolean
    CF_AllowDiagnosticDiscovery = Not CF_FastProductionMode()
End Function

Public Function CF_AllowAutomaticRetry() As Boolean
    CF_AllowAutomaticRetry = False
End Function

Public Function CF_FastFlowForDecision(ByVal decision As String) As String
    Select Case UCase$(Trim$(decision))
        Case "CF_ONLY"
            CF_FastFlowForDecision = ST_READY_CF
        Case "CHANGE_QTY"
            CF_FastFlowForDecision = "QTY>CREATE_WT>" & ST_READY_CF
        Case "CHANGE_BIN"
            CF_FastFlowForDecision = "CANCEL>CREATE_WT>" & ST_READY_CF
        Case "CHANGE_BATCH", "CHANGE_BIN_BATCH"
            CF_FastFlowForDecision = "CANCEL>POST_BATCH>CREATE_WT>" & ST_READY_CF
    End Select
End Function

Public Function CF_ConfigValue(ByVal keyName As String) As String
    Dim ws As Worksheet, lastRow As Long, r As Long
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("SYSTEM")
    lastRow = ws.Cells(ws.rows.count, 17).End(xlUp).row
    For r = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 17).Value2)), Trim$(keyName), vbTextCompare) = 0 Then
            CF_ConfigValue = Trim$(CStr(ws.Cells(r, 18).Value2))
            Exit Function
        End If
    Next r
ExitHere:
End Function

Public Function CF_ConfigBool(ByVal keyName As String, ByVal defaultValue As Boolean) As Boolean
    Dim valueText As String
    valueText = UCase$(CF_ConfigValue(keyName))
    Select Case valueText
        Case "TRUE", "YES", "1": CF_ConfigBool = True
        Case "FALSE", "NO", "0": CF_ConfigBool = False
        Case Else: CF_ConfigBool = defaultValue
    End Select
End Function

' Local-only gate. It resolves only this route's controls and never scans a
' SAP tree, opens a diagnostic session, or permits CF.
Public Function CF_FastPreflight(ByRef reason As String) As Boolean
    Dim ignored As String
    On Error GoTo Failed
    reason = vbNullString
    If Not CF_FastEnabled() Then
        reason = "STOP_FAST_FLOW_DISABLED"
        Exit Function
    End If
    If Not CF_FastProductionMode() Then
        reason = "STOP_NOT_RUN_MODE"
        Exit Function
    End If
    If Not CF_RealCFBlocked() Then
        reason = "STOP_CF_HANDOFF_POLICY_INVALID"
        Exit Function
    End If
    ignored = SAPAction_Command("WT_REQUEST_GRID", "S000")
    ignored = SAPAction_Command(AK_WT_STOCK_TAB, "S000")
    ignored = SAPAction_Command("WT_STOCK_GRID", "S000")
    ignored = SAPAction_Command("WT_STOCK_TOOLBAR", "S000")
    ignored = SAPAction_Command("WT_STOCK_SET_QTY", "S000")
    ignored = SAPAction_Command("WT_STOCK_SELECT_ROW", "S000")
    ignored = SAPAction_Command(AK_WT_ADOPT_QTY, "S000")
    ignored = SAPAction_Command("WT_TAB_WAREHOUSE_TASK", "S000")
    ignored = SAPAction_Command("WT_WAREHOUSE_TASK_GRID", "S000")
    ignored = SAPAction_Command("WT_REQUEST_TOOLBAR", "S000")
    ignored = SAPAction_Command(AK_WT_CREATE_FOLLOWON, "/SCWM/PRDO")
    ignored = SAPAction_Command(AK_WT_CREATE_QUERY, "S000")
    ignored = SAPAction_Command(AK_WT_CREATE_COMMAND, "S000")
    ignored = SAPAction_Command(AK_WT_SAVE_COMMAND, "S000")
    If Not SAPAction_AssertLivePass(AK_WT_CREATE_SAVE, "S000", reason) Then Exit Function
    reason = "PASS"
    CF_FastPreflight = True
    Exit Function
Failed:
    reason = Err.Description
End Function

' Keep the newest execution checkpoint cheap to recover: only FAST rows are
' written, so the executor never needs to parse the historical diagnostic log.
Public Sub CF_LogFast(ByVal stage As String, ByVal detail As String)
    Dim ws As Worksheet, r As Long
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("RUN_LOG")
    r = ws.Cells(ws.rows.count, 17).End(xlUp).row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 17).Value2 = Now
    ws.Cells(r, 18).Value2 = "FAST"
    ws.Cells(r, 3).Value2 = stage
    ws.Cells(r, 4).Value2 = detail
ExitHere:
End Sub

Public Sub CF_FastStaticCompileProbe()
    If CF_FastFlowForDecision("CHANGE_BIN_BATCH") <> "CANCEL>POST_BATCH>CREATE_WT>READY_FOR_CF" Then _
        Err.Raise vbObjectError + 9711, , "FAST_CONTRACT_ROUTE_INVALID"
    If CF_AllowAutomaticRetry() Then Err.Raise vbObjectError + 9712, , "FAST_RETRY_MUST_BE_OFF"
End Sub

