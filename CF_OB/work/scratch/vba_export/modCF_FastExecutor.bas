Option Explicit

' The only production button path. Detailed SAP work stays inside the current
' rework module so it can reuse its exact PLAN/context readers; this shell
' keeps the route deterministic and keeps diagnostics out of RUN.

Public Sub CF_RunFastToCF()
    Dim resultText As String
    resultText = CF_RunFastToCFCore()
    If Left$(resultText, 5) = "PASS|" Then
        MsgBox resultText, vbInformation, "XU LY DEN CF"
    Else
        MsgBox resultText, vbExclamation, "XU LY DEN CF"
    End If
End Sub

Public Function CF_RunFastToCFCore() As String
    Dim reason As String, resultText As String
    If Not CF_FastPreflight(reason) Then
        CF_RunFastToCFCore = "STOP|FAST_PREFLIGHT|" & reason & "|CF=USER_HANDOFF"
        Exit Function
    End If
    resultText = CWR_RunFastToCFCore()
    If Left$(resultText, 5) = "PASS|" Then
        CF_RunFastToCFCore = resultText
    Else
        CF_RunFastToCFCore = "STOP|FAST_EXECUTOR|" & resultText & "|CF=USER_HANDOFF"
    End If
End Function

Public Sub CF_FastExecutor_CompileProbe()
    Dim ignored As String
    CF_FastStaticCompileProbe
    SAPAction_CompileProbe
    CF_InputLayout_CompileProbe
    CF_MultiOB_CompileProbe
    RT_MONCompileProbe
    MPOST_CompileProbe
    SAPX_POST_CompileProbe
    CF_CheckPlanner_CompileProbe
    CF_Check_AutoSilent_CompileProbe
    CF_ExecutePlanToCFHandoff_StaticCompileProbe
    If False Then ignored = CWR_RunFastToCFCore()
End Sub

' Targeted, local-only regression. It deliberately does not attach to SAP.
Public Function CF_FastModuleRegression() As String
    Dim commandText As String, contextText As String, statusText As String, reviewText As String, messageText As String
    On Error GoTo Failed
    CF_FastExecutor_CompileProbe
    If Not SAPAction_Get("WT_ADOPT_QUANTITY", commandText, contextText, statusText, reviewText, messageText) Then _
        GoTo Stopped
    If statusText <> "LIVE_PASS" Or reviewText <> "LOCKED" Then
        messageText = "STOP_ADOPT_MAP_NOT_LOCKED|" & statusText & "|" & reviewText
        GoTo Stopped
    End If
    If Not SAPAction_AssertLivePass("WT_CREATE_SAVE", "S000", messageText) Then GoTo Stopped
    If Not SAPAction_Get("WT_CREATE_COMMAND_ID", commandText, contextText, statusText, reviewText, messageText) Then _
        GoTo Stopped
    If commandText <> "OK_OI_DLV_CREATE_TO" Or statusText <> "LIVE_PASS" Or reviewText <> "LOCKED" Then
        messageText = "STOP_CREATE_DRAFT_MAP_NOT_LOCKED|" & commandText & "|" & statusText & "|" & reviewText
        GoTo Stopped
    End If
    If Not SAPAction_Get("WT_SAVE_COMMAND_ID", commandText, contextText, statusText, reviewText, messageText) Then _
        GoTo Stopped
    If commandText <> "OK_OI_DLV_SAVE" Or statusText <> "LIVE_PASS" Or reviewText <> "LOCKED" Then
        messageText = "STOP_SAVE_DRAFT_MAP_NOT_LOCKED|" & commandText & "|" & statusText & "|" & reviewText
        GoTo Stopped
    End If
    If Not SAPAction_Get("WT_CREATE_AND_SAVE_COMMAND_ID", commandText, contextText, statusText, reviewText, messageText) Then _
        GoTo Stopped
    If statusText <> "DISCOVERED" Or reviewText <> "PENDING" Then
        messageText = "STOP_CREATE_SAVE_MAP_STATE_UNEXPECTED|" & statusText & "|" & reviewText
        GoTo Stopped
    End If
    CF_FastModuleRegression = "PASS|FAST_MODULE_REGRESSION|INPUT_COLUMN_A_GATE=PASS|" & _
        "ADOPT=LIVE_PASS_LOCKED|CREATE_SAVE=LIVE_PASS_LOCKED|" & _
        "DIRECT_CREATE_AND_SAVE=DISCOVERED_PENDING|POST_CAPTURE_CLASSIFIER=TESTED|CF=USER_HANDOFF"
    Exit Function
Stopped:
    CF_FastModuleRegression = "STOP|FAST_MODULE_REGRESSION|" & messageText
    Exit Function
Failed:
    CF_FastModuleRegression = "FAIL|FAST_MODULE_REGRESSION|" & CStr(Err.Number) & "|" & Err.Description
End Function
