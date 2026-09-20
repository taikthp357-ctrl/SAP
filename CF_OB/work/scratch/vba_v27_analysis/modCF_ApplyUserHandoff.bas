Attribute VB_Name = "modCF_ApplyUserHandoff"
Option Explicit

Public Sub CF_ApplyUserHandoffMode()
    On Error GoTo Failed
    CF_SetHandoffConfig "CF_USER_HANDOFF_REQUIRED", "TRUE", "BOOL"
    CF_SetHandoffConfig "CF_LIVE_ENABLED", "FALSE", "BOOL"
    CF_SetHandoffConfig "ALLOW_CF", "FALSE", "BOOL"
    CF_SetHandoffConfig "PROBE_ALLOW_CF", "FALSE", "BOOL"
    CF_SetHandoffConfig "CF_RUN_LEVEL", "OFFLINE_AUDIT", "TEXT"
    ModuleHelper.ResetToolConfigCache
    RT_SetStatus "STAGE", "CF HANDOFF - USER"
    RT_Log "CF_HANDOFF", "", "", "PASS", _
            "CF_USER_HANDOFF_REQUIRED=TRUE; qty/WH steps remain enabled; SAP CF=0"
    Exit Sub
Failed:
    RT_Log "CF_HANDOFF", "", "", "FAIL", "Err=" & CStr(Err.Number) & ": " & Err.Description
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub CF_SetHandoffConfig(ByVal keyName As String, ByVal valueText As String, ByVal typeText As String)
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("CONFIG")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value2 & "")), keyName, vbTextCompare) = 0 Then
            ws.Cells(r, 2).Value2 = valueText
            ws.Cells(r, 3).Value2 = typeText
            Exit Sub
        End If
    Next r
    r = lastRow + 1
    ws.Cells(r, 1).Value2 = keyName
    ws.Cells(r, 2).Value2 = valueText
    ws.Cells(r, 3).Value2 = typeText
    ws.Cells(r, 4).Value2 = "TRUE"
    ws.Cells(r, 5).Value2 = "LOCAL_HANDOFF_POLICY"
    ws.Cells(r, 6).Value2 = "Stop before real CF; user confirms WT manually in SAP."
End Sub
