Attribute VB_Name = "modCF_InputWorklist"
Option Explicit

' V3.0 operator worklist.  Completed WT rows stay in DATA_CORE/history but are
' hidden from the action table so the user sees only work still requiring a
' decision.  O is a hidden immutable Source Key used for safe write-back.

Private Const WL_CORE As String = "DATA_CORE"
Private Const WL_CONTROL As String = "CF_CONTROL"
Private Const WL_QUEUE As String = "OB_QUEUE"
Private Const WL_FIRST As Long = 9
Private Const WL_LAST As Long = 200
Private Const WL_KEY_COL As Long = 15

Private Const DC_OB As Long = 1
Private Const DC_WHO As Long = 4
Private Const DC_STATUS As Long = 10
Private Const DC_AREA As Long = 11
Private Const DC_ACTUAL_QTY As Long = 12
Private Const DC_BIN_NO As Long = 13
Private Const DC_RESULT As Long = 18
Private Const DC_ERROR As Long = 19
Private Const DC_SOURCE_KEY As Long = 20
Private Const DC_OB_STATE As Long = 25
Private Const DC_WT_STATE As Long = 26
Private Const DC_REMAINING As Long = 27
Private Const DC_QTY_MODE As Long = 28
Private Const DC_EXEC_MODE As Long = 29
Private Const DC_EXEC_STATE As Long = 30

Public Sub CF_RefreshInputWorklist()
    Dim wsCore As Worksheet, wsUI As Worksheet
    Dim lastCore As Long, coreRow As Long, uiRow As Long, rowCount As Long
    Dim vals As Variant
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Set wsCore = ThisWorkbook.Worksheets(WL_CORE)
    Set wsUI = ThisWorkbook.Worksheets(WL_CONTROL)
    wsUI.Range("E" & WL_FIRST & ":O" & WL_LAST).ClearContents

    lastCore = wsCore.Cells(wsCore.rows.count, DC_OB).End(xlUp).row
    If lastCore > 504 Then lastCore = 504
    If lastCore >= 5 Then
        vals = wsCore.Range("A5:AD" & lastCore).Value2
        uiRow = WL_FIRST
        For coreRow = 1 To UBound(vals, 1)
            If uiRow > WL_LAST Then Exit For
            If WL_Text(vals(coreRow, DC_OB)) <> "" Then
                If Not WL_IsCompleted(coreRow, vals) Then
                    wsUI.Cells(uiRow, 5).Value2 = vals(coreRow, DC_OB)
                    wsUI.Cells(uiRow, 6).Value2 = vals(coreRow, DC_WHO)
                    wsUI.Cells(uiRow, 7).Value2 = WL_StatusText(coreRow, vals)
                    wsUI.Cells(uiRow, 8).Value2 = WL_PlanText(coreRow, vals)
                    wsUI.Cells(uiRow, 9).Value2 = WL_GateText(coreRow, vals)
                    wsUI.Cells(uiRow, 10).Value2 = vals(coreRow, DC_ACTUAL_QTY)
                    wsUI.Cells(uiRow, 11).Value2 = vals(coreRow, DC_AREA)
                    wsUI.Cells(uiRow, 12).Value2 = vals(coreRow, DC_BIN_NO)
                    wsUI.Cells(uiRow, 13).Value2 = WL_ResultText(coreRow, vals)
                    wsUI.Cells(uiRow, 14).Value2 = vals(coreRow, DC_ERROR)
                    wsUI.Cells(uiRow, WL_KEY_COL).Value2 = vals(coreRow, DC_SOURCE_KEY)
                    uiRow = uiRow + 1
                    rowCount = rowCount + 1
                End If
            End If
        Next coreRow
    End If
    CF_RefreshOBStatuses
    WL_ApplyGateFormatting wsUI
    modCF_InputLayout.CF_UpdateStatusBar "V3.0: " & CStr(rowCount) & _
        " dong dang mo. Cot vang J:L la du lieu nguoi dung."
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    modCF_InputLayout.CF_UpdateStatusBar "LOI lam moi danh sach: " & Err.Description
End Sub

Public Sub CF_RefreshOBStatuses()
    Dim wsUI As Worksheet, wsCore As Worksheet, wsQ As Worksheet
    Dim r As Long, c As Long, q As Long, lastCore As Long, lastQ As Long
    Dim obText As String, coreOB As String, qOB As String, qState As String
    Dim openCount As Long, totalCount As Long, runID As String
    On Error GoTo SafeExit
    Set wsUI = ThisWorkbook.Worksheets(WL_CONTROL)
    Set wsCore = ThisWorkbook.Worksheets(WL_CORE)
    Set wsQ = ThisWorkbook.Worksheets(WL_QUEUE)
    lastCore = wsCore.Cells(wsCore.rows.count, DC_OB).End(xlUp).row
    lastQ = wsQ.Cells(wsQ.rows.count, 1).End(xlUp).row
    runID = WL_Text(wsCore.Range("W2").Value2)
    For r = 7 To 40
        obText = WL_Text(wsUI.Cells(r, 1).Value2)
        wsUI.Cells(r, 3).ClearContents
        If obText <> "" Then
            openCount = 0: totalCount = 0: qState = ""
            For c = 5 To lastCore
                coreOB = WL_Text(wsCore.Cells(c, DC_OB).Value2)
                If StrComp(coreOB, obText, vbTextCompare) = 0 Then
                    totalCount = totalCount + 1
                    If Not WL_CoreRowCompleted(wsCore, c) Then openCount = openCount + 1
                End If
            Next c
            For q = 2 To lastQ
                qOB = WL_Text(wsQ.Cells(q, 3).Value2)
                If StrComp(qOB, obText, vbTextCompare) = 0 And _
                   (runID = "" Or StrComp(WL_Text(wsQ.Cells(q, 18).Value2), runID, vbTextCompare) = 0) Then
                    qState = UCase$(WL_Text(wsQ.Cells(q, 7).Value2))
                End If
            Next q
            If qState = "FAIL" Then
                wsUI.Cells(r, 3).Value2 = "LOI TAI WH"
            ElseIf qState = "PASS" And totalCount > 0 Then
                If openCount = 0 Then
                    wsUI.Cells(r, 3).Value2 = "HOAN TAT"
                Else
                    wsUI.Cells(r, 3).Value2 = "DA TAI: " & CStr(openCount) & " MO / " & CStr(totalCount) & " WT"
                End If
            Else
                wsUI.Cells(r, 3).Value2 = "CHUA TAI"
            End If
        End If
    Next r
SafeExit:
End Sub

Public Sub CF_SyncControlToCore()
    Dim wsCore As Worksheet, wsUI As Worksheet, keyMap As Object
    Dim lastCore As Long, coreRow As Long, uiRow As Long, targetRow As Long
    Dim sourceKey As String, duplicateKey As String, syncCount As Long
    Dim errNo As Long, errText As String
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Set wsCore = ThisWorkbook.Worksheets(WL_CORE)
    Set wsUI = ThisWorkbook.Worksheets(WL_CONTROL)
    Set keyMap = CreateObject("Scripting.Dictionary")
    keyMap.CompareMode = vbTextCompare
    lastCore = wsCore.Cells(wsCore.rows.count, DC_OB).End(xlUp).row
    If lastCore > 504 Then lastCore = 504
    For coreRow = 5 To lastCore
        sourceKey = WL_Text(wsCore.Cells(coreRow, DC_SOURCE_KEY).Value2)
        If sourceKey <> "" Then
            If keyMap.Exists(sourceKey) Then
                duplicateKey = sourceKey
                Err.Raise vbObjectError + 2841, "CF_SyncControlToCore", _
                    "SOURCE_KEY trung trong DATA_CORE: " & duplicateKey
            End If
            keyMap.Add sourceKey, coreRow
        End If
    Next coreRow
    For uiRow = WL_FIRST To WL_LAST
        sourceKey = WL_Text(wsUI.Cells(uiRow, WL_KEY_COL).Value2)
        If sourceKey <> "" Then
            If Not keyMap.Exists(sourceKey) Then Err.Raise vbObjectError + 2842, _
                "CF_SyncControlToCore", "Khong tim thay SOURCE_KEY: " & sourceKey
            targetRow = CLng(keyMap(sourceKey))
            wsCore.Cells(targetRow, DC_ACTUAL_QTY).Value2 = wsUI.Cells(uiRow, 10).Value2
            wsCore.Cells(targetRow, DC_AREA).Value2 = wsUI.Cells(uiRow, 11).Value2
            wsCore.Cells(targetRow, DC_BIN_NO).Value2 = wsUI.Cells(uiRow, 12).Value2
            AP_NormalizeRow wsCore, targetRow, False
            syncCount = syncCount + 1
        End If
    Next uiRow
    Application.ScreenUpdating = True
    modCF_InputLayout.CF_UpdateStatusBar "Da dong bo " & CStr(syncCount) & _
        " dong theo SOURCE_KEY."
    Exit Sub
Failed:
    errNo = Err.Number: errText = Err.Description
    Application.ScreenUpdating = True
    modCF_InputLayout.CF_UpdateStatusBar "LOI dong bo: " & errText
    Err.Raise errNo, "CF_SyncControlToCore", errText
End Sub

Public Function CF_InputOBCount() As Long
    Dim ws As Worksheet, r As Long, obText As String, seen As Object
    Set ws = ThisWorkbook.Worksheets(WL_CONTROL)
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For r = 7 To 40
        obText = WL_Text(ws.Cells(r, 1).Value2)
        If obText <> "" And IsNumeric(obText) Then
            If Not seen.Exists(obText) Then seen.Add obText, True
        End If
    Next r
    CF_InputOBCount = seen.count
End Function

Public Sub CF_InvalidatePlanLock(Optional ByVal reasonText As String = "RECHECK_REQUIRED_V3")
    On Error Resume Next
    ThisWorkbook.Worksheets("PLAN").Range("AJ2").Value2 = reasonText
    On Error GoTo 0
End Sub

Public Function CF_CurrentPlanIsExecutable() As Boolean
    Dim wsPlan As Worksheet, wsCore As Worksheet, wsUI As Worksheet
    Dim inputOBs As Object, r As Long, lastPlan As Long, obText As String
    Dim hasPlan As Boolean, auditText As String
    On Error GoTo NotReady
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsCore = ThisWorkbook.Worksheets(WL_CORE)
    Set wsUI = ThisWorkbook.Worksheets(WL_CONTROL)
    If UCase$(WL_Text(wsPlan.Range("AJ2").Value2)) <> "LOCKED" Then Exit Function
    If StrComp(WL_Text(wsPlan.Range("AK2").Value2), WL_Text(wsCore.Range("W2").Value2), vbTextCompare) <> 0 Then Exit Function
    If WL_Text(wsPlan.Range("AT2").Value2) <> "CHECK_LOCKED_V2" Then Exit Function
    Set inputOBs = CreateObject("Scripting.Dictionary")
    inputOBs.CompareMode = vbTextCompare
    For r = 7 To 40
        obText = WL_Text(wsUI.Cells(r, 1).Value2)
        If obText <> "" Then inputOBs(obText) = True
    Next r
    lastPlan = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    For r = 2 To lastPlan
        If WL_Text(wsPlan.Cells(r, 1).Value2) <> "" Then
            obText = WL_Text(wsPlan.Cells(r, 2).Value2)
            If obText = "" Or Not inputOBs.Exists(obText) Then Exit Function
            hasPlan = True
        End If
    Next r
    If Not hasPlan Then Exit Function
    auditText = Codex_ValidateLockedPlanContractLocal()
    CF_CurrentPlanIsExecutable = (Left$(auditText, 5) = "PASS|")
NotReady:
End Function

Private Function WL_IsCompleted(ByVal dataRow As Long, ByVal vals As Variant) As Boolean
    WL_IsCompleted = (UCase$(WL_Text(vals(dataRow, DC_STATUS))) = "C" Or _
        UCase$(WL_Text(vals(dataRow, DC_OB_STATE))) = "OB_DONE" Or _
        UCase$(WL_Text(vals(dataRow, DC_WT_STATE))) = "WT_DONE")
End Function

Private Function WL_CoreRowCompleted(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    WL_CoreRowCompleted = (UCase$(WL_Text(ws.Cells(r, DC_STATUS).Value2)) = "C" Or _
        UCase$(WL_Text(ws.Cells(r, DC_OB_STATE).Value2)) = "OB_DONE" Or _
        UCase$(WL_Text(ws.Cells(r, DC_WT_STATE).Value2)) = "WT_DONE")
End Function

Private Function WL_StatusText(ByVal dataRow As Long, ByVal vals As Variant) As String
    Dim rawStatus As String, wtState As String
    rawStatus = UCase$(WL_Text(vals(dataRow, DC_STATUS)))
    wtState = UCase$(WL_Text(vals(dataRow, DC_WT_STATE)))
    If wtState = "WT_CANCELLED_REMAINING" Or (rawStatus = "A" And val(vals(dataRow, DC_REMAINING)) > 0) Then
        WL_StatusText = "A - CON LAI"
    ElseIf rawStatus = "A" Or wtState = "WT_CANCELLED_ZERO" Then
        WL_StatusText = "A - DA HUY"
    Else
        WL_StatusText = "DANG MO"
    End If
End Function

Private Function WL_PlanText(ByVal dataRow As Long, ByVal vals As Variant) As String
    Dim decisionText As String, qtyMode As String
    decisionText = UCase$(WL_Text(vals(dataRow, DC_EXEC_MODE)))
    qtyMode = UCase$(WL_Text(vals(dataRow, DC_QTY_MODE)))
    Select Case decisionText
        Case "CF_ONLY": WL_PlanText = "CF THU CONG"
        Case "CHANGE_QTY": WL_PlanText = "DOI SL"
        Case "CHANGE_BIN": WL_PlanText = "DOI BIN"
        Case "CHANGE_BATCH": WL_PlanText = "DOI BATCH"
        Case "CHANGE_BIN_BATCH": WL_PlanText = "DOI BIN+BATCH"
        Case "CANCEL_TO_ZERO": WL_PlanText = "HUY VE 0"
        Case "NO_ACTION_DONE": WL_PlanText = "HOAN TAT"
        Case Else
            If qtyMode = "DELTA_ADD" Then WL_PlanText = "THEM SL" Else WL_PlanText = decisionText
    End Select
End Function

Private Function WL_GateText(ByVal dataRow As Long, ByVal vals As Variant) As String
    Dim resultText As String, stateText As String
    resultText = UCase$(WL_Text(vals(dataRow, DC_RESULT)))
    stateText = UCase$(WL_Text(vals(dataRow, DC_EXEC_STATE)))
    If stateText = "READY_TO_EXECUTE" Or resultText = "READY_FOR_CF" Or resultText = "PASS" Then
        WL_GateText = "READY"
    ElseIf InStr(1, stateText & "|" & resultText, "STOP", vbTextCompare) > 0 Or _
           InStr(1, stateText & "|" & resultText, "BLOCKED", vbTextCompare) > 0 Or _
           InStr(1, stateText & "|" & resultText, "ERROR", vbTextCompare) > 0 Then
        WL_GateText = "BLOCKED"
    Else
        WL_GateText = "WAIT"
    End If
End Function

Private Function WL_ResultText(ByVal dataRow As Long, ByVal vals As Variant) As String
    Dim resultText As String, detailText As String
    resultText = WL_Text(vals(dataRow, DC_RESULT))
    detailText = WL_Text(vals(dataRow, DC_ERROR))
    On Error Resume Next
    WL_ResultText = CF_VN_GetMessage(resultText, detailText)
    On Error GoTo 0
    If WL_ResultText = "" Then WL_ResultText = resultText
End Function

Private Sub WL_ApplyGateFormatting(ByVal ws As Worksheet)
    Dim r As Long, gateText As String
    For r = WL_FIRST To WL_LAST
        gateText = UCase$(WL_Text(ws.Cells(r, 9).Value2))
        ws.Cells(r, 9).Font.Color = RGB(0, 0, 0)
        ws.Cells(r, 9).Font.Bold = False
        Select Case gateText
            Case "READY"
                ws.Cells(r, 9).Interior.Color = RGB(198, 239, 206)
                ws.Cells(r, 9).Font.Bold = True
            Case "BLOCKED"
                ws.Cells(r, 9).Interior.Color = RGB(255, 199, 206)
                ws.Cells(r, 9).Font.Bold = True
            Case "WAIT"
                ws.Cells(r, 9).Interior.Color = RGB(255, 235, 156)
            Case Else
                ws.Cells(r, 9).Interior.Color = RGB(230, 230, 231)
        End Select
    Next r
End Sub

Private Function WL_Text(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    WL_Text = Trim$(CStr(value))
End Function

Public Sub CF_InputWorklist_CompileProbe()
    If WL_FIRST <> 9 Or WL_LAST <> 200 Or WL_KEY_COL <> 15 Then _
        Err.Raise vbObjectError + 2890, , "V3 worklist constants are invalid."
    Debug.Print "PROBE_PASS: V3.0 multi-OB worklist"
End Sub
