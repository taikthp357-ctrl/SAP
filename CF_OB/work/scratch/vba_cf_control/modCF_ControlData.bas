Attribute VB_Name = "modCF_ControlData"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════════════
' modCF_ControlData — Data synchronization between DATA_CORE and CF_CONTROL table
' ═══════════════════════════════════════════════════════════════════════════════

' ═══════════════════════════════════════════════════════════════════════════════
' PUBLIC: Refresh CF_CONTROL table from DATA_CORE
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub CF_RefreshControlTable()
    Dim wsControl As Worksheet, wsCore As Worksheet
    Dim lastRow As Long, r As Long, outRow As Long
    Dim dataArr As Variant

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    Set wsControl = ThisWorkbook.Worksheets("CF_CONTROL")
    Set wsCore = ThisWorkbook.Worksheets("DATA_CORE")

    ' Clear existing table data (preserve headers)
    wsControl.Range("D12:P100").ClearContents

    ' Get DATA_CORE data (row 5 onwards, headers at row 4)
    lastRow = wsCore.Cells(wsCore.Rows.Count, 1).End(xlUp).Row
    If lastRow < 5 Then
        modCF_ControlLayout.CF_UpdateStatusBar "No data in DATA_CORE"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Read all data at once for performance
    dataArr = wsCore.Range("A5:AJ" & lastRow).Value2  ' Cols 1-36 (AJ = col 36)

    ' Map DATA_CORE → CF_CONTROL table
    outRow = 12  ' Start at row 12 (first data row)

    For r = 1 To UBound(dataArr, 1)
        If outRow > 100 Then Exit For  ' Table capacity limit

        ' SYSTEM group (read-only)
        wsControl.Cells(outRow, 4).Value = dataArr(r, 1)   ' D: OB (col 1)
        wsControl.Cells(outRow, 5).Value = dataArr(r, 2)   ' E: Document (col 2)
        wsControl.Cells(outRow, 6).Value = dataArr(r, 6)   ' F: Material (col 6)
        wsControl.Cells(outRow, 7).Value = dataArr(r, 7)   ' G: Sys Qty (col 7)
        wsControl.Cells(outRow, 8).Value = dataArr(r, 8)   ' H: Sys Bin (col 8)
        wsControl.Cells(outRow, 9).Value = dataArr(r, 9)   ' I: Sys Batch (col 9)
        wsControl.Cells(outRow, 10).Value = dataArr(r, 22) ' J: Batch (col 22 Actual Batch for SYSTEM view)

        ' ACTUAL group (user-editable)
        wsControl.Cells(outRow, 11).Value = dataArr(r, 12) ' K: Actual Qty (col 12)
        wsControl.Cells(outRow, 12).Value = dataArr(r, 14) ' L: Actual Bin (col 14)
        wsControl.Cells(outRow, 13).Value = dataArr(r, 22) ' M: Actual Batch (col 22)

        ' DECISION group (read-only)
        wsControl.Cells(outRow, 14).Value = CF_BuildPlanText(r, dataArr) ' N: PLAN (human-readable)
        wsControl.Cells(outRow, 15).Value = CF_BuildGateState(r, dataArr) ' O: GATE (READY/WAIT/BLOCKED/DONE)
        wsControl.Cells(outRow, 16).Value = CF_BuildResultText(r, dataArr) ' P: Result (Vietnamese)

        outRow = outRow + 1
    Next r

    modCF_ControlLayout.CF_UpdateStatusBar "Refreshed " & (outRow - 12) & " rows from DATA_CORE"
    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    modCF_ControlLayout.CF_UpdateStatusBar "ERROR refreshing table: " & Err.Description
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' Build human-readable PLAN text from cols 28-30
' ═══════════════════════════════════════════════════════════════════════════════
Private Function CF_BuildPlanText(dataRow As Long, dataArr As Variant) As String
    ' Col 28 = Process Allowed (QTY_MODE), Col 29 = Decision (EXEC_MODE), Col 30 = Reason (EXEC_STATE)
    Dim qtyMode As String, execMode As String, execState As String
    Dim planText As String

    qtyMode = Trim(CStr(dataArr(dataRow, 28)))
    execMode = Trim(CStr(dataArr(dataRow, 29)))
    execState = Trim(CStr(dataArr(dataRow, 30)))

    ' Build human-readable plan
    If execMode = "CF_ONLY" Then
        planText = "CF ONLY"
    ElseIf execMode = "WT_CF" Then
        If qtyMode = "INCREASE" Then
            Dim delta As Variant
            delta = dataArr(dataRow, 12)  ' Actual Qty
            If IsNumeric(delta) Then
                planText = "INCREASE +" & delta & " EA"
            Else
                planText = "INCREASE"
            End If
        ElseIf qtyMode = "REDUCE" Then
            planText = "REDUCE"
        ElseIf execState = "BIN_ONLY" Then
            planText = "CHANGE BIN"
        ElseIf execState = "BATCH_ONLY" Then
            planText = "CHANGE BATCH"
        ElseIf execState = "BIN_BATCH" Then
            planText = "CHANGE BIN+BATCH"
        Else
            planText = "WT + CF"
        End If
    ElseIf execMode = "SKIP" Then
        planText = "SKIP"
    Else
        planText = ""
    End If

    CF_BuildPlanText = planText
End Function

' ═══════════════════════════════════════════════════════════════════════════════
' Build GATE state from Status + Result
' ═══════════════════════════════════════════════════════════════════════════════
Private Function CF_BuildGateState(dataRow As Long, dataArr As Variant) As String
    ' Col 10 = Status (C=completed), Col 18 = Result (PASS/BLOCKED_*/STOP/etc)
    Dim status As String, result As String

    status = Trim(CStr(dataArr(dataRow, 10)))
    result = Trim(CStr(dataArr(dataRow, 18)))

    If status = "C" Then
        CF_BuildGateState = "DONE"
    ElseIf result = "PASS" Or result = "READY_FOR_CF" Then
        CF_BuildGateState = "READY"
    ElseIf InStr(1, result, "BLOCKED", vbTextCompare) > 0 Or result = "STOP" Then
        CF_BuildGateState = "BLOCKED"
    Else
        CF_BuildGateState = "WAIT"
    End If
End Function

' ═══════════════════════════════════════════════════════════════════════════════
' Build Vietnamese Result text (map from col 18 Result)
' ═══════════════════════════════════════════════════════════════════════════════
Private Function CF_BuildResultText(dataRow As Long, dataArr As Variant) As String
    ' Col 18 = Result, Col 19 = Error (detail)
    Dim result As String, detail As String

    result = Trim(CStr(dataArr(dataRow, 18)))
    detail = Trim(CStr(dataArr(dataRow, 19)))

    ' Call Vietnamese message mapper (from modCF_VN_Messages, CP6)
    On Error Resume Next
    CF_BuildResultText = CF_VN_GetMessage(result, detail)
    If Err.Number <> 0 Or CF_BuildResultText = "" Then
        ' Fallback: return English result if mapper not available
        CF_BuildResultText = result
    End If
    On Error GoTo 0
End Function

' ═══════════════════════════════════════════════════════════════════════════════
' PUBLIC: Sync user edits from CF_CONTROL → DATA_CORE
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub CF_SyncControlToCore()
    Dim wsControl As Worksheet, wsCore As Worksheet
    Dim r As Long, coreRow As Long
    Dim obKey As String, coreOB As String

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    Set wsControl = ThisWorkbook.Worksheets("CF_CONTROL")
    Set wsCore = ThisWorkbook.Worksheets("DATA_CORE")

    ' Loop through CF_CONTROL table rows
    For r = 12 To 100
        obKey = Trim(CStr(wsControl.Cells(r, 4).Value))  ' Col D = OB
        If obKey = "" Then Exit For

        ' Find matching row in DATA_CORE (simple linear search by OB in col 1)
        coreRow = 0
        Dim searchRow As Long
        For searchRow = 5 To wsCore.Cells(wsCore.Rows.Count, 1).End(xlUp).Row
            coreOB = Trim(CStr(wsCore.Cells(searchRow, 1).Value))
            If coreOB = obKey Then
                coreRow = searchRow
                Exit For
            End If
        Next searchRow

        If coreRow > 0 Then
            ' Sync ACTUAL cols (K:M → cols 12, 14, 22)
            wsCore.Cells(coreRow, 12).Value = wsControl.Cells(r, 11).Value  ' Actual Qty
            wsCore.Cells(coreRow, 14).Value = wsControl.Cells(r, 12).Value  ' Actual Bin
            wsCore.Cells(coreRow, 22).Value = wsControl.Cells(r, 13).Value  ' Actual Batch
        End If
    Next r

    modCF_ControlLayout.CF_UpdateStatusBar "Synced user edits to DATA_CORE"
    Application.ScreenUpdating = True

    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    modCF_ControlLayout.CF_UpdateStatusBar "ERROR syncing to DATA_CORE: " & Err.Description
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' PUBLIC: Auto-sync on CF_CONTROL sheet deactivate (optional)
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub CF_AutoSyncOnDeactivate()
    ' Called from Worksheet_Deactivate event
    CF_SyncControlToCore
End Sub
