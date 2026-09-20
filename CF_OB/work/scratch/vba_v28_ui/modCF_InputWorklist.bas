Attribute VB_Name = "modCF_InputWorklist"
Option Explicit

' ═════════════════════════════════════════════════════════════════════════════
' modCF_InputWorklist — DATA_CORE ↔ CF_CONTROL worklist sync (28-point spec)
' ═════════════════════════════════════════════════════════════════════════════
' Phase 1b worklist: mirrors DATA_CORE rows to CF_CONTROL with Vietnamese feedback
' Updated for new layout: E9:N40 (10 columns, 32 data rows)
' ═════════════════════════════════════════════════════════════════════════════

' ─────────────────────────────────────────────────────────────────────────────
' Constants (aligned with modCF_InputLayout 28-point spec)
' ─────────────────────────────────────────────────────────────────────────────
Private Const WL_WH_SHEET As String = "DATA_CORE"
Private Const WL_INPUT_SHEET As String = "CF_CONTROL"
Private Const WL_FIRST_DATA_ROW As Long = 9        ' First data row (was 5, now 9)
Private Const WL_HEADER_ROW As Long = 8            ' Header row (was 4, now 8)
Private Const WL_START_COL As Long = 5             ' E (unchanged)
Private Const WL_END_COL As Long = 14              ' N (was 15, now 14)
Private Const WL_MAX_ROWS As Long = 32             ' 9-40 = 32 rows (was 500)

' Column offsets (relative to WL_START_COL = 5)
Private Const WLO_OB As Long = 0           ' E (col 5)
Private Const WLO_DOC As Long = 1          ' F (col 6)
Private Const WLO_MAT As Long = 2          ' G (col 7)
Private Const WLO_PLAN As Long = 3         ' H (col 8) — NEW: derived PLAN text
Private Const WLO_GATE As Long = 4         ' I (col 9) — NEW: derived GATE state
Private Const WLO_TGT_QTY As Long = 5      ' J (col 10)
Private Const WLO_TGT_BIN As Long = 6      ' K (col 11)
Private Const WLO_TGT_BATCH As Long = 7    ' L (col 12)
Private Const WLO_RESULT_VN As Long = 8    ' M (col 13) — Vietnamese Result
Private Const WLO_DETAIL As Long = 9       ' N (col 14) — Chi tiết

' DATA_CORE column indices (1-based)
Private Const DC_COL_OB As Long = 1
Private Const DC_COL_DOCUMENT As Long = 2
Private Const DC_COL_MATERIAL As Long = 6
Private Const DC_COL_STATUS As Long = 10
Private Const DC_COL_ACTUAL_QTY As Long = 12
Private Const DC_COL_ACTUAL_BIN As Long = 14
Private Const DC_COL_RESULT As Long = 18
Private Const DC_COL_ERROR As Long = 19
Private Const DC_COL_ACTUAL_BATCH As Long = 22
Private Const DC_COL_QTY_MODE As Long = 28   ' Process Allowed
Private Const DC_COL_EXEC_MODE As Long = 29  ' Decision
Private Const DC_COL_EXEC_STATE As Long = 30 ' Reason

' ═════════════════════════════════════════════════════════════════════════════
' PUBLIC: Refresh CF_CONTROL worklist from DATA_CORE
' ═════════════════════════════════════════════════════════════════════════════
Public Sub CF_RefreshInputWorklist()
    Dim wsWH As Worksheet, wsInput As Worksheet
    Dim lastRow As Long, r As Long, outRow As Long
    Dim dataArr As Variant
    Dim rowsWritten As Long

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Set wsWH = ThisWorkbook.Worksheets(WL_WH_SHEET)
    Set wsInput = ThisWorkbook.Worksheets(WL_INPUT_SHEET)

    ' Clear existing worklist data (preserve headers at row 8)
    wsInput.Range(wsInput.Cells(WL_FIRST_DATA_ROW, WL_START_COL), _
                  wsInput.Cells(WL_FIRST_DATA_ROW + WL_MAX_ROWS - 1, WL_END_COL)).ClearContents

    ' Get DATA_CORE data (row 5 onwards, headers at row 4)
    lastRow = wsWH.Cells(wsWH.Rows.Count, DC_COL_OB).End(xlUp).Row
    If lastRow < 5 Then
        Call modCF_InputLayout.CF_UpdateStatusBar("No data in DATA_CORE")
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Read all data at once for performance (cols 1-30)
    dataArr = wsWH.Range("A5:AD" & lastRow).Value2  ' AD = col 30

    ' Map DATA_CORE → CF_CONTROL worklist
    outRow = WL_FIRST_DATA_ROW
    rowsWritten = 0

    For r = 1 To UBound(dataArr, 1)
        If rowsWritten >= WL_MAX_ROWS Then Exit For

        ' SYSTEM cols (read-only)
        wsInput.Cells(outRow, WL_START_COL + WLO_OB).Value = dataArr(r, DC_COL_OB)
        wsInput.Cells(outRow, WL_START_COL + WLO_DOC).Value = dataArr(r, DC_COL_DOCUMENT)
        wsInput.Cells(outRow, WL_START_COL + WLO_MAT).Value = dataArr(r, DC_COL_MATERIAL)

        ' DECISION cols (derived)
        wsInput.Cells(outRow, WL_START_COL + WLO_PLAN).Value = WL_BuildPlanText(r, dataArr)
        wsInput.Cells(outRow, WL_START_COL + WLO_GATE).Value = WL_BuildGateState(r, dataArr)

        ' ACTUAL cols (user-editable)
        wsInput.Cells(outRow, WL_START_COL + WLO_TGT_QTY).Value = dataArr(r, DC_COL_ACTUAL_QTY)
        wsInput.Cells(outRow, WL_START_COL + WLO_TGT_BIN).Value = dataArr(r, DC_COL_ACTUAL_BIN)
        wsInput.Cells(outRow, WL_START_COL + WLO_TGT_BATCH).Value = dataArr(r, DC_COL_ACTUAL_BATCH)

        ' RESULT cols (Vietnamese)
        wsInput.Cells(outRow, WL_START_COL + WLO_RESULT_VN).Value = WL_BuildResultText(r, dataArr)
        wsInput.Cells(outRow, WL_START_COL + WLO_DETAIL).Value = dataArr(r, DC_COL_ERROR)

        outRow = outRow + 1
        rowsWritten = rowsWritten + 1
    Next r

    ' Apply GATE conditional formatting (color badges)
    Call WL_ApplyGateFormatting(wsInput)

    Call modCF_InputLayout.CF_UpdateStatusBar("Refreshed " & rowsWritten & " rows from DATA_CORE")
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Call modCF_InputLayout.CF_UpdateStatusBar("ERROR refreshing worklist: " & Err.Description)
End Sub

' ═════════════════════════════════════════════════════════════════════════════
' Build human-readable PLAN text from QTY_MODE/EXEC_MODE/EXEC_STATE
' ═════════════════════════════════════════════════════════════════════════════
Private Function WL_BuildPlanText(dataRow As Long, dataArr As Variant) As String
    Dim qtyMode As String, execMode As String, execState As String
    Dim planText As String
    Dim delta As Variant

    qtyMode = Trim(CStr(dataArr(dataRow, DC_COL_QTY_MODE)))
    execMode = Trim(CStr(dataArr(dataRow, DC_COL_EXEC_MODE)))
    execState = Trim(CStr(dataArr(dataRow, DC_COL_EXEC_STATE)))

    ' Build short plan text (28-point spec: "CF ONLY", "+20 QTY", "CHANGE BIN+BATCH")
    If execMode = "CF_ONLY" Then
        planText = "CF ONLY"
    ElseIf execMode = "WT_CF" Then
        If qtyMode = "INCREASE" Then
            delta = dataArr(dataRow, DC_COL_ACTUAL_QTY)
            If IsNumeric(delta) And delta <> 0 Then
                planText = "+" & delta & " QTY"
            Else
                planText = "INCREASE"
            End If
        ElseIf qtyMode = "REDUCE" Then
            planText = "REDUCE QTY"
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

    WL_BuildPlanText = planText
End Function

' ═════════════════════════════════════════════════════════════════════════════
' Build GATE state badge (READY/WAIT/BLOCKED/DONE)
' ═════════════════════════════════════════════════════════════════════════════
Private Function WL_BuildGateState(dataRow As Long, dataArr As Variant) As String
    Dim status As String, result As String

    status = Trim(CStr(dataArr(dataRow, DC_COL_STATUS)))
    result = Trim(CStr(dataArr(dataRow, DC_COL_RESULT)))

    If status = "C" Then
        WL_BuildGateState = "DONE"
    ElseIf result = "PASS" Or result = "READY_FOR_CF" Then
        WL_BuildGateState = "READY"
    ElseIf InStr(1, result, "BLOCKED", vbTextCompare) > 0 Or result = "STOP" Then
        WL_BuildGateState = "BLOCKED"
    Else
        WL_BuildGateState = "WAIT"
    End If
End Function

' ═════════════════════════════════════════════════════════════════════════════
' Build Vietnamese Result text (map from DC Result col 18)
' ═════════════════════════════════════════════════════════════════════════════
Private Function WL_BuildResultText(dataRow As Long, dataArr As Variant) As String
    Dim result As String, detail As String

    result = Trim(CStr(dataArr(dataRow, DC_COL_RESULT)))
    detail = Trim(CStr(dataArr(dataRow, DC_COL_ERROR)))

    ' Call Vietnamese message mapper (from modCF_VietnameseMessages, CP6)
    On Error Resume Next
    WL_BuildResultText = CF_VN_GetMessage(result, detail)
    If Err.Number <> 0 Or WL_BuildResultText = "" Then
        ' Fallback: return English result if mapper not available
        WL_BuildResultText = result
    End If
    On Error GoTo 0
End Function

' ═════════════════════════════════════════════════════════════════════════════
' Apply GATE conditional formatting (color badges)
' ═════════════════════════════════════════════════════════════════════════════
Private Sub WL_ApplyGateFormatting(wsInput As Worksheet)
    Dim rng As Range
    Dim r As Long, gateCol As Long
    Dim gateValue As String

    gateCol = WL_START_COL + WLO_GATE

    For r = WL_FIRST_DATA_ROW To WL_FIRST_DATA_ROW + WL_MAX_ROWS - 1
        Set rng = wsInput.Cells(r, gateCol)
        gateValue = Trim(CStr(rng.Value))

        If gateValue = "" Then Exit For

        Select Case UCase(gateValue)
            Case "READY"
                rng.Interior.Color = RGB(144, 238, 144)  ' Light green
                rng.Font.Bold = True
            Case "WAIT"
                rng.Interior.Color = RGB(255, 255, 204)  ' Light yellow
            Case "BLOCKED"
                rng.Interior.Color = RGB(255, 182, 193)  ' Light red
                rng.Font.Bold = True
            Case "DONE"
                rng.Interior.Color = RGB(34, 139, 34)    ' Dark green
                rng.Font.Color = RGB(255, 255, 255)
                rng.Font.Bold = True
        End Select
    Next r
End Sub

' ═════════════════════════════════════════════════════════════════════════════
' PUBLIC: Sync user edits from CF_CONTROL → DATA_CORE
' ═════════════════════════════════════════════════════════════════════════════
Public Sub CF_SyncControlToCore()
    Dim wsWH As Worksheet, wsInput As Worksheet
    Dim r As Long, coreRow As Long
    Dim obKey As String, coreOB As String
    Dim syncCount As Long

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Set wsWH = ThisWorkbook.Worksheets(WL_WH_SHEET)
    Set wsInput = ThisWorkbook.Worksheets(WL_INPUT_SHEET)

    syncCount = 0

    ' Loop through CF_CONTROL worklist rows
    For r = WL_FIRST_DATA_ROW To WL_FIRST_DATA_ROW + WL_MAX_ROWS - 1
        obKey = Trim(CStr(wsInput.Cells(r, WL_START_COL + WLO_OB).Value))
        If obKey = "" Then Exit For

        ' Find matching row in DATA_CORE (simple linear search by OB)
        coreRow = 0
        Dim searchRow As Long
        For searchRow = 5 To wsWH.Cells(wsWH.Rows.Count, DC_COL_OB).End(xlUp).Row
            coreOB = Trim(CStr(wsWH.Cells(searchRow, DC_COL_OB).Value))
            If coreOB = obKey Then
                coreRow = searchRow
                Exit For
            End If
        Next searchRow

        If coreRow > 0 Then
            ' Sync ACTUAL cols (J:L → DC cols 12, 14, 22)
            wsWH.Cells(coreRow, DC_COL_ACTUAL_QTY).Value = _
                wsInput.Cells(r, WL_START_COL + WLO_TGT_QTY).Value
            wsWH.Cells(coreRow, DC_COL_ACTUAL_BIN).Value = _
                wsInput.Cells(r, WL_START_COL + WLO_TGT_BIN).Value
            wsWH.Cells(coreRow, DC_COL_ACTUAL_BATCH).Value = _
                wsInput.Cells(r, WL_START_COL + WLO_TGT_BATCH).Value
            syncCount = syncCount + 1
        End If
    Next r

    Call modCF_InputLayout.CF_UpdateStatusBar("Synced " & syncCount & " rows to DATA_CORE")
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Call modCF_InputLayout.CF_UpdateStatusBar("ERROR syncing to DATA_CORE: " & Err.Description)
End Sub

' ═════════════════════════════════════════════════════════════════════════════
' OFFLINE PROBE: Verify worklist constants and mapping
' ═════════════════════════════════════════════════════════════════════════════
Public Sub CF_InputWorklist_CompileProbe()
    On Error GoTo Failed

    ' Verify constants match layout
    If WL_START_COL <> 5 Then
        Debug.Print "PROBE_FAIL: WL_START_COL=" & WL_START_COL & " (expected 5)"
    End If
    If WL_END_COL <> 14 Then
        Debug.Print "PROBE_FAIL: WL_END_COL=" & WL_END_COL & " (expected 14)"
    End If
    If WL_FIRST_DATA_ROW <> 9 Then
        Debug.Print "PROBE_FAIL: WL_FIRST_DATA_ROW=" & WL_FIRST_DATA_ROW & " (expected 9)"
    End If

    ' Verify column offsets
    If WLO_PLAN <> 3 Or WLO_GATE <> 4 Then
        Debug.Print "PROBE_FAIL: PLAN/GATE offset mismatch"
    End If

    Debug.Print "PROBE_PASS: Worklist constants aligned with 28-point spec"
    Debug.Print "CF_InputWorklist_CompileProbe: DONE"
    Exit Sub

Failed:
    Debug.Print "CF_InputWorklist_CompileProbe: FAIL - " & Err.Description
End Sub
