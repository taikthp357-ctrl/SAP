Attribute VB_Name = "modCF_InputWorklist"
Option Explicit

' Phase 1b worklist: mirrors DATA_CORE rows to INPUT with Vietnamese feedback
' and AutoFilter on Status=C.

Private Const WL_WH_SHEET As String = "DATA_CORE"
Private Const WL_INPUT_SHEET As String = "INPUT"
Private Const WL_FIRST_DATA_ROW As Long = 5
Private Const WL_HEADER_ROW As Long = 4
Private Const WL_START_COL As Long = 5      ' col E
Private Const WL_END_COL As Long = 15        ' col O
Private Const WL_MAX_ROWS As Long = 500      ' rows 5..504

' DATA_CORE column indices (1-based)
Private Const DC_COL_OB As Long = 1
Private Const DC_COL_DOC As Long = 2
Private Const DC_COL_MATERIAL As Long = 6
Private Const DC_COL_ACTUAL_QTY As Long = 12
Private Const DC_COL_ACTUAL_BIN As Long = 14
Private Const DC_COL_ACTUAL_BATCH As Long = 22
Private Const DC_COL_STATUS As Long = 10
Private Const DC_COL_WT As Long = 5
Private Const DC_COL_SYS_QTY As Long = 7
Private Const DC_COL_RESULT As Long = 18
Private Const DC_COL_ERROR As Long = 19
Private Const DC_COL_SYS_BIN As Long = 8
Private Const DC_COL_SYS_BATCH As Long = 9

' Worklist output column offsets (0-based from WL_START_COL)
' E=OB, F=Doc, G=Material, H=Target Qty, I=Target Bin, J=Target Batch,
' K=Status, L=WT cu, M=Qty cu, N=Ket qua (VN), O=Chi tiet
Private Const WLO_OB As Long = 0
Private Const WLO_DOC As Long = 1
Private Const WLO_MAT As Long = 2
Private Const WLO_TGT_QTY As Long = 3
Private Const WLO_TGT_BIN As Long = 4
Private Const WLO_TGT_BATCH As Long = 5
Private Const WLO_STATUS As Long = 6
Private Const WLO_WT_CU As Long = 7
Private Const WLO_QTY_CU As Long = 8
Private Const WLO_KQ_VN As Long = 9
Private Const WLO_CHI_TIET As Long = 10

Public Sub CF_RefreshInputWorklist()
    Dim wsWH As Worksheet, wsInput As Worksheet
    Dim lastRow As Long, r As Long, outRow As Long
    Dim resultText As String, detailText As String, vnMessage As String
    Dim arrOut() As Variant

    On Error GoTo Failed
    Set wsWH = ThisWorkbook.Worksheets(WL_WH_SHEET)
    Set wsInput = ThisWorkbook.Worksheets(WL_INPUT_SHEET)

    ' Clear existing worklist data
    wsInput.Range(wsInput.Cells(WL_FIRST_DATA_ROW, WL_START_COL), _
                  wsInput.Cells(WL_FIRST_DATA_ROW + WL_MAX_ROWS - 1, WL_END_COL)).ClearContents

    ' Find last row in DATA_CORE
    lastRow = wsWH.Cells(wsWH.Rows.Count, DC_COL_OB).End(xlUp).Row
    If lastRow < WL_FIRST_DATA_ROW Then Exit Sub

    ' Build output array
    ReDim arrOut(1 To WL_MAX_ROWS, 1 To WL_END_COL - WL_START_COL + 1)
    outRow = 0

    For r = WL_FIRST_DATA_ROW To lastRow
        If outRow >= WL_MAX_ROWS Then Exit For

        ' Read source data
        resultText = Trim$(CStr(wsWH.Cells(r, DC_COL_RESULT).Value2 & ""))
        detailText = Trim$(CStr(wsWH.Cells(r, DC_COL_ERROR).Value2 & ""))

        ' Map Result to Vietnamese
        vnMessage = CF_VN_GetMessage(resultText, detailText)

        outRow = outRow + 1
        arrOut(outRow, WLO_OB + 1) = wsWH.Cells(r, DC_COL_OB).Value2
        arrOut(outRow, WLO_DOC + 1) = wsWH.Cells(r, DC_COL_DOC).Value2
        arrOut(outRow, WLO_MAT + 1) = wsWH.Cells(r, DC_COL_MATERIAL).Value2
        arrOut(outRow, WLO_TGT_QTY + 1) = wsWH.Cells(r, DC_COL_ACTUAL_QTY).Value2
        arrOut(outRow, WLO_TGT_BIN + 1) = wsWH.Cells(r, DC_COL_ACTUAL_BIN).Value2
        arrOut(outRow, WLO_TGT_BATCH + 1) = wsWH.Cells(r, DC_COL_ACTUAL_BATCH).Value2
        arrOut(outRow, WLO_STATUS + 1) = wsWH.Cells(r, DC_COL_STATUS).Value2
        arrOut(outRow, WLO_WT_CU + 1) = wsWH.Cells(r, DC_COL_WT).Value2
        arrOut(outRow, WLO_QTY_CU + 1) = wsWH.Cells(r, DC_COL_SYS_QTY).Value2
        arrOut(outRow, WLO_KQ_VN + 1) = vnMessage
        arrOut(outRow, WLO_CHI_TIET + 1) = detailText
    Next r

    ' Write array to INPUT
    If outRow > 0 Then
        wsInput.Range(wsInput.Cells(WL_FIRST_DATA_ROW, WL_START_COL), _
                      wsInput.Cells(WL_FIRST_DATA_ROW + outRow - 1, WL_END_COL)).Value2 = arrOut
    End If

    ' Apply AutoFilter on Status column (col K = WL_START_COL + WLO_STATUS)
    On Error Resume Next
    wsInput.AutoFilterMode = False
    On Error GoTo 0

    If outRow > 0 Then
        wsInput.Range(wsInput.Cells(WL_HEADER_ROW, WL_START_COL), _
                      wsInput.Cells(WL_FIRST_DATA_ROW + outRow - 1, WL_END_COL)).AutoFilter _
                      Field:=WLO_STATUS + 1, Criteria1:="<>C"
    End If

    Exit Sub
Failed:
    RT_Log "INPUT_WORKLIST", "", "", "FAIL", "CF_RefreshInputWorklist: " & Err.Description
End Sub

Public Sub CF_SyncInputToCore()
    ' Sync user-edited Source Bin/Batch from INPUT back to DATA_CORE.
    ' Match by OB + Document (cols E + F in INPUT → cols 1 + 2 in DATA_CORE).

    Dim wsWH As Worksheet, wsInput As Worksheet
    Dim lastRowInput As Long, lastRowWH As Long
    Dim rInput As Long, rWH As Long
    Dim obInput As String, docInput As String
    Dim obWH As String, docWH As String
    Dim syncCount As Long

    On Error GoTo Failed
    Set wsWH = ThisWorkbook.Worksheets(WL_WH_SHEET)
    Set wsInput = ThisWorkbook.Worksheets(WL_INPUT_SHEET)

    lastRowInput = wsInput.Cells(wsInput.Rows.Count, WL_START_COL).End(xlUp).Row
    If lastRowInput < WL_FIRST_DATA_ROW Then Exit Sub

    lastRowWH = wsWH.Cells(wsWH.Rows.Count, DC_COL_OB).End(xlUp).Row
    If lastRowWH < WL_FIRST_DATA_ROW Then Exit Sub

    syncCount = 0

    ' For each INPUT row, find matching DATA_CORE row by OB + Document
    For rInput = WL_FIRST_DATA_ROW To lastRowInput
        obInput = Trim$(CStr(wsInput.Cells(rInput, WL_START_COL + WLO_OB).Value2 & ""))
        docInput = Trim$(CStr(wsInput.Cells(rInput, WL_START_COL + WLO_DOC).Value2 & ""))

        If obInput <> "" Then
            For rWH = WL_FIRST_DATA_ROW To lastRowWH
                obWH = Trim$(CStr(wsWH.Cells(rWH, DC_COL_OB).Value2 & ""))
                docWH = Trim$(CStr(wsWH.Cells(rWH, DC_COL_DOC).Value2 & ""))

                If StrComp(obWH, obInput, vbTextCompare) = 0 And _
                   StrComp(docWH, docInput, vbTextCompare) = 0 Then
                    ' Match found — sync editable columns (Target Bin/Batch → Sys Bin/Batch)
                    wsWH.Cells(rWH, DC_COL_SYS_BIN).Value2 = _
                        wsInput.Cells(rInput, WL_START_COL + WLO_TGT_BIN).Value2
                    wsWH.Cells(rWH, DC_COL_SYS_BATCH).Value2 = _
                        wsInput.Cells(rInput, WL_START_COL + WLO_TGT_BATCH).Value2
                    syncCount = syncCount + 1
                    Exit For
                End If
            Next rWH
        End If
    Next rInput

    RT_Log "INPUT_WORKLIST", "", "", "PASS", "CF_SyncInputToCore: synced " & CStr(syncCount) & " rows"
    Exit Sub
Failed:
    RT_Log "INPUT_WORKLIST", "", "", "FAIL", "CF_SyncInputToCore: " & Err.Description
End Sub

Public Sub CF_InputWorklist_CompileProbe()
    ' Offline compile probe
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WL_INPUT_SHEET)
    If ws Is Nothing Then Err.Raise vbObjectError + 9920, , "INPUT_WORKLIST_PROBE: INPUT missing"
End Sub
