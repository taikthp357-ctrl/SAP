Attribute VB_Name = "modCF_ActualPicking"
Option Explicit

' Actual Picking input contract (WH_ACTUAL).
' K = Area / prefix, L = Actual Qty, M = Bin No. or full bin, N = Actual Bin (resolved).
' No SAP transaction is called from this module.

Public Const AP_WH_SHEET As String = "WH_ACTUAL"
Public Const AP_FIRST_ROW As Long = 5
Public Const AP_LAST_ROW As Long = 504
Public Const AP_COL_AREA As Long = 11
Public Const AP_COL_QTY As Long = 12
Public Const AP_COL_BIN_INPUT As Long = 13
Public Const AP_COL_BIN_RESOLVED As Long = 14
Public Const AP_COL_SYS_BIN As Long = 8
Public Const AP_COL_MATERIAL As Long = 6
Public Const AP_COL_STATE As Long = 17
Public Const AP_COL_RESULT As Long = 18
Public Const AP_COL_ERROR As Long = 19
Public Const AP_COL_OB_STATE As Long = 25
Public Const AP_COL_WT_STATE As Long = 26

Public Sub AP_PrepareActualPickingSheet()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(AP_WH_SHEET)
    AP_EnsureConfig
    AP_WriteHeaders ws
    AP_ApplyValidations ws

    lastRow = AP_LastRow(ws)
    For r = AP_FIRST_ROW To lastRow
        If AP_Text(ws.Cells(r, AP_COL_MATERIAL).Value2) <> "" Then
            AP_NormalizeRow ws, r, False
        End If
    Next r
    Exit Sub
Failed:
    ' A sheet-formatting failure must not alter SAP/WH data or start SAP.
    RT_Log "ACTUAL_PICKING", "", "", "FAIL", "PREPARE_FAIL: " & Err.Description
    Err.Raise vbObjectError + 9902, , "ACTUAL_PICKING_PREPARE_FAIL: " & Err.Description
End Sub

Public Function AP_NormalizeAndValidateAll(ByVal ws As Worksheet, ByRef messageText As String) As Boolean
    Dim lastRow As Long, r As Long, rowError As String
    Dim errorCount As Long

    On Error GoTo Failed
    messageText = ""
    If ws Is Nothing Then
        messageText = "STOP_ACTUAL_INPUT_SHEET_MISSING: WH_ACTUAL is missing."
        Exit Function
    End If

    AP_EnsureConfig
    AP_WriteHeaders ws
    lastRow = AP_LastRow(ws)
    For r = AP_FIRST_ROW To lastRow
        If AP_Text(ws.Cells(r, AP_COL_MATERIAL).Value2) <> "" Then
            rowError = ""
            AP_NormalizeRow ws, r, True, rowError
            If rowError <> "" Then
                errorCount = errorCount + 1
                If messageText = "" Then messageText = "Row " & CStr(r) & ": " & rowError
            End If
        End If
    Next r
    If errorCount > 0 Then
        messageText = "STOP_ACTUAL_BIN_INPUT_INVALID: " & messageText & "; errors=" & CStr(errorCount)
        Exit Function
    End If
    AP_NormalizeAndValidateAll = True
    Exit Function
Failed:
    messageText = "STOP_ACTUAL_BIN_NORMALIZE_FAIL: " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Sub AP_NormalizeRow(ByVal ws As Worksheet, ByVal rowNo As Long, _
                           Optional ByVal strictValidation As Boolean = False, _
                           Optional ByRef rowError As String)
    Dim areaText As String, binInput As String, sysBin As String, resolvedBin As String
    Dim statusText As String, processNeeded As Boolean, inputWasFull As Boolean

    If ws Is Nothing Then Exit Sub
    rowError = ""
    sysBin = AP_Text(ws.Cells(rowNo, AP_COL_SYS_BIN).Value2)
    statusText = UCase$(AP_Text(ws.Cells(rowNo, AP_COL_STATE).Value2))
    processNeeded = (AP_Text(ws.Cells(rowNo, AP_COL_MATERIAL).Value2) <> "" And _
                     UCase$(AP_Text(ws.Cells(rowNo, AP_COL_OB_STATE).Value2)) = "OB_OPEN_OR_PARTIAL" And _
                     (UCase$(AP_Text(ws.Cells(rowNo, AP_COL_WT_STATE).Value2)) = "WT_ACTIVE" Or _
                      UCase$(AP_Text(ws.Cells(rowNo, AP_COL_WT_STATE).Value2)) = "WT_CANCELLED_REMAINING"))

    areaText = AP_NormalizeArea(AP_Text(ws.Cells(rowNo, AP_COL_AREA).Value2))
    binInput = AP_Text(ws.Cells(rowNo, AP_COL_BIN_INPUT).Value2)

    ' Accept legacy full-bin data in M once, then split it visibly to Area + Bin No.
    inputWasFull = AP_IsFullBin(binInput)
    If inputWasFull Then
        AP_SplitFullBin binInput, areaText, binInput
    End If
    If areaText = "" And binInput = "" And sysBin <> "" Then
        AP_SplitFullBin sysBin, areaText, binInput
    End If

    If areaText <> "" Then
        ws.Cells(rowNo, AP_COL_AREA).Value2 = areaText
    Else
        ws.Cells(rowNo, AP_COL_AREA).ClearContents
    End If
    If binInput <> "" Then ws.Cells(rowNo, AP_COL_BIN_INPUT).Value2 = UCase$(binInput)

    If inputWasFull Then
        resolvedBin = AP_BuildBin(areaText, binInput)
    ElseIf binInput = "" Then
        resolvedBin = sysBin
    ElseIf areaText = "" Then
        If AP_IsFullBin(binInput) Then
            resolvedBin = UCase$(binInput)
            AP_SplitFullBin resolvedBin, areaText, binInput
            ws.Cells(rowNo, AP_COL_AREA).Value2 = areaText
            ws.Cells(rowNo, AP_COL_BIN_INPUT).Value2 = binInput
        Else
            If strictValidation And processNeeded Then rowError = "Area must be ZF, P, or HR when Bin No. is entered."
        End If
    ElseIf Not AP_IsAllowedArea(areaText) Then
        If strictValidation And processNeeded Then rowError = "Area '" & areaText & "' is not configured in ACTUAL_PICKING_AREAS."
    ElseIf Not AP_IsValidBinNo(binInput) Then
        If strictValidation And processNeeded Then rowError = "Bin No. must be NN-NN-NN, for example 04-02-10."
    Else
        resolvedBin = AP_BuildBin(areaText, binInput)
    End If

    If resolvedBin <> "" Then
        ws.Cells(rowNo, AP_COL_BIN_RESOLVED).Value2 = resolvedBin
    ElseIf strictValidation And processNeeded Then
        ws.Cells(rowNo, AP_COL_BIN_RESOLVED).ClearContents
    Else
        ws.Cells(rowNo, AP_COL_BIN_RESOLVED).Value2 = sysBin
    End If

    If rowError <> "" Then
        ws.Cells(rowNo, AP_COL_ERROR).Value2 = "ACTUAL_INPUT: " & rowError
        ws.Cells(rowNo, AP_COL_RESULT).Value2 = "STOP_INPUT"
    End If
End Sub

Public Function AP_ResolvedBin(ByVal ws As Worksheet, ByVal rowNo As Long, Optional ByVal fallbackSysBin As String = "") As String
    Dim rowError As String
    If ws Is Nothing Then Exit Function
    AP_NormalizeRow ws, rowNo, False, rowError
    AP_ResolvedBin = AP_Text(ws.Cells(rowNo, AP_COL_BIN_RESOLVED).Value2)
    If AP_ResolvedBin = "" Then AP_ResolvedBin = fallbackSysBin
End Function

Public Function AP_NormalizeArea(ByVal rawText As String) As String
    Dim keyText As String
    keyText = UCase$(Replace(Trim$(rawText), " ", ""))
    If Left$(keyText, 5) = "APFL-" Then keyText = Mid$(keyText, 6)
    Select Case keyText
        Case "P", "ZF", "HR": AP_NormalizeArea = keyText
    End Select
End Function

Public Function AP_BuildBin(ByVal areaText As String, ByVal binNo As String) As String
    Dim areaKey As String, noKey As String, rootText As String
    areaKey = AP_NormalizeArea(areaText)
    noKey = UCase$(Replace(Trim$(binNo), " ", ""))
    If areaKey = "" Or Not AP_IsValidBinNo(noKey) Then Exit Function
    rootText = AP_BinRoot()
    AP_BuildBin = rootText & areaKey & noKey
End Function

Public Function AP_BinRoot() As String
    Dim rootText As String
    rootText = UCase$(Trim$(ModuleHelper.ToolConfig("ACTUAL_PICKING_BIN_ROOT", "APFL-")))
    If rootText = "" Then rootText = "APFL-"
    If Right$(rootText, 1) <> "-" Then rootText = rootText & "-"
    AP_BinRoot = rootText
End Function

Public Function AP_IsAllowedArea(ByVal areaText As String) As Boolean
    Dim allowedText As String, parts() As String, i As Long, oneArea As String
    allowedText = ModuleHelper.ToolConfig("ACTUAL_PICKING_AREAS", "ZF,P,HR")
    parts = Split(allowedText, ",")
    For i = LBound(parts) To UBound(parts)
        oneArea = AP_NormalizeArea(parts(i))
        If oneArea <> "" And StrComp(oneArea, AP_NormalizeArea(areaText), vbTextCompare) = 0 Then
            AP_IsAllowedArea = True
            Exit Function
        End If
    Next i
End Function

Public Function AP_IsValidBinNo(ByVal binNo As String) As Boolean
    Dim rawText As String
    rawText = UCase$(Replace(Trim$(binNo), " ", ""))
    If Len(rawText) <> 8 Then Exit Function
    If Mid$(rawText, 3, 1) <> "-" Or Mid$(rawText, 6, 1) <> "-" Then Exit Function
    If Not IsNumeric(Left$(rawText, 2)) Then Exit Function
    If Not IsNumeric(Mid$(rawText, 4, 2)) Then Exit Function
    If Not IsNumeric(Right$(rawText, 2)) Then Exit Function
    AP_IsValidBinNo = True
End Function

Public Function AP_IsFullBin(ByVal binText As String) As Boolean
    Dim areaText As String, noText As String
    AP_SplitFullBin binText, areaText, noText
    AP_IsFullBin = (areaText <> "" And AP_IsValidBinNo(noText))
End Function

Public Sub AP_SplitFullBin(ByVal fullBin As String, ByRef areaText As String, ByRef binNo As String)
    Dim rawText As String, rootText As String, remainderText As String
    Dim candidate As Variant
    areaText = "": binNo = ""
    rawText = UCase$(Replace(Trim$(fullBin), " ", ""))
    rootText = AP_BinRoot()
    If Left$(rawText, Len(rootText)) <> rootText Then Exit Sub
    remainderText = Mid$(rawText, Len(rootText) + 1)
    For Each candidate In Array("ZF", "HR", "P")
        If Left$(remainderText, Len(CStr(candidate))) = CStr(candidate) Then
            If AP_IsValidBinNo(Mid$(remainderText, Len(CStr(candidate)) + 1)) Then
                areaText = CStr(candidate)
                binNo = Mid$(remainderText, Len(CStr(candidate)) + 1)
                Exit Sub
            End If
        End If
    Next candidate
End Sub

Public Sub AP_WriteHeaders(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Range("K3:N3").UnMerge
    ws.Range("K3:N3").Merge
    ws.Range("K3").Value2 = "ACTUAL PICKING - USER INPUT"
    ws.Range("K4:N4").Value = Array("Area", "Actual Qty", "Bin No.", "Actual Bin")
    ws.Range("K3").HorizontalAlignment = xlCenter
    ws.Range("K3").VerticalAlignment = xlCenter
    ws.Range("K4:N4").HorizontalAlignment = xlCenter
    ws.Range("K:N").NumberFormat = "@"
    ws.columns("L").NumberFormat = "#,##0.###"
    ws.columns("K").ColumnWidth = 10
    ws.columns("L").ColumnWidth = 12
    ws.columns("M").ColumnWidth = 13
    ws.columns("N").ColumnWidth = 18
    On Error GoTo 0
End Sub

Public Sub AP_ApplyValidations(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ' Do not create any Validation rule here.  Excel parses Validation.Add
    ' differently by locale and can raise 1004 while the sheet is loading.
    ' CHECK is the single, fail-closed validator for Area and Bin No.
    On Error Resume Next
    ws.Range("K" & CStr(AP_FIRST_ROW) & ":K" & CStr(AP_LAST_ROW)).Validation.Delete
    ws.Range("M" & CStr(AP_FIRST_ROW) & ":M" & CStr(AP_LAST_ROW)).Validation.Delete
    ws.Range("N" & CStr(AP_FIRST_ROW) & ":N" & CStr(AP_LAST_ROW)).Validation.Delete
    ws.Range("N" & CStr(AP_FIRST_ROW) & ":N" & CStr(AP_LAST_ROW)).Locked = True
    ws.Range("K" & CStr(AP_FIRST_ROW) & ":M" & CStr(AP_LAST_ROW)).Locked = False
    On Error GoTo 0
End Sub

Public Sub AP_EnsureConfig()
    Dim ws As Worksheet, lastRow As Long
    On Error GoTo Done
    Set ws = ThisWorkbook.Worksheets("CONFIG")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    AP_UpsertConfig ws, lastRow, "ACTUAL_PICKING_BIN_ROOT", "APFL-", "TEXT", "YES", "USER", "Fixed root used to build Actual Bin."
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    AP_UpsertConfig ws, lastRow, "ACTUAL_PICKING_AREAS", "ZF,P,HR", "TEXT", "YES", "USER", "Allowed Area values for Actual Picking."
Done:
    On Error Resume Next
    ModuleHelper.ResetToolConfigCache
End Sub

Private Sub AP_UpsertConfig(ByVal ws As Worksheet, ByVal lastRow As Long, _
                            ByVal keyText As String, ByVal valueText As String, _
                            ByVal typeText As String, ByVal requiredText As String, _
                            ByVal sourceText As String, ByVal descriptionText As String)
    Dim r As Long, found As Boolean
    For r = 4 To lastRow
        If StrComp(AP_Text(ws.Cells(r, 1).Value2), keyText, vbTextCompare) = 0 Then
            found = True
            If AP_Text(ws.Cells(r, 2).Value2) = "" Then ws.Cells(r, 2).Value2 = valueText
            Exit For
        End If
    Next r
    If Not found Then
        r = lastRow + 1
        ws.Cells(r, 1).Value2 = keyText
        ws.Cells(r, 2).Value2 = valueText
        ws.Cells(r, 3).Value2 = typeText
        ws.Cells(r, 4).Value2 = requiredText
        ws.Cells(r, 5).Value2 = sourceText
        ws.Cells(r, 6).Value2 = descriptionText
    End If
End Sub

Private Function AP_ValidationList() As String
    Dim allowedText As String, parts() As String, i As Long, oneArea As String
    allowedText = ModuleHelper.ToolConfig("ACTUAL_PICKING_AREAS", "ZF,P,HR")
    parts = Split(allowedText, ",")
    For i = LBound(parts) To UBound(parts)
        oneArea = AP_NormalizeArea(parts(i))
        If oneArea <> "" Then
            If AP_ValidationList <> "" Then AP_ValidationList = AP_ValidationList & ","
            AP_ValidationList = AP_ValidationList & oneArea
        End If
    Next i
    If AP_ValidationList = "" Then AP_ValidationList = "ZF,P,HR"
End Function

Private Function AP_LastRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If r < AP_FIRST_ROW Then r = AP_FIRST_ROW - 1
    AP_LastRow = r
End Function

Private Function AP_Text(ByVal rawValue As Variant) As String
    On Error Resume Next
    AP_Text = Trim$(CStr(rawValue & ""))
End Function

Public Sub AP_CompileProbe()
    Dim areaText As String, binNo As String, binText As String, msg As String
    areaText = AP_NormalizeArea("hr")
    AP_Assert areaText = "HR", "AREA_CASE"
    binText = AP_BuildBin("zf", "04-02-10")
    AP_Assert binText = "APFL-ZF04-02-10", "BUILD_BIN"
    AP_SplitFullBin "APFL-HR27-01-04", areaText, binNo
    AP_Assert areaText = "HR" And binNo = "27-01-04", "SPLIT_BIN"
    AP_Assert AP_IsValidBinNo("04-02-10"), "VALID_BIN"
    AP_Assert Not AP_IsValidBinNo("4-2-10"), "INVALID_BIN"
    AP_Assert AP_IsAllowedArea("P"), "AREA_CONFIG"
End Sub

Private Sub AP_Assert(ByVal conditionOK As Boolean, ByVal testName As String)
    If Not conditionOK Then Err.Raise vbObjectError + 9901, , "ACTUAL_PICKING_PROBE_FAIL: " & testName
End Sub
