Attribute VB_Name = "modCF_InputLayout"
Option Explicit

' Layout-only migration for INPUT. The three normal workflow buttons are placed
' outside INPUT A:C so a long OB list cannot cover them. INPUT column A is the
' only OB input channel; this module performs no SAP action.

Public Sub CF_ApplyInputLayout()
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("INPUT")

    CF_RemoveLegacyPanel ws

    ws.Range("A3").Value2 = "TAI WH -> USER NHAP ACTUAL -> CHECK + POSTING -> XU LY DOI BATCH -> CF USER"
    CF_StyleOBInput ws
    ws.Range("P4:S4").Merge
    ws.Range("P4").Value2 = "CONTROL PANEL"
    ws.Range("P5:S5").Merge
    ws.Range("P5").Value2 = "1. TAI WH  ->  2. CHECK  ->  3. XU LY  ->  4. LAM MOI"
    ws.Range("P6:S6").Merge
    ws.Range("P6").Value2 = "Nhap Qty/Bin/Batch o cot vang; 2.CHECK chi doc SAP va lap PLAN (SAP_WRITE=0)"
    ws.Range("P7:S7").Merge
    ws.Range("P7").Value2 = "3.XU LY chi chay 1 PLAN da khoa; dung truoc CF do user xu ly. 4.LAM MOI xoa khoa cu."

    ' V2.8 canonical 4-button action bar (2x2 grid). Captions/macros/colors per
    ' PROMPT V2.8 §4.2. Button 4 (LAM MOI) is new in V2.8.
    CF_EnsureButton ws, "btnCF_ExecuteToCF"
    CF_EnsureButton ws, "btnCF_Reset"
    CF_MoveButton ws, "btnCF_LoadWH", 16, 8, "1. TAI WH", True, "CF_LoadWH"
    CF_MoveButton ws, "btnCF_CheckActual", 18, 8, "2. CHECK", True, "CF_CheckAndLoadPosting"
    CF_MoveButton ws, "btnCF_ExecuteToCF", 16, 10, "3. XU LY", True, "CF_ExecutePlanToCFHandoff"
    CF_MoveButton ws, "btnCF_Reset", 18, 10, "4. LAM MOI", True, "CF_ResetUI"

    ' Diagnostics hidden from the end user (still callable by name for support).
    CF_MoveButton ws, "btnCF_LoadPosting", 16, 14, "TAI POSTING", False, "CF_LoadPosting"
    CF_MoveButton ws, "btnCF_CheckSAP", 18, 14, "KIEM TRA SAP", False, "CF_CheckSAP"
    CF_MoveButton ws, "btnCF_OpenLog", 16, 16, "MO RUN LOG", False, "CF_OpenLog"
    CF_MoveButton ws, "btnCF_MapPOST", 18, 16, "MAP POST", False, "CF_MapSAP_POST"

    ws.columns("P:S").ColumnWidth = 15
    CF_StyleControlPanel ws
    CF_ApplyWorklistHeaders ws
    Exit Sub
Failed:
    RT_Log "INPUT_LAYOUT", "", "", "FAIL", "INPUT_LAYOUT_FAIL: " & Err.Description
End Sub

' Makes the single valid OB entry path visually explicit.  C is deliberately
' reserved: a numeric value there is rejected by the multi-OB reader before
' it can run an older value still present in column A.
Private Sub CF_StyleOBInput(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range("A4:C4").UnMerge
    ws.Range("A4:C4").Merge
    ws.Range("A4").Value2 = "NHAP OB VAO O VANG BEN DUOI -> bam 1. TAI WH. Khong co OB mac dinh."
    With ws.Range("A4:C4")
        .Interior.Color = RGB(255, 244, 204)
        .Font.Color = RGB(100, 70, 0)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ws.Range("A6:C6").UnMerge
    ws.Range("A7:C7").UnMerge
    ws.Range("A9:C9").UnMerge
    ws.Range("A6:C6").Merge
    ws.Range("A6").Value2 = "OB CAN TAI WH - NHAP VAO O VANG BEN DUOI"
    With ws.Range("A6:C6")
        .Interior.Color = RGB(31, 58, 83)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A7:C7").Merge
    With ws.Range("A7:C7")
        .NumberFormat = "@"
        .Locked = False
        .Interior.Color = RGB(255, 242, 204)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A9:C9").Merge
    ws.Range("A9").Value2 = "Danh sach OB bo sung (tuy chon): nhap tu dong 11, cot A."
    With ws.Range("A9:C9")
        .Interior.Color = RGB(236, 242, 247)
        .Font.Color = RGB(31, 41, 55)
        .HorizontalAlignment = xlLeft
    End With
    ws.Range("A10").Value2 = "OB"
    ws.Range("B10").Value2 = "GHI CHU"
    ws.Range("C10").Value2 = "TRANG THAI - KHONG NHAP OB"
    With ws.Range("A10:C10")
        .Interior.Color = RGB(217, 225, 242)
        .Font.Bold = True
    End With
    With ws.Range("A11:A506")
        .NumberFormat = "@"
        .Locked = False
        .Interior.Color = RGB(226, 239, 218)
    End With
    With ws.Range("B11:B506")
        .Locked = False
        .Interior.Color = RGB(255, 255, 255)
    End With
    With ws.Range("C8:C506")
        .Locked = True
        .Interior.Color = RGB(242, 242, 242)
        .Font.Bold = False
    End With
    On Error GoTo 0
End Sub

Private Sub CF_RemoveLegacyPanel(ByVal ws As Worksheet)
    Dim oldPanel As Boolean
    oldPanel = ws.Range("A17").MergeCells And _
               ws.Range("A22").MergeCells And _
               ws.Range("A28").MergeCells
    If Not oldPanel Then Exit Sub

    On Error Resume Next
    ws.Range("A17:B19").UnMerge
    ws.Range("C17:D19").UnMerge
    ws.Range("E17:F19").UnMerge
    ws.Range("G17:I19").UnMerge
    ws.Range("A22:I22").UnMerge
    ws.Range("A23:B25").UnMerge
    ws.Range("C23:D25").UnMerge
    ws.Range("E23:F25").UnMerge
    ws.Range("G23:H25").UnMerge
    ws.Range("I23:I25").UnMerge
    ws.Range("A28:I28").UnMerge
    ws.Range("A17:I31").ClearContents
    ws.Range("A17:I31").ClearFormats
    On Error GoTo 0

    With ws.Range("A17:C31")
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Color = RGB(31, 41, 55)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub CF_StyleControlPanel(ByVal ws As Worksheet)
    With ws.Range("P4:S4")
        .Interior.Color = RGB(31, 58, 83)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("P5:S5")
        .Interior.Color = RGB(55, 110, 168)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("P6:S6")
        .Interior.Color = RGB(255, 244, 204)
        .Font.Color = RGB(100, 70, 0)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("P7:S7")
        .Interior.Color = RGB(236, 242, 247)
        .Font.Color = RGB(31, 41, 55)
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub CF_MoveButton(ByVal ws As Worksheet, ByVal shapeName As String, _
                          ByVal startCol As Long, ByVal startRow As Long, _
                          ByVal captionText As String, ByVal isVisible As Boolean, _
                          ByVal macroName As String, Optional ByVal spanCols As Long = 2)
    Dim sh As Shape, leftPos As Double, topPos As Double, widthPos As Double, macroPrefix As String
    On Error Resume Next
    Set sh = ws.Shapes.item(shapeName)
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub

    leftPos = ws.Cells(startRow, startCol).Left
    topPos = ws.Cells(startRow, startCol).Top
    If spanCols < 1 Then spanCols = 1
    widthPos = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 1, startCol + spanCols - 1)).Width
    sh.Left = leftPos
    sh.Top = topPos
    sh.Width = widthPos
    sh.Height = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 1, startCol + 1)).Height
    sh.Visible = IIf(isVisible, msoTrue, msoFalse)
    sh.Placement = xlMoveAndSize
    macroPrefix = "'" & ThisWorkbook.Name & "'!"
    sh.OnAction = macroPrefix & macroName
    On Error Resume Next
    sh.TextFrame2.TextRange.text = captionText
    On Error GoTo 0

    If isVisible Then
        ' V2.8 color contract (PROMPT §4.2): 1 blue, 2 amber, 3 green, 4 gray.
        If shapeName = "btnCF_LoadWH" Then
            CF_StyleButton sh, RGB(31, 78, 121)          ' blue  - TAI WH
        ElseIf shapeName = "btnCF_CheckActual" Then
            CF_StyleButton sh, RGB(191, 143, 0)          ' amber - CHECK (read-only)
        ElseIf shapeName = "btnCF_ExecuteToCF" Then
            CF_StyleButton sh, RGB(56, 118, 29)          ' green - XU LY
        ElseIf shapeName = "btnCF_Reset" Then
            CF_StyleButton sh, RGB(89, 89, 89)           ' gray  - LAM MOI
        End If
    End If
End Sub

Private Sub CF_EnsureButton(ByVal ws As Worksheet, ByVal shapeName As String)
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes.item(shapeName)
    On Error GoTo 0
    If Not sh Is Nothing Then Exit Sub
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Cells(10, 16).Left, ws.Cells(10, 16).Top, 120, 30)
    sh.Name = shapeName
End Sub

Private Sub CF_StyleButton(ByVal sh As Shape, ByVal fillColor As Long)
    On Error Resume Next
    sh.Fill.Visible = msoTrue
    sh.Fill.Solid
    sh.Fill.ForeColor.RGB = fillColor
    sh.Line.Visible = msoTrue
    sh.Line.ForeColor.RGB = RGB(255, 255, 255)
    sh.Line.Weight = 1.25
    sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    sh.TextFrame2.TextRange.Font.Size = 11
    sh.TextFrame2.TextRange.Font.Bold = msoTrue
    sh.TextFrame2.VerticalAnchor = msoAnchorMiddle
    sh.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    On Error GoTo 0
End Sub

Private Sub CF_ApplyWorklistHeaders(ByVal ws As Worksheet)
    ' Phase 1b worklist header row E4:O4 (Vietnamese labels, col E=5..O=15).
    ' Cols E-J and L-N are read-only mirror; K=Status (AutoFilter); O=Chi tiet.

    ' Clear old control panel at K:N (moved to P:S) and status labels at E:J rows 5-10.
    ' Both areas contain merged cells that would block worklist writes.
    On Error Resume Next
    ws.Range("E4:O10").UnMerge
    ws.Range("E4:O10").ClearContents
    ws.Range("E4:O10").Interior.ColorIndex = xlNone
    On Error GoTo 0

    Dim headers(0 To 10) As String
    Dim c As Long
    headers(0)  = "OB"
    headers(1)  = "Document"
    headers(2)  = "Material"
    headers(3)  = "Target Qty"
    headers(4)  = "Target Bin"
    headers(5)  = "Target Batch"
    headers(6)  = "Status"
    headers(7)  = "WT " & ChrW(272) & ChrW(227)
    headers(8)  = "Qty " & ChrW(272) & ChrW(227)
    headers(9)  = "K" & ChrW(7871) & "t qu" & ChrW(7843)
    headers(10) = "Chi ti" & ChrW(7871) & "t"

    For c = 0 To 10
        ws.Cells(4, 5 + c).Value2 = headers(c)
    Next c

    ' Header row style
    With ws.Range(ws.Cells(4, 5), ws.Cells(4, 15))
        .Interior.Color = RGB(31, 58, 83)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Column widths for worklist
    ws.Columns("E:F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H:J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 8
    ws.Columns("L:M").ColumnWidth = 10
    ws.Columns("N").ColumnWidth = 24
    ws.Columns("O").ColumnWidth = 28
End Sub

Public Sub CF_InputLayout_CompileProbe()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("INPUT")
    If ws Is Nothing Then Err.Raise vbObjectError + 9910, , "INPUT_LAYOUT_PROBE_FAIL: INPUT missing"
    If Not ws.Range("A7").MergeCells Then Err.Raise vbObjectError + 9911, , "INPUT_LAYOUT_PROBE_FAIL: INPUT!A7:C7 must be one entry box"
    If ws.Range("A7").Locked Then Err.Raise vbObjectError + 9911, , "INPUT_LAYOUT_PROBE_FAIL: INPUT!A7 must be editable"
End Sub
