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
    ws.Range("K4:N4").Merge
    ws.Range("K4").Value2 = "CONTROL PANEL"
    ws.Range("K5:N5").Merge
    ws.Range("K5").Value2 = "1. TAI WH  ->  2. USER NHAP  ->  3. CHECK + POSTING  ->  4. XU LY DOI BATCH"
    ws.Range("K6:N6").Merge
    ws.Range("K6").Value2 = "Nhap Qty + Bin; CHECK + POSTING chi doc SAP va lap PLAN"
    ws.Range("K7:N7").Merge
    ws.Range("K7").Value2 = "Nut 4 xu ly 1 PLAN PASS (qty/bin/batch); dung truoc CF do user xu ly"

    CF_EnsureButton ws, "btnCF_ExecuteToCF"
    CF_MoveButton ws, "btnCF_LoadWH", 11, 8, "1. TAI WH", True, "CF_LoadWH"
    CF_MoveButton ws, "btnCF_CheckActual", 13, 8, "3. CHECK + POSTING", True, "CF_CheckAndLoadPosting"
    CF_MoveButton ws, "btnCF_ExecuteToCF", 11, 10, "4. XU LY DEN CF", True, "CF_ExecutePlanToCFHandoff", 4

    CF_MoveButton ws, "btnCF_LoadPosting", 11, 11, "TAI POSTING", False, "CF_LoadPosting"
    CF_MoveButton ws, "btnCF_CheckSAP", 13, 11, "KIEM TRA SAP", False, "CF_CheckSAP"
    CF_MoveButton ws, "btnCF_OpenLog", 11, 14, "MO RUN LOG", False, "CF_OpenLog"
    CF_MoveButton ws, "btnCF_MapPOST", 13, 14, "MAP POST", False, "CF_MapSAP_POST"

    ws.columns("K:N").ColumnWidth = 15
    CF_StyleControlPanel ws
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
    With ws.Range("K4:N4")
        .Interior.Color = RGB(31, 58, 83)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("K5:N5")
        .Interior.Color = RGB(55, 110, 168)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("K6:N6")
        .Interior.Color = RGB(255, 244, 204)
        .Font.Color = RGB(100, 70, 0)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("K7:N7")
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
        If shapeName = "btnCF_LoadWH" Then
            CF_StyleButton sh, RGB(0, 121, 140)
        ElseIf shapeName = "btnCF_CheckActual" Then
            CF_StyleButton sh, RGB(23, 126, 137)
        ElseIf shapeName = "btnCF_ExecuteToCF" Then
            CF_StyleButton sh, RGB(198, 93, 0)
        End If
    End If
End Sub

Private Sub CF_EnsureButton(ByVal ws As Worksheet, ByVal shapeName As String)
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes.item(shapeName)
    On Error GoTo 0
    If Not sh Is Nothing Then Exit Sub
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Cells(10, 11).Left, ws.Cells(10, 11).Top, 120, 30)
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

Public Sub CF_InputLayout_CompileProbe()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("INPUT")
    If ws Is Nothing Then Err.Raise vbObjectError + 9910, , "INPUT_LAYOUT_PROBE_FAIL: INPUT missing"
    If Not ws.Range("A7").MergeCells Then Err.Raise vbObjectError + 9911, , "INPUT_LAYOUT_PROBE_FAIL: INPUT!A7:C7 must be one entry box"
    If ws.Range("A7").Locked Then Err.Raise vbObjectError + 9911, , "INPUT_LAYOUT_PROBE_FAIL: INPUT!A7 must be editable"
End Sub
