Attribute VB_Name = "modCF_ControlLayout"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════════════
' modCF_ControlLayout — SAP Operation Control Center UI (V2.8 CF_CONTROL)
' Design: 5-zone layout without touching business engine
' ═══════════════════════════════════════════════════════════════════════════════

Private Const DEVELOPER_MODE As Boolean = False

' Color scheme (Navy/Blue professional palette)
Private Const COLOR_NAVY As Long = &H5F3A16         ' Header #163A5F
Private Const COLOR_BLUE As Long = &H8C6E1F         ' Accents #1F6E8C
Private Const COLOR_BG As Long = &HF9F6F3           ' Background #F3F6F9
Private Const COLOR_SYSTEM As Long = &HF8F4E8       ' SYSTEM cols #E8F4F8
Private Const COLOR_ACTUAL As Long = &HE6F9FF       ' ACTUAL cols #FFF9E6
Private Const COLOR_PASS As Long = &H549819         ' READY green #198754
Private Const COLOR_WAIT As Long = &H00A8E0         ' WAIT yellow #E0A800
Private Const COLOR_ERROR As Long = &H2B39C0        ' ERROR red #C0392B
Private Const COLOR_DONE As Long = &H2F8443         ' DONE dark green #438F2F

' ═══════════════════════════════════════════════════════════════════════════════
' PUBLIC: Build entire CF_CONTROL layout from scratch
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub CF_BuildControlLayout()
    Dim ws As Worksheet
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Get or rename INPUT → CF_CONTROL
    Set ws = CF_GetControlSheet()

    ' Clear old content
    CF_ClearOldLayout ws

    ' Build 5 zones
    CF_BuildZone1_Header ws
    CF_BuildZone2_Buttons ws
    CF_BuildZone3_OBInput ws
    CF_BuildZone4_DataTable ws
    CF_BuildZone5_StatusBar ws

    ' Apply formatting
    CF_ApplyControlFormatting ws

    ' Set view defaults
    ws.Activate
    ActiveWindow.Zoom = 95
    ws.Range("A1").Select
    ActiveWindow.FreezePanes = False
    ws.Range("A11").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "CF_CONTROL layout built successfully.", vbInformation, "Layout Complete"
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error building layout: " & Err.Description, vbCritical
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' Get or create CF_CONTROL sheet
' ═══════════════════════════════════════════════════════════════════════════════
Private Function CF_GetControlSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")
    On Error GoTo 0

    If ws Is Nothing Then
        ' Try INPUT sheet (rename it)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("INPUT")
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Name = "CF_CONTROL"
        Else
            ' Create new
            Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
            ws.Name = "CF_CONTROL"
        End If
    End If

    Set CF_GetControlSheet = ws
End Function

' ═══════════════════════════════════════════════════════════════════════════════
' Clear old buttons and content (preserve nothing — clean slate)
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub CF_ClearOldLayout(ws As Worksheet)
    Dim shp As Shape
    ' Delete all shapes/buttons
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    ' Clear all cells
    ws.Cells.Clear
    ws.Cells.Interior.ColorIndex = xlNone
    ws.Cells.Font.ColorIndex = xlAutomatic
    ws.Cells.Borders.LineStyle = xlNone
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' ZONE 1: Header (A1:Z3)
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub CF_BuildZone1_Header(ws As Worksheet)
    ' Row 1: Title
    ws.Range("A1:Z1").Merge
    ws.Range("A1").Value = "CF OB CONTROL CENTER"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1").Interior.Color = COLOR_NAVY
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").VerticalAlignment = xlCenter
    ws.Rows("1:1").RowHeight = 30

    ' Row 2: SAP Status + Version
    ws.Range("A2").Value = "SAP Status:"
    ws.Range("B2").Value = "●"
    ws.Range("B2").Name = "SAP_STATUS_INDICATOR"
    ws.Range("B2").Font.Size = 14
    ws.Range("B2").Font.Color = COLOR_PASS  ' Green by default

    ws.Range("C2").Value = "Connected"
    ws.Range("C2").Name = "SAP_STATUS_TEXT"

    ws.Range("Y2").Value = "Version:"
    ws.Range("Z2").Value = "V2.8"
    ws.Range("Z2").Font.Bold = True
    ws.Rows("2:2").RowHeight = 20

    ' Row 3: Empty separator
    ws.Rows("3:3").RowHeight = 5
    ws.Range("A3:Z3").Interior.Color = COLOR_BG
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' ZONE 2: Buttons (3 large buttons at row 5-8)
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub CF_BuildZone2_Buttons(ws As Worksheet)
    Dim btn As Shape
    Dim btnTop As Double, btnHeight As Double, btnWidth As Double

    btnTop = ws.Range("A5").Top
    btnHeight = 50
    btnWidth = 180

    ' Button 1: TẢI WH (Load Warehouse)
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("B5").Left, btnTop, btnWidth, btnHeight)
    btn.Name = "btnLoadWH"
    btn.TextFrame2.TextRange.Text = "TẢI WH"
    btn.TextFrame2.TextRange.Font.Size = 14
    btn.TextFrame2.TextRange.Font.Bold = msoTrue
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btn.Fill.ForeColor.RGB = RGB(31, 110, 140)  ' COLOR_BLUE
    btn.Line.Visible = msoFalse
    btn.TextFrame.HorizontalAlignment = xlHAlignCenter
    btn.TextFrame.VerticalAlignment = xlVAlignCenter
    btn.OnAction = "CF_ButtonLoadWH"

    ' Button 2: CHECK + XỬ LÝ (state-aware)
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("F5").Left, btnTop, btnWidth, btnHeight)
    btn.Name = "btnCheckExecute"
    btn.TextFrame2.TextRange.Text = "CHECK"
    btn.TextFrame2.TextRange.Font.Size = 14
    btn.TextFrame2.TextRange.Font.Bold = msoTrue
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btn.Fill.ForeColor.RGB = RGB(25, 135, 84)  ' COLOR_PASS green
    btn.Line.Visible = msoFalse
    btn.TextFrame.HorizontalAlignment = xlHAlignCenter
    btn.TextFrame.VerticalAlignment = xlVAlignCenter
    btn.OnAction = "CF_ButtonCheckExecute"

    ' Button 3: LÀM MỚI (Reset)
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("J5").Left, btnTop, btnWidth, btnHeight)
    btn.Name = "btnReset"
    btn.TextFrame2.TextRange.Text = "LÀM MỚI"
    btn.TextFrame2.TextRange.Font.Size = 14
    btn.TextFrame2.TextRange.Font.Bold = msoTrue
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btn.Fill.ForeColor.RGB = RGB(192, 57, 43)  ' COLOR_ERROR red
    btn.Line.Visible = msoFalse
    btn.TextFrame.HorizontalAlignment = xlHAlignCenter
    btn.TextFrame.VerticalAlignment = xlVAlignCenter
    btn.OnAction = "CF_ButtonReset"

    ws.Rows("5:8").RowHeight = 15  ' Space for buttons
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' ZONE 3: OB Input Panel (A10:C30, left side)
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub CF_BuildZone3_OBInput(ws As Worksheet)
    ' Header
    ws.Range("A10:C10").Merge
    ws.Range("A10").Value = "NHẬP OB"
    ws.Range("A10").Font.Size = 12
    ws.Range("A10").Font.Bold = True
    ws.Range("A10").Interior.Color = COLOR_BLUE
    ws.Range("A10").Font.Color = RGB(255, 255, 255)
    ws.Range("A10").HorizontalAlignment = xlCenter
    ws.Rows("10:10").RowHeight = 25

    ' Input area A11:C30 (user paste OB list here)
    ws.Range("A11:C30").Interior.Color = RGB(255, 255, 230)  ' Light yellow
    ws.Range("A11:C30").Borders.LineStyle = xlContinuous
    ws.Range("A11:C30").Borders.Color = RGB(200, 200, 200)
    ws.Range("A11").Value = "Paste OB list here"
    ws.Range("A11").Font.Italic = True
    ws.Range("A11").Font.Color = RGB(150, 150, 150)

    ' Stats cells below
    ws.Range("A32").Value = "Count:"
    ws.Range("B32").Value = "=COUNTA(A11:A30)"
    ws.Range("B32").Name = "OB_COUNT"

    ws.Range("A33").Value = "Duplicates:"
    ws.Range("B33").Value = "0"
    ws.Range("B33").Name = "OB_DUPLICATES"
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' ZONE 4: Data Table (D10:Z100+, main worklist)
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub CF_BuildZone4_DataTable(ws As Worksheet)
    Dim startCol As Long, r As Range
    startCol = 4  ' Column D

    ' Group header row 10 (merged across groups)
    ws.Range("D10:J10").Merge
    ws.Range("D10").Value = "SYSTEM (Hệ thống)"
    ws.Range("D10").Interior.Color = COLOR_SYSTEM
    ws.Range("D10").Font.Bold = True
    ws.Range("D10").HorizontalAlignment = xlCenter

    ws.Range("K10:M10").Merge
    ws.Range("K10").Value = "ACTUAL (Thực tế)"
    ws.Range("K10").Interior.Color = COLOR_ACTUAL
    ws.Range("K10").Font.Bold = True
    ws.Range("K10").HorizontalAlignment = xlCenter

    ws.Range("N10:P10").Merge
    ws.Range("N10").Value = "DECISION (Quyết định)"
    ws.Range("N10").Interior.Color = RGB(255, 255, 255)
    ws.Range("N10").Font.Bold = True
    ws.Range("N10").HorizontalAlignment = xlCenter

    ' Column headers row 11
    Dim headers As Variant
    headers = Array("OB", "Document", "Material", "Sys Qty", "Sys Bin", "Sys Batch", "Batch", _
                    "Actual Qty", "Actual Bin", "Actual Batch", _
                    "PLAN", "GATE", "Result")

    Dim colPositions As Variant
    colPositions = Array(4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)  ' D through P

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(11, colPositions(i)).Value = headers(i)
        ws.Cells(11, colPositions(i)).Font.Bold = True
        ws.Cells(11, colPositions(i)).Interior.Color = RGB(230, 230, 230)
        ws.Cells(11, colPositions(i)).HorizontalAlignment = xlCenter
    Next i

    ws.Rows("11:11").RowHeight = 25

    ' Data area D12:P100 (formatted, empty initially)
    Set r = ws.Range("D12:P100")
    r.Borders.LineStyle = xlContinuous
    r.Borders.Color = RGB(200, 200, 200)
    r.Borders.Weight = xlThin

    ' SYSTEM cols (D:J, read-only, light blue)
    ws.Range("D12:J100").Interior.Color = COLOR_SYSTEM
    ws.Range("D12:J100").Locked = True

    ' ACTUAL cols (K:M, editable, light yellow)
    ws.Range("K12:M100").Interior.Color = COLOR_ACTUAL
    ws.Range("K12:M100").Locked = False

    ' DECISION cols (N:P, read-only, white)
    ws.Range("N12:P100").Interior.Color = RGB(255, 255, 255)
    ws.Range("N12:P100").Locked = True

    ' GATE col conditional formatting (col O = 15)
    ws.Range("O12:O100").FormatConditions.Delete
    With ws.Range("O12:O100").FormatConditions.Add(xlCellValue, xlEqual, "=""READY""")
        .Interior.Color = COLOR_PASS
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    With ws.Range("O12:O100").FormatConditions.Add(xlCellValue, xlEqual, "=""WAIT""")
        .Interior.Color = COLOR_WAIT
        .Font.Bold = True
    End With
    With ws.Range("O12:O100").FormatConditions.Add(xlCellValue, xlEqual, "=""BLOCKED""")
        .Interior.Color = COLOR_ERROR
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    With ws.Range("O12:O100").FormatConditions.Add(xlCellValue, xlEqual, "=""DONE""")
        .Interior.Color = COLOR_DONE
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Set column widths
    ws.Columns("D:D").ColumnWidth = 12  ' OB
    ws.Columns("E:E").ColumnWidth = 14  ' Document
    ws.Columns("F:F").ColumnWidth = 12  ' Material
    ws.Columns("G:G").ColumnWidth = 10  ' Sys Qty
    ws.Columns("H:I").ColumnWidth = 12  ' Bins
    ws.Columns("J:J").ColumnWidth = 12  ' Batch
    ws.Columns("K:M").ColumnWidth = 12  ' ACTUAL
    ws.Columns("N:N").ColumnWidth = 20  ' PLAN
    ws.Columns("O:O").ColumnWidth = 10  ' GATE
    ws.Columns("P:P").ColumnWidth = 20  ' Result
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' ZONE 5: Status Bar (A105:Z105, bottom)
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub CF_BuildZone5_StatusBar(ws As Worksheet)
    ws.Range("A105:Z105").Merge
    ws.Range("A105").Value = "Ready"
    ws.Range("A105").Name = "STATUS_BAR"
    ws.Range("A105").Interior.Color = RGB(240, 240, 240)
    ws.Range("A105").Font.Size = 10
    ws.Range("A105").Font.Italic = True
    ws.Range("A105").HorizontalAlignment = xlLeft
    ws.Range("A105").VerticalAlignment = xlCenter
    ws.Rows("105:105").RowHeight = 20
    ws.Range("A105").Borders(xlEdgeTop).LineStyle = xlContinuous
    ws.Range("A105").Borders(xlEdgeTop).Weight = xlMedium
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' Apply overall formatting
' ═══════════════════════════════════════════════════════════════════════════════
Private Sub CF_ApplyControlFormatting(ws As Worksheet)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Tab.Color = COLOR_BLUE

    ' Protect sheet (allow editing ACTUAL cols only)
    ws.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' PUBLIC: Refresh data table from DATA_CORE
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub CF_RefreshControlTable()
    ' Placeholder — will read DATA_CORE and populate D12:P100
    ' Maps columns per 28-point spec
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' PUBLIC: Update status bar
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub CF_UpdateStatusBar(message As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")
    If Not ws Is Nothing Then
        ws.Range("STATUS_BAR").Value = message
    End If
    On Error GoTo 0
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' PUBLIC: Update CHECK/XỬ LÝ button text based on state
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub CF_UpdateCheckExecuteButton()
    Dim ws As Worksheet, btn As Shape, hasPlan As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")
    If ws Is Nothing Then Exit Sub
    Set btn = ws.Shapes("btnCheckExecute")
    If btn Is Nothing Then Exit Sub
    On Error GoTo 0

    ' Check if PLAN sheet has data (simple proxy for "plan exists")
    Dim wsPlan As Worksheet
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    hasPlan = (wsPlan.Range("A5").Value <> "")

    If hasPlan Then
        btn.TextFrame2.TextRange.Text = "XỬ LÝ"
        btn.Fill.ForeColor.RGB = RGB(31, 110, 140)  ' Blue for execute
    Else
        btn.TextFrame2.TextRange.Text = "CHECK"
        btn.Fill.ForeColor.RGB = RGB(25, 135, 84)  ' Green for check
    End If
End Sub