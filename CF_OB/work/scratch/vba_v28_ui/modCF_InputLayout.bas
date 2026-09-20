Attribute VB_Name = "modCF_InputLayout"
Option Explicit

' ═════════════════════════════════════════════════════════════════════════════
' modCF_InputLayout — CF_CONTROL UI Layout Builder (28-point spec)
' ═════════════════════════════════════════════════════════════════════════════
' DESIGN GOALS:
' - Single visible sheet: CF_CONTROL (all technical sheets VeryHidden)
' - Professional operations dashboard: fits A:N at 90-100% zoom, NO horizontal scroll
' - Navy/blue palette, tall headers, large action buttons
' - Clear visual hierarchy: SYSTEM (gray, read-only) | ACTUAL (yellow, editable) | DECISION (white)
' ═════════════════════════════════════════════════════════════════════════════

' ─────────────────────────────────────────────────────────────────────────────
' Layout constants (28-point spec)
' ─────────────────────────────────────────────────────────────────────────────
Private Const LAYOUT_FIRST_COL As Long = 1   ' A
Private Const LAYOUT_LAST_COL As Long = 14   ' N
Private Const LAYOUT_ZOOM As Long = 95

' Zone coordinates
Private Const Z1_HEADER_ROW As Long = 1
Private Const Z1_HEADER_HEIGHT As Double = 36#

Private Const Z2_SUBHEADER_ROW As Long = 2
Private Const Z2_SUBHEADER_HEIGHT As Double = 24#

Private Const Z3_BUTTONS_ROW As Long = 4
Private Const Z3_BUTTONS_HEIGHT As Double = 48#

Private Const Z4_LEFT_PANEL_START_ROW As Long = 6
Private Const Z4_LEFT_PANEL_END_ROW As Long = 40

Private Const Z5_WORKLIST_HEADER_ROW As Long = 8
Private Const Z5_WORKLIST_DATA_START As Long = 9
Private Const Z5_WORKLIST_DATA_END As Long = 40

Private Const Z6_STATUS_BAR_ROW As Long = 42
Private Const Z6_STATUS_BAR_HEIGHT As Double = 22#

' Column layout
Private Const COL_OB_INPUT As Long = 1       ' A: OB input (LEFT panel)
Private Const COL_OB_COUNT As Long = 2       ' B: Count label
Private Const COL_OB_DUP As Long = 3         ' C: Dup count

Private Const COL_WORKLIST_START As Long = 5 ' E: Worklist starts
Private Const COL_WL_OB As Long = 5          ' E
Private Const COL_WL_DOC As Long = 6         ' F
Private Const COL_WL_MAT As Long = 7         ' G
Private Const COL_WL_PLAN As Long = 8        ' H
Private Const COL_WL_GATE As Long = 9        ' I
Private Const COL_WL_TGT_QTY As Long = 10    ' J
Private Const COL_WL_TGT_BIN As Long = 11    ' K
Private Const COL_WL_TGT_BATCH As Long = 12  ' L
Private Const COL_WL_RESULT_VN As Long = 13  ' M
Private Const COL_WL_DETAIL As Long = 14     ' N

' Colors (Navy/Blue palette)
Private Const CLR_NAVY As Long = &H5C3D14       ' RGB(20,61,92) — header
Private Const CLR_BLUE_MEDIUM As Long = &H8B6914 ' RGB(20,105,139) — subheader
Private Const CLR_GRAY_SYSTEM As Long = &HE7E6E6 ' RGB(230,230,231) — SYSTEM cols
Private Const CLR_YELLOW_ACTUAL As Long = &HC0FFFF ' RGB(255,255,192) — ACTUAL cols
Private Const CLR_WHITE As Long = &HFFFFFF
Private Const CLR_TEXT_WHITE As Long = &HFFFFFF
Private Const CLR_TEXT_BLACK As Long = &H0

' Button colors
Private Const BTN_CLR_LOAD As Long = &HC86400     ' RGB(0,100,200) — blue
Private Const BTN_CLR_CHECK As Long = &H00C000    ' RGB(0,192,0) — green
Private Const BTN_CLR_EXECUTE As Long = &HFF8000  ' RGB(0,128,255) — orange-blue
Private Const BTN_CLR_RESET As Long = &H808080    ' RGB(128,128,128) — gray

' ═════════════════════════════════════════════════════════════════════════════
' PUBLIC: Build complete CF_CONTROL layout
' ═════════════════════════════════════════════════════════════════════════════
Public Sub CF_ApplyInputLayout()
    Dim ws As Worksheet

    On Error GoTo Failed
    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")

    ' Clear existing layout (preserve data in worklist area if any)
    Call CF_ClearFormatting(ws)

    ' Build zones
    Call CF_BuildZone1_Header(ws)
    Call CF_BuildZone2_Subheader(ws)
    Call CF_BuildZone3_Buttons(ws)
    Call CF_BuildZone4_LeftPanel(ws)
    Call CF_BuildZone5_Worklist(ws)
    Call CF_BuildZone6_StatusBar(ws)

    ' Apply window settings
    Call CF_ApplyWindowSettings(ws)

    ' Create Named Range for status cell (replaces old H1:N3)
    Call CF_DefineNamedRange(ws)

    Application.ScreenUpdating = True
    If Application.UserControl Then
        MsgBox "CF_CONTROL layout built successfully", vbInformation, "SAP CF OB"
    End If
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    If Application.UserControl Then
        MsgBox "Layout build failed: " & Err.Description, vbCritical, "SAP CF OB"
    End If
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Zone 1: Main header (row 1, A1:N1 merged, Navy, 36px tall)
' ─────────────────────────────────────────────────────────────────────────────
Private Sub CF_BuildZone1_Header(ws As Worksheet)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(Z1_HEADER_ROW, LAYOUT_FIRST_COL), _
                       ws.Cells(Z1_HEADER_ROW, LAYOUT_LAST_COL))

    rng.Merge
    rng.Value = "CF OB OPERATIONS DASHBOARD"

    With rng
        .Interior.Color = CLR_NAVY
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = CLR_TEXT_WHITE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Rows(Z1_HEADER_ROW).RowHeight = Z1_HEADER_HEIGHT
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Zone 2: Subheader (row 2, split: A2:H2 status, I2:N2 version)
' ─────────────────────────────────────────────────────────────────────────────
Private Sub CF_BuildZone2_Subheader(ws As Worksheet)
    Dim rngStatus As Range, rngVersion As Range

    Set rngStatus = ws.Range("A2:H2")
    Set rngVersion = ws.Range("I2:N2")

    ' Left side: SAP connection status
    rngStatus.Merge
    rngStatus.Value = "SAP Status: ● Connected"
    With rngStatus
        .Interior.Color = CLR_BLUE_MEDIUM
        .Font.Size = 10
        .Font.Color = CLR_TEXT_WHITE
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' Right side: Version
    rngVersion.Merge
    rngVersion.Value = "Version: V2.8"
    With rngVersion
        .Interior.Color = CLR_BLUE_MEDIUM
        .Font.Size = 10
        .Font.Color = CLR_TEXT_WHITE
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

    ws.Rows(Z2_SUBHEADER_ROW).RowHeight = Z2_SUBHEADER_HEIGHT
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Zone 3: Action buttons row (row 4, 48px tall)
' ─────────────────────────────────────────────────────────────────────────────
Private Sub CF_BuildZone3_Buttons(ws As Worksheet)
    ws.Rows(Z3_BUTTONS_ROW).RowHeight = Z3_BUTTONS_HEIGHT

    ' Clear old buttons first
    Call CF_ClearOldButtons(ws)

    ' Create 3 buttons: TẢI WH | CHECK+XỬ LÝ | LÀM MỚI
    Call CF_CreateButton(ws, "btnCF_LoadWH", "TẢI WH", _
                         Z3_BUTTONS_ROW, 5, BTN_CLR_LOAD, "CF_ButtonLoadWH")

    Call CF_CreateButton(ws, "btnCF_CheckExecute", "CHECK", _
                         Z3_BUTTONS_ROW, 8, BTN_CLR_CHECK, "CF_ButtonCheckExecute")

    Call CF_CreateButton(ws, "btnCF_Reset", "LÀM MỚI", _
                         Z3_BUTTONS_ROW, 11, BTN_CLR_RESET, "CF_ButtonReset")
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Zone 4: LEFT panel — OB input area (A6:C40)
' ─────────────────────────────────────────────────────────────────────────────
Private Sub CF_BuildZone4_LeftPanel(ws As Worksheet)
    Dim rngHeader As Range, rngInput As Range, rngStats As Range

    ' Header: "NHẬP OB" (A6:C6)
    Set rngHeader = ws.Range("A6:C6")
    rngHeader.Merge
    rngHeader.Value = "NHẬP OB"
    With rngHeader
        .Interior.Color = CLR_BLUE_MEDIUM
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = CLR_TEXT_WHITE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(6).RowHeight = 22

    ' OB input: A7 (CRITICAL: engine reads A7 for OB)
    Set rngInput = ws.Range("A7")
    With rngInput
        .Interior.Color = CLR_YELLOW_ACTUAL
        .Font.Size = 10
        .Locked = False
        .BorderAround Weight:=xlThin
    End With

    ' OB list: A8:A40 (multi-OB, yellow, editable)
    Set rngInput = ws.Range("A8:A40")
    With rngInput
        .Interior.Color = CLR_YELLOW_ACTUAL
        .Font.Size = 9
        .Locked = False
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
    End With

    ' Stats labels: B6:C6 merged header, B41:C41 stats display
    ws.Range("B41").Value = "Count:"
    ws.Range("C41").Value = "0"
    ws.Range("B41:C41").Font.Size = 9
    ws.Range("B41:C41").Interior.Color = CLR_GRAY_SYSTEM
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Zone 5: MAIN worklist table (E8:N40)
' ─────────────────────────────────────────────────────────────────────────────
Private Sub CF_BuildZone5_Worklist(ws As Worksheet)
    Dim rngHeader As Range, rngData As Range

    ' Column group headers (E6:N6, merged groups)
    ' SYSTEM group: E6:G6
    ws.Range("E6:G6").Merge
    ws.Range("E6").Value = "HỆ THỐNG"
    With ws.Range("E6")
        .Interior.Color = CLR_GRAY_SYSTEM
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' DECISION group: H6:I6
    ws.Range("H6:I6").Merge
    ws.Range("H6").Value = "KẾ HOẠCH"
    With ws.Range("H6")
        .Interior.Color = CLR_GRAY_SYSTEM
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' ACTUAL group: J6:L6
    ws.Range("J6:L6").Merge
    ws.Range("J6").Value = "THỰC TẾ"
    With ws.Range("J6")
        .Interior.Color = CLR_YELLOW_ACTUAL
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' RESULT group: M6:N6
    ws.Range("M6:N6").Merge
    ws.Range("M6").Value = "KẾT QUẢ"
    With ws.Range("M6")
        .Interior.Color = CLR_WHITE
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ws.Rows(6).RowHeight = 20

    ' Individual column headers (row 8)
    Dim headers As Variant
    headers = Array("OB", "Document", "Material", "PLAN", "GATE", _
                    "Target Qty", "Target Bin", "Target Batch", "Kết quả", "Chi tiết")

    Dim col As Long
    For col = COL_WORKLIST_START To COL_WL_DETAIL
        ws.Cells(Z5_WORKLIST_HEADER_ROW, col).Value = headers(col - COL_WORKLIST_START)
        With ws.Cells(Z5_WORKLIST_HEADER_ROW, col)
            .Font.Size = 10
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .BorderAround Weight:=xlThin

            ' Color code by group
            If col >= COL_WL_OB And col <= COL_WL_MAT Then
                .Interior.Color = CLR_GRAY_SYSTEM ' SYSTEM
            ElseIf col = COL_WL_PLAN Or col = COL_WL_GATE Then
                .Interior.Color = CLR_GRAY_SYSTEM ' DECISION (Plan/Gate)
            ElseIf col >= COL_WL_TGT_QTY And col <= COL_WL_TGT_BATCH Then
                .Interior.Color = CLR_YELLOW_ACTUAL ' ACTUAL
            Else
                .Interior.Color = CLR_WHITE ' RESULT
            End If
        End With
    Next col

    ws.Rows(Z5_WORKLIST_HEADER_ROW).RowHeight = 22

    ' Data area (E9:N40) — borders and locking
    Set rngData = ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WORKLIST_START), _
                           ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_DETAIL))

    With rngData
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlHairline
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Font.Size = 9
        .RowHeight = 20
    End With

    ' SYSTEM cols (E:G) — gray, locked
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_OB), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_MAT)).Interior.Color = CLR_GRAY_SYSTEM
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_OB), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_MAT)).Locked = True

    ' PLAN/GATE cols (H:I) — gray, locked
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_PLAN), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_GATE)).Interior.Color = CLR_GRAY_SYSTEM
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_PLAN), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_GATE)).Locked = True

    ' ACTUAL cols (J:L) — yellow, editable
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_TGT_QTY), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_TGT_BATCH)).Interior.Color = CLR_YELLOW_ACTUAL
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_TGT_QTY), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_TGT_BATCH)).Locked = False

    ' RESULT cols (M:N) — white, locked
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_RESULT_VN), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_DETAIL)).Interior.Color = CLR_WHITE
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_RESULT_VN), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_DETAIL)).Locked = True
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Zone 6: Status bar (row 42, A42:N42 merged)
' ─────────────────────────────────────────────────────────────────────────────
Private Sub CF_BuildZone6_StatusBar(ws As Worksheet)
    Dim rng As Range
    Set rng = ws.Range("A42:N42")

    rng.Merge
    rng.Value = "Ready"

    With rng
        .Interior.Color = CLR_GRAY_SYSTEM
        .Font.Size = 9
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .BorderAround Weight:=xlThin
    End With

    ws.Rows(Z6_STATUS_BAR_ROW).RowHeight = Z6_STATUS_BAR_HEIGHT
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Helper: Clear old formatting (preserve data in worklist)
' ─────────────────────────────────────────────────────────────────────────────
Private Sub CF_ClearFormatting(ws As Worksheet)
    ' Clear formats only in non-data areas
    ws.Range("A1:N5").ClearFormats
    ws.Range("A6:D40").ClearFormats
    ws.Range("A41:N50").ClearFormats
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Helper: Clear old buttons (remove CP10/CP11 legacy buttons)
' ─────────────────────────────────────────────────────────────────────────────
Private Sub CF_ClearOldButtons(ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 5) = "btnCF" Then
            shp.Delete
        End If
    Next shp
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Helper: Create a styled button
' ─────────────────────────────────────────────────────────────────────────────
Private Sub CF_CreateButton(ws As Worksheet, btnName As String, btnText As String, _
                            btnRow As Long, btnCol As Long, btnColor As Long, macroName As String)
    Dim btn As Shape
    Dim cellTop As Double, cellLeft As Double

    ' Remove old button with same name
    On Error Resume Next
    ws.Shapes(btnName).Delete
    On Error GoTo 0

    ' Get cell position
    cellTop = ws.Cells(btnRow, btnCol).Top + 4
    cellLeft = ws.Cells(btnRow, btnCol).Left

    ' Create rounded rectangle button
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, cellLeft, cellTop, 90, 36)
    btn.Name = btnName
    btn.TextFrame2.TextRange.Text = btnText

    With btn
        .Fill.ForeColor.RGB = btnColor
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Bold = True
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_TEXT_WHITE
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = macroName
    End With
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Apply window settings (zoom, freeze, hide gridlines/headings)
' ─────────────────────────────────────────────────────────────────────────────
Private Sub CF_ApplyWindowSettings(ws As Worksheet)
    ws.Activate

    With ActiveWindow
        .Zoom = LAYOUT_ZOOM
        .DisplayGridlines = False
        .DisplayHeadings = False

        ' Freeze pane after row 8 (worklist header)
        .FreezePanes = False
        ws.Range("A9").Select
        .FreezePanes = True

        ' Return to top
        ws.Range("A1").Select
    End With
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Define Named Range for status cell (replaces old H1:N3)
' ─────────────────────────────────────────────────────────────────────────────
Private Sub CF_DefineNamedRange(ws As Worksheet)
    On Error Resume Next
    ThisWorkbook.Names("CF_UI_STATUS_CELL").Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:="CF_UI_STATUS_CELL", _
                           RefersTo:=ws.Range("A42:N42")
End Sub

' ═════════════════════════════════════════════════════════════════════════════
' PUBLIC: Update CHECK/XỬ LÝ button state (state-aware button)
' ═════════════════════════════════════════════════════════════════════════════
Public Sub CF_UpdateCheckExecuteButton()
    Dim ws As Worksheet, wsPlan As Worksheet
    Dim btn As Shape
    Dim hasPlan As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set btn = ws.Shapes("btnCF_CheckExecute")

    If btn Is Nothing Then Exit Sub

    ' Check if PLAN sheet has data (PLAN!AJ2 <> empty)
    hasPlan = False
    If Not wsPlan Is Nothing Then
        hasPlan = (Trim(CStr(wsPlan.Range("AJ2").Value2)) <> "")
    End If

    If hasPlan Then
        ' Change to XỬ LÝ mode (orange-blue)
        btn.TextFrame2.TextRange.Text = "XỬ LÝ"
        btn.Fill.ForeColor.RGB = BTN_CLR_EXECUTE
    Else
        ' CHECK mode (green)
        btn.TextFrame2.TextRange.Text = "CHECK"
        btn.Fill.ForeColor.RGB = BTN_CLR_CHECK
    End If

    On Error GoTo 0
End Sub

' ═════════════════════════════════════════════════════════════════════════════
' PUBLIC: Update status bar message
' ═════════════════════════════════════════════════════════════════════════════
Public Sub CF_UpdateStatusBar(message As String)
    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.Names("CF_UI_STATUS_CELL").RefersToRange
    If Not rng Is Nothing Then
        rng.Value = message
    End If
    On Error GoTo 0
End Sub

' ═════════════════════════════════════════════════════════════════════════════
' OFFLINE PROBE: Verify layout constants and structure
' ═════════════════════════════════════════════════════════════════════════════
Public Sub CF_InputLayout_CompileProbe()
    Dim ws As Worksheet
    On Error GoTo Failed

    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")

    ' Verify A7 is OB input cell (engine anchor)
    If Not ws.Range("A7").Locked Then
        Debug.Print "PROBE_PASS: A7 is editable (engine anchor)"
    Else
        Debug.Print "PROBE_FAIL: A7 is locked"
    End If

    ' Verify worklist area exists
    If ws.Cells(Z5_WORKLIST_HEADER_ROW, COL_WORKLIST_START).Value <> "" Then
        Debug.Print "PROBE_PASS: Worklist headers present"
    Else
        Debug.Print "PROBE_WARN: Worklist headers empty (run CF_ApplyInputLayout first)"
    End If

    ' Verify Named Range exists
    On Error Resume Next
    Dim testRng As Range
    Set testRng = ThisWorkbook.Names("CF_UI_STATUS_CELL").RefersToRange
    If Not testRng Is Nothing Then
        Debug.Print "PROBE_PASS: CF_UI_STATUS_CELL Named Range exists at " & testRng.Address
    Else
        Debug.Print "PROBE_FAIL: CF_UI_STATUS_CELL Named Range missing"
    End If
    On Error GoTo 0

    Debug.Print "CF_InputLayout_CompileProbe: DONE"
    Exit Sub

Failed:
    Debug.Print "CF_InputLayout_CompileProbe: FAIL - " & Err.Description
End Sub
