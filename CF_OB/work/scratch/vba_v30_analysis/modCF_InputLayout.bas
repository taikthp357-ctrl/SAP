Attribute VB_Name = "modCF_InputLayout"
Option Explicit

' -----------------------------------------------------------------------------
' modCF_InputLayout � CF_CONTROL UI Layout Builder (28-point spec)
' -----------------------------------------------------------------------------
' DESIGN GOALS:
' - Single visible sheet: CF_CONTROL (all technical sheets VeryHidden)
' - Professional operations dashboard: fits A:N at 90-100% zoom, NO horizontal scroll
' - Navy/blue palette, tall headers, large action buttons
' - Clear visual hierarchy: SYSTEM (gray, read-only) | ACTUAL (yellow, editable) | DECISION (white)
' -----------------------------------------------------------------------------

' -----------------------------------------------------------------------------
' Layout constants (28-point spec)
' -----------------------------------------------------------------------------
Private Const LAYOUT_FIRST_COL As Long = 1   ' A
Private Const LAYOUT_LAST_COL As Long = 16   ' P (was N=14; +2 for WT+Material)
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
Private Const Z5_WORKLIST_DATA_END As Long = 200

Private Const Z6_STATUS_BAR_ROW As Long = 3
Private Const Z6_STATUS_BAR_HEIGHT As Double = 22#

' Column layout
Private Const COL_OB_INPUT As Long = 1       ' A: OB input (LEFT panel)
Private Const COL_OB_COUNT As Long = 2       ' B: Count label
Private Const COL_OB_DUP As Long = 3         ' C: Dup count

Private Const COL_WORKLIST_START As Long = 5 ' E: Worklist starts
Private Const COL_WL_OB As Long = 5          ' E  OB
Private Const COL_WL_WHO As Long = 6         ' F  Whse Order
Private Const COL_WL_WT As Long = 7          ' G  WT (ma can sua)   NEW
Private Const COL_WL_MAT As Long = 8         ' H  Material          NEW pos
Private Const COL_WL_STATUS As Long = 9      ' I  Tinh trang (trong/A)
Private Const COL_WL_PLAN As Long = 10       ' J  KE HOACH
Private Const COL_WL_GATE As Long = 11       ' K  GATE
Private Const COL_WL_TGT_QTY As Long = 12    ' L  SL thuc te  (editable)
Private Const COL_WL_TGT_BIN As Long = 13    ' M  Khu vuc     (editable; stale name = Area)
Private Const COL_WL_TGT_BATCH As Long = 14  ' N  So bin      (editable; stale name = Bin No)
Private Const COL_WL_RESULT_VN As Long = 15  ' O  Ket qua
Private Const COL_WL_DETAIL As Long = 16     ' P  Chi tiet

' Colors (Navy/Blue palette)
Private Const CLR_NAVY As Long = &H5C3D14       ' RGB(20,61,92) � header
Private Const CLR_BLUE_MEDIUM As Long = &H8B6914 ' RGB(20,105,139) � subheader
Private Const CLR_GRAY_SYSTEM As Long = &HE7E6E6 ' RGB(230,230,231) � SYSTEM cols
Private Const CLR_YELLOW_ACTUAL As Long = &HC0FFFF ' RGB(255,255,192) � ACTUAL cols
Private Const CLR_WHITE As Long = &HFFFFFF
Private Const CLR_TEXT_WHITE As Long = &HFFFFFF
Private Const CLR_TEXT_BLACK As Long = &H0

' Button colors
Private Const BTN_CLR_LOAD As Long = &HC86400     ' RGB(0,100,200) � blue
Private Const BTN_CLR_CHECK As Long = &HC000      ' RGB(0,192,0) � green
Private Const BTN_CLR_EXECUTE As Long = &HFF8000  ' RGB(0,128,255) � orange-blue
Private Const BTN_CLR_RESET As Long = &H808080    ' RGB(128,128,128) � gray

' -----------------------------------------------------------------------------
' PUBLIC: Build complete CF_CONTROL layout
' -----------------------------------------------------------------------------
Public Sub CF_ApplyInputLayout()
    Dim ws As Worksheet

    On Error GoTo Failed
    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")

    ' Clear existing layout (preserve data in worklist area if any)
    Call CF_ClearLegacyLayout(ws)

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
    Call CF_EnforceSingleVisibleSheet(ws)

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

' -----------------------------------------------------------------------------
' Zone 1: Main header (row 1, A1:P1 merged, Navy, 36px tall)
' -----------------------------------------------------------------------------
Private Sub CF_BuildZone1_Header(ws As Worksheet)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(Z1_HEADER_ROW, LAYOUT_FIRST_COL), _
                       ws.Cells(Z1_HEADER_ROW, LAYOUT_LAST_COL))

    rng.Merge
    rng.value = "CF OB - V3.0 MULTI-OB CONTROL"

    With rng
        .Interior.Color = CLR_NAVY
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = CLR_TEXT_WHITE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.rows(Z1_HEADER_ROW).RowHeight = Z1_HEADER_HEIGHT
End Sub

' -----------------------------------------------------------------------------
' Zone 2: Subheader (row 2, split: A2:J2 status, K2:P2 version)
' -----------------------------------------------------------------------------
Private Sub CF_BuildZone2_Subheader(ws As Worksheet)
    Dim rngStatus As Range, rngVersion As Range

    Set rngStatus = ws.Range("A2:J2")
    Set rngVersion = ws.Range("K2:P2")

    ' Left side: SAP connection status
    rngStatus.Merge
    rngStatus.value = "SAP Status: Chua kiem tra"
    With rngStatus
        .Interior.Color = CLR_BLUE_MEDIUM
        .Font.Size = 10
        .Font.Color = CLR_TEXT_WHITE
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    ' Right side: Version
    rngVersion.Merge
    rngVersion.value = "Version: V3.0"
    With rngVersion
        .Interior.Color = CLR_BLUE_MEDIUM
        .Font.Size = 10
        .Font.Color = CLR_TEXT_WHITE
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

    ws.rows(Z2_SUBHEADER_ROW).RowHeight = Z2_SUBHEADER_HEIGHT
End Sub

' -----------------------------------------------------------------------------
' Zone 3: Action buttons row (row 4, 48px tall)
' -----------------------------------------------------------------------------
Private Sub CF_BuildZone3_Buttons(ws As Worksheet)
    ws.rows(Z3_BUTTONS_ROW).RowHeight = Z3_BUTTONS_HEIGHT

    ' Clear old buttons first
    Call CF_ClearOldButtons(ws)

    ' Create 3 buttons: T?I WH | CHECK+X? L� | L�M M?I
    Call CF_CreateButton(ws, "btnCF_LoadWH", CF_UIText("LOAD_WH"), _
                         Z3_BUTTONS_ROW, 5, BTN_CLR_LOAD, "CF_ButtonLoadWH")

    Call CF_CreateButton(ws, "btnCF_CheckExecute", "CHECK", _
                         Z3_BUTTONS_ROW, 8, BTN_CLR_CHECK, "CF_ButtonCheckExecute")

    Call CF_CreateButton(ws, "btnCF_Reset", CF_UIText("RESET"), _
                         Z3_BUTTONS_ROW, 11, BTN_CLR_RESET, "CF_ButtonReset")
End Sub

' -----------------------------------------------------------------------------
' Zone 4: LEFT panel � OB input area (A6:C40)
' -----------------------------------------------------------------------------
Private Sub CF_BuildZone4_LeftPanel(ws As Worksheet)
    Dim rngHeader As Range, rngInput As Range, rngStats As Range

    ' Header: "NH?P OB" (A6:C6)
    Set rngHeader = ws.Range("A6:B6")
    rngHeader.Merge
    rngHeader.value = "DANH SACH OB"
    With rngHeader
        .Interior.Color = CLR_BLUE_MEDIUM
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = CLR_TEXT_WHITE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.rows(6).RowHeight = 22
    ws.Range("C6").Value2 = "Tinh trang"
    ws.Range("C6").Interior.Color = CLR_BLUE_MEDIUM
    ws.Range("C6").Font.Color = CLR_TEXT_WHITE
    ws.Range("C6").Font.Bold = True
    ws.Range("C6").HorizontalAlignment = xlCenter

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
    ws.Range("A41").value = "Tong OB:"
    ws.Range("C41").value = "0"
    ws.Range("A41:C41").Font.Size = 9
    ws.Range("A41:C41").Interior.Color = CLR_GRAY_SYSTEM
End Sub

' -----------------------------------------------------------------------------
' Zone 5: MAIN worklist table (E8:P200)
' -----------------------------------------------------------------------------
Private Sub CF_BuildZone5_Worklist(ws As Worksheet)
    Dim rngHeader As Range, rngData As Range

    ' Column group headers (E6:P6, merged groups)
    ' SYSTEM group: E6:I6 (OB, Whse Order, WT, Material, Tinh trang)
    ws.Range("E6:I6").Merge
    ws.Range("E6").value = CF_UIText("SYSTEM")
    With ws.Range("E6")
        .Interior.Color = CLR_GRAY_SYSTEM
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' DECISION group: J6:K6 (KE HOACH, GATE)
    ws.Range("J6:K6").Merge
    ws.Range("J6").value = CF_UIText("PLAN")
    With ws.Range("J6")
        .Interior.Color = CLR_GRAY_SYSTEM
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' ACTUAL group: L6:N6 (SL thuc te, Khu vuc, So bin)
    ws.Range("L6:N6").Merge
    ws.Range("L6").value = CF_UIText("ACTUAL")
    With ws.Range("L6")
        .Interior.Color = CLR_YELLOW_ACTUAL
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' RESULT group: O6:P6 (Ket qua, Chi tiet)
    ws.Range("O6:P6").Merge
    ws.Range("O6").value = CF_UIText("RESULT")
    With ws.Range("O6")
        .Interior.Color = CLR_WHITE
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ws.rows(6).RowHeight = 20

    ' Individual column headers (row 8) � 12 cols E:P
    Dim headers As Variant
    headers = Array("OB", "Whse Order", "WT", "Material", "Tinh trang", "KE HOACH", "GATE", _
                    "SL thuc te", "Khu vuc", "So bin", CF_UIText("RESULT"), CF_UIText("DETAIL"))

    Dim col As Long
    For col = COL_WORKLIST_START To COL_WL_DETAIL
        ws.Cells(Z5_WORKLIST_HEADER_ROW, col).value = headers(col - COL_WORKLIST_START)
        With ws.Cells(Z5_WORKLIST_HEADER_ROW, col)
            .Font.Size = 10
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .BorderAround Weight:=xlThin

            ' Color code by group
            If col >= COL_WL_OB And col <= COL_WL_STATUS Then
                .Interior.Color = CLR_GRAY_SYSTEM ' SYSTEM (E:I)
            ElseIf col = COL_WL_PLAN Or col = COL_WL_GATE Then
                .Interior.Color = CLR_GRAY_SYSTEM ' DECISION (Plan/Gate J:K)
            ElseIf col >= COL_WL_TGT_QTY And col <= COL_WL_TGT_BATCH Then
                .Interior.Color = CLR_YELLOW_ACTUAL ' ACTUAL (L:N)
            Else
                .Interior.Color = CLR_WHITE ' RESULT (O:P)
            End If
        End With
    Next col

    ws.rows(Z5_WORKLIST_HEADER_ROW).RowHeight = 22
    ws.Range(ws.Cells(Z5_WORKLIST_HEADER_ROW, COL_WORKLIST_START), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_DETAIL)).AutoFilter

    ' Data area (E9:P200) � borders and locking
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

    ' SYSTEM cols (E:I) � gray, locked
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_OB), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_STATUS)).Interior.Color = CLR_GRAY_SYSTEM
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_OB), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_STATUS)).Locked = True

    ' PLAN/GATE cols (J:K) � gray, locked
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_PLAN), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_GATE)).Interior.Color = CLR_GRAY_SYSTEM
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_PLAN), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_GATE)).Locked = True

    ' ACTUAL cols (L:N) � yellow, editable
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_TGT_QTY), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_TGT_BATCH)).Interior.Color = CLR_YELLOW_ACTUAL
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_TGT_QTY), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_TGT_BATCH)).Locked = False

    ' RESULT cols (O:P) � white, locked
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_RESULT_VN), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_DETAIL)).Interior.Color = CLR_WHITE
    ws.Range(ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_RESULT_VN), _
             ws.Cells(Z5_WORKLIST_DATA_END, COL_WL_DETAIL)).Locked = True
End Sub

' -----------------------------------------------------------------------------
' Zone 6: Status bar (row 42, A42:N42 merged)
' -----------------------------------------------------------------------------
Private Sub CF_BuildZone6_StatusBar(ws As Worksheet)
    Dim rng As Range
    Set rng = ws.Range("A3:P3")

    rng.Merge
    rng.value = "Ready"

    With rng
        .Interior.Color = CLR_GRAY_SYSTEM
        .Font.Size = 9
        .Font.Italic = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .BorderAround Weight:=xlThin
    End With

    ws.rows(Z6_STATUS_BAR_ROW).RowHeight = Z6_STATUS_BAR_HEIGHT
End Sub

' -----------------------------------------------------------------------------
' Helper: Clear old formatting (preserve data in worklist)
' -----------------------------------------------------------------------------
Private Sub CF_ClearFormatting(ws As Worksheet)
    ' Clear formats only in non-data areas
    ws.Range("A1:P5").ClearFormats
    ws.Range("A6:D40").ClearFormats
    ws.Range("A41:P50").ClearFormats
End Sub

' -----------------------------------------------------------------------------
' Helper: Clear old buttons (remove CP10/CP11 legacy buttons)
' -----------------------------------------------------------------------------
Private Sub CF_ClearOldButtons(ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 5) = "btnCF" Then
            shp.Delete
        End If
    Next shp
End Sub

' -----------------------------------------------------------------------------
' Helper: Create a styled button
' -----------------------------------------------------------------------------
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
    btn.TextFrame2.TextRange.text = btnText

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

' -----------------------------------------------------------------------------
' Apply window settings (zoom, freeze, hide gridlines/headings)
' -----------------------------------------------------------------------------
Private Sub CF_ApplyWindowSettings(ws As Worksheet)
    ws.Activate

    ws.columns("B:B").Hidden = True
    ws.columns("D:D").Hidden = True
    ws.columns("Q:S").Hidden = True          ' Q = hidden Source Key; O:P now visible (Ket qua/Chi tiet)
    ws.ScrollArea = "A1:P200"
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

' -----------------------------------------------------------------------------
' Define Named Range for status cell (replaces old H1:N3)
' -----------------------------------------------------------------------------
Private Sub CF_DefineNamedRange(ws As Worksheet)
    On Error Resume Next
    ThisWorkbook.names("CF_UI_STATUS_CELL").Delete
    On Error GoTo 0

    ThisWorkbook.names.Add Name:="CF_UI_STATUS_CELL", _
                           RefersTo:=ws.Range("A3:P3")
End Sub

' -----------------------------------------------------------------------------
' PUBLIC: Update CHECK/X? L� button state (state-aware button)
' -----------------------------------------------------------------------------
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
        hasPlan = CF_CurrentPlanIsExecutable()
    End If

    If hasPlan Then
        ' Change to X? L� mode (orange-blue)
        btn.TextFrame2.TextRange.text = CF_UIText("EXECUTE")
        btn.Fill.ForeColor.RGB = BTN_CLR_EXECUTE
    Else
        ' CHECK mode (green)
        btn.TextFrame2.TextRange.text = "CHECK"
        btn.Fill.ForeColor.RGB = BTN_CLR_CHECK
    End If

    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' PUBLIC: Update status bar message
' -----------------------------------------------------------------------------
Public Sub CF_UpdateStatusBar(message As String)
    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.names("CF_UI_STATUS_CELL").RefersToRange
    If Not rng Is Nothing Then
        rng.value = message
    End If
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' OFFLINE PROBE: Verify layout constants and structure
' -----------------------------------------------------------------------------
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
    If ws.Cells(Z5_WORKLIST_HEADER_ROW, COL_WORKLIST_START).value <> "" Then
        Debug.Print "PROBE_PASS: Worklist headers present"
    Else
        Debug.Print "PROBE_WARN: Worklist headers empty (run CF_ApplyInputLayout first)"
    End If

    ' Verify Named Range exists
    On Error Resume Next
    Dim testRng As Range
    Set testRng = ThisWorkbook.names("CF_UI_STATUS_CELL").RefersToRange
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

' -----------------------------------------------------------------------------
' §4 REGRESSION: TEST_UI_V3_LAYOUT � invoke the ONE canonical V3 layout builder,
' then assert the shipped E:P worklist, hidden Q SourceKey, A7 OB anchor, and the
' system/editable lock map. Returns "PASS|..." or "FAIL|...". Offline, no SAP.
' -----------------------------------------------------------------------------
Public Function TEST_UI_V3_LAYOUT() As String
    Dim ws As Worksheet, r As Long, c As Long, exp As Variant
    On Error GoTo Failed
    CF_ApplyInputLayout                       ' the single canonical builder
    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")
    r = Z5_WORKLIST_HEADER_ROW

    ' Row-8 headers E:N are fixed ASCII; O:P come from CF_UIText (diacritics) � non-empty only
    exp = Array("OB", "Whse Order", "WT", "Material", "Tinh trang", _
                "KE HOACH", "GATE", "SL thuc te", "Khu vuc", "So bin")
    For c = 0 To 9
        If CStr(ws.Cells(r, COL_WORKLIST_START + c).value) <> exp(c) Then
            TEST_UI_V3_LAYOUT = "FAIL|hdr col " & (COL_WORKLIST_START + c) & " got '" & _
                CStr(ws.Cells(r, COL_WORKLIST_START + c).value) & "' want '" & exp(c) & "'"
            Exit Function
        End If
    Next c
    If Trim$(CStr(ws.Cells(r, COL_WL_RESULT_VN).value)) = "" Then TEST_UI_V3_LAYOUT = "FAIL|O8 empty": Exit Function
    If Trim$(CStr(ws.Cells(r, COL_WL_DETAIL).value)) = "" Then TEST_UI_V3_LAYOUT = "FAIL|P8 empty": Exit Function

    If Not ws.columns(17).Hidden Then TEST_UI_V3_LAYOUT = "FAIL|Q(17) SourceKey not hidden": Exit Function
    If Not ws.columns(2).Hidden Then TEST_UI_V3_LAYOUT = "FAIL|B not hidden": Exit Function
    If Not ws.columns(4).Hidden Then TEST_UI_V3_LAYOUT = "FAIL|D not hidden": Exit Function
    If ws.Cells(7, 1).Locked Then TEST_UI_V3_LAYOUT = "FAIL|A7 OB anchor locked": Exit Function
    If Not ws.AutoFilterMode Then TEST_UI_V3_LAYOUT = "FAIL|worklist AutoFilter missing": Exit Function
    If ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_TGT_QTY).Locked Then TEST_UI_V3_LAYOUT = "FAIL|L9 not editable": Exit Function
    If ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_TGT_BATCH).Locked Then TEST_UI_V3_LAYOUT = "FAIL|N9 not editable": Exit Function
    If Not ws.Cells(Z5_WORKLIST_DATA_START, COL_WL_OB).Locked Then TEST_UI_V3_LAYOUT = "FAIL|E9 not locked": Exit Function

    TEST_UI_V3_LAYOUT = "PASS|E:P headers, Q hidden SourceKey, A7 anchor unlocked, L:N editable, filter on"
    Exit Function
Failed:
    TEST_UI_V3_LAYOUT = "FAIL|ERR " & Err.Number & " " & Err.Description
End Function

Private Sub CF_ClearLegacyLayout(ByVal ws As Worksheet)
    Dim savedOBs As Variant, r As Long, obText As String
    savedOBs = ws.Range("A7:A40").Value2
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:S506").UnMerge
    On Error GoTo 0
    ws.Range("A1:S506").ClearContents
    ws.Range("A1:S506").ClearFormats
    For r = 1 To UBound(savedOBs, 1)
        obText = Trim$(CStr(savedOBs(r, 1) & ""))
        If obText <> "" And IsNumeric(obText) Then ws.Cells(r + 6, 1).Value2 = savedOBs(r, 1)
    Next r
End Sub

Private Sub CF_EnforceSingleVisibleSheet(ByVal controlSheet As Worksheet)
    Dim ws As Worksheet
    controlSheet.Visible = xlSheetVisible
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> controlSheet.Name Then ws.Visible = xlSheetVeryHidden
    Next ws
End Sub


