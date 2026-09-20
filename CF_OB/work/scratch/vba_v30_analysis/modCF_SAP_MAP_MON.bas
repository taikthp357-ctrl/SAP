Attribute VB_Name = "modCF_SAP_MAP_MON"
Option Explicit

' MON mapping helper derived from the stable SAP map/dump pattern.
'
' Image-confirmed semantic targets:
'   MON_MATERIAL_ID       Product / Material input
'   MON_BIN_ID            Storage Bin input
'   MON_STATUS_ID         Warehouse Task Status filter
'   MON_RESULT_GRID_ID    Warehouse Task result grid
'   MON_WT_COL             Whse Order column
'   MON_OB_COL             Document / OB column
'   MON_MATERIAL_COL       Product / Material column
'   MON_BATCH_COL          Batch column
'   MON_BIN_COL            Source Bin column
'   MON_QTY_COL            SrcTgtQty / Dest.Qty candidate column
'   MON_STATUS_COL         Status column
'   MON_CANCEL_ID          Cancel WT action
'   MON_CANCEL_CONFIRM_ID  Cancel confirmation button/popup
'
' The routine writes real SAP GUI IDs discovered from the live scripting tree
' into the RAW SAP MAP section of SAP_MAP. It does not lock CONFIG values.
' A user review is required before copying a candidate into the locked table.

Public Sub CF_MapSAP_MON()
    Dim ses As Object, root As Object
    Dim ws As Worksheet, errorText As String, tcode As String
    Dim NextRow As Long, i As Long, mappedRows As Long, guardMessage As String
    Dim scanStarted As Double, scanLimit As Long, depthLimit As Long, timeoutSeconds As Double

    On Error GoTo Failed
    If Not SAPRun_RequireMapMode("CF_MapSAP_MON", guardMessage) Then
        MsgBox guardMessage, vbExclamation, "SAP MAP MON"
        Exit Sub
    End If
    tcode = "/SCWM/MON"
    Set ses = CF_OpenTCodeSession(tcode, errorText)
    If ses Is Nothing Then Err.Raise vbObjectError + 4101, , "OPEN_MON_FAILED: " & errorText

    Set ws = ThisWorkbook.Worksheets("SYSTEM")
    scanLimit = SAPMap_MaxControlScan()
    depthLimit = SAPMap_MaxScanDepth()
    timeoutSeconds = SAPMap_ScanTimeoutSeconds()
    scanStarted = Timer
    NextRow = MAP_RawStartRow(ws)

    For i = 0 To 3
        SAPMap_RequireScanBudget scanStarted, mappedRows, scanLimit, timeoutSeconds
        Set root = MAP_FindById(ses, "wnd[" & CStr(i) & "]")
        If Not root Is Nothing Then
            MAP_DumpTree root, ws, NextRow, 0, "wnd[" & CStr(i) & "]", "", mappedRows, _
                         scanStarted, scanLimit, depthLimit, timeoutSeconds
        End If
    Next i

    ws.columns("A:P").AutoFit
    ws.Range("D6").Value2 = "MAPPED_NOT_LOCKED"
    ws.Range("B6").Value2 = Now
    If Not gRT_Silent Then
        MsgBox "Da map SAP MON vao RAW SAP MAP. Rows=" & CStr(mappedRows) & _
               vbCrLf & "Loc cot P = Match Key de xem candidate MON.", _
               vbInformation, "SAP MAP MON"
    End If
    Exit Sub

Failed:
    If Not gRT_Silent Then
        MsgBox "CF_MapSAP_MON failed: Err " & CStr(Err.Number) & ": " & Err.Description, _
               vbExclamation, "SAP MAP MON"
    End If
End Sub

Private Function MAP_RawStartRow(ByVal ws As Worksheet) As Long
    Dim r As Long, lastRow As Long
    For r = 1 To 500
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value2))) = "SESSION" And _
           UCase$(Trim$(CStr(ws.Cells(r, 2).Value2))) = "TCODE" Then
            lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
            If lastRow < r Then lastRow = r
            MAP_RawStartRow = lastRow + 1
            Exit Function
        End If
    Next r

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow < 1 Then lastRow = 1
    MAP_WriteRawHeaders ws, lastRow + 2
    MAP_RawStartRow = lastRow + 3
End Function

Private Sub MAP_WriteRawHeaders(ByVal ws As Worksheet, ByVal headerRow As Long)
    Dim headers As Variant, i As Long
    headers = Array("Session", "TCode", "Window", "Depth", "Technical ID", _
                    "Type", "Name", "Text / Title", "Tooltip", "Parent ID", _
                    "Column Key", "Column Header", "Visible", "Enabled", _
                    "Selected", "Match Key")
    For i = LBound(headers) To UBound(headers)
        ws.Cells(headerRow, i + 1).value = headers(i)
    Next i
    ws.rows(headerRow).Font.Bold = True
End Sub

Private Sub MAP_DumpTree(ByVal obj As Object, ByVal ws As Worksheet, _
                         ByRef NextRow As Long, ByVal depth As Long, _
                         ByVal windowName As String, ByVal parentId As String, _
                         ByRef mappedRows As Long, ByVal scanStarted As Double, _
                         ByVal scanLimit As Long, ByVal depthLimit As Long, _
                         ByVal timeoutSeconds As Double)
    Dim idText As String, typeText As String, matchKey As String
    Dim children As Object, child As Object, columns As Object
    Dim i As Long, n As Long

    If obj Is Nothing Then Exit Sub
    SAPMap_RequireScanBudget scanStarted, mappedRows, scanLimit, timeoutSeconds
    If depth > depthLimit Then Exit Sub

    idText = MAP_Property(obj, "Id")
    typeText = MAP_Property(obj, "Type")
    matchKey = MAP_ControlMatchKey(obj)
    If matchKey = "" Then
        If MAP_GridScore(obj) >= 2 Then matchKey = "MON_RESULT_GRID_ID"
    End If
    MAP_WriteControlRow ws, NextRow, windowName, depth, obj, parentId, matchKey
    NextRow = NextRow + 1
    mappedRows = mappedRows + 1

    Set columns = MAP_ColumnOrder(obj)
    If Not columns Is Nothing Then
        MAP_WriteGridRows ws, NextRow, obj, idText, depth, columns, mappedRows
    End If

    Set children = MAP_Children(obj)
    n = MAP_CollectionCount(children)
    For i = 0 To n - 1
        SAPMap_RequireScanBudget scanStarted, mappedRows, scanLimit, timeoutSeconds
        Set child = MAP_CollectionItem(children, i)
        If Not child Is Nothing Then
            MAP_DumpTree child, ws, NextRow, depth + 1, windowName, idText, mappedRows, _
                         scanStarted, scanLimit, depthLimit, timeoutSeconds
        End If
        Set child = Nothing
    Next i
End Sub

Private Sub MAP_WriteControlRow(ByVal ws As Worksheet, ByVal rowNo As Long, _
                                ByVal windowName As String, ByVal depth As Long, _
                                ByVal obj As Object, ByVal parentId As String, _
                                ByVal matchKey As String)
    ws.Cells(rowNo, 1).value = Now
    ws.Cells(rowNo, 2).value = "/SCWM/MON"
    ws.Cells(rowNo, 3).value = windowName
    ws.Cells(rowNo, 4).value = depth
    ws.Cells(rowNo, 5).value = MAP_Property(obj, "Id")
    ws.Cells(rowNo, 6).value = MAP_Property(obj, "Type")
    ws.Cells(rowNo, 7).value = MAP_Property(obj, "Name")
    ws.Cells(rowNo, 8).value = MAP_Property(obj, "Text")
    ws.Cells(rowNo, 9).value = MAP_Property(obj, "Tooltip")
    ws.Cells(rowNo, 10).value = parentId
    ws.Cells(rowNo, 13).value = MAP_Property(obj, "Visible")
    ws.Cells(rowNo, 14).value = MAP_Property(obj, "Enabled")
    ws.Cells(rowNo, 15).value = MAP_Property(obj, "Selected")
    ws.Cells(rowNo, 16).value = matchKey
End Sub

Private Sub MAP_WriteGridRows(ByVal ws As Worksheet, ByRef NextRow As Long, _
                              ByVal grid As Object, ByVal gridId As String, _
                              ByVal depth As Long, ByVal columns As Object, _
                              ByRef mappedRows As Long)
    Dim i As Long, n As Long, colKey As String, caption As String, matchKey As String
    n = MAP_CollectionCount(columns)
    For i = 0 To n - 1
        colKey = MAP_CollectionValue(columns, i)
        If colKey <> "" Then
            caption = MAP_GridCaption(grid, colKey)
            matchKey = MAP_GridMatchKey(colKey, caption, grid)
            ws.Cells(NextRow, 1).value = Now
            ws.Cells(NextRow, 2).value = "/SCWM/MON"
            ws.Cells(NextRow, 3).value = "GRID:" & gridId
            ws.Cells(NextRow, 4).value = depth + 1
            ws.Cells(NextRow, 5).value = colKey
            ws.Cells(NextRow, 6).value = "GRID_COLUMN"
            ws.Cells(NextRow, 8).value = caption
            ws.Cells(NextRow, 10).value = gridId
            ws.Cells(NextRow, 11).value = colKey
            ws.Cells(NextRow, 12).value = caption
            ws.Cells(NextRow, 16).value = matchKey
            NextRow = NextRow + 1
            mappedRows = mappedRows + 1
        End If
        If mappedRows >= SAPMap_MaxControlScan() Then Exit For
    Next i
End Sub

Private Function MAP_ControlMatchKey(ByVal obj As Object) As String
    Dim allText As String, typ As String
    typ = UCase$(MAP_Property(obj, "Type"))
    allText = UCase$(MAP_Property(obj, "Id") & "|" & MAP_Property(obj, "Name") & _
                     "|" & MAP_Property(obj, "Text") & "|" & MAP_Property(obj, "Tooltip"))

    If InStr(allText, "CANCEL") > 0 And _
       (InStr(typ, "BUTTON") > 0 Or InStr(typ, "MENU") > 0 Or InStr(typ, "CONTEXT") > 0) Then
        MAP_ControlMatchKey = "MON_CANCEL_ID"
    ElseIf InStr(allText, "CONFIRM") > 0 And InStr(typ, "BUTTON") > 0 Then
        MAP_ControlMatchKey = "MON_CANCEL_CONFIRM_ID"
    ElseIf (InStr(allText, "PRODUCT") > 0 Or InStr(allText, "MATERIAL") > 0 Or _
            InStr(allText, "MATNR") > 0) And MAP_IsInputType(typ) Then
        MAP_ControlMatchKey = "MON_MATERIAL_ID"
    ElseIf (InStr(allText, "STORAGE BIN") > 0 Or InStr(allText, "SOURCE BIN") > 0 Or _
            InStr(allText, "LGPLA") > 0 Or InStr(allText, "BIN") > 0) And MAP_IsInputType(typ) Then
        MAP_ControlMatchKey = "MON_BIN_ID"
    ElseIf (InStr(allText, "WAREHOUSE TASK STATUS") > 0 Or _
            InStr(allText, "WT STATUS") > 0 Or InStr(allText, "STATUS") > 0) And _
            (MAP_IsInputType(typ) Or InStr(typ, "COMBO") > 0) Then
        MAP_ControlMatchKey = "MON_STATUS_ID"
    End If
End Function

Private Function MAP_GridMatchKey(ByVal colKey As String, ByVal caption As String, _
                                  ByVal grid As Object) As String
    Dim s As String
    s = UCase$(colKey & "|" & caption)
    If InStr(s, "WHSE ORDER") > 0 Or InStr(s, "WAREHOUSE ORDER") > 0 Or _
       InStr(s, "WHO") > 0 Then
        MAP_GridMatchKey = "MON_WT_COL"
    ElseIf InStr(s, "PRODUCT") > 0 Or InStr(s, "MATERIAL") > 0 Or InStr(s, "MATNR") > 0 Then
        MAP_GridMatchKey = "MON_MATERIAL_COL"
    ElseIf InStr(s, "BATCH") > 0 Or InStr(s, "CHARG") > 0 Then
        MAP_GridMatchKey = "MON_BATCH_COL"
    ElseIf InStr(s, "SOURCE BIN") > 0 Or InStr(s, "STORAGE BIN") > 0 Or _
           InStr(s, "LGPLA") > 0 Then
        MAP_GridMatchKey = "MON_BIN_COL"
    ElseIf InStr(s, "STATUS") > 0 Or InStr(s, "STAT") > 0 Then
        MAP_GridMatchKey = "MON_STATUS_COL"
    ElseIf InStr(s, "DOCUMENT") > 0 Or InStr(s, "DELIVERY") > 0 Or _
           InStr(s, "OUTBOUND") > 0 Or InStr(s, "VBELN") > 0 Then
        MAP_GridMatchKey = "MON_OB_COL"
    ElseIf InStr(s, "SRCTGT") > 0 Or InStr(s, "SRC_TGT") > 0 Or _
           InStr(s, "DEST.QTY") > 0 Or InStr(s, "QUANTITY") > 0 Or _
           InStr(s, "QTY") > 0 Then
        MAP_GridMatchKey = "MON_QTY_COL"
    End If
End Function

Private Function MAP_GridScore(ByVal grid As Object) As Long
    Dim columns As Object, i As Long, n As Long, keyText As String, caption As String
    Set columns = MAP_ColumnOrder(grid)
    n = MAP_CollectionCount(columns)
    For i = 0 To n - 1
        keyText = MAP_CollectionValue(columns, i)
        caption = MAP_GridCaption(grid, keyText)
        If InStr(UCase$(keyText & "|" & caption), "PRODUCT") > 0 Or _
           InStr(UCase$(keyText & "|" & caption), "MATERIAL") > 0 Then MAP_GridScore = MAP_GridScore + 1
        If InStr(UCase$(keyText & "|" & caption), "SOURCE BIN") > 0 Or _
           InStr(UCase$(keyText & "|" & caption), "LGPLA") > 0 Then MAP_GridScore = MAP_GridScore + 1
        If InStr(UCase$(keyText & "|" & caption), "STATUS") > 0 Then MAP_GridScore = MAP_GridScore + 1
    Next i
End Function

Private Function MAP_IsInputType(ByVal typ As String) As Boolean
    MAP_IsInputType = (InStr(typ, "TEXTFIELD") > 0 Or InStr(typ, "CTEXTFIELD") > 0 Or _
                       InStr(typ, "COMBOBOX") > 0 Or InStr(typ, "COMBO") > 0)
End Function

Private Function MAP_FindById(ByVal ses As Object, ByVal idText As String) As Object
    If ses Is Nothing Or Trim$(idText) = "" Then Exit Function
    On Error Resume Next
    Set MAP_FindById = ses.FindById(idText, False)
    On Error GoTo 0
End Function

Private Function MAP_Property(ByVal obj As Object, ByVal propertyName As String) As String
    Dim v As Variant
    If obj Is Nothing Then Exit Function
    On Error Resume Next
    v = CallByName(obj, propertyName, VbGet)
    MAP_Property = CStr(v & "")
    Err.Clear
    On Error GoTo 0
End Function

Private Function MAP_Children(ByVal obj As Object) As Object
    If obj Is Nothing Then Exit Function
    On Error Resume Next
    Set MAP_Children = obj.children
    On Error GoTo 0
End Function

Private Function MAP_ColumnOrder(ByVal obj As Object) As Object
    If obj Is Nothing Then Exit Function
    On Error Resume Next
    Set MAP_ColumnOrder = obj.columnOrder
    On Error GoTo 0
End Function

Private Function MAP_CollectionCount(ByVal items As Object) As Long
    If items Is Nothing Then Exit Function
    On Error Resume Next
    MAP_CollectionCount = CLng(items.count)
    If Err.Number <> 0 Then
        Err.Clear
        MAP_CollectionCount = CLng(items.Length)
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function MAP_CollectionItem(ByVal items As Object, ByVal indexNo As Long) As Object
    If items Is Nothing Then Exit Function
    On Error Resume Next
    Set MAP_CollectionItem = items.ElementAt(indexNo)
    If MAP_CollectionItem Is Nothing Then Set MAP_CollectionItem = items.item(indexNo)
    If MAP_CollectionItem Is Nothing Then Set MAP_CollectionItem = items(indexNo)
    On Error GoTo 0
End Function

Private Function MAP_CollectionValue(ByVal items As Object, ByVal indexNo As Long) As String
    Dim v As Variant, item As Object
    If items Is Nothing Then Exit Function
    On Error Resume Next
    v = items.ElementAt(indexNo)
    If Err.Number <> 0 Then
        Err.Clear
        v = items.item(indexNo)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        Set item = MAP_CollectionItem(items, indexNo)
        If Not item Is Nothing Then v = MAP_Property(item, "Name")
    End If
    MAP_CollectionValue = Trim$(CStr(v & ""))
    Err.Clear
    On Error GoTo 0
End Function

Private Function MAP_GridCaption(ByVal grid As Object, ByVal colKey As String) As String
    Dim v As Variant, i As Long, s As String
    If grid Is Nothing Or colKey = "" Then Exit Function
    On Error Resume Next
    v = grid.GetDisplayedColumnTitle(colKey)
    s = CStr(v & "")
    If s = "" Then
        v = grid.GetColumnTitles(colKey)
        If IsArray(v) Then
            For i = LBound(v) To UBound(v)
                If Trim$(CStr(v(i) & "")) <> "" Then s = s & " " & CStr(v(i))
            Next i
        End If
    End If
    MAP_GridCaption = Trim$(s)
    Err.Clear
    On Error GoTo 0
End Function
Public Sub CF_MapSAP_POST()
    Dim ses As Object, root As Object
    Dim ws As Worksheet, errorText As String, tcode As String
    Dim NextRow As Long, i As Long, mappedRows As Long, guardMessage As String
    Dim scanStarted As Double, scanLimit As Long, depthLimit As Long, timeoutSeconds As Double

    On Error GoTo Failed
    If Not SAPRun_RequireMapMode("CF_MapSAP_POST", guardMessage) Then
        MsgBox guardMessage, vbExclamation, "SAP MAP POST"
        Exit Sub
    End If
    tcode = "/SCWM/POST"
    Set ses = CF_OpenTCodeSession(tcode, errorText)
    If ses Is Nothing Then Err.Raise vbObjectError + 4201, , "OPEN_POST_FAILED: " & errorText

    Set ws = ThisWorkbook.Worksheets("SYSTEM")
    scanLimit = SAPMap_MaxControlScan()
    depthLimit = SAPMap_MaxScanDepth()
    timeoutSeconds = SAPMap_ScanTimeoutSeconds()
    scanStarted = Timer
    NextRow = MAP_RawStartRow(ws)

    For i = 0 To 3
        SAPMap_RequireScanBudget scanStarted, mappedRows, scanLimit, timeoutSeconds
        Set root = MAP_FindById(ses, "wnd[" & CStr(i) & "]")
        If Not root Is Nothing Then
            MAP_DumpPOSTTree root, ws, NextRow, 0, "wnd[" & CStr(i) & "]", "", mappedRows, _
                             scanStarted, scanLimit, depthLimit, timeoutSeconds
        End If
    Next i

    ws.columns("A:P").AutoFit
    ws.Range("D6").Value2 = "MAPPED_NOT_LOCKED"
    ws.Range("B6").Value2 = Now
    If Not gRT_Silent Then
        MsgBox "Da map SAP /SCWM/POST vao RAW SAP MAP. Rows=" & CStr(mappedRows) & _
               vbCrLf & "Loc cot P = Match Key de xem candidate POST.", _
               vbInformation, "SAP MAP POST"
    End If
    Exit Sub

Failed:
    If Not gRT_Silent Then
        MsgBox "CF_MapSAP_POST failed: Err " & CStr(Err.Number) & ": " & Err.Description, _
               vbExclamation, "SAP MAP POST"
    End If
End Sub

Private Sub MAP_DumpPOSTTree(ByVal obj As Object, ByVal ws As Worksheet, _
                             ByRef NextRow As Long, ByVal depth As Long, _
                             ByVal windowName As String, ByVal parentId As String, _
                             ByRef mappedRows As Long, ByVal scanStarted As Double, _
                             ByVal scanLimit As Long, ByVal depthLimit As Long, _
                             ByVal timeoutSeconds As Double)
    Dim idText As String, matchKey As String
    Dim children As Object, child As Object, columns As Object
    Dim i As Long, n As Long

    If obj Is Nothing Then Exit Sub
    SAPMap_RequireScanBudget scanStarted, mappedRows, scanLimit, timeoutSeconds
    If depth > depthLimit Then Exit Sub

    idText = MAP_Property(obj, "Id")
    matchKey = MAP_POSTControlMatchKey(obj)
    If matchKey = "" Then
        If MAP_POSTGridScore(obj) >= 3 Then matchKey = "POST_GRID_ID"
    End If

    MAP_WritePOSTControlRow ws, NextRow, windowName, depth, obj, parentId, matchKey
    NextRow = NextRow + 1
    mappedRows = mappedRows + 1

    Set columns = MAP_ColumnOrder(obj)
    If Not columns Is Nothing Then
        MAP_WritePOSTGridRows ws, NextRow, obj, idText, depth, columns, mappedRows
    End If

    Set children = MAP_Children(obj)
    n = MAP_CollectionCount(children)
    For i = 0 To n - 1
        SAPMap_RequireScanBudget scanStarted, mappedRows, scanLimit, timeoutSeconds
        Set child = MAP_CollectionItem(children, i)
        If Not child Is Nothing Then
            MAP_DumpPOSTTree child, ws, NextRow, depth + 1, windowName, idText, mappedRows, _
                             scanStarted, scanLimit, depthLimit, timeoutSeconds
        End If
        Set child = Nothing
    Next i
End Sub

Private Sub MAP_WritePOSTControlRow(ByVal ws As Worksheet, ByVal rowNo As Long, _
                                    ByVal windowName As String, ByVal depth As Long, _
                                    ByVal obj As Object, ByVal parentId As String, _
                                    ByVal matchKey As String)
    ws.Cells(rowNo, 1).value = Now
    ws.Cells(rowNo, 2).value = "/SCWM/POST"
    ws.Cells(rowNo, 3).value = windowName
    ws.Cells(rowNo, 4).value = depth
    ws.Cells(rowNo, 5).value = MAP_Property(obj, "Id")
    ws.Cells(rowNo, 6).value = MAP_Property(obj, "Type")
    ws.Cells(rowNo, 7).value = MAP_Property(obj, "Name")
    ws.Cells(rowNo, 8).value = MAP_Property(obj, "Text")
    ws.Cells(rowNo, 9).value = MAP_Property(obj, "Tooltip")
    ws.Cells(rowNo, 10).value = parentId
    ws.Cells(rowNo, 13).value = MAP_Property(obj, "Visible")
    ws.Cells(rowNo, 14).value = MAP_Property(obj, "Enabled")
    ws.Cells(rowNo, 15).value = MAP_Property(obj, "Selected")
    ws.Cells(rowNo, 16).value = matchKey
End Sub

Private Sub MAP_WritePOSTGridRows(ByVal ws As Worksheet, ByRef NextRow As Long, _
                                  ByVal grid As Object, ByVal gridId As String, _
                                  ByVal depth As Long, ByVal columns As Object, _
                                  ByRef mappedRows As Long)
    Dim i As Long, n As Long, colKey As String, caption As String, matchKey As String

    n = MAP_CollectionCount(columns)
    For i = 0 To n - 1
        colKey = MAP_CollectionValue(columns, i)
        If colKey <> "" Then
            caption = MAP_GridCaption(grid, colKey)
            matchKey = MAP_POSTGridMatchKey(colKey, caption)
            ws.Cells(NextRow, 1).value = Now
            ws.Cells(NextRow, 2).value = "/SCWM/POST"
            ws.Cells(NextRow, 3).value = "GRID:" & gridId
            ws.Cells(NextRow, 4).value = depth + 1
            ws.Cells(NextRow, 5).value = colKey
            ws.Cells(NextRow, 6).value = "GRID_COLUMN"
            ws.Cells(NextRow, 8).value = caption
            ws.Cells(NextRow, 10).value = gridId
            ws.Cells(NextRow, 11).value = colKey
            ws.Cells(NextRow, 12).value = caption
            ws.Cells(NextRow, 16).value = matchKey
            NextRow = NextRow + 1
            mappedRows = mappedRows + 1
        End If
        If mappedRows >= SAPMap_MaxControlScan() Then Exit For
    Next i
End Sub

Private Function SAPMap_MaxControlScan() As Long
    SAPMap_MaxControlScan = ModuleHelper.ToolConfigLong("MAP_MAX_CONTROL_SCAN", 4000, 100, 25000)
End Function

Private Function SAPMap_MaxScanDepth() As Long
    SAPMap_MaxScanDepth = ModuleHelper.ToolConfigLong("MAP_MAX_SCAN_DEPTH", 12, 1, 30)
End Function

Private Function SAPMap_ScanTimeoutSeconds() As Double
    SAPMap_ScanTimeoutSeconds = ModuleHelper.ToolConfigDouble("MAP_SCAN_TIMEOUT_SEC", 20#, 3#, 120#)
End Function

Private Sub SAPMap_RequireScanBudget(ByVal started As Double, ByVal mappedRows As Long, _
                                      ByVal scanLimit As Long, ByVal timeoutSeconds As Double)
    If mappedRows >= scanLimit Then
        Err.Raise vbObjectError + 4260, , "STOP_MAP_SCAN_TIMEOUT: control scan reached MAP_MAX_CONTROL_SCAN=" & CStr(scanLimit)
    End If
    If SAPMap_Elapsed(started) >= timeoutSeconds Then
        Err.Raise vbObjectError + 4261, , "STOP_MAP_SCAN_TIMEOUT: control scan exceeded MAP_SCAN_TIMEOUT_SEC=" & CStr(timeoutSeconds)
    End If
    DoEvents
End Sub

Private Function SAPMap_Elapsed(ByVal started As Double) As Double
    If Timer >= started Then
        SAPMap_Elapsed = Timer - started
    Else
        SAPMap_Elapsed = 86400# - started + Timer
    End If
End Function

Private Function MAP_POSTControlMatchKey(ByVal obj As Object) As String
    Dim allText As String, typ As String
    typ = UCase$(MAP_Property(obj, "Type"))
    allText = UCase$(MAP_Property(obj, "Id") & "|" & MAP_Property(obj, "Name") & _
                     "|" & MAP_Property(obj, "Text") & "|" & MAP_Property(obj, "Tooltip"))

    If InStr(allText, "SAVE") > 0 And _
       (InStr(typ, "BUTTON") > 0 Or InStr(typ, "MENU") > 0 Or _
        InStr(typ, "TOOLBAR") > 0) Then
        MAP_POSTControlMatchKey = "POST_SAVE_ID"
    End If
End Function

Private Function MAP_POSTGridMatchKey(ByVal colKey As String, ByVal caption As String) As String
    Dim s As String
    s = UCase$(colKey & "|" & caption)
    s = Replace(s, "_", " ")
    s = Replace(s, ".", " ")

    If InStr(s, "DEST BATCH") > 0 Or InStr(s, "TO CHARG") > 0 Or _
       InStr(s, "DESTINATION BATCH") > 0 Then
        MAP_POSTGridMatchKey = "POST_DEST_BATCH_COL"
    ElseIf InStr(s, "QTYPOSTCHG") > 0 Or InStr(s, "QTY POST CHG") > 0 Or _
           InStr(s, "POST QUAN") > 0 Or InStr(s, "POST QTY") > 0 Then
        MAP_POSTGridMatchKey = "POST_QTY_CHG_COL"
    ElseIf InStr(s, "DESTINATION BIN") > 0 Or InStr(s, "DEST BIN") > 0 Or _
           InStr(s, "NLPLA") > 0 Or InStr(s, "TO BIN") > 0 Then
        MAP_POSTGridMatchKey = "POST_DEST_BIN_COL"
    ElseIf InStr(s, "TOTAL STK") > 0 Or InStr(s, "TOTAL STOCK") > 0 Or _
           InStr(s, "TOTAL QUANTITY") > 0 Or Left$(s, 5) = "QUAN|" Then
        MAP_POSTGridMatchKey = "POST_TOTAL_STOCK_COL"
    ElseIf InStr(s, "A QTY AUOM") > 0 Or InStr(s, "A QTY") > 0 Or _
           InStr(s, "AQUAN") > 0 Or InStr(s, "AVAILABLE QTY") > 0 Or _
           InStr(s, "AVAIL QTY") > 0 Then
        MAP_POSTGridMatchKey = "POST_AVAILABLE_QTY_COL"
    ElseIf InStr(s, "SOURCE STORAGE BIN") > 0 Or InStr(s, "SOURCE BIN") > 0 Or _
           InStr(s, "STORAGE BIN") > 0 Or InStr(s, "LGPLA") > 0 Then
        MAP_POSTGridMatchKey = "POST_SOURCE_BIN_COL"
    ElseIf InStr(s, "STOCK TYPE") > 0 Or InStr(s, "STOCKTYPE") > 0 Or _
           InStr(s, "TST") > 0 Or Left$(s, 4) = "CAT|" Or InStr(s, "|CAT|") > 0 Then
        MAP_POSTGridMatchKey = "POST_TST_COL"
    ElseIf InStr(s, "BATCH") > 0 Or InStr(s, "CHARG") > 0 Then
        MAP_POSTGridMatchKey = "POST_BATCH_COL"
    ElseIf InStr(s, "PRODUCT") > 0 Or InStr(s, "MATERIAL") > 0 Or _
           InStr(s, "MATNR") > 0 Then
        MAP_POSTGridMatchKey = "POST_PRODUCT_COL"
    End If
End Function

Private Function MAP_POSTGridScore(ByVal grid As Object) As Long
    Dim columns As Object, i As Long, n As Long
    Dim keyText As String, caption As String, matchKey As String

    Set columns = MAP_ColumnOrder(grid)
    n = MAP_CollectionCount(columns)
    For i = 0 To n - 1
        keyText = MAP_CollectionValue(columns, i)
        caption = MAP_GridCaption(grid, keyText)
        matchKey = MAP_POSTGridMatchKey(keyText, caption)
        If matchKey <> "" Then MAP_POSTGridScore = MAP_POSTGridScore + 1
    Next i
End Function



