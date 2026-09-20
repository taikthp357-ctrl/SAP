Attribute VB_Name = "modCodexSAPReadOnly"
Option Explicit

' Grid lazy-load fix: force SAP GUI to render all rows before reading.
' Issue #6/#2: weak machine + many rows causes unrendered cells to return blank.
Private Sub Codex_EnsureGridRendered(ByVal grid As Object, ByVal ses As Object)
    Dim rowCount As Long
    On Error Resume Next
    rowCount = CLng(grid.rowCount)
    On Error GoTo 0
    If rowCount > 20 Then
        grid.SetCurrentCell rowCount - 1, 0
        CWR_WaitIdle ses, 0.5
        grid.SetCurrentCell 0, 0
    End If
End Sub

Public Function CodexOpenOBItemAndDump(ByVal expectedOB As String) As String
    Dim gui As Object, app As Object, con As Object, mon As Object, upper As Object
    Dim i As Long, r As Long, n As Long, targetRow As Long, docText As String, obText As String
    Dim t0 As Double, ses As Object, tcode As String, resultText As String

    On Error GoTo Failed
    expectedOB = Trim$(expectedOB)
    If expectedOB = "" Then CodexOpenOBItemAndDump = "STOP_EXPECTED_OB_EMPTY": Exit Function

    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    Set mon = Nothing
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then Set mon = ses: Exit For
    Next i
    If mon Is Nothing Then CodexOpenOBItemAndDump = "STOP_MON_SESSION_NOT_FOUND": Exit Function

    Set upper = mon.FindById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell", False)
    If upper Is Nothing Then CodexOpenOBItemAndDump = "STOP_MON_UPPER_GRID_NOT_FOUND": Exit Function
    Codex_EnsureGridRendered upper, mon
    n = CLng(upper.rowCount)
    targetRow = -1
    For r = 0 To n - 1
        obText = Trim$(CStr(upper.GetCellValue(r, "DOCNO_H") & ""))
        If obText = "" Then obText = Trim$(CStr(upper.GetCellValue(r, "ERP_NUMBER_H") & ""))
        If StrComp(CodexNormalizeNumber(obText), CodexNormalizeNumber(expectedOB), vbTextCompare) = 0 Then
            targetRow = r
            docText = Trim$(CStr(upper.GetCellValue(r, "DOCNO_H") & ""))
            Exit For
        End If
    Next r
    If targetRow < 0 Then
        CodexOpenOBItemAndDump = "STOP_OB_NOT_IN_UPPER_GRID=" & expectedOB & ";ROWS=" & CStr(n)
        Exit Function
    End If

    upper.selectedRows = CStr(targetRow)
    upper.SetCurrentCell targetRow, "DOCNO_H"
    upper.PressToolbarButton "N000000011"

    t0 = Timer
    Do
        DoEvents
        For i = 0 To con.children.count - 1
            Set ses = con.children.ElementAt(i)
            tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
            If UCase$(tcode) = "S000" Then
                resultText = resultText & "S000_SESSION=" & CStr(i) & "|ID=" & CStr(ses.id) & "|DOC=" & docText & vbLf
                resultText = resultText & CodexDumpRelevant(ses, "wnd[0]", 0)
            End If
        Next i
        If resultText <> "" Then
            CodexOpenOBItemAndDump = "OPENED_OB=" & expectedOB & ";DOC=" & docText & vbLf & resultText
            Exit Function
        End If
        If CodexElapsed(t0) >= 20# Then Exit Do
    Loop
    CodexOpenOBItemAndDump = "STOP_S000_NOT_OPENED|OB=" & expectedOB & "|DOC=" & docText
    Exit Function

Failed:
    CodexOpenOBItemAndDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexDumpCurrentS000Tree() As String
    Dim gui As Object, app As Object, con As Object, ses As Object
    Dim i As Long, tcode As String, titleText As String, props As Variant
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": titleText = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction)
        titleText = CStr(ses.FindById("wnd[0]").text)
        On Error GoTo Failed
        If UCase$(tcode) = "S000" Then
            CodexDumpCurrentS000Tree = "S000=" & CStr(ses.id) & "|TITLE=" & CodexClean(titleText) & vbLf & _
                                       CStr(ses.GetObjectTree("wnd[0]", props))
            Exit Function
        End If
    Next i
    CodexDumpCurrentS000Tree = "S000_NOT_FOUND"
    Exit Function
Failed:
    CodexDumpCurrentS000Tree = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexDumpCurrentS000Grids() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, r As Long, c As Variant, n As Long, tcode As String, pathText As String
    Dim rowCount As Long, lineText As String, titleText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "S000" Then
            pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell"
            Set grid = ses.FindById(pathText, False)
            If grid Is Nothing Then
                CodexDumpCurrentS000Grids = "S000_GRID_NOT_FOUND|ID=" & CStr(ses.id)
                Exit Function
            End If
            Codex_EnsureGridRendered grid, ses
            rowCount = CLng(grid.rowCount)
            CodexDumpCurrentS000Grids = "S000_GRID=" & CStr(grid.id) & "|ROWS=" & CStr(rowCount) & vbLf
            n = 0
            For Each c In grid.columnOrder
                titleText = "": On Error Resume Next: titleText = CStr(grid.GetColumnTitles(CStr(c))): On Error GoTo Failed
                lineText = "COL=" & CStr(c) & "|TITLE=" & CodexClean(titleText)
                For r = 0 To rowCount - 1
                    lineText = lineText & "|R" & CStr(r) & "=" & CodexClean(CStr(grid.GetCellValue(r, CStr(c)) & ""))
                Next r
                CodexDumpCurrentS000Grids = CodexDumpCurrentS000Grids & lineText & vbLf
                n = n + 1
                If n >= 180 Then Exit For
            Next c
            Exit Function
        End If
    Next i
    CodexDumpCurrentS000Grids = "S000_NOT_FOUND"
    Exit Function
Failed:
    CodexDumpCurrentS000Grids = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexSelectS000ItemAndDump() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, tcode As String, pathText As String, props As Variant
    Dim materialText As String, qtyText As String, resultText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "S000" Then
            pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell"
            Set grid = ses.FindById(pathText, False)
            If grid Is Nothing Then GoTo NextSession
            materialText = CStr(grid.GetCellValue(0, "PRODUCTNO") & "")
            qtyText = CStr(grid.GetCellValue(0, "QTY_UI") & "")
            grid.selectedRows = "0"
            grid.SetCurrentCell 0, "ITEMNO"
            On Error Resume Next
            grid.ClickCurrentCell
            Err.Clear
            On Error GoTo Failed
            props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
            resultText = CStr(ses.GetObjectTree("wnd[0]", props))
            CodexSelectS000ItemAndDump = "S000=" & CStr(ses.id) & "|MATERIAL=" & materialText & "|QTY=" & qtyText & vbLf & resultText
            Exit Function
        End If
NextSession:
    Next i
    CodexSelectS000ItemAndDump = "S000_ITEM_GRID_NOT_FOUND"
    Exit Function
Failed:
    CodexSelectS000ItemAndDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexExpandS000DetailsAndDump() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, btn As Object
    Dim i As Long, tcode As String, props As Variant
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "S000" Then
            Set btn = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_ODP2:/SCWM/SAPLUI_DLV_CORE:4100/btnCMD_ODP2_SIZE", False)
            If btn Is Nothing Then
                CodexExpandS000DetailsAndDump = "S000_DETAIL_EXPAND_BUTTON_NOT_FOUND"
                Exit Function
            End If
            btn.Press
            CodexExpandS000DetailsAndDump = CStr(ses.GetObjectTree("wnd[0]", props))
            Exit Function
        End If
    Next i
    CodexExpandS000DetailsAndDump = "S000_NOT_FOUND"
    Exit Function
Failed:
    CodexExpandS000DetailsAndDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexSelectS000PickingBinAndDump() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, tab As Object
    Dim i As Long, tcode As String, props As Variant
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "S000" Then
            Set tab = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_ODP2:/SCWM/SAPLUI_DLV_PRD:4000/tabsTABSTRIP_ODP2/tabpOK_ODP2_TAB11", False)
            If tab Is Nothing Then
                CodexSelectS000PickingBinAndDump = "S000_PICKING_BIN_TAB_NOT_FOUND"
                Exit Function
            End If
            tab.Select
            If Not CodexWaitIdle(ses, 10#) Then
                CodexSelectS000PickingBinAndDump = "S000_TAB_BUSY_TIMEOUT"
                Exit Function
            End If
            CodexSelectS000PickingBinAndDump = CStr(ses.GetObjectTree(CStr(tab.Id), props))
            Exit Function
        End If
    Next i
    CodexSelectS000PickingBinAndDump = "S000_NOT_FOUND"
    Exit Function
Failed:
    CodexSelectS000PickingBinAndDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexOpenWarehouseTaskFollowOn() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, menuItem As Object
    Dim i As Long, j As Long, tcode As String, titleText As String, t0 As Double
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "S000" Then
            titleText = "": On Error Resume Next: titleText = CStr(ses.FindById("wnd[0]").text): On Error GoTo Failed
            If InStr(1, titleText, "Maintain Outb. Deliv. Order", vbTextCompare) > 0 Then
                Set menuItem = ses.FindById("wnd[0]/mbar/menu[0]/menu[2]/menu[0]", False)
                If menuItem Is Nothing Then
                    CodexOpenWarehouseTaskFollowOn = "FOLLOW_ON_MENU_NOT_FOUND"
                    Exit Function
                End If
                menuItem.Select
                t0 = Timer
                Do
                    DoEvents
                    For j = 0 To con.children.count - 1
                        Set ses = con.children.ElementAt(j)
                        tcode = "": titleText = ""
                        On Error Resume Next
                        tcode = CStr(ses.Info.Transaction)
                        titleText = CStr(ses.FindById("wnd[0]").text)
                        On Error GoTo Failed
                        If UCase$(tcode) = "S000" And InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then
                            CodexOpenWarehouseTaskFollowOn = "OPENED=" & CStr(ses.id) & "|TITLE=" & CodexClean(titleText)
                            Exit Function
                        End If
                    Next j
                    If CodexElapsed(t0) >= 20# Then Exit Do
                Loop
                CodexOpenWarehouseTaskFollowOn = "STOP_CREATE_WT_NOT_OPENED"
                Exit Function
            End If
        End If
    Next i
    CodexOpenWarehouseTaskFollowOn = "S000_ORDER_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexOpenWarehouseTaskFollowOn = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexDumpCreateWTGrids() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, r As Long, n As Long, tcode As String, pathText As String
    Dim labelText As String, lineText As String, titleText As String, rowCount As Long, c As Variant
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "S000" Then
            For Each labelText In Array("WR", "STOCK")
                If labelText = "WR" Then
                    pathText = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
                Else
                    pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
                End If
                Set grid = ses.FindById(pathText, False)
                If grid Is Nothing Then
                    CodexDumpCreateWTGrids = CodexDumpCreateWTGrids & labelText & "_GRID_NOT_FOUND|" & pathText & vbLf
                Else
                    Codex_EnsureGridRendered grid, ses
                    rowCount = CLng(grid.rowCount)
                    CodexDumpCreateWTGrids = CodexDumpCreateWTGrids & labelText & "_GRID=" & CStr(grid.id) & "|ROWS=" & CStr(rowCount) & vbLf
                    n = 0
                    For Each c In grid.columnOrder
                        titleText = "": On Error Resume Next: titleText = CStr(grid.GetColumnTitles(CStr(c))): On Error GoTo Failed
                        lineText = "COL=" & CStr(c) & "|TITLE=" & CodexClean(titleText)
                        For r = 0 To rowCount - 1
                            lineText = lineText & "|R" & CStr(r) & "=" & CodexClean(CStr(grid.GetCellValue(r, CStr(c)) & ""))
                        Next r
                        CodexDumpCreateWTGrids = CodexDumpCreateWTGrids & labelText & "|" & lineText & vbLf
                        n = n + 1
                        If n >= 180 Then Exit For
                    Next c
                End If
            Next labelText
            Exit Function
        End If
    Next i
    CodexDumpCreateWTGrids = "S000_NOT_FOUND"
    Exit Function
Failed:
    CodexDumpCreateWTGrids = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexExecuteCreateWTQuery() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, btn As Object
    Dim i As Long, tcode As String, titleText As String, t0 As Double, busy As Boolean
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": titleText = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction)
        titleText = CStr(ses.FindById("wnd[0]").text)
        On Error GoTo Failed
        If (UCase$(tcode) = "S000" Or UCase$(tcode) = "/SCWM/PRDO") And InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then
            Set btn = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_TODLV:1001/btnCMD_QUERY_EXECUTE", False)
            If btn Is Nothing Then
                CodexExecuteCreateWTQuery = "CREATE_WT_QUERY_BUTTON_NOT_FOUND"
                Exit Function
            End If
            btn.Press
            t0 = Timer
            Do
                DoEvents
                busy = True: On Error Resume Next: busy = CBool(ses.busy): On Error GoTo Failed
                If Not busy Then Exit Do
                If CodexElapsed(t0) >= 30# Then
                    CodexExecuteCreateWTQuery = "STOP_CREATE_WT_QUERY_BUSY_TIMEOUT"
                    Exit Function
                End If
            Loop
            CodexExecuteCreateWTQuery = "CREATE_WT_QUERY_EXECUTED|STATUS=" & CodexClean(CStr(ses.FindById("wnd[0]/sbar", False).text))
            Exit Function
        End If
    Next i
    CodexExecuteCreateWTQuery = "CREATE_WT_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexExecuteCreateWTQuery = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexOpenPRDOAndDump() As String
    Dim gui As Object, app As Object, con As Object, ses As Object
    Dim i As Long, tcode As String, titleText As String, props As Variant
    Dim t0 As Double, busy As Boolean
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": titleText = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction)
        titleText = CStr(ses.FindById("wnd[0]").text)
        On Error GoTo Failed
        If UCase$(tcode) = "S000" Then
            ses.StartTransaction "/SCWM/PRDO"
            t0 = Timer
            Do
                DoEvents
                busy = True
                On Error Resume Next
                busy = CBool(ses.busy)
                tcode = CStr(ses.Info.Transaction)
                titleText = CStr(ses.FindById("wnd[0]").text)
                On Error GoTo Failed
                If Not busy Then Exit Do
                If CodexElapsed(t0) >= 30# Then Exit Do
            Loop
            props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
            CodexOpenPRDOAndDump = "PRDO=" & CStr(ses.id) & "|TCODE=" & tcode & "|TITLE=" & CodexClean(titleText) & "|BUSY=" & CStr(busy) & vbLf & _
                                   CStr(ses.GetObjectTree("wnd[0]", props))
            Exit Function
        End If
    Next i
    CodexOpenPRDOAndDump = "S000_NOT_FOUND"
    Exit Function
Failed:
    CodexOpenPRDOAndDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexQueryPRDOAndDump(ByVal docNumber As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object
    Dim field As Object, btn As Object, grid As Object
    Dim i As Long, r As Long, k As Long, n As Long, tcode As String
    Dim t0 As Double, busy As Boolean, rowCount As Long, colId As Variant
    Dim lineText As String, titleText As String, pathText As String
    On Error GoTo Failed
    docNumber = Trim$(docNumber)
    If docNumber = "" Then CodexQueryPRDOAndDump = "STOP_PRDO_DOC_EMPTY": Exit Function
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then
            Set field = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_DLV_PRD:2001/txt/SCWM/S_SP_Q_HEAD-DOCNO_ODO", False)
            Set btn = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_DLV_PRD:2001/btnCMD_GO", False)
            If field Is Nothing Or btn Is Nothing Then
                CodexQueryPRDOAndDump = "STOP_PRDO_SEARCH_CONTROLS_NOT_FOUND"
                Exit Function
            End If
            field.text = docNumber
            btn.Press
            t0 = Timer
            Do
                DoEvents
                busy = True: On Error Resume Next: busy = CBool(ses.busy): On Error GoTo Failed
                If Not busy Then Exit Do
                If CodexElapsed(t0) >= 30# Then Exit Do
            Loop
            pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell"
            Set grid = ses.FindById(pathText, False)
            If grid Is Nothing Then
                CodexQueryPRDOAndDump = "PRDO_QUERY_DONE|DOC=" & docNumber & "|STATUS=" & CodexClean(CStr(ses.FindById("wnd[0]/sbar", False).text)) & "|GRID_NOT_FOUND"
                Exit Function
            End If
            Codex_EnsureGridRendered grid, ses
            rowCount = CLng(grid.rowCount)
            CodexQueryPRDOAndDump = "PRDO_QUERY_DONE|DOC=" & docNumber & "|STATUS=" & CodexClean(CStr(ses.FindById("wnd[0]/sbar", False).text)) & "|ROWS=" & CStr(rowCount) & vbLf
            For r = 0 To rowCount - 1
                lineText = "ROW=" & CStr(r)
                For Each colId In Array("ITEMNO", "PRODUCTNO", "BATCHNO_UI", "QTY_UI", "UOM", "CHANGEABLE_ICON", "STATUS_PICKING", "STATUS_GI")
                    lineText = lineText & "|" & CStr(colId) & "=" & CodexClean(CStr(grid.GetCellValue(r, CStr(colId)) & ""))
                Next colId
                CodexQueryPRDOAndDump = CodexQueryPRDOAndDump & lineText & vbLf
            Next r
            On Error Resume Next
            n = grid.ToolbarButtonCount
            On Error GoTo Failed
            CodexQueryPRDOAndDump = CodexQueryPRDOAndDump & "TOOLBAR_COUNT=" & CStr(n) & vbLf
            For k = 0 To n - 1
                Dim tbId As String, tbText As String, tbTip As String
                On Error Resume Next
                tbId = CStr(grid.GetToolbarButtonId(k) & "")
                tbText = CStr(grid.GetToolbarButtonText(k) & "")
                tbTip = CStr(grid.GetToolbarButtonTooltip(k) & "")
                On Error GoTo Failed
                CodexQueryPRDOAndDump = CodexQueryPRDOAndDump & "TB=" & CStr(k) & "|ID=" & CodexClean(tbId) & "|TEXT=" & CodexClean(tbText) & "|TIP=" & CodexClean(tbTip) & vbLf
            Next k
            Exit Function
        End If
    Next i
    CodexQueryPRDOAndDump = "PRDO_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexQueryPRDOAndDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexPRDOItemDetailsDump() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, tcode As String, t0 As Double, busy As Boolean, props As Variant
    Dim pathText As String, titleText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then
            pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell"
            Set grid = ses.FindById(pathText, False)
            If grid Is Nothing Then CodexPRDOItemDetailsDump = "PRDO_ITEM_GRID_NOT_FOUND": Exit Function
            If CLng(grid.rowCount) < 1 Then CodexPRDOItemDetailsDump = "PRDO_ITEM_GRID_EMPTY": Exit Function
            grid.selectedRows = "0"
            grid.SetCurrentCell 0, "ITEMNO"
            grid.PressToolbarButton "&DETAIL"
            t0 = Timer
            Do
                DoEvents
                busy = True: On Error Resume Next: busy = CBool(ses.busy): On Error GoTo Failed
                If Not busy Then Exit Do
                If CodexElapsed(t0) >= 20# Then Exit Do
            Loop
            titleText = "": On Error Resume Next: titleText = CStr(ses.FindById("wnd[0]").text): On Error GoTo Failed
            CodexPRDOItemDetailsDump = "PRDO_DETAILS|TITLE=" & CodexClean(titleText) & "|BUSY=" & CStr(busy) & vbLf & _
                                       CodexDumpRelevant(ses.FindById("wnd[0]"), "wnd[0]", 0)
            Exit Function
        End If
    Next i
    CodexPRDOItemDetailsDump = "PRDO_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexPRDOItemDetailsDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexOpenPRDOOutboundDelivery() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, menuItem As Object
    Dim i As Long, j As Long, tcode As String, titleText As String, t0 As Double, busy As Boolean
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": titleText = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction)
        titleText = CStr(ses.FindById("wnd[0]").text)
        On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then
            Set menuItem = ses.FindById("wnd[0]/mbar/menu[3]/menu[1]", False)
            If menuItem Is Nothing Then CodexOpenPRDOOutboundDelivery = "PRDO_OUTBOUND_DELIVERY_MENU_NOT_FOUND": Exit Function
            menuItem.Select
            t0 = Timer
            Do
                DoEvents
                busy = False
                On Error Resume Next
                For j = 0 To con.children.count - 1
                    Set ses = con.children.ElementAt(j)
                    If CStr(ses.Info.Transaction) <> "" Then busy = CBool(ses.busy)
                Next j
                On Error GoTo Failed
                If Not busy Then Exit Do
                If CodexElapsed(t0) >= 30# Then Exit Do
            Loop
            CodexOpenPRDOOutboundDelivery = "OUTBOUND_DELIVERY_NAVIGATED|SESSIONS=" & CStr(con.children.count) & vbLf
            For j = 0 To con.children.count - 1
                Set ses = con.children.ElementAt(j)
                On Error Resume Next
                tcode = CStr(ses.Info.Transaction)
                titleText = CStr(ses.FindById("wnd[0]").text)
                On Error GoTo Failed
                CodexOpenPRDOOutboundDelivery = CodexOpenPRDOOutboundDelivery & "SESSION=" & CStr(j) & "|ID=" & CStr(ses.id) & "|TCODE=" & CodexClean(tcode) & "|TITLE=" & CodexClean(titleText) & vbLf
            Next j
            Exit Function
        End If
    Next i
    CodexOpenPRDOOutboundDelivery = "PRDO_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexOpenPRDOOutboundDelivery = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexDumpCurrentPRDORelevant() As String
    Dim gui As Object, app As Object, con As Object, ses As Object
    Dim i As Long, tcode As String, titleText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": titleText = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction)
        titleText = CStr(ses.FindById("wnd[0]").text)
        On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then
            CodexDumpCurrentPRDORelevant = "PRDO_CURRENT|TITLE=" & CodexClean(titleText) & vbLf & _
                                           CodexDumpRelevant(ses.FindById("wnd[0]"), "wnd[0]", 0)
            Exit Function
        End If
    Next i
    CodexDumpCurrentPRDORelevant = "PRDO_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexDumpCurrentPRDORelevant = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexQueryOutboundDeliveryAndDump(ByVal docNumber As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, field As Object, btn As Object, grid As Object
    Dim i As Long, r As Long, n As Long, k As Long, tcode As String, t0 As Double, busy As Boolean
    Dim pathText As String, lineText As String, titleText As String, c As Variant
    On Error GoTo Failed
    docNumber = Trim$(docNumber)
    If docNumber = "" Then CodexQueryOutboundDeliveryAndDump = "STOP_OUTBOUND_DOC_EMPTY": Exit Function
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then
            Set field = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_DLV_FD:2001/txt/SCWM/S_SP_Q_HEAD-DOCNO_OD", False)
            Set btn = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_FD:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_DLV_FD:2001/btnCMD_GO", False)
            If field Is Nothing Or btn Is Nothing Then CodexQueryOutboundDeliveryAndDump = "STOP_OUTBOUND_SEARCH_CONTROLS_NOT_FOUND": Exit Function
            field.text = docNumber
            btn.Press
            t0 = Timer
            Do
                DoEvents
                busy = True: On Error Resume Next: busy = CBool(ses.busy): On Error GoTo Failed
                If Not busy Then Exit Do
                If CodexElapsed(t0) >= 30# Then Exit Do
            Loop
            pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_FD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell"
            Set grid = ses.FindById(pathText, False)
            If grid Is Nothing Then CodexQueryOutboundDeliveryAndDump = "OUTBOUND_QUERY_DONE|STATUS=" & CodexClean(CStr(ses.FindById("wnd[0]/sbar", False).text)) & "|GRID_NOT_FOUND": Exit Function
            Codex_EnsureGridRendered grid, ses
            n = CLng(grid.rowCount)
            CodexQueryOutboundDeliveryAndDump = "OUTBOUND_QUERY_DONE|DOC=" & docNumber & "|STATUS=" & CodexClean(CStr(ses.FindById("wnd[0]/sbar", False).text)) & "|ROWS=" & CStr(n) & vbLf
            For r = 0 To n - 1
                lineText = "ROW=" & CStr(r)
                For Each c In grid.columnOrder
                    lineText = lineText & "|" & CStr(c) & "=" & CodexClean(CStr(grid.GetCellValue(r, CStr(c)) & ""))
                Next c
                CodexQueryOutboundDeliveryAndDump = CodexQueryOutboundDeliveryAndDump & lineText & vbLf
                If r >= 10 Then Exit For
            Next r
            On Error Resume Next: k = grid.ToolbarButtonCount: On Error GoTo Failed
            CodexQueryOutboundDeliveryAndDump = CodexQueryOutboundDeliveryAndDump & "TOOLBAR_COUNT=" & CStr(k) & vbLf
            Exit Function
        End If
    Next i
    CodexQueryOutboundDeliveryAndDump = "OUTBOUND_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexQueryOutboundDeliveryAndDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexReturnPRDOOrder() As String
    Dim gui As Object, app As Object, con As Object, ses As Object
    Dim i As Long, tcode As String, titleText As String, t0 As Double, busy As Boolean
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then
            ses.FindById("wnd[0]/tbar[1]/btn[21]").Press
            t0 = Timer
            Do
                DoEvents
                busy = True: On Error Resume Next: busy = CBool(ses.busy): titleText = CStr(ses.FindById("wnd[0]").text): On Error GoTo Failed
                If Not busy Then Exit Do
                If CodexElapsed(t0) >= 20# Then Exit Do
            Loop
            CodexReturnPRDOOrder = "PRDO_ORDER_RETURNED|TITLE=" & CodexClean(titleText) & "|BUSY=" & CStr(busy)
            Exit Function
        End If
    Next i
    CodexReturnPRDOOrder = "PRDO_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexReturnPRDOOrder = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexSelectPRDOReferenceAndDump() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, tab As Object
    Dim i As Long, tcode As String, titleText As String, t0 As Double, busy As Boolean
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then
            Set tab = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB10", False)
            If tab Is Nothing Then CodexSelectPRDOReferenceAndDump = "PRDO_REFERENCE_TAB_NOT_FOUND": Exit Function
            tab.Select
            t0 = Timer
            Do
                DoEvents
                busy = True: On Error Resume Next: busy = CBool(ses.busy): On Error GoTo Failed
                If Not busy Then Exit Do
                If CodexElapsed(t0) >= 20# Then Exit Do
            Loop
            titleText = "": On Error Resume Next: titleText = CStr(ses.FindById("wnd[0]").text): On Error GoTo Failed
            CodexSelectPRDOReferenceAndDump = "PRDO_REFERENCE_SELECTED|TITLE=" & CodexClean(titleText) & "|BUSY=" & CStr(busy) & vbLf & _
                                              CodexDumpRelevant(tab, CStr(tab.Id), 0)
            Exit Function
        End If
    Next i
    CodexSelectPRDOReferenceAndDump = "PRDO_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexSelectPRDOReferenceAndDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexVerifyMONWT(ByVal expectedWT As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, r As Long, n As Long, tcode As String, wtText As String, lineText As String, c As Variant
    Dim pathText As String
    On Error GoTo Failed
    expectedWT = Trim$(expectedWT)
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then
            pathText = "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell"
            Set grid = ses.FindById(pathText, False)
            If grid Is Nothing Then CodexVerifyMONWT = "MON_WT_GRID_NOT_FOUND": Exit Function
            Codex_EnsureGridRendered grid, ses
            n = CLng(grid.rowCount)
            CodexVerifyMONWT = "MON_VERIFY|SESSION=" & CStr(ses.id) & "|ROWS=" & CStr(n) & vbLf
            For r = 0 To n - 1
                wtText = "": On Error Resume Next: wtText = CStr(grid.GetCellValue(r, "TANUM")): On Error GoTo Failed
                If expectedWT = "" Or StrComp(Trim$(wtText), expectedWT, vbTextCompare) = 0 Then
                    lineText = "ROW=" & CStr(r)
                    For Each c In Array("DOCNO", "WHO", "TANUM", "ITEMNO", "MATNR", "CHARG", "VLPLA", "VSOLM", "TOSTAT", "NLPLA", "PROCTY")
                        lineText = lineText & "|" & CStr(c) & "=" & CodexClean(CStr(grid.GetCellValue(r, CStr(c)) & ""))
                    Next c
                    CodexVerifyMONWT = CodexVerifyMONWT & lineText & vbLf
                End If
            Next r
            Exit Function
        End If
    Next i
    CodexVerifyMONWT = "MON_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexVerifyMONWT = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexDumpMONTanums() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, r As Long, n As Long, tcode As String, v As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then
            Set grid = ses.FindById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell", False)
            If grid Is Nothing Then CodexDumpMONTanums = "MON_WT_GRID_NOT_FOUND": Exit Function
            Codex_EnsureGridRendered grid, ses
            n = CLng(grid.rowCount)
            CodexDumpMONTanums = "MON_TANUMS|ROWS=" & CStr(n) & vbLf
            For r = 0 To n - 1
                On Error Resume Next
                v = CStr(grid.GetCellValue(r, "TANUM") & "")
                If Err.Number <> 0 Then v = "ERR:" & CStr(Err.Number): Err.Clear
                On Error GoTo Failed
                CodexDumpMONTanums = CodexDumpMONTanums & "R=" & CStr(r) & "|TANUM=" & CodexClean(v) & vbLf
            Next r
            Exit Function
        End If
    Next i
    CodexDumpMONTanums = "MON_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexDumpMONTanums = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexRefreshMON() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, tcode As String, t0 As Double, busy As Boolean
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then
            Set grid = ses.FindById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell", False)
            If grid Is Nothing Then CodexRefreshMON = "MON_WT_GRID_NOT_FOUND": Exit Function
            grid.PressToolbarButton "REFRESH"
            t0 = Timer
            Do
                DoEvents
                busy = True: On Error Resume Next: busy = CBool(ses.busy): On Error GoTo Failed
                If Not busy Then Exit Do
                If CodexElapsed(t0) >= 30# Then Exit Do
            Loop
            CodexRefreshMON = "MON_REFRESHED|BUSY=" & CStr(busy) & "|STATUS=" & CodexClean(CStr(ses.FindById("wnd[0]/sbar", False).text))
            Exit Function
        End If
    Next i
    CodexRefreshMON = "MON_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexRefreshMON = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexProbePRDOReferenceTree() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, tab As Object
    Dim i As Long, tcode As String, props As Variant, tabId As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then
            Set tab = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB10", False)
            If tab Is Nothing Then CodexProbePRDOReferenceTree = "PRDO_REFERENCE_TAB_NOT_FOUND": Exit Function
            tabId = CStr(tab.Id)
            CodexProbePRDOReferenceTree = "REF_TAB=" & CodexClean(tabId) & "|TEXT=" & CodexClean(CStr(tab.Text)) & vbLf & _
                                          CStr(ses.GetObjectTree(tabId, props))
            Exit Function
        End If
    Next i
    CodexProbePRDOReferenceTree = "PRDO_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexProbePRDOReferenceTree = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexDumpPOSTRows(ByVal expectedMaterial As String, ByVal expectedBatch As String, ByVal expectedBin As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, r As Long, n As Long, tcode As String, lineText As String, mat As String, bat As String, bin As String
    On Error GoTo Failed
    expectedMaterial = Trim$(expectedMaterial): expectedBatch = Trim$(expectedBatch): expectedBin = Trim$(expectedBin)
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/POST" Then
            Set grid = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/subSUB_OIP_1_DATA:/SCWM/SAPLUI_POST:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell", False)
            If grid Is Nothing Then CodexDumpPOSTRows = "POST_GRID_NOT_FOUND": Exit Function
            Codex_EnsureGridRendered grid, ses
            n = CLng(grid.rowCount)
            CodexDumpPOSTRows = "POST_VERIFY|ROWS=" & CStr(n) & vbLf
            For r = 0 To n - 1
                mat = "": bat = "": bin = ""
                On Error Resume Next
                mat = CStr(grid.GetCellValue(r, "MATNR") & "")
                bat = CStr(grid.GetCellValue(r, "CHARG") & "")
                bin = CStr(grid.GetCellValue(r, "VLPLA") & "")
                On Error GoTo Failed
                If expectedMaterial = "" Or (StrComp(Trim$(mat), expectedMaterial, vbTextCompare) = 0 And _
                   (expectedBatch = "" Or StrComp(Trim$(bat), expectedBatch, vbTextCompare) = 0) And _
                   (expectedBin = "" Or StrComp(Trim$(bin), expectedBin, vbTextCompare) = 0)) Then
                    lineText = "ROW=" & CStr(r) & "|MATNR=" & CodexClean(mat) & "|AVL_ALT_QTY=" & CodexClean(CStr(grid.GetCellValue(r, "AVL_ALT_QTY") & "")) & _
                               "|CHARG=" & CodexClean(bat) & "|VLPLA=" & CodexClean(bin) & "|TO_CAT=" & CodexClean(CStr(grid.GetCellValue(r, "TO_CAT") & "")) & _
                               "|LIME_QUAN=" & CodexClean(CStr(grid.GetCellValue(r, "LIME_QUAN") & "")) & "|NLPLA=" & CodexClean(CStr(grid.GetCellValue(r, "NLPLA") & "")) & _
                               "|POST_QUAN=" & CodexClean(CStr(grid.GetCellValue(r, "POST_QUAN") & "")) & "|TO_CHARG=" & CodexClean(CStr(grid.GetCellValue(r, "TO_CHARG") & ""))
                    CodexDumpPOSTRows = CodexDumpPOSTRows & lineText & vbLf
                End If
            Next r
            Exit Function
        End If
    Next i
    CodexDumpPOSTRows = "POST_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexDumpPOSTRows = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexOpenWTForegroundExact(ByVal expectedOB As String, ByVal expectedDoc As String, ByVal expectedWT As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, tree As Object, keys As Object
    Dim upper As Object, lower As Object, menuItem As Object, root As Object, ctl As Object
    Dim i As Long, j As Long, r As Long, n As Long, keyCount As Long, matchCount As Long
    Dim tcode As String, titleText As String, nodeText As String, keyText As String, docText As String, wtText As String
    Dim t0 As Double, busy As Boolean, props As Variant, dumpText As String
    Dim methodsErr As Long, methodsDesc As String, selectErr As Long, selectDesc As String, contextErr As Long, contextDesc As String
    On Error GoTo Failed
    expectedOB = Trim$(expectedOB): expectedDoc = Trim$(expectedDoc): expectedWT = Trim$(expectedWT)
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then Exit For
    Next i
    If ses Is Nothing Or UCase$(Replace(tcode, "/", "")) <> "SCWMMON" Then CodexOpenWTForegroundExact = "MON_SESSION_NOT_FOUND": Exit Function

    Set tree = ses.FindById("wnd[0]/usr/shell/shellcont[0]/shell", False)
    If tree Is Nothing Then CodexOpenWTForegroundExact = "MON_TREE_NOT_FOUND": Exit Function
    On Error Resume Next
    Set keys = CallByName(tree, "GetAllNodeKeys", VbMethod)
    If Err.Number <> 0 Or keys Is Nothing Then Err.Clear: On Error GoTo Failed: CodexOpenWTForegroundExact = "MON_TREE_KEYS_NOT_AVAILABLE": Exit Function
    keyCount = CodexCollectionCount(keys)
    On Error GoTo Failed
    For j = 0 To keyCount - 1
        keyText = CodexCollectionValue(keys, j)
        nodeText = ""
        On Error Resume Next: nodeText = CStr(CallByName(tree, "GetNodeTextByKey", VbMethod, keyText) & ""): On Error GoTo Failed
        If StrComp(Trim$(nodeText), "OUTBOUND DELIVERY ORDER", vbTextCompare) = 0 Then
            tree.DoubleClickNode keyText
            Exit For
        End If
    Next j
    If j >= keyCount Then CodexOpenWTForegroundExact = "MON_OUTBOUND_NODE_NOT_FOUND": Exit Function
    If Not CodexWaitSessionIdle(ses, 15#) Then CodexOpenWTForegroundExact = "MON_SELECTION_SCREEN_BUSY_TIMEOUT": Exit Function

    Set ctl = ses.FindById("wnd[1]/usr/txtS_ERP_H-LOW", False)
    If ctl Is Nothing Then Set ctl = ses.FindById("wnd[0]/usr/txtS_ERP_H-LOW", False)
    If ctl Is Nothing Then CodexOpenWTForegroundExact = "MON_OB_INPUT_NOT_FOUND": Exit Function
    ctl.text = expectedOB
    If StrComp(Trim$(CStr(ctl.text & "")), expectedOB, vbTextCompare) <> 0 Then CodexOpenWTForegroundExact = "MON_OB_INPUT_READBACK_MISMATCH": Exit Function
    Set ctl = ses.FindById("wnd[1]/tbar[0]/btn[8]", False)
    If ctl Is Nothing Then Set ctl = ses.FindById("wnd[0]/tbar[0]/btn[8]", False)
    If ctl Is Nothing Then CodexOpenWTForegroundExact = "MON_EXECUTE_NOT_FOUND": Exit Function
    ctl.Press
    If Not CodexWaitSessionIdle(ses, 30#) Then CodexOpenWTForegroundExact = "MON_EXECUTE_BUSY_TIMEOUT": Exit Function

    Set upper = ses.FindById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell", False)
    If upper Is Nothing Then CodexOpenWTForegroundExact = "MON_UPPER_GRID_NOT_FOUND": Exit Function
    Codex_EnsureGridRendered upper, ses
    n = CLng(upper.rowCount)
    matchCount = 0
    For r = 0 To n - 1
        docText = "": On Error Resume Next: docText = CStr(upper.GetCellValue(r, "DOCNO_H") & ""): On Error GoTo Failed
        If (expectedDoc = "" And docText <> "") Or StrComp(Trim$(docText), expectedDoc, vbTextCompare) = 0 Then matchCount = matchCount + 1: j = r
    Next r
    If matchCount <> 1 Then CodexOpenWTForegroundExact = "MON_DOC_MATCH_COUNT=" & CStr(matchCount) & "|ROWS=" & CStr(n): Exit Function
    upper.selectedRows = CStr(j)
    upper.SetCurrentCell j, "DOCNO_H"
    upper.PressToolbarButton "N000000096"
    If Not CodexWaitSessionIdle(ses, 30#) Then CodexOpenWTForegroundExact = "MON_WT_NAV_BUSY_TIMEOUT": Exit Function

    Set lower = ses.FindById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell", False)
    If lower Is Nothing Then CodexOpenWTForegroundExact = "MON_WT_GRID_NOT_FOUND": Exit Function
    Codex_EnsureGridRendered lower, ses
    n = CLng(lower.rowCount): matchCount = 0
    For r = 0 To n - 1
        wtText = "": On Error Resume Next: wtText = CStr(lower.GetCellValue(r, "TANUM") & ""): On Error GoTo Failed
        If StrComp(Trim$(wtText), expectedWT, vbTextCompare) = 0 Then matchCount = matchCount + 1: j = r
    Next r
    If matchCount <> 1 Then CodexOpenWTForegroundExact = "MON_WT_MATCH_COUNT=" & CStr(matchCount) & "|ROWS=" & CStr(n): Exit Function
    lower.selectedRows = CStr(j)
    lower.SetCurrentCell j, "TANUM"
    On Error Resume Next
    Err.Clear: lower.PressToolbarContextButton "METHODS": methodsErr = Err.Number: methodsDesc = Err.Description: Err.Clear
    lower.SelectContextMenuItemByText "Confirm WT in Foreground": selectErr = Err.Number: selectDesc = Err.Description: Err.Clear
    If selectErr <> 0 Then
        lower.SelectContextMenuItem "Confirm WT in Foreground": contextErr = Err.Number: contextDesc = Err.Description: Err.Clear
    End If
    On Error GoTo Failed
    If Not CodexWaitSessionIdle(ses, 30#) Then CodexOpenWTForegroundExact = "WT_FOREGROUND_BUSY_TIMEOUT": Exit Function

    titleText = "": On Error Resume Next: titleText = CStr(ses.FindById("wnd[0]").text): On Error GoTo Failed
    CodexOpenWTForegroundExact = "WT_FOREGROUND_PROBE|TITLE=" & CodexClean(titleText) & "|STATUS=" & CodexClean(CStr(ses.FindById("wnd[0]/sbar", False).text)) & _
                                 "|METHODS_ERR=" & CStr(methodsErr) & "|METHODS_DESC=" & CodexClean(methodsDesc) & _
                                 "|SELECT_ERR=" & CStr(selectErr) & "|SELECT_DESC=" & CodexClean(selectDesc) & _
                                 "|CONTEXT_ERR=" & CStr(contextErr) & "|CONTEXT_DESC=" & CodexClean(contextDesc) & vbLf
    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
    Set root = ses.FindById("wnd[0]", False)
    If Not root Is Nothing Then
        dumpText = CStr(ses.GetObjectTree("wnd[0]", props))
        CodexOpenWTForegroundExact = CodexOpenWTForegroundExact & dumpText
    End If
    Exit Function
Failed:
    CodexOpenWTForegroundExact = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexDumpConfirmWTGrids() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, r As Long, c As Variant, n As Long, rowCount As Long
    Dim tcode As String, titleText As String, lineText As String, valueText As String
    Dim resultPath As String, detailPath As String, changeableText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        titleText = "": tcode = ""
        On Error Resume Next
        titleText = CStr(ses.FindById("wnd[0]", False).text & "")
        tcode = CStr(ses.Info.Transaction & "")
        On Error GoTo Failed
        If InStr(1, titleText, "Confirm Warehouse Task", vbTextCompare) > 0 Or UCase$(Replace(tcode, "/", "")) = "SCWMTOCONF" Then Exit For
    Next i
    If ses Is Nothing Then CodexDumpConfirmWTGrids = "CONFIRM_SESSION_NOT_FOUND": Exit Function

    resultPath = "wnd[0]/usr/subSUB_OIP:/SCWM/SAPLUI_TO_CONF:0120/subSUB_SEARCH_RESULT:/SCWM/SAPLUI_TO_CONF:0130/cntlCC_OIP/shellcont/shell"
    detailPath = "wnd[0]/usr/subSUB_ODP_1:/SCWM/SAPLUI_TO_CONF:0150/tabsTS_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_1_TAB:/SCWM/SAPLUI_TO_CONF:0155/subSUB_ODP1_1_CONT:/SCWM/SAPLUI_TO_CONF:0160/cntlCC_ODP1_MAT/shellcont/shell"
    CodexDumpConfirmWTGrids = "CONFIRM_SCREEN|TITLE=" & CodexClean(titleText) & "|TCODE=" & CodexClean(tcode) & vbLf

    Set grid = ses.FindById(resultPath, False)
    If grid Is Nothing Then
        CodexDumpConfirmWTGrids = CodexDumpConfirmWTGrids & "RESULT_GRID_NOT_FOUND|ID=" & resultPath & vbLf
    Else
        Codex_EnsureGridRendered grid, ses
        rowCount = CLng(grid.rowCount)
        CodexDumpConfirmWTGrids = CodexDumpConfirmWTGrids & "RESULT_GRID=" & CStr(grid.id) & "|ROWS=" & CStr(rowCount) & vbLf
        n = 0
        For Each c In grid.columnOrder
            titleText = "": On Error Resume Next: titleText = CStr(grid.GetColumnTitles(CStr(c)) & ""): On Error GoTo Failed
            lineText = "RESULT_COL=" & CStr(c) & "|TITLE=" & CodexClean(titleText)
            For r = 0 To rowCount - 1
                valueText = "": On Error Resume Next: valueText = CStr(grid.GetCellValue(r, CStr(c)) & ""): On Error GoTo Failed
                lineText = lineText & "|R" & CStr(r) & "=" & CodexClean(valueText)
            Next r
            CodexDumpConfirmWTGrids = CodexDumpConfirmWTGrids & lineText & vbLf
            n = n + 1
            If n >= 160 Then Exit For
        Next c
    End If

    Set grid = ses.FindById(detailPath, False)
    If grid Is Nothing Then
        CodexDumpConfirmWTGrids = CodexDumpConfirmWTGrids & "DETAIL_GRID_NOT_FOUND|ID=" & detailPath & vbLf
    Else
        Codex_EnsureGridRendered grid, ses
        rowCount = CLng(grid.rowCount)
        changeableText = ""
        On Error Resume Next: changeableText = CStr(CallByName(grid, "Changeable", VbGet) & ""): On Error GoTo Failed
        CodexDumpConfirmWTGrids = CodexDumpConfirmWTGrids & "DETAIL_GRID=" & CStr(grid.id) & "|ROWS=" & CStr(rowCount) & "|CHANGEABLE=" & CodexClean(changeableText) & vbLf
        n = 0
        For Each c In grid.columnOrder
            titleText = "": On Error Resume Next: titleText = CStr(grid.GetColumnTitles(CStr(c)) & ""): On Error GoTo Failed
            lineText = "DETAIL_COL=" & CStr(c) & "|TITLE=" & CodexClean(titleText)
            For r = 0 To rowCount - 1
                valueText = "": On Error Resume Next: valueText = CStr(grid.GetCellValue(r, CStr(c)) & ""): On Error GoTo Failed
                lineText = lineText & "|R" & CStr(r) & "=" & CodexClean(valueText)
            Next r
            CodexDumpConfirmWTGrids = CodexDumpConfirmWTGrids & lineText & vbLf
            n = n + 1
            If n >= 200 Then Exit For
        Next c
    End If
    Exit Function
Failed:
    CodexDumpConfirmWTGrids = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexProbeConfirmQtyCells() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, r As Long, j As Long, tcode As String, titleText As String
    Dim colId As Variant, cellValue As String, probeText As String, detailPath As String
    Dim candidates As Variant, methodNames As Variant, methodName As Variant, resultValue As Variant
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        titleText = "": tcode = ""
        On Error Resume Next
        titleText = CStr(ses.FindById("wnd[0]", False).text & "")
        tcode = CStr(ses.Info.Transaction & "")
        On Error GoTo Failed
        If InStr(1, titleText, "Confirm Warehouse Task", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexProbeConfirmQtyCells = "CONFIRM_SESSION_NOT_FOUND": Exit Function
    detailPath = "wnd[0]/usr/subSUB_ODP_1:/SCWM/SAPLUI_TO_CONF:0150/tabsTS_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_1_TAB:/SCWM/SAPLUI_TO_CONF:0155/subSUB_ODP1_1_CONT:/SCWM/SAPLUI_TO_CONF:0160/cntlCC_ODP1_MAT/shellcont/shell"
    Set grid = ses.FindById(detailPath, False)
    If grid Is Nothing Then CodexProbeConfirmQtyCells = "DETAIL_GRID_NOT_FOUND": Exit Function
    candidates = Array("VSOLA", "VSOLM", "NISTA", "NISTM", "NDIFA", "NDIFM", "VRESA", "VRESA_CWQUAN")
    methodNames = Array("GetCellChangeable", "IsCellChangeable", "GetCellEditable", "IsCellEditable")
    CodexProbeConfirmQtyCells = "QTY_CELL_PROBE|GRID=" & CStr(grid.id) & "|ROWS=" & CStr(grid.rowCount) & vbLf
    For j = LBound(candidates) To UBound(candidates)
        colId = CStr(candidates(j))
        cellValue = "": On Error Resume Next: cellValue = CStr(grid.GetCellValue(0, CStr(colId)) & ""): On Error GoTo Failed
        probeText = "COL=" & CStr(colId) & "|VALUE=" & CodexClean(cellValue)
        For Each methodName In methodNames
            resultValue = Empty
            On Error Resume Next
            Err.Clear
            resultValue = CallByName(grid, CStr(methodName), VbMethod, 0, CStr(colId))
            If Err.Number = 0 Then
                probeText = probeText & "|" & CStr(methodName) & "=" & CodexClean(CStr(resultValue & ""))
            Else
                Err.Clear
            End If
            On Error GoTo Failed
        Next methodName
        CodexProbeConfirmQtyCells = CodexProbeConfirmQtyCells & probeText & vbLf
    Next j
    Exit Function
Failed:
    CodexProbeConfirmQtyCells = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexOpenWTExceptionExact(ByVal expectedOB As String, ByVal expectedDoc As String, ByVal expectedWT As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, tree As Object, ctl As Object, upper As Object, lower As Object, root As Object
    Dim keys As Variant, props As Variant, i As Long, j As Long, r As Long, n As Long, keyCount As Long, matchCount As Long
    Dim tcode As String, nodeText As String, keyText As String, docText As String, wtText As String, titleText As String, dumpText As String
    Dim methodsErr As Long, selectErr As Long, selectDesc As String, statusText As String
    On Error GoTo Failed
    expectedOB = Trim$(expectedOB): expectedDoc = Trim$(expectedDoc): expectedWT = Trim$(expectedWT)
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then Exit For
    Next i
    If ses Is Nothing Or UCase$(Replace(tcode, "/", "")) <> "SCWMMON" Then CodexOpenWTExceptionExact = "MON_SESSION_NOT_FOUND": Exit Function
    Set tree = ses.FindById("wnd[0]/usr/shell/shellcont[0]/shell", False)
    If tree Is Nothing Then CodexOpenWTExceptionExact = "MON_TREE_NOT_FOUND": Exit Function
    On Error Resume Next
    Set keys = CallByName(tree, "GetAllNodeKeys", VbMethod)
    If Err.Number <> 0 Or keys Is Nothing Then Err.Clear: On Error GoTo Failed: CodexOpenWTExceptionExact = "MON_TREE_KEYS_NOT_AVAILABLE": Exit Function
    keyCount = CodexCollectionCount(keys)
    On Error GoTo Failed
    For j = 0 To keyCount - 1
        keyText = CodexCollectionValue(keys, j)
        nodeText = "": On Error Resume Next: nodeText = CStr(CallByName(tree, "GetNodeTextByKey", VbMethod, keyText) & ""): On Error GoTo Failed
        If StrComp(Trim$(nodeText), "OUTBOUND DELIVERY ORDER", vbTextCompare) = 0 Then tree.DoubleClickNode keyText: Exit For
    Next j
    If j >= keyCount Then CodexOpenWTExceptionExact = "MON_OUTBOUND_NODE_NOT_FOUND": Exit Function
    If Not CodexWaitSessionIdle(ses, 15#) Then CodexOpenWTExceptionExact = "MON_SELECTION_SCREEN_BUSY_TIMEOUT": Exit Function
    Set ctl = ses.FindById("wnd[1]/usr/txtS_ERP_H-LOW", False)
    If ctl Is Nothing Then Set ctl = ses.FindById("wnd[0]/usr/txtS_ERP_H-LOW", False)
    If ctl Is Nothing Then CodexOpenWTExceptionExact = "MON_OB_INPUT_NOT_FOUND": Exit Function
    ctl.text = expectedOB
    Set ctl = ses.FindById("wnd[1]/tbar[0]/btn[8]", False)
    If ctl Is Nothing Then Set ctl = ses.FindById("wnd[0]/tbar[0]/btn[8]", False)
    If ctl Is Nothing Then CodexOpenWTExceptionExact = "MON_EXECUTE_NOT_FOUND": Exit Function
    ctl.Press
    If Not CodexWaitSessionIdle(ses, 30#) Then CodexOpenWTExceptionExact = "MON_EXECUTE_BUSY_TIMEOUT": Exit Function
    Set upper = ses.FindById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell", False)
    If upper Is Nothing Then CodexOpenWTExceptionExact = "MON_UPPER_GRID_NOT_FOUND": Exit Function
    Codex_EnsureGridRendered upper, ses
    n = CLng(upper.rowCount): matchCount = 0
    For r = 0 To n - 1
        docText = "": On Error Resume Next: docText = CStr(upper.GetCellValue(r, "DOCNO_H") & ""): On Error GoTo Failed
        If StrComp(Trim$(docText), expectedDoc, vbTextCompare) = 0 Then matchCount = matchCount + 1: j = r
    Next r
    If matchCount <> 1 Then CodexOpenWTExceptionExact = "MON_DOC_MATCH_COUNT=" & CStr(matchCount) & "|ROWS=" & CStr(n): Exit Function
    upper.selectedRows = CStr(j): upper.SetCurrentCell j, "DOCNO_H": upper.PressToolbarButton "N000000096"
    If Not CodexWaitSessionIdle(ses, 30#) Then CodexOpenWTExceptionExact = "MON_WT_NAV_BUSY_TIMEOUT": Exit Function
    Set lower = ses.FindById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell", False)
    If lower Is Nothing Then CodexOpenWTExceptionExact = "MON_WT_GRID_NOT_FOUND": Exit Function
    Codex_EnsureGridRendered lower, ses
    n = CLng(lower.rowCount): matchCount = 0
    For r = 0 To n - 1
        wtText = "": On Error Resume Next: wtText = CStr(lower.GetCellValue(r, "TANUM") & ""): On Error GoTo Failed
        If StrComp(Trim$(wtText), expectedWT, vbTextCompare) = 0 Then matchCount = matchCount + 1: j = r
    Next r
    If matchCount <> 1 Then CodexOpenWTExceptionExact = "MON_WT_MATCH_COUNT=" & CStr(matchCount) & "|ROWS=" & CStr(n): Exit Function
    lower.selectedRows = CStr(j): lower.SetCurrentCell j, "TANUM"
    On Error Resume Next
    Err.Clear: lower.PressToolbarContextButton "METHODS": methodsErr = Err.Number: Err.Clear
    lower.SelectContextMenuItemByText "Confirm WT with Exception": selectErr = Err.Number: selectDesc = Err.Description: Err.Clear
    On Error GoTo Failed
    If Not CodexWaitSessionIdle(ses, 30#) Then CodexOpenWTExceptionExact = "WT_EXCEPTION_BUSY_TIMEOUT": Exit Function
    titleText = "": statusText = ""
    On Error Resume Next
    titleText = CStr(ses.FindById("wnd[0]", False).text & "")
    statusText = CStr(ses.FindById("wnd[0]/sbar", False).text & "")
    On Error GoTo Failed
    CodexOpenWTExceptionExact = "WT_EXCEPTION_PROBE|TITLE=" & CodexClean(titleText) & "|STATUS=" & CodexClean(statusText) & "|METHODS_ERR=" & CStr(methodsErr) & "|SELECT_ERR=" & CStr(selectErr) & "|SELECT_DESC=" & CodexClean(selectDesc) & vbLf
    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
    Set root = ses.FindById("wnd[0]", False)
    If Not root Is Nothing Then dumpText = CStr(ses.GetObjectTree("wnd[0]", props)): CodexOpenWTExceptionExact = CodexOpenWTExceptionExact & dumpText
    Exit Function
Failed:
    CodexOpenWTExceptionExact = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexReturnMONFromConfirm() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, btn As Object
    Dim i As Long, titleText As String, tcode As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        titleText = "": tcode = ""
        On Error Resume Next
        titleText = CStr(ses.FindById("wnd[0]", False).text & "")
        tcode = CStr(ses.Info.Transaction & "")
        On Error GoTo Failed
        If InStr(1, titleText, "Confirm Warehouse Task", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexReturnMONFromConfirm = "CONFIRM_SESSION_NOT_FOUND": Exit Function
    Set btn = ses.FindById("wnd[0]/tbar[0]/btn[3]", False)
    If btn Is Nothing Then CodexReturnMONFromConfirm = "BACK_BUTTON_NOT_FOUND": Exit Function
    btn.Press
    If Not CodexWaitSessionIdle(ses, 30#) Then CodexReturnMONFromConfirm = "MON_RETURN_BUSY_TIMEOUT": Exit Function
    titleText = "": On Error Resume Next: titleText = CStr(ses.FindById("wnd[0]", False).text & ""): On Error GoTo Failed
    CodexReturnMONFromConfirm = "MON_RETURNED|TITLE=" & CodexClean(titleText) & "|TCODE=" & CodexClean(tcode)
    Exit Function
Failed:
    CodexReturnMONFromConfirm = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexReadWTMethodsMenu() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, tcode As String, titleText As String, itemId As String, itemText As String, tipText As String, lineText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then Exit For
    Next i
    If ses Is Nothing Then CodexReadWTMethodsMenu = "MON_SESSION_NOT_FOUND": Exit Function
    Set grid = ses.FindById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell", False)
    If grid Is Nothing Then CodexReadWTMethodsMenu = "MON_WT_GRID_NOT_FOUND": Exit Function
    grid.selectedRows = "0": grid.SetCurrentCell 0, "TANUM": grid.PressToolbarButton "METHODS"
    CodexReadWTMethodsMenu = "METHODS_MENU|GRID=" & CStr(grid.id) & vbLf
    For i = 0 To 100
        itemId = "": itemText = "": tipText = ""
        On Error Resume Next
        itemId = CStr(grid.GetToolbarMenuItemId(i) & "")
        If Err.Number <> 0 Then Err.Clear: Exit For
        itemText = CStr(grid.GetToolbarMenuItemText(i) & "")
        tipText = CStr(grid.GetToolbarMenuItemTooltip(i) & "")
        On Error GoTo Failed
        lineText = "ITEM=" & CStr(i) & "|ID=" & CodexClean(itemId) & "|TEXT=" & CodexClean(itemText) & "|TIP=" & CodexClean(tipText)
        CodexReadWTMethodsMenu = CodexReadWTMethodsMenu & lineText & vbLf
    Next i
    On Error Resume Next: grid.SelectToolbarMenuItem "__CLOSE_MENU__": On Error GoTo Failed
    Exit Function
Failed:
    CodexReadWTMethodsMenu = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexDumpSAPSessionWindows() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, win As Object
    Dim i As Long, j As Long, tcode As String, winText As String, winType As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then
            CodexDumpSAPSessionWindows = "SESSION=" & CStr(ses.id) & "|TCODE=" & CodexClean(tcode) & "|COUNT=" & CStr(ses.children.count) & vbLf
            For j = 0 To ses.children.count - 1
                Set win = ses.children.ElementAt(j)
                winText = "": winType = ""
                On Error Resume Next
                winText = CStr(win.text & "")
                winType = CStr(win.Type & "")
                On Error GoTo Failed
                CodexDumpSAPSessionWindows = CodexDumpSAPSessionWindows & "WIN=" & CStr(j) & "|ID=" & CStr(win.id) & "|TYPE=" & CodexClean(winType) & "|TEXT=" & CodexClean(winText) & vbLf
            Next j
            Exit Function
        End If
    Next i
    CodexDumpSAPSessionWindows = "MON_SESSION_NOT_FOUND"
    Exit Function
Failed:
    CodexDumpSAPSessionWindows = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexDumpWTExceptionWindow() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, win As Object
    Dim i As Long, j As Long, tcode As String, winText As String, props As Variant
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then
            For j = 0 To ses.children.count - 1
                Set win = ses.children.ElementAt(j)
                winText = "": On Error Resume Next: winText = CStr(win.text & ""): On Error GoTo Failed
                If InStr(1, winText, "Confirm Warehouse Task with Exception", vbTextCompare) > 0 Then
                    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
                    CodexDumpWTExceptionWindow = "EXCEPTION_WINDOW=" & CStr(win.id) & "|TITLE=" & CodexClean(winText) & vbLf & CStr(ses.GetObjectTree(CStr(win.id), props))
                    Exit Function
                End If
            Next j
        End If
    Next i
    CodexDumpWTExceptionWindow = "EXCEPTION_WINDOW_NOT_FOUND"
    Exit Function
Failed:
    CodexDumpWTExceptionWindow = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexReadWTExceptionInputs() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, win As Object, ctl As Object
    Dim i As Long, j As Long, tcode As String, winText As String, idText As String, valueText As String, selectedText As String
    Dim ids As Variant, idItem As Variant
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then
            For j = 0 To ses.children.count - 1
                Set win = ses.children.ElementAt(j)
                winText = "": On Error Resume Next: winText = CStr(win.text & ""): On Error GoTo Failed
                If InStr(1, winText, "Confirm Warehouse Task with Exception", vbTextCompare) > 0 Then
                    ids = Array("txt/SCWM/S_CONF_WT_FG_MON-TANUM", "txt/SCWM/S_CONF_WT_FG_MON-WHO", "ctxt/SCWM/S_CONF_WT_FG_MON-MATNR", "txt/SCWM/S_CONF_WT_FG_MON-CHARG", "txt/SCWM/S_CONF_WT_FG_MON-NISTA", "ctxt/SCWM/S_CONF_WT_FG_MON-ALTME", "txt/SCWM/S_CONF_WT_FG_MON-VLPLA", "ctxt/SCWM/S_CONF_WT_FG_MON-NLPLA")
                    CodexReadWTExceptionInputs = "EXCEPTION_INPUTS|WINDOW=" & CStr(win.id) & vbLf
                    For Each idItem In ids
                        Set ctl = ses.FindById(CStr(win.id) & "/usr/" & CStr(idItem), False)
                        If ctl Is Nothing Then
                            CodexReadWTExceptionInputs = CodexReadWTExceptionInputs & "ID=" & CStr(idItem) & "|NOT_FOUND" & vbLf
                        Else
                            valueText = "": selectedText = ""
                            On Error Resume Next
                            valueText = CStr(ctl.text & "")
                            selectedText = CStr(CallByName(ctl, "Selected", VbGet) & "")
                            On Error GoTo Failed
                            CodexReadWTExceptionInputs = CodexReadWTExceptionInputs & "ID=" & CStr(idItem) & "|TEXT=" & CodexClean(valueText) & "|CHANGEABLE=" & CodexClean(CStr(ctl.Changeable & "")) & "|SELECTED=" & CodexClean(selectedText) & vbLf
                        End If
                    Next idItem
                    For Each idItem In Array("rad/SCWM/S_CONF_WT_FG_MON-FLG_NO_QTY_DEV", "rad/SCWM/S_CONF_WT_FG_MON-FLG_SPLIT_WT", "rad/SCWM/S_CONF_WT_FG_MON-FLG_QTY_DIFF_WHSE", "rad/SCWM/S_CONF_WT_FG_MON-FLG_QTY_DIFF_SBIN")
                        Set ctl = ses.FindById(CStr(win.id) & "/usr/" & CStr(idItem), False)
                        If Not ctl Is Nothing Then
                            selectedText = "": On Error Resume Next: selectedText = CStr(CallByName(ctl, "Selected", VbGet) & ""): On Error GoTo Failed
                            CodexReadWTExceptionInputs = CodexReadWTExceptionInputs & "ID=" & CStr(idItem) & "|SELECTED=" & CodexClean(selectedText) & "|CHANGEABLE=" & CodexClean(CStr(ctl.Changeable & "")) & vbLf
                        End If
                    Next idItem
                    Exit Function
                End If
            Next j
        End If
    Next i
    CodexReadWTExceptionInputs = "EXCEPTION_WINDOW_NOT_FOUND"
    Exit Function
Failed:
    CodexReadWTExceptionInputs = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexSetWTExceptionQtyAndCheck(ByVal expectedOB As String, ByVal expectedDoc As String, ByVal expectedWT As String, ByVal expectedWH As String, ByVal expectedMaterial As String, ByVal expectedBin As String, ByVal expectedBatch As String, ByVal targetQtyText As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, field As Object, btn As Object
    Dim i As Long, j As Long, tcode As String, titleText As String, winText As String, statusText As String, beforeQty As String, afterQty As String
    Dim wtText As String, whoText As String, matText As String, binText As String, batchText As String, targetText As String, radioText As String, radioName As Variant
    On Error GoTo Failed
    expectedOB = Trim$(expectedOB): expectedDoc = Trim$(expectedDoc): expectedWT = Trim$(expectedWT): expectedWH = Trim$(expectedWH)
    expectedMaterial = Trim$(expectedMaterial): expectedBin = Trim$(expectedBin): expectedBatch = Trim$(expectedBatch): targetText = Trim$(targetQtyText)
    If targetText <> "20" Then CodexSetWTExceptionQtyAndCheck = "STOP_TARGET_QTY_NOT_20": Exit Function
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then Exit For
    Next i
    If ses Is Nothing Then CodexSetWTExceptionQtyAndCheck = "MON_SESSION_NOT_FOUND": Exit Function

    ' Open the exact WT method. This is navigation only; the Confirm WT button is not pressed here.
    If ses.children.count < 2 Then
        CodexSetWTExceptionQtyAndCheck = CodexOpenWTExceptionExact(expectedOB, expectedDoc, expectedWT)
        If InStr(1, CodexSetWTExceptionQtyAndCheck, "WT_EXCEPTION_PROBE", vbTextCompare) = 0 Then Exit Function
    End If
    Set field = ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-NISTA", False)
    If field Is Nothing Then CodexSetWTExceptionQtyAndCheck = "EXCEPTION_QTY_CONTROL_NOT_FOUND": Exit Function
    wtText = CStr(ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-TANUM", False).text & "")
    whoText = CStr(ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-WHO", False).text & "")
    matText = CStr(ses.FindById("wnd[1]/usr/ctxt/SCWM/S_CONF_WT_FG_MON-MATNR", False).text & "")
    batchText = CStr(ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-CHARG", False).text & "")
    binText = CStr(ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-VLPLA", False).text & "")
    beforeQty = CStr(field.text & "")
    If StrComp(Trim$(wtText), expectedWT, vbTextCompare) <> 0 Or _
       StrComp(Trim$(whoText), expectedWH, vbTextCompare) <> 0 Or _
       StrComp(Trim$(matText), expectedMaterial, vbTextCompare) <> 0 Or _
       StrComp(Trim$(batchText), expectedBatch, vbTextCompare) <> 0 Or _
       StrComp(Trim$(binText), expectedBin, vbTextCompare) <> 0 Then
        CodexSetWTExceptionQtyAndCheck = "STOP_EXCEPTION_PREVERIFY_MISMATCH|WT=" & CodexClean(wtText) & "|WH=" & CodexClean(whoText) & "|MAT=" & CodexClean(matText) & "|BATCH=" & CodexClean(batchText) & "|BIN=" & CodexClean(binText)
        Exit Function
    End If
    If Trim$(beforeQty) <> "10" Then CodexSetWTExceptionQtyAndCheck = "STOP_EXCEPTION_CURRENT_QTY_NOT_10|QTY=" & CodexClean(beforeQty): Exit Function
    If Not CBool(field.Changeable) Then CodexSetWTExceptionQtyAndCheck = "STOP_EXCEPTION_QTY_NOT_CHANGEABLE": Exit Function

    field.text = targetText
    afterQty = CStr(field.text & "")
    If Trim$(afterQty) <> targetText Then CodexSetWTExceptionQtyAndCheck = "STOP_EXCEPTION_QTY_READBACK_MISMATCH|QTY=" & CodexClean(afterQty): Exit Function

    Set btn = ses.FindById("wnd[1]/tbar[0]/btn[0]", False)
    If btn Is Nothing Then CodexSetWTExceptionQtyAndCheck = "EXCEPTION_CHECK_BUTTON_NOT_FOUND": Exit Function
    btn.Press
    If Not CodexWaitSessionIdle(ses, 30#) Then CodexSetWTExceptionQtyAndCheck = "EXCEPTION_CHECK_BUSY_TIMEOUT": Exit Function
    statusText = "": On Error Resume Next: statusText = CStr(ses.FindById("wnd[1]/sbar", False).text & ""): On Error GoTo Failed
    titleText = "": On Error Resume Next: titleText = CStr(ses.FindById("wnd[1]", False).text & ""): On Error GoTo Failed
    radioText = ""
    For Each radioName In Array("FLG_NO_QTY_DEV", "FLG_SPLIT_WT", "FLG_QTY_DIFF_WHSE", "FLG_QTY_DIFF_SBIN")
        Set field = ses.FindById("wnd[1]/usr/rad/SCWM/S_CONF_WT_FG_MON-" & CStr(radioName), False)
        If Not field Is Nothing Then
            On Error Resume Next
            If CBool(CallByName(field, "Selected", VbGet)) Then radioText = CStr(radioName)
            On Error GoTo Failed
        End If
    Next win
    CodexSetWTExceptionQtyAndCheck = "EXCEPTION_QTY_CHECKED|WT=" & CodexClean(wtText) & "|WHO=" & CodexClean(whoText) & "|MAT=" & CodexClean(matText) & "|BIN=" & CodexClean(binText) & "|BATCH=" & CodexClean(batchText) & "|BEFORE=" & CodexClean(beforeQty) & "|AFTER=" & CodexClean(afterQty) & "|STATUS=" & CodexClean(statusText) & "|TITLE=" & CodexClean(titleText) & "|RADIO=" & CodexClean(radioText)
    Exit Function
Failed:
    CodexSetWTExceptionQtyAndCheck = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexSetWTExceptionQtyOnly(ByVal expectedOB As String, ByVal expectedDoc As String, ByVal expectedWT As String, ByVal expectedWH As String, ByVal expectedMaterial As String, ByVal expectedBin As String, ByVal expectedBatch As String, ByVal targetQtyText As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, field As Object, i As Long, tcode As String
    Dim wtText As String, whoText As String, matText As String, binText As String, batchText As String, beforeQty As String, afterQty As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then Exit For
    Next i
    If ses Is Nothing Then CodexSetWTExceptionQtyOnly = "MON_SESSION_NOT_FOUND": Exit Function
    If ses.children.count < 2 Then
        CodexSetWTExceptionQtyOnly = CodexOpenWTExceptionExact(expectedOB, expectedDoc, expectedWT)
        If InStr(1, CodexSetWTExceptionQtyOnly, "WT_EXCEPTION_PROBE", vbTextCompare) = 0 Then Exit Function
    End If
    wtText = CStr(ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-TANUM", False).text & "")
    whoText = CStr(ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-WHO", False).text & "")
    matText = CStr(ses.FindById("wnd[1]/usr/ctxt/SCWM/S_CONF_WT_FG_MON-MATNR", False).text & "")
    batchText = CStr(ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-CHARG", False).text & "")
    binText = CStr(ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-VLPLA", False).text & "")
    Set field = ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-NISTA", False)
    beforeQty = CStr(field.text & "")
    If StrComp(Trim$(wtText), expectedWT, vbTextCompare) <> 0 Or StrComp(Trim$(whoText), expectedWH, vbTextCompare) <> 0 Or StrComp(Trim$(matText), expectedMaterial, vbTextCompare) <> 0 Or StrComp(Trim$(batchText), expectedBatch, vbTextCompare) <> 0 Or StrComp(Trim$(binText), expectedBin, vbTextCompare) <> 0 Then
        CodexSetWTExceptionQtyOnly = "STOP_EXCEPTION_PREVERIFY_MISMATCH": Exit Function
    End If
    If Trim$(beforeQty) <> "10" Then CodexSetWTExceptionQtyOnly = "STOP_EXCEPTION_CURRENT_QTY_NOT_10|QTY=" & CodexClean(beforeQty): Exit Function
    field.text = Trim$(targetQtyText)
    afterQty = CStr(field.text & "")
    CodexSetWTExceptionQtyOnly = "EXCEPTION_QTY_READBACK|WT=" & CodexClean(wtText) & "|WHO=" & CodexClean(whoText) & "|MAT=" & CodexClean(matText) & "|BIN=" & CodexClean(binText) & "|BATCH=" & CodexClean(batchText) & "|BEFORE=" & CodexClean(beforeQty) & "|AFTER=" & CodexClean(afterQty)
    Exit Function
Failed:
    CodexSetWTExceptionQtyOnly = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexCheckWTExceptionInputOnly() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, btn As Object, field As Object, win As Object
    Dim i As Long, j As Long, tcode As String, titleText As String, statusText As String, qtyText As String, t0 As Double, childText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then Exit For
    Next i
    If ses Is Nothing Then CodexCheckWTExceptionInputOnly = "MON_SESSION_NOT_FOUND": Exit Function
    Set field = ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-NISTA", False)
    If field Is Nothing Then CodexCheckWTExceptionInputOnly = "EXCEPTION_QTY_CONTROL_NOT_FOUND": Exit Function
    qtyText = CStr(field.text & "")
    If Trim$(qtyText) <> "20" Then CodexCheckWTExceptionInputOnly = "STOP_EXCEPTION_QTY_NOT_20|QTY=" & CodexClean(qtyText): Exit Function
    Set btn = ses.FindById("wnd[1]/tbar[0]/btn[0]", False)
    If btn Is Nothing Then CodexCheckWTExceptionInputOnly = "EXCEPTION_CHECK_BUTTON_NOT_FOUND": Exit Function
    btn.Press
    t0 = Timer
    Do
        DoEvents
        If CodexElapsed(t0) >= 3# Then Exit Do
    Loop
    statusText = "": On Error Resume Next: statusText = CStr(ses.FindById("wnd[1]/sbar", False).text & ""): On Error GoTo Failed
    titleText = "": On Error Resume Next: titleText = CStr(ses.FindById("wnd[1]", False).text & ""): On Error GoTo Failed
    CodexCheckWTExceptionInputOnly = "EXCEPTION_CHECK_SENT|QTY=" & CodexClean(qtyText) & "|STATUS=" & CodexClean(statusText) & "|TITLE=" & CodexClean(titleText) & "|WINDOWS=" & CStr(ses.children.count)
    For j = 0 To ses.children.count - 1
        Set win = ses.children.ElementAt(j): childText = "": On Error Resume Next: childText = CStr(win.text & ""): On Error GoTo Failed
        CodexCheckWTExceptionInputOnly = CodexCheckWTExceptionInputOnly & "|W" & CStr(j) & "=" & CodexClean(childText)
    Next j
    Exit Function
Failed:
    CodexCheckWTExceptionInputOnly = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexSelectWarehouseDifferenceAndCheck() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, radio As Object, qtyField As Object, btn As Object, win As Object
    Dim i As Long, j As Long, tcode As String, statusText As String, titleText As String, childText As String, t0 As Double
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then Exit For
    Next i
    If ses Is Nothing Then CodexSelectWarehouseDifferenceAndCheck = "MON_SESSION_NOT_FOUND": Exit Function
    Set qtyField = ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-NISTA", False)
    Set radio = ses.FindById("wnd[1]/usr/rad/SCWM/S_CONF_WT_FG_MON-FLG_QTY_DIFF_WHSE", False)
    Set btn = ses.FindById("wnd[1]/tbar[0]/btn[0]", False)
    If qtyField Is Nothing Or radio Is Nothing Or btn Is Nothing Then CodexSelectWarehouseDifferenceAndCheck = "EXCEPTION_CONTROLS_NOT_FOUND": Exit Function
    If Trim$(CStr(qtyField.text & "")) <> "20" Then CodexSelectWarehouseDifferenceAndCheck = "STOP_EXCEPTION_QTY_NOT_20": Exit Function
    radio.Press
    btn.Press
    t0 = Timer
    Do
        DoEvents
        If CodexElapsed(t0) >= 3# Then Exit Do
    Loop
    statusText = "": On Error Resume Next: statusText = CStr(ses.FindById("wnd[1]/sbar", False).text & ""): On Error GoTo Failed
    titleText = "": On Error Resume Next: titleText = CStr(ses.FindById("wnd[1]", False).text & ""): On Error GoTo Failed
    CodexSelectWarehouseDifferenceAndCheck = "WAREHOUSE_DIFF_CHECKED|QTY=" & CodexClean(CStr(qtyField.text & "")) & "|SELECTED=" & CodexClean(CStr(CallByName(radio, "Selected", VbGet) & "")) & "|STATUS=" & CodexClean(statusText) & "|TITLE=" & CodexClean(titleText) & "|WINDOWS=" & CStr(ses.children.count)
    For j = 0 To ses.children.count - 1
        Set win = ses.children.ElementAt(j): childText = "": On Error Resume Next: childText = CStr(win.text & ""): On Error GoTo Failed
        CodexSelectWarehouseDifferenceAndCheck = CodexSelectWarehouseDifferenceAndCheck & "|W" & CStr(j) & "=" & CodexClean(childText)
    Next j
    Exit Function
Failed:
    CodexSelectWarehouseDifferenceAndCheck = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexProbeWarehouseDifferenceRadio() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, radio As Object
    Dim i As Long, tcode As String, beforeText As String, afterText As String, errSelect As String, errPress As String, errSet As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then Exit For
    Next i
    Set radio = ses.FindById("wnd[1]/usr/rad/SCWM/S_CONF_WT_FG_MON-FLG_QTY_DIFF_WHSE", False)
    If radio Is Nothing Then CodexProbeWarehouseDifferenceRadio = "RADIO_NOT_FOUND": Exit Function
    beforeText = CStr(CallByName(radio, "Selected", VbGet) & "")
    On Error Resume Next
    Err.Clear: radio.Select: errSelect = CStr(Err.Number) & ":" & Err.Description: Err.Clear
    Err.Clear: radio.Press: errPress = CStr(Err.Number) & ":" & Err.Description: Err.Clear
    Err.Clear: radio.selected = True: errSet = CStr(Err.Number) & ":" & Err.Description: Err.Clear
    afterText = CStr(CallByName(radio, "Selected", VbGet) & "")
    On Error GoTo Failed
    CodexProbeWarehouseDifferenceRadio = "RADIO_PROBE|BEFORE=" & CodexClean(beforeText) & "|AFTER=" & CodexClean(afterText) & "|SELECT=" & CodexClean(errSelect) & "|PRESS=" & CodexClean(errPress) & "|SET=" & CodexClean(errSet)
    Exit Function
Failed:
    CodexProbeWarehouseDifferenceRadio = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexConfirmWTExceptionAndReadStatus(ByVal expectedWT As String, ByVal expectedWH As String, ByVal expectedMaterial As String, ByVal expectedBin As String, ByVal expectedBatch As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, btn As Object, field As Object, radio As Object, win As Object, grid As Object
    Dim i As Long, j As Long, r As Long, tcode As String, qtyText As String, wtText As String, whoText As String, matText As String, binText As String, batchText As String
    Dim titleText As String, statusText As String, childText As String, t0 As Double, lowerPath As String, wtRow As Long, matches As Long
    Dim liveWT As String, liveWH As String, liveMat As String, liveBin As String, liveBatch As String, liveQty As String, liveStatus As String, liveDoc As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then Exit For
    Next i
    If ses Is Nothing Then CodexConfirmWTExceptionAndReadStatus = "MON_SESSION_NOT_FOUND": Exit Function
    Set field = ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-NISTA", False)
    Set radio = ses.FindById("wnd[1]/usr/rad/SCWM/S_CONF_WT_FG_MON-FLG_QTY_DIFF_WHSE", False)
    Set btn = ses.FindById("wnd[1]/tbar[0]/btn[8]", False)
    If field Is Nothing Or radio Is Nothing Or btn Is Nothing Then CodexConfirmWTExceptionAndReadStatus = "STOP_CONFIRM_CONTROLS_NOT_FOUND": Exit Function
    qtyText = CStr(field.text & "")
    wtText = CStr(ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-TANUM", False).text & "")
    whoText = CStr(ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-WHO", False).text & "")
    matText = CStr(ses.FindById("wnd[1]/usr/ctxt/SCWM/S_CONF_WT_FG_MON-MATNR", False).text & "")
    batchText = CStr(ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-CHARG", False).text & "")
    binText = CStr(ses.FindById("wnd[1]/usr/txt/SCWM/S_CONF_WT_FG_MON-VLPLA", False).text & "")
    If Trim$(qtyText) <> "20" Or StrComp(Trim$(wtText), expectedWT, vbTextCompare) <> 0 Or StrComp(Trim$(whoText), expectedWH, vbTextCompare) <> 0 Or StrComp(Trim$(matText), expectedMaterial, vbTextCompare) <> 0 Or StrComp(Trim$(batchText), expectedBatch, vbTextCompare) <> 0 Or StrComp(Trim$(binText), expectedBin, vbTextCompare) <> 0 Then
        CodexConfirmWTExceptionAndReadStatus = "STOP_CONFIRM_PREVERIFY_MISMATCH|WT=" & CodexClean(wtText) & "|WHO=" & CodexClean(whoText) & "|MAT=" & CodexClean(matText) & "|BIN=" & CodexClean(binText) & "|BATCH=" & CodexClean(batchText) & "|QTY=" & CodexClean(qtyText)
        Exit Function
    End If
    If Not CBool(CallByName(radio, "Selected", VbGet)) Then CodexConfirmWTExceptionAndReadStatus = "STOP_CONFIRM_DIFF_RADIO_NOT_SELECTED": Exit Function

    ' Single SAP mutation: Confirm WT. No retry is performed after this point.
    btn.Press
    t0 = Timer
    Do
        DoEvents
        If CodexElapsed(t0) >= 15# Then Exit Do
    Loop
    statusText = "": On Error Resume Next: statusText = CStr(ses.FindById("wnd[0]/sbar", False).text & ""): On Error GoTo Failed
    CodexConfirmWTExceptionAndReadStatus = "CONFIRM_SENT|STATUS=" & CodexClean(statusText) & "|WINDOWS=" & CStr(ses.children.count)
    For j = 0 To ses.children.count - 1
        Set win = ses.children.ElementAt(j): childText = "": On Error Resume Next: childText = CStr(win.text & ""): On Error GoTo Failed
        CodexConfirmWTExceptionAndReadStatus = CodexConfirmWTExceptionAndReadStatus & "|W" & CStr(j) & "=" & CodexClean(childText)
    Next j
    Set field = ses.FindById("wnd[1]", False)
    If Not field Is Nothing Then
        CodexConfirmWTExceptionAndReadStatus = CodexConfirmWTExceptionAndReadStatus & "|STOP_CONFIRM_MODAL_REMAINS"
        Exit Function
    End If
    lowerPath = "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell"
    Set grid = ses.FindById(lowerPath, False)
    If grid Is Nothing Then CodexConfirmWTExceptionAndReadStatus = CodexConfirmWTExceptionAndReadStatus & "|STOP_MON_WT_GRID_NOT_FOUND": Exit Function
    matches = 0: wtRow = -1
    Codex_EnsureGridRendered grid, ses
    For r = 0 To CLng(grid.rowCount) - 1
        liveWT = CStr(grid.GetCellValue(r, "TANUM") & "")
        If StrComp(Trim$(liveWT), expectedWT, vbTextCompare) = 0 Then matches = matches + 1: wtRow = r
    Next r
    If matches <> 1 Then CodexConfirmWTExceptionAndReadStatus = CodexConfirmWTExceptionAndReadStatus & "|STOP_WT_MATCH_COUNT=" & CStr(matches): Exit Function
    liveDoc = CStr(grid.GetCellValue(wtRow, "DOCNO") & "")
    liveWH = CStr(grid.GetCellValue(wtRow, "WHO") & "")
    liveMat = CStr(grid.GetCellValue(wtRow, "MATNR") & "")
    liveBatch = CStr(grid.GetCellValue(wtRow, "CHARG") & "")
    liveBin = CStr(grid.GetCellValue(wtRow, "VLPLA") & "")
    liveQty = CStr(grid.GetCellValue(wtRow, "VSOLM") & "")
    liveStatus = CStr(grid.GetCellValue(wtRow, "TOSTAT") & "")
    CodexConfirmWTExceptionAndReadStatus = CodexConfirmWTExceptionAndReadStatus & "|LIVE_WT=" & CodexClean(liveWT) & "|LIVE_DOC=" & CodexClean(liveDoc) & "|LIVE_WHO=" & CodexClean(liveWH) & "|LIVE_MAT=" & CodexClean(liveMat) & "|LIVE_BIN=" & CodexClean(liveBin) & "|LIVE_BATCH=" & CodexClean(liveBatch) & "|LIVE_QTY=" & CodexClean(liveQty) & "|LIVE_STATUS=" & CodexClean(liveStatus)
    Exit Function
Failed:
    CodexConfirmWTExceptionAndReadStatus = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexShowWTExceptionsExact(ByVal expectedWT As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object, win As Object
    Dim i As Long, j As Long, r As Long, tcode As String, wtText As String, titleText As String, childText As String, statusText As String
    Dim matchCount As Long, rowNo As Long, t0 As Double
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(Replace(tcode, "/", "")) = "SCWMMON" Then Exit For
    Next i
    If ses Is Nothing Then CodexShowWTExceptionsExact = "MON_SESSION_NOT_FOUND": Exit Function
    Set grid = ses.FindById("wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell", False)
    If grid Is Nothing Then CodexShowWTExceptionsExact = "MON_WT_GRID_NOT_FOUND": Exit Function
    matchCount = 0: rowNo = -1
    Codex_EnsureGridRendered grid, ses
    For r = 0 To CLng(grid.rowCount) - 1
        wtText = CStr(grid.GetCellValue(r, "TANUM") & "")
        If StrComp(Trim$(wtText), Trim$(expectedWT), vbTextCompare) = 0 Then matchCount = matchCount + 1: rowNo = r
    Next r
    If matchCount <> 1 Then CodexShowWTExceptionsExact = "MON_WT_MATCH_COUNT=" & CStr(matchCount): Exit Function
    grid.selectedRows = CStr(rowNo): grid.SetCurrentCell rowNo, "TANUM": grid.PressToolbarContextButton "METHODS"
    grid.SelectContextMenuItemByText "Show Exceptions"
    t0 = Timer
    Do
        DoEvents
        If CodexElapsed(t0) >= 5# Then Exit Do
    Loop
    statusText = "": On Error Resume Next: statusText = CStr(ses.FindById("wnd[0]/sbar", False).text & ""): On Error GoTo Failed
    CodexShowWTExceptionsExact = "SHOW_EXCEPTIONS_SENT|STATUS=" & CodexClean(statusText) & "|WINDOWS=" & CStr(ses.children.count)
    For j = 0 To ses.children.count - 1
        Set win = ses.children.ElementAt(j): childText = "": On Error Resume Next: childText = CStr(win.text & ""): On Error GoTo Failed
        CodexShowWTExceptionsExact = CodexShowWTExceptionsExact & "|W" & CStr(j) & "=" & CodexClean(childText)
    Next j
    Exit Function
Failed:
    CodexShowWTExceptionsExact = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexReadPRDOMenus() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, menuItem As Object
    Dim i As Long, j As Long, tcode As String, textValue As String, idValue As String, changeValue As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then Exit For
    Next i
    If ses Is Nothing Then CodexReadPRDOMenus = "PRDO_SESSION_NOT_FOUND": Exit Function
    CodexReadPRDOMenus = "PRDO_MENUS|SESSION=" & CStr(ses.id) & vbLf
    For i = 0 To 12
        Set menuItem = ses.FindById("wnd[0]/mbar/menu[" & CStr(i) & "]", False)
        If Not menuItem Is Nothing Then
            CodexReadPRDOMenus = CodexReadPRDOMenus & "TOP=" & CStr(i) & "|TEXT=" & CodexClean(CStr(menuItem.text & "")) & vbLf
            For j = 0 To 40
                Set menuItem = ses.FindById("wnd[0]/mbar/menu[" & CStr(i) & "]/menu[" & CStr(j) & "]", False)
                If menuItem Is Nothing Then Exit For
                textValue = "": idValue = "": changeValue = ""
                On Error Resume Next
                textValue = CStr(menuItem.text & "")
                idValue = CStr(menuItem.id & "")
                changeValue = CStr(menuItem.Changeable & "")
                On Error GoTo Failed
                CodexReadPRDOMenus = CodexReadPRDOMenus & "ITEM=" & CStr(i) & "." & CStr(j) & "|ID=" & CodexClean(idValue) & "|TEXT=" & CodexClean(textValue) & "|CHANGEABLE=" & CodexClean(changeValue) & vbLf
            Next j
        End If
    Next i
    Exit Function
Failed:
    CodexReadPRDOMenus = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexExpandConfirmWTDetailsAndDump() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, btn As Object
    Dim i As Long, tcode As String, titleText As String, props As Variant
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": titleText = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction & "")
        titleText = CStr(ses.FindById("wnd[0]", False).text & "")
        On Error GoTo Failed
        If InStr(1, titleText, "Confirm Warehouse Task", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexExpandConfirmWTDetailsAndDump = "CONFIRM_SESSION_NOT_FOUND": Exit Function
    Set btn = ses.FindById("wnd[0]/usr/subSUB_ODP_2:/SCWM/SAPLUI_TO_CONF:0245/btnCMD_ODP2_EXPAND", False)
    If btn Is Nothing Then CodexExpandConfirmWTDetailsAndDump = "CONFIRM_DETAILS_EXPAND_NOT_FOUND": Exit Function
    btn.Press
    If Not CodexWaitSessionIdle(ses, 20#) Then CodexExpandConfirmWTDetailsAndDump = "CONFIRM_DETAILS_BUSY_TIMEOUT": Exit Function
    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
    CodexExpandConfirmWTDetailsAndDump = "CONFIRM_DETAILS_EXPANDED|TITLE=" & CodexClean(CStr(ses.FindById("wnd[0]", False).text & "")) & vbLf & CStr(ses.GetObjectTree("wnd[0]", props))
    Exit Function
Failed:
    CodexExpandConfirmWTDetailsAndDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexReadPRDOToolbars() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, ctl As Object, toolbar As Object
    Dim i As Long, j As Long, tcode As String, titleText As String, textValue As String, tipValue As String, typeValue As String, enabledValue As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": titleText = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction & "")
        titleText = CStr(ses.FindById("wnd[0]", False).text & "")
        On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then Exit For
    Next i
    If ses Is Nothing Then CodexReadPRDOToolbars = "PRDO_SESSION_NOT_FOUND": Exit Function
    CodexReadPRDOToolbars = "PRDO_TOOLBARS|SESSION=" & CStr(ses.id) & "|TITLE=" & CodexClean(titleText) & vbLf
    For i = 0 To 4
        Set toolbar = ses.FindById("wnd[0]/tbar[" & CStr(i) & "]", False)
        If Not toolbar Is Nothing Then
            CodexReadPRDOToolbars = CodexReadPRDOToolbars & "TOOLBAR=" & CStr(i) & vbLf
            For j = 0 To toolbar.children.count - 1
                Set ctl = toolbar.children.ElementAt(j)
                textValue = "": tipValue = "": typeValue = "": enabledValue = ""
                On Error Resume Next
                textValue = CStr(ctl.text & "")
                tipValue = CStr(ctl.Tooltip & "")
                typeValue = CStr(ctl.Type & "")
                enabledValue = CStr(ctl.Changeable & "")
                On Error GoTo Failed
                CodexReadPRDOToolbars = CodexReadPRDOToolbars & "BTN=" & CStr(i) & "." & CStr(j) & "|ID=" & CStr(ctl.id) & "|TYPE=" & CodexClean(typeValue) & "|TEXT=" & CodexClean(textValue) & "|TIP=" & CodexClean(tipValue) & "|CHANGEABLE=" & CodexClean(enabledValue) & vbLf
            Next j
        End If
    Next i
    Exit Function
Failed:
    CodexReadPRDOToolbars = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexFindPRDOEditControls() As String
    Dim gui As Object, app As Object, con As Object, ses As Object
    Dim i As Long, tcode As String, titleText As String, props As Variant, treeText As String, oneLine As Variant, q As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": titleText = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction & "")
        titleText = CStr(ses.FindById("wnd[0]", False).text & "")
        On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then Exit For
    Next i
    If ses Is Nothing Then CodexFindPRDOEditControls = "PRDO_SESSION_NOT_FOUND": Exit Function
    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
    treeText = CStr(ses.GetObjectTree("wnd[0]", props))
    q = Chr$(34)
    CodexFindPRDOEditControls = "PRDO_EDIT_CONTROLS|TITLE=" & CodexClean(titleText) & vbLf
    For Each oneLine In Split(treeText, vbLf)
        If InStr(1, CStr(oneLine), q & "Tooltip" & q & ":" & q & "Change", vbTextCompare) > 0 Or _
           InStr(1, CStr(oneLine), q & "Text" & q & ":" & q & "Change", vbTextCompare) > 0 Or _
           InStr(1, CStr(oneLine), q & "Name" & q & ":" & q & "CHANGE", vbTextCompare) > 0 Or _
           InStr(1, CStr(oneLine), q & "Tooltip" & q & ":" & q & "Edit", vbTextCompare) > 0 Or _
           InStr(1, CStr(oneLine), q & "Text" & q & ":" & q & "Edit", vbTextCompare) > 0 Then
            CodexFindPRDOEditControls = CodexFindPRDOEditControls & CStr(oneLine) & vbLf
        End If
    Next oneLine
    Exit Function
Failed:
    CodexFindPRDOEditControls = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexReadPRDOCustomToolbars() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, candidateSes As Object, shell As Object, child As Object
    Dim i As Long, k As Long, tcode As String, titleText As String, pathText As String, labelText As Variant
    Dim countValue As Long, idText As String, textValue As String, tipValue As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    Set ses = Nothing
    For i = 0 To con.children.count - 1
        Set candidateSes = con.children.ElementAt(i): tcode = "": titleText = ""
        On Error Resume Next
        tcode = CStr(candidateSes.Info.Transaction & "")
        On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then
            Set ses = candidateSes
            On Error Resume Next
            titleText = CStr(ses.FindById("wnd[0]", False).text & "")
            On Error GoTo Failed
            Exit For
        End If
    Next i
    If ses Is Nothing Then CodexReadPRDOCustomToolbars = "PRDO_SESSION_NOT_FOUND": Exit Function
    CodexReadPRDOCustomToolbars = "PRDO_CUSTOM_TOOLBARS|TITLE=" & CodexClean(titleText) & vbLf
    For Each labelText In Array("OIP", "ODP1", "ODP2")
        If labelText = "OIP" Then
            pathText = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell"
        ElseIf labelText = "ODP1" Then
            pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/cntlCONTAINER_TB_ODP1_1/shellcont/shell"
        Else
            pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP2:/SCWM/SAPLUI_DLV_PRD:4000/tabsTABSTRIP_ODP2/tabpOK_ODP2_TAB1/ssubSUB_ODP2_TAB1:/SCWM/SAPLUI_DLV_CORE:4210/cntlCONTAINER_TB_ODP2_1/shellcont/shell"
        End If
        Set shell = ses.FindById(pathText, False)
        If shell Is Nothing Then
            CodexReadPRDOCustomToolbars = CodexReadPRDOCustomToolbars & labelText & "_NOT_FOUND|ID=" & pathText & vbLf
        Else
            countValue = 0
            On Error Resume Next
            Err.Clear: countValue = CLng(shell.buttonCount)
            On Error GoTo Failed
            CodexReadPRDOCustomToolbars = CodexReadPRDOCustomToolbars & labelText & "_TOOLBAR=" & CStr(shell.id) & "|COUNT=" & CStr(countValue) & vbLf
            For k = 0 To countValue - 1
                idText = "": textValue = "": tipValue = ""
                On Error Resume Next
                idText = CStr(shell.GetToolbarButtonId(k) & "")
                textValue = CStr(shell.GetToolbarButtonText(k) & "")
                tipValue = CStr(shell.GetToolbarButtonTooltip(k) & "")
                If idText = "" Then idText = CStr(shell.GetButtonId(k) & "")
                If textValue = "" Then textValue = CStr(shell.GetButtonText(k) & "")
                If tipValue = "" Then tipValue = CStr(shell.GetButtonTooltip(k) & "")
                If idText = "" And k < shell.children.count Then
                    Set child = shell.children.ElementAt(k)
                    idText = CStr(child.id & "")
                    textValue = CStr(child.text & "")
                    tipValue = CStr(child.Tooltip & "")
                End If
                On Error GoTo Failed
                CodexReadPRDOCustomToolbars = CodexReadPRDOCustomToolbars & labelText & "_BTN=" & CStr(k) & "|ID=" & CodexClean(idText) & "|TEXT=" & CodexClean(textValue) & "|TIP=" & CodexClean(tipValue) & vbLf
            Next k
        End If
    Next labelText
    Exit Function
Failed:
    CodexReadPRDOCustomToolbars = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexEnterPRDOEditAndDump() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, toolbar As Object, grid As Object
    Dim i As Long, r As Long, k As Long, tcode As String, titleText As String, statusText As String, idText As String, valueText As String, changeText As String
    Dim pathText As String, countValue As Long, colId As Variant, lineText As String, t0 As Double
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then Exit For
    Next i
    If ses Is Nothing Then CodexEnterPRDOEditAndDump = "PRDO_SESSION_NOT_FOUND": Exit Function
    Set toolbar = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell", False)
    If toolbar Is Nothing Then CodexEnterPRDOEditAndDump = "PRDO_OIP_TOOLBAR_NOT_FOUND": Exit Function
    toolbar.PressButton "OIP_CHANGE"
    If Not CodexWaitSessionIdle(ses, 20#) Then CodexEnterPRDOEditAndDump = "PRDO_EDIT_BUSY_TIMEOUT": Exit Function
    titleText = "": statusText = ""
    On Error Resume Next
    titleText = CStr(ses.FindById("wnd[0]", False).text & "")
    statusText = CStr(ses.FindById("wnd[0]/sbar", False).text & "")
    On Error GoTo Failed
    pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell"
    Set grid = ses.FindById(pathText, False)
    If grid Is Nothing Then CodexEnterPRDOEditAndDump = "PRDO_ITEM_GRID_NOT_FOUND|TITLE=" & CodexClean(titleText) & "|STATUS=" & CodexClean(statusText): Exit Function
    CodexEnterPRDOEditAndDump = "PRDO_EDIT_ENTERED|TITLE=" & CodexClean(titleText) & "|STATUS=" & CodexClean(statusText) & "|ROWS=" & CStr(grid.rowCount) & vbLf
    For Each colId In Array("ITEMNO", "PRODUCTNO", "BATCHNO_UI", "QTY_UI", "UOM", "CHANGEABLE_ICON", "STATUS_PICKING", "STATUS_GI")
        valueText = "": changeText = ""
        On Error Resume Next
        valueText = CStr(grid.GetCellValue(0, CStr(colId)) & "")
        changeText = CStr(grid.GetCellChangeable(0, CStr(colId)) & "")
        On Error GoTo Failed
        CodexEnterPRDOEditAndDump = CodexEnterPRDOEditAndDump & "COL=" & CStr(colId) & "|VALUE=" & CodexClean(valueText) & "|CHANGEABLE=" & CodexClean(changeText) & vbLf
    Next colId
    On Error Resume Next
    countValue = CLng(toolbar.buttonCount)
    On Error GoTo Failed
    For k = 0 To countValue - 1
        idText = "": valueText = ""
        On Error Resume Next
        idText = CStr(toolbar.GetButtonId(k) & "")
        valueText = CStr(toolbar.GetButtonTooltip(k) & "")
        On Error GoTo Failed
        If InStr(1, idText, "OIP_", vbTextCompare) > 0 Or InStr(1, valueText, "Change", vbTextCompare) > 0 Or InStr(1, valueText, "Display", vbTextCompare) > 0 Then CodexEnterPRDOEditAndDump = CodexEnterPRDOEditAndDump & "OIP_BTN=" & CStr(k) & "|ID=" & CodexClean(idText) & "|TIP=" & CodexClean(valueText) & vbLf
    Next k
    Exit Function
Failed:
    CodexEnterPRDOEditAndDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexModifyPRDOQtyAndReadback(ByVal expectedDoc As String, ByVal expectedItem As String, ByVal expectedMaterial As String, ByVal expectedBatch As String, ByVal targetQtyText As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object, toolbar As Object
    Dim i As Long, tcode As String, pathText As String, itemText As String, matText As String, batchText As String, beforeQty As String, afterQty As String, statusText As String
    Dim changeableText As String
    On Error GoTo Failed
    expectedDoc = Trim$(expectedDoc): expectedItem = Trim$(expectedItem): expectedMaterial = Trim$(expectedMaterial): expectedBatch = Trim$(expectedBatch): targetQtyText = Trim$(targetQtyText)
    If targetQtyText <> "20" Then CodexModifyPRDOQtyAndReadback = "STOP_TARGET_QTY_NOT_20": Exit Function
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then Exit For
    Next i
    If ses Is Nothing Then CodexModifyPRDOQtyAndReadback = "PRDO_SESSION_NOT_FOUND": Exit Function
    pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell"
    Set grid = ses.FindById(pathText, False)
    If grid Is Nothing Then CodexModifyPRDOQtyAndReadback = "PRDO_ITEM_GRID_NOT_FOUND": Exit Function
    If CLng(grid.rowCount) <> 1 Then CodexModifyPRDOQtyAndReadback = "STOP_PRDO_ITEM_ROW_COUNT=" & CStr(grid.rowCount): Exit Function
    itemText = CStr(grid.GetCellValue(0, "ITEMNO") & ""): matText = CStr(grid.GetCellValue(0, "PRODUCTNO") & ""): batchText = CStr(grid.GetCellValue(0, "BATCHNO_UI") & ""): beforeQty = CStr(grid.GetCellValue(0, "QTY_UI") & "")
    changeableText = CStr(grid.GetCellChangeable(0, "QTY_UI") & "")
    If StrComp(Trim$(itemText), expectedItem, vbTextCompare) <> 0 Or StrComp(Trim$(matText), expectedMaterial, vbTextCompare) <> 0 Or StrComp(Trim$(batchText), expectedBatch, vbTextCompare) <> 0 Or Trim$(beforeQty) <> "10" Then
        CodexModifyPRDOQtyAndReadback = "STOP_PRDO_QTY_PREVERIFY_MISMATCH|DOC=" & CodexClean(expectedDoc) & "|ITEM=" & CodexClean(itemText) & "|MAT=" & CodexClean(matText) & "|BATCH=" & CodexClean(batchText) & "|QTY=" & CodexClean(beforeQty): Exit Function
    End If
    If UCase$(Trim$(changeableText)) <> "TRUE" Then CodexModifyPRDOQtyAndReadback = "STOP_PRDO_QTY_NOT_CHANGEABLE|VALUE=" & CodexClean(changeableText): Exit Function

    grid.SetCurrentCell 0, "QTY_UI"
    grid.ModifyCell 0, "QTY_UI", targetQtyText
    afterQty = CStr(grid.GetCellValue(0, "QTY_UI") & "")
    statusText = "": On Error Resume Next: statusText = CStr(ses.FindById("wnd[0]/sbar", False).text & ""): On Error GoTo Failed
    CodexModifyPRDOQtyAndReadback = "PRDO_QTY_MODIFIED|DOC=" & CodexClean(expectedDoc) & "|ITEM=" & CodexClean(itemText) & "|MAT=" & CodexClean(matText) & "|BATCH=" & CodexClean(batchText) & "|BEFORE=" & CodexClean(beforeQty) & "|AFTER=" & CodexClean(afterQty) & "|CHANGEABLE=" & CodexClean(changeableText) & "|STATUS=" & CodexClean(statusText)
    Exit Function
Failed:
    CodexModifyPRDOQtyAndReadback = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexCheckPRDOEditedQty(ByVal expectedItem As String, ByVal expectedMaterial As String, ByVal expectedBatch As String, ByVal expectedQtyText As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object, toolbar As Object
    Dim i As Long, tcode As String, pathText As String, itemText As String, matText As String, batchText As String, qtyText As String, statusText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then Exit For
    Next i
    If ses Is Nothing Then CodexCheckPRDOEditedQty = "PRDO_SESSION_NOT_FOUND": Exit Function
    pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell"
    Set grid = ses.FindById(pathText, False)
    Set toolbar = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell", False)
    If grid Is Nothing Or toolbar Is Nothing Then CodexCheckPRDOEditedQty = "PRDO_CHECK_CONTROLS_NOT_FOUND": Exit Function
    itemText = CStr(grid.GetCellValue(0, "ITEMNO") & ""): matText = CStr(grid.GetCellValue(0, "PRODUCTNO") & ""): batchText = CStr(grid.GetCellValue(0, "BATCHNO_UI") & ""): qtyText = CStr(grid.GetCellValue(0, "QTY_UI") & "")
    If StrComp(Trim$(itemText), Trim$(expectedItem), vbTextCompare) <> 0 Or StrComp(Trim$(matText), Trim$(expectedMaterial), vbTextCompare) <> 0 Or StrComp(Trim$(batchText), Trim$(expectedBatch), vbTextCompare) <> 0 Or StrComp(Trim$(qtyText), Trim$(expectedQtyText), vbTextCompare) <> 0 Then
        CodexCheckPRDOEditedQty = "STOP_PRDO_CHECK_PREVERIFY_MISMATCH|ITEM=" & CodexClean(itemText) & "|MAT=" & CodexClean(matText) & "|BATCH=" & CodexClean(batchText) & "|QTY=" & CodexClean(qtyText): Exit Function
    End If
    toolbar.PressButton "OIP_CHECK"
    If Not CodexWaitSessionIdle(ses, 20#) Then CodexCheckPRDOEditedQty = "PRDO_CHECK_BUSY_TIMEOUT": Exit Function
    statusText = "": On Error Resume Next: statusText = CStr(ses.FindById("wnd[0]/sbar", False).text & ""): On Error GoTo Failed
    qtyText = CStr(grid.GetCellValue(0, "QTY_UI") & "")
    CodexCheckPRDOEditedQty = "PRDO_CHECKED|ITEM=" & CodexClean(itemText) & "|MAT=" & CodexClean(matText) & "|BATCH=" & CodexClean(batchText) & "|QTY=" & CodexClean(qtyText) & "|STATUS=" & CodexClean(statusText)
    Exit Function
Failed:
    CodexCheckPRDOEditedQty = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexSavePRDOEditedQtyAndReadback(ByVal expectedItem As String, ByVal expectedMaterial As String, ByVal expectedBatch As String, ByVal expectedQtyText As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object, toolbar As Object, win As Object
    Dim i As Long, j As Long, tcode As String, pathText As String, itemText As String, matText As String, batchText As String, qtyText As String, statusText As String, childText As String, t0 As Double
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" Then Exit For
    Next i
    If ses Is Nothing Then CodexSavePRDOEditedQtyAndReadback = "PRDO_SESSION_NOT_FOUND": Exit Function
    pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell"
    Set grid = ses.FindById(pathText, False)
    Set toolbar = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell", False)
    If grid Is Nothing Or toolbar Is Nothing Then CodexSavePRDOEditedQtyAndReadback = "PRDO_SAVE_CONTROLS_NOT_FOUND": Exit Function
    itemText = CStr(grid.GetCellValue(0, "ITEMNO") & ""): matText = CStr(grid.GetCellValue(0, "PRODUCTNO") & ""): batchText = CStr(grid.GetCellValue(0, "BATCHNO_UI") & ""): qtyText = CStr(grid.GetCellValue(0, "QTY_UI") & "")
    If StrComp(Trim$(itemText), Trim$(expectedItem), vbTextCompare) <> 0 Or StrComp(Trim$(matText), Trim$(expectedMaterial), vbTextCompare) <> 0 Or StrComp(Trim$(batchText), Trim$(expectedBatch), vbTextCompare) <> 0 Or StrComp(Trim$(qtyText), Trim$(expectedQtyText), vbTextCompare) <> 0 Then
        CodexSavePRDOEditedQtyAndReadback = "STOP_PRDO_SAVE_PREVERIFY_MISMATCH|ITEM=" & CodexClean(itemText) & "|MAT=" & CodexClean(matText) & "|BATCH=" & CodexClean(batchText) & "|QTY=" & CodexClean(qtyText): Exit Function
    End If
    toolbar.PressButton "SAVE"
    t0 = Timer
    Do
        DoEvents
        If CodexElapsed(t0) >= 15# Then Exit Do
    Loop
    statusText = "": On Error Resume Next: statusText = CStr(ses.FindById("wnd[0]/sbar", False).text & ""): On Error GoTo Failed
    qtyText = CStr(grid.GetCellValue(0, "QTY_UI") & "")
    CodexSavePRDOEditedQtyAndReadback = "PRDO_SAVED|ITEM=" & CodexClean(itemText) & "|MAT=" & CodexClean(matText) & "|BATCH=" & CodexClean(batchText) & "|QTY=" & CodexClean(qtyText) & "|STATUS=" & CodexClean(statusText) & "|WINDOWS=" & CStr(ses.children.count)
    For j = 0 To ses.children.count - 1
        Set win = ses.children.ElementAt(j): childText = "": On Error Resume Next: childText = CStr(win.text & ""): On Error GoTo Failed
        CodexSavePRDOEditedQtyAndReadback = CodexSavePRDOEditedQtyAndReadback & "|W" & CStr(j) & "=" & CodexClean(childText)
    Next j
    Exit Function
Failed:
    CodexSavePRDOEditedQtyAndReadback = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexOpenPRDOWarehouseTaskFollowOn() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, menuItem As Object
    Dim i As Long, j As Long, tcode As String, titleText As String, childTitle As String, t0 As Double
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": titleText = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction & "")
        titleText = CStr(ses.FindById("wnd[0]", False).text & "")
        On Error GoTo Failed
        If UCase$(tcode) = "/SCWM/PRDO" And InStr(1, titleText, "Maintain Outb. Deliv. Order", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexOpenPRDOWarehouseTaskFollowOn = "PRDO_SESSION_NOT_FOUND": Exit Function
    Set menuItem = ses.FindById("wnd[0]/mbar/menu[0]/menu[2]/menu[0]", False)
    If menuItem Is Nothing Then CodexOpenPRDOWarehouseTaskFollowOn = "PRDO_FOLLOW_ON_MENU_NOT_FOUND": Exit Function
    menuItem.Select
    t0 = Timer
    Do
        DoEvents
        For j = 0 To con.children.count - 1
            Set ses = con.children.ElementAt(j): childTitle = ""
            On Error Resume Next: childTitle = CStr(ses.FindById("wnd[0]", False).text & ""): On Error GoTo Failed
            If InStr(1, childTitle, "Create Warehouse Task", vbTextCompare) > 0 Then
                CodexOpenPRDOWarehouseTaskFollowOn = "CREATE_WT_OPENED|SESSION=" & CStr(ses.id) & "|TITLE=" & CodexClean(childTitle)
                Exit Function
            End If
        Next j
        If CodexElapsed(t0) >= 20# Then Exit Do
    Loop
    CodexOpenPRDOWarehouseTaskFollowOn = "STOP_CREATE_WT_NOT_OPENED"
    Exit Function
Failed:
    CodexOpenPRDOWarehouseTaskFollowOn = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexReadSAPSessions() As String
    Dim gui As Object, app As Object, con As Object, ses As Object
    Dim i As Long, tcode As String, titleText As String, busyText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    CodexReadSAPSessions = "SAP_SESSIONS|COUNT=" & CStr(con.children.count) & vbLf
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": titleText = "": busyText = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction & "")
        titleText = CStr(ses.FindById("wnd[0]", False).text & "")
        busyText = CStr(CBool(ses.busy))
        On Error GoTo Failed
        CodexReadSAPSessions = CodexReadSAPSessions & "SESSION=" & CStr(i) & "|ID=" & CStr(ses.id) & "|TCODE=" & CodexClean(tcode) & "|TITLE=" & CodexClean(titleText) & "|BUSY=" & CodexClean(busyText) & vbLf
    Next i
    Exit Function
Failed:
    CodexReadSAPSessions = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexReadCreateWTRowsLimited() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, r As Long, tcode As String, titleText As String, pathText As String, labelText As Variant, colId As Variant, valueText As String, lineText As String
    Dim rowCount As Long, maxRows As Long
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": titleText = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction & "")
        titleText = CStr(ses.FindById("wnd[0]", False).text & "")
        On Error GoTo Failed
        If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexReadCreateWTRowsLimited = "CREATE_WT_SESSION_NOT_FOUND": Exit Function
    CodexReadCreateWTRowsLimited = "CREATE_WT_ROWS|TITLE=" & CodexClean(titleText) & vbLf
    For Each labelText In Array("WR", "STOCK")
        If labelText = "WR" Then
            pathText = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
        Else
            pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
        End If
        Set grid = ses.FindById(pathText, False)
        If grid Is Nothing Then
            CodexReadCreateWTRowsLimited = CodexReadCreateWTRowsLimited & CStr(labelText) & "_GRID_NOT_FOUND" & vbLf
        Else
            Codex_EnsureGridRendered grid, ses
            rowCount = CLng(grid.rowCount): maxRows = rowCount: If maxRows > 20 Then maxRows = 20
            CodexReadCreateWTRowsLimited = CodexReadCreateWTRowsLimited & CStr(labelText) & "_GRID=" & CStr(grid.id) & "|ROWS=" & CStr(rowCount) & vbLf
            For r = 0 To maxRows - 1
                lineText = CStr(labelText) & "|ROW=" & CStr(r)
                For Each colId In Array("DOCNO", "DOCID", "DOCNO_H", "ITEMNO", "ITEMID", "MATNR", "PRODUCTNO", "CHARG", "QTY", "QUAN", "VSOLM", "VLPLA", "VLTYP", "NLPLA", "WHO", "TANUM", "STATUS", "PROCTY", "CAT", "OWNER", "ENTITLED")
                    valueText = "": On Error Resume Next: valueText = CStr(grid.GetCellValue(r, CStr(colId)) & ""): On Error GoTo Failed
                    If valueText <> "" Then lineText = lineText & "|" & CStr(colId) & "=" & CodexClean(valueText)
                Next colId
                CodexReadCreateWTRowsLimited = CodexReadCreateWTRowsLimited & lineText & vbLf
            Next r
        End If
    Next labelText
    Exit Function
Failed:
    CodexReadCreateWTRowsLimited = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexReadCreateWTWRMap() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object
    Dim i As Long, k As Long, n As Long, tcode As String, titleText As String, c As Variant, valueText As String, titleValue As String, idText As String, tipText As String
    Dim pathText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": titleText = ""
        On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): titleText = CStr(ses.FindById("wnd[0]", False).text & ""): On Error GoTo Failed
        If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexReadCreateWTWRMap = "CREATE_WT_SESSION_NOT_FOUND": Exit Function
    pathText = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Set grid = ses.FindById(pathText, False)
    If grid Is Nothing Then CodexReadCreateWTWRMap = "CREATE_WT_WR_GRID_NOT_FOUND": Exit Function
    CodexReadCreateWTWRMap = "CREATE_WT_WR_MAP|GRID=" & CStr(grid.id) & "|ROWS=" & CStr(grid.rowCount) & vbLf
    n = 0
    For Each c In grid.columnOrder
        valueText = "": titleValue = "": On Error Resume Next: valueText = CStr(grid.GetCellValue(0, CStr(c)) & ""): titleValue = CStr(grid.GetColumnTitles(CStr(c)) & ""): On Error GoTo Failed
        CodexReadCreateWTWRMap = CodexReadCreateWTWRMap & "COL=" & CStr(c) & "|TITLE=" & CodexClean(titleValue) & "|VALUE=" & CodexClean(valueText) & vbLf
        n = n + 1: If n >= 120 Then Exit For
    Next c
    n = 0: On Error Resume Next: n = CLng(grid.ToolbarButtonCount): On Error GoTo Failed
    CodexReadCreateWTWRMap = CodexReadCreateWTWRMap & "TOOLBAR_COUNT=" & CStr(n) & vbLf
    For k = 0 To n - 1
        idText = "": tipText = "": On Error Resume Next: idText = CStr(grid.GetToolbarButtonId(k) & ""): tipText = CStr(grid.GetToolbarButtonTooltip(k) & ""): On Error GoTo Failed
        CodexReadCreateWTWRMap = CodexReadCreateWTWRMap & "BTN=" & CStr(k) & "|ID=" & CodexClean(idText) & "|TIP=" & CodexClean(tipText) & vbLf
    Next k
    Exit Function
Failed:
    CodexReadCreateWTWRMap = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexReadCreateWTToolbars() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, shell As Object
    Dim i As Long, k As Long, tcode As String, titleText As String, pathText As String, labelText As Variant, countValue As Long, idText As String, tipText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": titleText = ""
        On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): titleText = CStr(ses.FindById("wnd[0]", False).text & ""): On Error GoTo Failed
        If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexReadCreateWTToolbars = "CREATE_WT_SESSION_NOT_FOUND": Exit Function
    CodexReadCreateWTToolbars = "CREATE_WT_TOOLBARS|TITLE=" & CodexClean(titleText) & vbLf
    For Each labelText In Array("WR", "STOCK")
        If labelText = "WR" Then
            pathText = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/cntlCC_TB_OIP_DLV/shellcont/shell"
        Else
            pathText = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/cntlCC_TB_OD_AVSTKDLV/shellcont/shell"
        End If
        Set shell = ses.FindById(pathText, False)
        If shell Is Nothing Then
            CodexReadCreateWTToolbars = CodexReadCreateWTToolbars & CStr(labelText) & "_NOT_FOUND" & vbLf
        Else
            countValue = 0: On Error Resume Next: countValue = CLng(shell.buttonCount): On Error GoTo Failed
            CodexReadCreateWTToolbars = CodexReadCreateWTToolbars & CStr(labelText) & "_TOOLBAR=" & CStr(shell.id) & "|COUNT=" & CStr(countValue) & vbLf
            For k = 0 To countValue - 1
                idText = "": tipText = ""
                On Error Resume Next
                idText = CStr(shell.GetButtonId(k) & "")
                tipText = CStr(shell.GetButtonTooltip(k) & "")
                If idText = "" Then idText = CStr(shell.GetToolbarButtonId(k) & "")
                On Error GoTo Failed
                CodexReadCreateWTToolbars = CodexReadCreateWTToolbars & CStr(labelText) & "_BTN=" & CStr(k) & "|ID=" & CodexClean(idText) & "|TIP=" & CodexClean(tipText) & vbLf
            Next k
        End If
    Next labelText
    Exit Function
Failed:
    CodexReadCreateWTToolbars = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexCreateRemainingWTExact(ByVal expectedDoc As String, ByVal expectedItem As String, ByVal expectedMaterial As String, ByVal expectedBatch As String, ByVal expectedDestBin As String, ByVal expectedOpenQty As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object, toolbar As Object, win As Object
    Dim i As Long, j As Long, r As Long, tcode As String, titleText As String, statusText As String, childText As String, t0 As Double
    Dim pathText As String, docText As String, itemText As String, matText As String, batchText As String, destText As String, openQtyText As String, rowNo As Long, matches As Long
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": titleText = ""
        On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): titleText = CStr(ses.FindById("wnd[0]", False).text & ""): On Error GoTo Failed
        If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexCreateRemainingWTExact = "CREATE_WT_SESSION_NOT_FOUND": Exit Function
    pathText = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Set grid = ses.FindById(pathText, False)
    Set toolbar = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/cntlCC_TB_OIP_DLV/shellcont/shell", False)
    If grid Is Nothing Or toolbar Is Nothing Then CodexCreateRemainingWTExact = "CREATE_WT_CONTROLS_NOT_FOUND": Exit Function
    matches = 0: rowNo = -1
    Codex_EnsureGridRendered grid, ses
    For r = 0 To CLng(grid.rowCount) - 1
        docText = CStr(grid.GetCellValue(r, "DLVNO") & ""): itemText = CStr(grid.GetCellValue(r, "ITMNO") & ""): matText = CStr(grid.GetCellValue(r, "MATNR") & ""): batchText = CStr(grid.GetCellValue(r, "CHARG") & ""): destText = CStr(grid.GetCellValue(r, "NLPLA") & ""): openQtyText = CStr(grid.GetCellValue(r, "UI_OPNQTY") & "")
        If StrComp(Trim$(docText), Trim$(expectedDoc), vbTextCompare) = 0 And StrComp(Trim$(itemText), Trim$(expectedItem), vbTextCompare) = 0 And StrComp(Trim$(matText), Trim$(expectedMaterial), vbTextCompare) = 0 And StrComp(Trim$(batchText), Trim$(expectedBatch), vbTextCompare) = 0 And StrComp(Trim$(destText), Trim$(expectedDestBin), vbTextCompare) = 0 And StrComp(Trim$(openQtyText), Trim$(expectedOpenQty), vbTextCompare) = 0 Then matches = matches + 1: rowNo = r
    Next r
    If matches <> 1 Then CodexCreateRemainingWTExact = "STOP_CREATE_WT_MATCH_COUNT=" & CStr(matches): Exit Function
    grid.selectedRows = CStr(rowNo): grid.SetCurrentCell rowNo, "DLVNO"
    toolbar.PressButton "OK_OI_DLV_CREATE_TO"
    t0 = Timer
    Do
        DoEvents
        If CodexElapsed(t0) >= 15# Then Exit Do
    Loop
    statusText = "": On Error Resume Next: statusText = CStr(ses.FindById("wnd[0]/sbar", False).text & ""): On Error GoTo Failed
    CodexCreateRemainingWTExact = "CREATE_WT_SENT|DOC=" & CodexClean(docText) & "|ITEM=" & CodexClean(itemText) & "|MAT=" & CodexClean(matText) & "|BATCH=" & CodexClean(batchText) & "|DEST=" & CodexClean(destText) & "|OPEN_QTY=" & CodexClean(openQtyText) & "|STATUS=" & CodexClean(statusText) & "|WINDOWS=" & CStr(ses.children.count)
    For j = 0 To ses.children.count - 1
        Set win = ses.children.ElementAt(j): childText = "": On Error Resume Next: childText = CStr(win.text & ""): On Error GoTo Failed
        CodexCreateRemainingWTExact = CodexCreateRemainingWTExact & "|W" & CStr(j) & "=" & CodexClean(childText)
    Next j
    Exit Function
Failed:
    CodexCreateRemainingWTExact = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexDumpCreateWTRelevantLimited() As String
    Dim gui As Object, app As Object, con As Object, ses As Object
    Dim i As Long, tcode As String, titleText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": titleText = ""
        On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): titleText = CStr(ses.FindById("wnd[0]", False).text & ""): On Error GoTo Failed
        If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexDumpCreateWTRelevantLimited = "CREATE_WT_SESSION_NOT_FOUND": Exit Function
    CodexDumpCreateWTRelevantLimited = "CREATE_WT_RELEVANT|TITLE=" & CodexClean(titleText) & vbLf & CodexDumpRelevant(ses.FindById("wnd[0]", False), "wnd[0]", 0)
    Exit Function
Failed:
    CodexDumpCreateWTRelevantLimited = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexCreateRemainingWTAndSaveExact(ByVal expectedDoc As String, ByVal expectedItem As String, ByVal expectedMaterial As String, ByVal expectedBatch As String, ByVal expectedDestBin As String, ByVal expectedOpenQty As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object, toolbar As Object, win As Object
    Dim i As Long, j As Long, r As Long, tcode As String, titleText As String, statusText As String, childText As String, t0 As Double
    Dim pathText As String, docText As String, itemText As String, matText As String, batchText As String, destText As String, openQtyText As String, matches As Long, rowNo As Long
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": titleText = ""
        On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): titleText = CStr(ses.FindById("wnd[0]", False).text & ""): On Error GoTo Failed
        If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexCreateRemainingWTAndSaveExact = "CREATE_WT_SESSION_NOT_FOUND": Exit Function
    pathText = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Set grid = ses.FindById(pathText, False)
    Set toolbar = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/cntlCC_TB_OIP_DLV/shellcont/shell", False)
    If grid Is Nothing Or toolbar Is Nothing Then CodexCreateRemainingWTAndSaveExact = "CREATE_WT_CONTROLS_NOT_FOUND": Exit Function
    matches = 0: rowNo = -1
    Codex_EnsureGridRendered grid, ses
    For r = 0 To CLng(grid.rowCount) - 1
        docText = CStr(grid.GetCellValue(r, "DLVNO") & ""): itemText = CStr(grid.GetCellValue(r, "ITMNO") & ""): matText = CStr(grid.GetCellValue(r, "MATNR") & ""): batchText = CStr(grid.GetCellValue(r, "CHARG") & ""): destText = CStr(grid.GetCellValue(r, "NLPLA") & ""): openQtyText = CStr(grid.GetCellValue(r, "UI_OPNQTY") & "")
        If StrComp(Trim$(docText), Trim$(expectedDoc), vbTextCompare) = 0 And StrComp(Trim$(itemText), Trim$(expectedItem), vbTextCompare) = 0 And StrComp(Trim$(matText), Trim$(expectedMaterial), vbTextCompare) = 0 And StrComp(Trim$(batchText), Trim$(expectedBatch), vbTextCompare) = 0 And StrComp(Trim$(destText), Trim$(expectedDestBin), vbTextCompare) = 0 And StrComp(Trim$(openQtyText), Trim$(expectedOpenQty), vbTextCompare) = 0 Then matches = matches + 1: rowNo = r
    Next r
    If matches <> 1 Then CodexCreateRemainingWTAndSaveExact = "STOP_CREATE_WT_MATCH_COUNT=" & CStr(matches): Exit Function
    grid.selectedRows = CStr(rowNo): grid.SetCurrentCell rowNo, "DLVNO"
    toolbar.PressButton "OK_OI_DLV_CREATE_AND_SAVE"
    t0 = Timer
    Do
        DoEvents
        If CodexElapsed(t0) >= 20# Then Exit Do
    Loop
    statusText = "": On Error Resume Next: statusText = CStr(ses.FindById("wnd[0]/sbar", False).text & ""): On Error GoTo Failed
    CodexCreateRemainingWTAndSaveExact = "CREATE_WT_SAVE_SENT|DOC=" & CodexClean(docText) & "|ITEM=" & CodexClean(itemText) & "|MAT=" & CodexClean(matText) & "|BATCH=" & CodexClean(batchText) & "|DEST=" & CodexClean(destText) & "|OPEN_QTY=" & CodexClean(openQtyText) & "|STATUS=" & CodexClean(statusText) & "|WINDOWS=" & CStr(ses.children.count)
    For j = 0 To ses.children.count - 1
        Set win = ses.children.ElementAt(j): childText = "": On Error Resume Next: childText = CStr(win.text & ""): On Error GoTo Failed
        CodexCreateRemainingWTAndSaveExact = CodexCreateRemainingWTAndSaveExact & "|W" & CStr(j) & "=" & CodexClean(childText)
    Next j
    Exit Function
Failed:
    CodexCreateRemainingWTAndSaveExact = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexSaveCreatedWTExact(ByVal expectedDoc As String, ByVal expectedItem As String, ByVal expectedMaterial As String, ByVal expectedBatch As String) As String
    Dim gui As Object, app As Object, con As Object, ses As Object, grid As Object, toolbar As Object, win As Object
    Dim i As Long, j As Long, tcode As String, titleText As String, statusText As String, childText As String, t0 As Double
    Dim pathText As String, docText As String, itemText As String, matText As String, batchText As String, openQtyText As String, stateText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): tcode = "": titleText = ""
        On Error Resume Next: tcode = CStr(ses.Info.Transaction & ""): titleText = CStr(ses.FindById("wnd[0]", False).text & ""): On Error GoTo Failed
        If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexSaveCreatedWTExact = "CREATE_WT_SESSION_NOT_FOUND": Exit Function
    pathText = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Set grid = ses.FindById(pathText, False)
    Set toolbar = ses.FindById("wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/cntlCC_TB_OIP_DLV/shellcont/shell", False)
    If grid Is Nothing Or toolbar Is Nothing Then CodexSaveCreatedWTExact = "CREATE_WT_SAVE_CONTROLS_NOT_FOUND": Exit Function
    If CLng(grid.rowCount) <> 1 Then CodexSaveCreatedWTExact = "STOP_CREATE_WT_ROW_COUNT=" & CStr(grid.rowCount): Exit Function
    docText = CStr(grid.GetCellValue(0, "DLVNO") & ""): itemText = CStr(grid.GetCellValue(0, "ITMNO") & ""): matText = CStr(grid.GetCellValue(0, "MATNR") & ""): batchText = CStr(grid.GetCellValue(0, "CHARG") & ""): openQtyText = CStr(grid.GetCellValue(0, "UI_OPNQTY") & ""): stateText = CStr(grid.GetCellValue(0, "STATUS_TO") & "")
    If StrComp(Trim$(docText), Trim$(expectedDoc), vbTextCompare) <> 0 Or StrComp(Trim$(itemText), Trim$(expectedItem), vbTextCompare) <> 0 Or StrComp(Trim$(matText), Trim$(expectedMaterial), vbTextCompare) <> 0 Or StrComp(Trim$(batchText), Trim$(expectedBatch), vbTextCompare) <> 0 Or Trim$(openQtyText) <> "0" Or InStr(1, stateText, "Green", vbTextCompare) = 0 Then
        CodexSaveCreatedWTExact = "STOP_CREATE_WT_SAVE_PREVERIFY|DOC=" & CodexClean(docText) & "|ITEM=" & CodexClean(itemText) & "|MAT=" & CodexClean(matText) & "|BATCH=" & CodexClean(batchText) & "|OPEN=" & CodexClean(openQtyText) & "|STATE=" & CodexClean(stateText): Exit Function
    End If
    toolbar.PressButton "OK_OI_DLV_SAVE"
    t0 = Timer
    Do
        DoEvents
        If CodexElapsed(t0) >= 20# Then Exit Do
    Loop
    statusText = "": On Error Resume Next: statusText = CStr(ses.FindById("wnd[0]/sbar", False).text & ""): On Error GoTo Failed
    CodexSaveCreatedWTExact = "CREATE_WT_SAVED|DOC=" & CodexClean(docText) & "|ITEM=" & CodexClean(itemText) & "|MAT=" & CodexClean(matText) & "|BATCH=" & CodexClean(batchText) & "|OPEN=" & CodexClean(openQtyText) & "|STATE=" & CodexClean(stateText) & "|STATUS=" & CodexClean(statusText) & "|WINDOWS=" & CStr(ses.children.count)
    For j = 0 To ses.children.count - 1
        Set win = ses.children.ElementAt(j): childText = "": On Error Resume Next: childText = CStr(win.text & ""): On Error GoTo Failed
        CodexSaveCreatedWTExact = CodexSaveCreatedWTExact & "|W" & CStr(j) & "=" & CodexClean(childText)
    Next j
    Exit Function
Failed:
    CodexSaveCreatedWTExact = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexDumpCreateWTWindows() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, win As Object
    Dim i As Long, j As Long, tcode As String, titleText As String, winText As String, props As Variant
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): titleText = "": On Error Resume Next: titleText = CStr(ses.FindById("wnd[0]", False).text & ""): On Error GoTo Failed
        If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexDumpCreateWTWindows = "CREATE_WT_SESSION_NOT_FOUND": Exit Function
    CodexDumpCreateWTWindows = "CREATE_WT_WINDOWS|SESSION=" & CStr(ses.id) & "|COUNT=" & CStr(ses.children.count) & vbLf
    For j = 0 To ses.children.count - 1
        Set win = ses.children.ElementAt(j): winText = "": On Error Resume Next: winText = CStr(win.text & ""): On Error GoTo Failed
        CodexDumpCreateWTWindows = CodexDumpCreateWTWindows & "WIN=" & CStr(j) & "|ID=" & CStr(win.id) & "|TEXT=" & CodexClean(winText) & vbLf
        If j > 0 Then
            props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
            CodexDumpCreateWTWindows = CodexDumpCreateWTWindows & CStr(ses.GetObjectTree(CStr(win.id), props)) & vbLf
        End If
    Next j
    Exit Function
Failed:
    CodexDumpCreateWTWindows = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CodexConfirmCreateWTYes() As String
    Dim gui As Object, app As Object, con As Object, ses As Object, btn As Object, win As Object
    Dim i As Long, j As Long, tcode As String, titleText As String, statusText As String, childText As String, t0 As Double
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i): titleText = "": On Error Resume Next: titleText = CStr(ses.FindById("wnd[0]", False).text & ""): On Error GoTo Failed
        If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then Exit For
    Next i
    If ses Is Nothing Then CodexConfirmCreateWTYes = "CREATE_WT_SESSION_NOT_FOUND": Exit Function
    Set btn = ses.FindById("wnd[1]/usr/btnBUTTON_1", False)
    If btn Is Nothing Then CodexConfirmCreateWTYes = "CREATE_WT_YES_BUTTON_NOT_FOUND": Exit Function
    btn.Press
    t0 = Timer
    Do
        DoEvents
        If CodexElapsed(t0) >= 20# Then Exit Do
    Loop
    statusText = "": On Error Resume Next: statusText = CStr(ses.FindById("wnd[0]/sbar", False).text & ""): On Error GoTo Failed
    CodexConfirmCreateWTYes = "CREATE_WT_CONFIRM_YES_SENT|STATUS=" & CodexClean(statusText) & "|WINDOWS=" & CStr(ses.children.count)
    For j = 0 To ses.children.count - 1
        Set win = ses.children.ElementAt(j): childText = "": On Error Resume Next: childText = CStr(win.text & ""): On Error GoTo Failed
        CodexConfirmCreateWTYes = CodexConfirmCreateWTYes & "|W" & CStr(j) & "=" & CodexClean(childText)
    Next j
    Exit Function
Failed:
    CodexConfirmCreateWTYes = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CodexCollectionCount(ByVal obj As Object) As Long
    On Error Resume Next
    CodexCollectionCount = CLng(obj.count)
    If Err.Number <> 0 Then Err.Clear: CodexCollectionCount = 0
End Function

Private Function CodexCollectionValue(ByVal obj As Object, ByVal indexNo As Long) As String
    Dim v As Variant
    On Error Resume Next
    v = obj.item(indexNo)
    If Err.Number <> 0 Then Err.Clear: v = obj.ElementAt(indexNo)
    CodexCollectionValue = CStr(v & "")
End Function

Private Function CodexWaitSessionIdle(ByVal ses As Object, ByVal timeoutSeconds As Double) As Boolean
    Dim t0 As Double, busy As Boolean
    t0 = Timer
    Do
        DoEvents
        busy = True: On Error Resume Next: busy = CBool(ses.busy): On Error GoTo 0
        If Not busy Then CodexWaitSessionIdle = True: Exit Function
        If CodexElapsed(t0) >= timeoutSeconds Then Exit Do
    Loop
End Function

Private Function CodexWaitIdle(ByVal ses As Object, ByVal timeoutSeconds As Double) As Boolean
    Dim t0 As Double, busy As Boolean
    t0 = Timer
    Do
        DoEvents
        busy = True
        On Error Resume Next
        busy = CBool(ses.busy)
        On Error GoTo 0
        If Not busy Then CodexWaitIdle = True: Exit Function
        If CodexElapsed(t0) >= timeoutSeconds Then Exit Do
    Loop
End Function

Private Function CodexDumpRelevant(ByVal obj As Object, ByVal pathText As String, ByVal depth As Long) As String
    Dim typ As String, idText As String, txt As String, valueText As String, tip As String, ch As Object
    Dim kids As Object, i As Long, count As Long, lineText As String, allText As String
    If obj Is Nothing Or depth > 14 Then Exit Function
    If CodexElapsed(0#) < -1# Then Exit Function

    On Error Resume Next
    typ = CStr(obj.Type & "")
    idText = CStr(obj.id & "")
    txt = CStr(obj.text & "")
    tip = CStr(obj.Tooltip & "")
    valueText = CStr(obj.Value & "")
    Set kids = obj.children
    If Not kids Is Nothing Then count = CLng(kids.count)
    On Error GoTo 0

    allText = UCase$(idText & "|" & txt & "|" & tip)
    If InStr(allText, "QUANTITY") > 0 Or InStr(allText, "QTY") > 0 Or _
       InStr(allText, "QUAN") > 0 Or InStr(allText, "SOURCE BIN") > 0 Or _
       InStr(allText, "STORAGE BIN") > 0 Or InStr(allText, "BATCH") > 0 Or _
       InStr(allText, "SAVE") > 0 Or InStr(allText, "CREATE") > 0 Or _
       InStr(UCase$(typ), "TEXTFIELD") > 0 Or InStr(UCase$(typ), "CTEXTFIELD") > 0 Or _
       InStr(UCase$(typ), "COMBOBOX") > 0 Or InStr(UCase$(typ), "BUTTON") > 0 Or _
       InStr(UCase$(typ), "TAB") > 0 Or InStr(UCase$(typ), "LABEL") > 0 Then
        lineText = "NODE|D=" & CStr(depth) & "|TYPE=" & CodexClean(typ) & _
                   "|ID=" & CodexClean(idText) & "|TEXT=" & CodexClean(txt) & _
                   "|VALUE=" & CodexClean(valueText) & "|TIP=" & CodexClean(tip)
        CodexDumpRelevant = CodexDumpRelevant & lineText & vbLf
    End If

    If kids Is Nothing Then Exit Function
    For i = 0 To count - 1
        Set ch = Nothing
        On Error Resume Next
        Set ch = kids.ElementAt(i)
        If ch Is Nothing Then Set ch = kids.item(i)
        On Error GoTo 0
        If Not ch Is Nothing Then CodexDumpRelevant = CodexDumpRelevant & CodexDumpRelevant(ch, pathText & "/" & CStr(i), depth + 1)
        Set ch = Nothing
        If Len(CodexDumpRelevant) > 60000 Then Exit For
    Next i
End Function

Private Function CodexNormalizeNumber(ByVal valueText As String) As String
    Dim s As String
    s = Replace(Trim$(valueText), " ", "")
    Do While Len(s) > 1 And Left$(s, 1) = "0"
        s = Mid$(s, 2)
    Loop
    CodexNormalizeNumber = UCase$(s)
End Function

Private Function CodexElapsed(ByVal started As Double) As Double
    If started = 0# Then CodexElapsed = 0#: Exit Function
    If Timer >= started Then CodexElapsed = Timer - started Else CodexElapsed = 86400# - started + Timer
End Function

Private Function CodexClean(ByVal valueText As String) As String
    CodexClean = Replace(Replace(Replace(valueText, vbCr, " "), vbLf, " "), "|", "/")
End Function
