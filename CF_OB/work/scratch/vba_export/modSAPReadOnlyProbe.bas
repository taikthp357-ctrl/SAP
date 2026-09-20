Option Explicit

' V1.3B OPEN/READ-only evidence layer.
'
' Each entry point reuses an already-open SAP session, verifies only exact
' locked/candidate IDs, reads evidence, writes a local log, and stops.  No
' navigation, no new session/tab, no SendKeys, no Cancel/Save/Create/CF.

Private Const SP_LOG As String = "SAP_PROBE_LOG"
Private Const SP_PLAN As String = "PLAN"
Private Const SP_LEGS As String = "PLAN_SOURCE_LEGS"
Private Const SP_WH As String = "WH_ACTUAL"
Private Const SP_EPS As Double = 0.000001

Public Sub CF_RunReadOnlyProbes()
    Dim messageText As String
    If Not SAPProbe_Posting(messageText) Then GoTo Failed
    If Not SAPProbe_MONWT(messageText) Then GoTo Failed
    If Not SAPProbe_Release(messageText) Then GoTo Failed
    MsgBox "READ-ONLY PROBES PASS. SAP write actions were not called.", vbInformation, "SAP PROBE"
    Exit Sub
Failed:
    MsgBox messageText, vbExclamation, "SAP PROBE"
End Sub

Public Sub CF_ProbePostingReadOnly()
    Dim messageText As String
    If SAPProbe_Posting(messageText) Then
        MsgBox messageText, vbInformation, "POST PROBE"
    Else
        MsgBox messageText, vbExclamation, "POST PROBE"
    End If
End Sub

Public Sub CF_ProbeMONWTReadOnly()
    Dim messageText As String
    If SAPProbe_MONWT(messageText) Then
        MsgBox messageText, vbInformation, "MON WT PROBE"
    Else
        MsgBox messageText, vbExclamation, "MON WT PROBE"
    End If
End Sub

Public Sub CF_ProbeReleaseReadOnly()
    Dim messageText As String
    If SAPProbe_Release(messageText) Then
        MsgBox messageText, vbInformation, "RELEASE PROBE"
    Else
        MsgBox messageText, vbExclamation, "RELEASE PROBE"
    End If
End Sub

Public Sub SAPProbe_CompileProbe()
    If SAPProbe_SourceKey("M", "B", "C") <> "M|B|C" Then _
        Err.Raise vbObjectError + 9990, , "SAP_PROBE_COMPILE_FAIL: SourceKey"
    If SAPProbe_Elapsed(86399#) < 0# Then _
        Err.Raise vbObjectError + 9991, , "SAP_PROBE_COMPILE_FAIL: elapsed"
End Sub

Private Function SAPProbe_Posting(ByRef messageText As String) As Boolean
    Dim requiredKeys As String, ses As Object, grid As Object
    Dim material As String, binText As String, batchText As String, sourceKey As String
    Dim r As Long, foundRow As Long, matchCount As Long, rows As Long
    Dim msg As String, runID As String, t0 As Double, findMs As Double, readMs As Double
    Dim available As String, totalStock As String, tst As String
    On Error GoTo Failed
    requiredKeys = "POST_GRID_ID,POST_PRODUCT_COL,POST_AVAILABLE_QTY_COL,POST_BATCH_COL,POST_SOURCE_BIN_COL," & _
                   "POST_TST_COL,POST_TOTAL_STOCK_COL,POST_DEST_BIN_COL,POST_QTY_CHG_COL,POST_DEST_BATCH_COL"
    If Not SAPRun_RequireProbeRuntime("POST_PROBE", requiredKeys, messageText, False) Then Exit Function
    t0 = Timer
    If Not SAPProbe_GetFirstSource(material, binText, batchText, messageText) Then Exit Function
    sourceKey = SAPProbe_SourceKey(material, binText, batchText)
    Set ses = SAPProbe_FindExistingSession("/SCWM/POST", messageText)
    If ses Is Nothing Then GoTo Failed
    If Not SAPProbe_VerifyTCode(ses, "/SCWM/POST", messageText) Then GoTo Failed
    Set grid = SAPRun_FindExact(ses, "POST_GRID_ID", msg)
    If grid Is Nothing Then messageText = msg: GoTo Failed
    If Not SAPProbe_RequireGridColumns(grid, "POST_PRODUCT_COL,POST_AVAILABLE_QTY_COL,POST_BATCH_COL,POST_SOURCE_BIN_COL,POST_TST_COL,POST_TOTAL_STOCK_COL,POST_DEST_BIN_COL,POST_QTY_CHG_COL,POST_DEST_BATCH_COL", messageText) Then GoTo Failed
    rows = SAPProbe_RowCount(grid, messageText)
    findMs = SAPProbe_Elapsed(t0) * 1000#
    For r = 0 To rows - 1
        If SAPProbe_Text(grid.GetCellValue(r, ModuleHelper.ToolConfig("POST_PRODUCT_COL", ""))) = material And _
           SAPProbe_Text(grid.GetCellValue(r, ModuleHelper.ToolConfig("POST_SOURCE_BIN_COL", ""))) = binText And _
           SAPProbe_Text(grid.GetCellValue(r, ModuleHelper.ToolConfig("POST_BATCH_COL", ""))) = batchText Then
            matchCount = matchCount + 1
            foundRow = r
        End If
    Next r
    If matchCount <> 1 Then
        messageText = "STOP_TARGET_MISMATCH: POST SourceKey=" & sourceKey & "; matches=" & CStr(matchCount)
        GoTo Failed
    End If
    available = SAPProbe_Text(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("POST_AVAILABLE_QTY_COL", "")))
    totalStock = SAPProbe_Text(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("POST_TOTAL_STOCK_COL", "")))
    tst = SAPProbe_Text(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("POST_TST_COL", "")))
    If UCase$(tst) <> "FT" Then
        messageText = "STOP_TARGET_MISMATCH: POST SourceKey=" & sourceKey & "; TST=" & tst & "; expected FT."
        GoTo Failed
    End If
    ' Read the write-target columns only. No SetCellValue, no Save.
    available = available & "; Total=" & totalStock & "; TST=" & tst & _
                "; DestBin=" & SAPProbe_Text(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("POST_DEST_BIN_COL", ""))) & _
                "; QtyPostChg=" & SAPProbe_Text(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("POST_QTY_CHG_COL", ""))) & _
                "; DestBatch=" & SAPProbe_Text(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("POST_DEST_BATCH_COL", "")))
    readMs = SAPProbe_Elapsed(t0) * 1000# - findMs
    runID = SAPProbe_RunID()
    SAPProbe_Log runID, "POST", "POST_PROBE", sourceKey, "", "", "", _
                 "SourceKey=" & sourceKey, available, "WRITE=FALSE", "PASS", "", findMs, readMs
    messageText = "POST_PROBE_PASS: " & sourceKey & "; " & available & "; WRITE=FALSE"
    SAPProbe_Posting = True
    Exit Function
Failed:
    If messageText = "" Then messageText = "POST_PROBE_FAIL: Err " & CStr(Err.Number) & ": " & Err.Description
    SAPProbe_Log SAPProbe_RunID(), "POST", "POST_PROBE", sourceKey, "", "", "", "", "", "WRITE=FALSE", "FAIL", messageText, 0#, 0#
End Function

Private Function SAPProbe_MONWT(ByRef messageText As String) As Boolean
    Dim requiredKeys As String, ses As Object, grid As Object, msg As String
    Dim ob As String, doc As String, whNo As String, wt As String, material As String, binText As String, batchText As String
    Dim expectedQty As Double, r As Long, rows As Long, foundRow As Long, matchCount As Long
    Dim liveDoc As String, liveWH As String, liveMaterial As String, liveBin As String, liveBatch As String, liveStatus As String, liveQty As Double
    Dim t0 As Double, findMs As Double, readMs As Double, runID As String
    On Error GoTo Failed
    requiredKeys = "MON_WT_GRID_ID,MON_WT_DOC_COL,MON_WT_WHO_COL,MON_WT_COL,MON_WT_MATERIAL_COL,MON_WT_BATCH_COL,MON_WT_BIN_COL,MON_WT_QTY_COL,MON_WT_STATUS_COL"
    If Not SAPRun_RequireProbeRuntime("MON_WT_PROBE", requiredKeys, messageText, False) Then Exit Function
    t0 = Timer
    If Not SAPProbe_GetFirstWHActual(ob, doc, whNo, wt, material, expectedQty, binText, batchText, messageText) Then Exit Function
    Set ses = SAPProbe_FindExistingSession("/SCWM/MON", messageText)
    If ses Is Nothing Then GoTo Failed
    If Not SAPProbe_VerifyTCode(ses, "/SCWM/MON", messageText) Then GoTo Failed
    Set grid = SAPRun_FindExact(ses, "MON_WT_GRID_ID", msg)
    If grid Is Nothing Then messageText = msg: GoTo Failed
    If Not SAPProbe_RequireGridColumns(grid, "MON_WT_DOC_COL,MON_WT_WHO_COL,MON_WT_COL,MON_WT_MATERIAL_COL,MON_WT_BATCH_COL,MON_WT_BIN_COL,MON_WT_QTY_COL,MON_WT_STATUS_COL", messageText) Then GoTo Failed
    rows = SAPProbe_RowCount(grid, messageText)
    For r = 0 To rows - 1
        If SAPProbe_Text(grid.GetCellValue(r, ModuleHelper.ToolConfig("MON_WT_COL", ""))) = wt Then
            matchCount = matchCount + 1
            foundRow = r
        End If
    Next r
    findMs = SAPProbe_Elapsed(t0) * 1000#
    If matchCount <> 1 Then
        messageText = "STOP_TARGET_MISMATCH: MON WT=" & wt & "; matches=" & CStr(matchCount)
        GoTo Failed
    End If
    liveDoc = SAPProbe_Text(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("MON_WT_DOC_COL", "")))
    liveWH = SAPProbe_Text(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("MON_WT_WHO_COL", "")))
    liveMaterial = SAPProbe_Text(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("MON_WT_MATERIAL_COL", "")))
    liveBatch = SAPProbe_Text(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("MON_WT_BATCH_COL", "")))
    liveBin = SAPProbe_Text(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("MON_WT_BIN_COL", "")))
    liveQty = SAPProbe_Number(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("MON_WT_QTY_COL", "")))
    liveStatus = SAPProbe_Text(grid.GetCellValue(foundRow, ModuleHelper.ToolConfig("MON_WT_STATUS_COL", "")))
    If liveDoc <> doc Or liveWH <> whNo Or liveMaterial <> material Or liveBatch <> batchText Or liveBin <> binText Or Abs(liveQty - expectedQty) > SP_EPS Then
        messageText = "STOP_TARGET_MISMATCH: WT=" & wt & "; expected Doc/WH/Mat/Qty/Bin/Batch differs from live row."
        GoTo Failed
    End If
    readMs = SAPProbe_Elapsed(t0) * 1000# - findMs
    runID = SAPProbe_RunID()
    SAPProbe_Log runID, "MON", "MON_WT_PROBE", wt, ob, doc, wt, _
                 "OB=" & ob & ";WT=" & wt & ";Qty=" & CStr(expectedQty), _
                 "Doc=" & liveDoc & ";WH=" & liveWH & ";Mat=" & liveMaterial & ";Qty=" & CStr(liveQty) & ";Bin=" & liveBin & ";Batch=" & liveBatch & ";Status=" & liveStatus, _
                 "WRITE=FALSE; FOCUS=NOT_CALLED", "PASS", "", findMs, readMs
    messageText = "MON_WT_PROBE_PASS: WT=" & wt & "; status=" & liveStatus & "; WRITE=FALSE"
    SAPProbe_MONWT = True
    Exit Function
Failed:
    If messageText = "" Then messageText = "MON_WT_PROBE_FAIL: Err " & CStr(Err.Number) & ": " & Err.Description
    SAPProbe_Log SAPProbe_RunID(), "MON", "MON_WT_PROBE", wt, ob, doc, wt, "", "", "WRITE=FALSE", "FAIL", messageText, 0#, 0#
End Function

Private Function SAPProbe_Release(ByRef messageText As String) As Boolean
    Dim requiredKeys As String, ws As Worksheet, r As Long, lastRow As Long, sourceKey As String
    Dim material As String, binText As String, batchText As String, ses As Object, grid As Object, msg As String
    Dim rows As Long, rawRows As Long, validRows As Long, expectedRows As Long, expectedText As String
    Dim runID As String, t0 As Double, readMs As Double, hasPlanSource As Boolean
    On Error GoTo Failed
    requiredKeys = "RELEASE_MON_RESULT_GRID_ID,RELEASE_MON_RESULT_COUNT_ID,RELEASE_MON_RESULT_COUNT_PATTERN,RELEASE_MON_OB_COL,RELEASE_MON_WH_COL,RELEASE_MON_WT_COL,RELEASE_MON_MATERIAL_COL,RELEASE_MON_BATCH_COL,RELEASE_MON_BIN_COL,RELEASE_MON_QTY_COL,RELEASE_MON_STATUS_COL"
    If Not SAPRun_RequireProbeRuntime("RELEASE_PROBE", requiredKeys, messageText, False) Then Exit Function
    t0 = Timer
    Set ws = ThisWorkbook.Worksheets(SP_LEGS)
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If UCase$(SAPProbe_Text(ws.Cells(r, 7).Value2)) = "NEED_RELEASE" Then
            sourceKey = SAPProbe_Text(ws.Cells(r, 3).Value2)
            Exit For
        End If
    Next r
    If sourceKey = "" Then
        ' A normal pilot often has no release requirement.  In that case the
        ' probe still proves the mapped context/count/grid are readable, but
        ' does not claim a holder match.
        sourceKey = SAPProbe_FirstAnySource(ws)
        If sourceKey = "" Then
            sourceKey = SAPProbe_FirstPostingSource()
            If sourceKey = "" Then
                messageText = "STOP_RELEASE_PROBE_TARGET_REQUIRED: no SourceKey exists in PLAN_SOURCE_LEGS or POSTING."
                GoTo Failed
            End If
        End If
    Else
        hasPlanSource = True
    End If
    material = SAPProbe_KeyPart(sourceKey, 1): binText = SAPProbe_KeyPart(sourceKey, 2): batchText = SAPProbe_KeyPart(sourceKey, 3)
    Set ses = SAPProbe_FindExistingSession("/SCWM/MON", messageText)
    If ses Is Nothing Then GoTo Failed
    If Not SAPProbe_VerifyTCode(ses, "/SCWM/MON", messageText) Then GoTo Failed
    ' Safety contract: this probe deliberately does not clear/set filters or press Execute.
    Set grid = SAPRun_FindExact(ses, "RELEASE_MON_RESULT_GRID_ID", msg)
    If grid Is Nothing Then messageText = msg: GoTo Failed
    If Not SAPProbe_RequireGridColumns(grid, "RELEASE_MON_OB_COL,RELEASE_MON_WH_COL,RELEASE_MON_WT_COL,RELEASE_MON_MATERIAL_COL,RELEASE_MON_BATCH_COL,RELEASE_MON_BIN_COL,RELEASE_MON_QTY_COL,RELEASE_MON_STATUS_COL", messageText) Then GoTo Failed
    expectedText = SAPProbe_ControlText(SAPRun_FindExact(ses, "RELEASE_MON_RESULT_COUNT_ID", msg))
    If Not SAPProbe_ParseCount(expectedText, ModuleHelper.ToolConfig("RELEASE_MON_RESULT_COUNT_PATTERN", ""), expectedRows) Then
        messageText = "STOP_RELEASE_CAPTURE_UNPROVEN: result count control/pattern is not readable."
        GoTo Failed
    End If
    rows = SAPProbe_RowCount(grid, messageText)
    If rows <> expectedRows Then
        messageText = "STOP_RELEASE_CAPTURE_UNPROVEN: ExpectedHolderRows=" & CStr(expectedRows) & "; CapturedRawHolderRows=" & CStr(rows)
        GoTo Failed
    End If
    rawRows = rows
    If hasPlanSource Then
        For r = 0 To rows - 1
            If SAPProbe_Text(grid.GetCellValue(r, ModuleHelper.ToolConfig("RELEASE_MON_MATERIAL_COL", ""))) = material And _
               SAPProbe_Text(grid.GetCellValue(r, ModuleHelper.ToolConfig("RELEASE_MON_BIN_COL", ""))) = binText And _
               SAPProbe_Text(grid.GetCellValue(r, ModuleHelper.ToolConfig("RELEASE_MON_BATCH_COL", ""))) = batchText Then
                validRows = validRows + 1
            End If
        Next r
    End If
    readMs = SAPProbe_Elapsed(t0) * 1000#
    runID = SAPProbe_RunID()
    SAPProbe_Log runID, "RELEASE", "RELEASE_PROBE", sourceKey, "", "", "", _
                 "SourceKey=" & sourceKey & ";NeedRelease=" & CStr(hasPlanSource), "Expected=" & CStr(expectedRows) & ";CapturedRaw=" & CStr(rawRows) & ";Matching=" & CStr(validRows), _
                 "WRITE=FALSE; FILTER/EXECUTE=NOT_CALLED", "PASS", "", 0#, readMs
    messageText = "RELEASE_PROBE_PASS: " & sourceKey & "; Expected=" & CStr(expectedRows) & "; CapturedRaw=" & CStr(rawRows) & "; WRITE=FALSE"
    SAPProbe_Release = True
    Exit Function
Failed:
    If messageText = "" Then messageText = "RELEASE_PROBE_FAIL: Err " & CStr(Err.Number) & ": " & Err.Description
    SAPProbe_Log SAPProbe_RunID(), "RELEASE", "RELEASE_PROBE", sourceKey, "", "", "", "", "", "WRITE=FALSE", "FAIL", messageText, 0#, 0#
End Function

Private Function SAPProbe_FirstAnySource(ByVal ws As Worksheet) As String
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        SAPProbe_FirstAnySource = SAPProbe_Text(ws.Cells(r, 3).Value2)
        If SAPProbe_FirstAnySource <> "" Then Exit Function
    Next r
End Function

Private Function SAPProbe_FirstPostingSource() As String
    Dim ws As Worksheet, r As Long, lastRow As Long
    On Error GoTo Done
    Set ws = ThisWorkbook.Worksheets("POSTING")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If UCase$(SAPProbe_Text(ws.Cells(r, 5).Value2)) = "FT" Then
            SAPProbe_FirstPostingSource = SAPProbe_SourceKey(SAPProbe_Text(ws.Cells(r, 1).Value2), _
                                                               SAPProbe_Text(ws.Cells(r, 4).Value2), _
                                                               SAPProbe_Text(ws.Cells(r, 3).Value2))
            If SAPProbe_FirstPostingSource <> "||" Then Exit Function
        End If
    Next r
Done:
End Function

Private Function SAPProbe_GetFirstSource(ByRef material As String, ByRef binText As String, ByRef batchText As String, ByRef messageText As String) As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(SP_LEGS)
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If SAPProbe_Text(ws.Cells(r, 3).Value2) <> "" Then
            material = SAPProbe_KeyPart(SAPProbe_Text(ws.Cells(r, 3).Value2), 1)
            binText = SAPProbe_KeyPart(SAPProbe_Text(ws.Cells(r, 3).Value2), 2)
            batchText = SAPProbe_KeyPart(SAPProbe_Text(ws.Cells(r, 3).Value2), 3)
            SAPProbe_GetFirstSource = (material <> "" And binText <> "" And batchText <> "")
            Exit Function
        End If
    Next r
    Set ws = ThisWorkbook.Worksheets("POSTING")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If UCase$(SAPProbe_Text(ws.Cells(r, 5).Value2)) = "FT" Then
            material = SAPProbe_Text(ws.Cells(r, 1).Value2)
            binText = SAPProbe_Text(ws.Cells(r, 4).Value2)
            batchText = SAPProbe_Text(ws.Cells(r, 3).Value2)
            SAPProbe_GetFirstSource = (material <> "" And binText <> "" And batchText <> "")
            If SAPProbe_GetFirstSource Then Exit Function
        End If
    Next r
    messageText = "STOP_POST_PROBE_SOURCE_REQUIRED: build PLAN_SOURCE_LEGS or load a FT row in POSTING first."
    Exit Function
Failed:
    messageText = "STOP_POST_PROBE_SOURCE_REQUIRED: PLAN_SOURCE_LEGS is missing."
End Function

Private Function SAPProbe_GetFirstWHActual(ByRef ob As String, ByRef doc As String, ByRef whNo As String, _
                                            ByRef wt As String, ByRef material As String, ByRef qty As Double, _
                                            ByRef binText As String, ByRef batchText As String, ByRef messageText As String) As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(SP_WH)
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 5 To lastRow
        wt = SAPProbe_Text(ws.Cells(r, 5).Value2)
        If wt <> "" Then
            ob = SAPProbe_Text(ws.Cells(r, 1).Value2)
            doc = SAPProbe_Text(ws.Cells(r, 2).Value2)
            whNo = SAPProbe_Text(ws.Cells(r, 4).Value2)
            material = SAPProbe_Text(ws.Cells(r, 6).Value2)
            qty = SAPProbe_Number(ws.Cells(r, 7).Value2)
            binText = SAPProbe_Text(ws.Cells(r, 8).Value2)
            batchText = SAPProbe_Text(ws.Cells(r, 9).Value2)
            SAPProbe_GetFirstWHActual = (ob <> "" And doc <> "" And whNo <> "" And material <> "" And binText <> "" And batchText <> "")
            Exit Function
        End If
    Next r
    messageText = "STOP_MON_PROBE_TARGET_REQUIRED: WH_ACTUAL has no loaded WT row."
    Exit Function
Failed:
    messageText = "STOP_MON_PROBE_TARGET_REQUIRED: WH_ACTUAL is missing."
End Function

Private Function SAPProbe_FindExistingSession(ByVal expectedTcode As String, ByRef messageText As String) As Object
    Dim guiAuto As Object, app As Object, conns As Object, conn As Object, sessions As Object, ses As Object
    Dim ci As Long, si As Long, cn As Long, sn As Long
    On Error GoTo Failed
    messageText = ""
    Set guiAuto = GetObject("SAPGUI")
    If guiAuto Is Nothing Then GoTo NotFound
    Set app = guiAuto.GetScriptingEngine
    If app Is Nothing Then GoTo NotFound
    Set conns = app.children
    cn = CLng(conns.count)
    For ci = 0 To cn - 1
        Set conn = SAPProbe_ElementAt(conns, ci)
        If Not conn Is Nothing Then
            Set sessions = conn.children
            sn = CLng(sessions.count)
            For si = 0 To sn - 1
                Set ses = SAPProbe_ElementAt(sessions, si)
                If Not ses Is Nothing Then
                    If Not SAPProbe_SessionBusy(ses) Then
                        If SAPProbe_TCodeIs(ses, expectedTcode) Then
                            Set SAPProbe_FindExistingSession = ses
                            Exit Function
                        End If
                    End If
                End If
            Next si
        End If
    Next ci
NotFound:
    messageText = "STOP_SESSION_NOT_FOUND: open an idle " & expectedTcode & " session; probe will not create or navigate a SAP tab."
    Exit Function
Failed:
    messageText = "STOP_SESSION_NOT_FOUND: SAP session scan failed. Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function SAPProbe_VerifyTCode(ByVal ses As Object, ByVal expectedTcode As String, ByRef messageText As String) As Boolean
    If SAPProbe_TCodeIs(ses, expectedTcode) Then
        SAPProbe_VerifyTCode = True
    Else
        messageText = "STOP_CONTEXT_MISMATCH: expected " & expectedTcode & "; current session is different."
    End If
End Function

Private Function SAPProbe_TCodeIs(ByVal ses As Object, ByVal expectedTcode As String) As Boolean
    Dim actualText As String, expectedText As String
    On Error Resume Next
    actualText = UCase$(Trim$(CStr(ses.Info.Transaction & "")))
    On Error GoTo 0
    expectedText = UCase$(Trim$(expectedTcode))
    If Left$(actualText, 1) = "/" Then actualText = Mid$(actualText, 2)
    If Left$(expectedText, 1) = "/" Then expectedText = Mid$(expectedText, 2)
    SAPProbe_TCodeIs = (actualText = expectedText)
End Function

Private Function SAPProbe_SessionBusy(ByVal ses As Object) As Boolean
    On Error Resume Next
    SAPProbe_SessionBusy = CBool(ses.busy)
    If Err.Number <> 0 Then SAPProbe_SessionBusy = True: Err.Clear
    On Error GoTo 0
End Function

Private Function SAPProbe_ElementAt(ByVal items As Object, ByVal indexNo As Long) As Object
    If items Is Nothing Then Exit Function
    On Error Resume Next
    Set SAPProbe_ElementAt = items.ElementAt(indexNo)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Private Function SAPProbe_RequireGridColumns(ByVal grid As Object, ByVal keysCsv As String, ByRef messageText As String) As Boolean
    Dim keyText As Variant, columnID As String
    For Each keyText In Split(keysCsv, ",")
        columnID = Trim$(ModuleHelper.ToolConfig(CStr(keyText), ""))
        If columnID = "" Or Not SAPProbe_HasColumn(grid, columnID) Then
            messageText = "STOP_CONTROL_NOT_FOUND: exact technical column " & CStr(keyText)
            Exit Function
        End If
    Next keyText
    SAPProbe_RequireGridColumns = True
End Function

Private Function SAPProbe_HasColumn(ByVal grid As Object, ByVal columnID As String) As Boolean
    Dim order As Object, i As Long, n As Long, v As Variant
    On Error Resume Next
    Set order = grid.columnOrder
    n = CLng(order.count)
    On Error GoTo 0
    If order Is Nothing Then Exit Function
    For i = 0 To n - 1
        On Error Resume Next
        v = order.ElementAt(i)
        On Error GoTo 0
        If StrComp(Trim$(CStr(v & "")), columnID, vbTextCompare) = 0 Then
            SAPProbe_HasColumn = True
            Exit Function
        End If
    Next i
End Function

Private Function SAPProbe_RowCount(ByVal grid As Object, ByRef messageText As String) As Long
    On Error Resume Next
    SAPProbe_RowCount = CLng(grid.rowCount)
    If Err.Number <> 0 Then
        messageText = "STOP_CONTROL_NOT_FOUND: mapped grid RowCount could not be read."
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function SAPProbe_ControlText(ByVal ctl As Object) As String
    If ctl Is Nothing Then Exit Function
    On Error Resume Next
    SAPProbe_ControlText = Trim$(CStr(ctl.text & ""))
    If SAPProbe_ControlText = "" Then SAPProbe_ControlText = Trim$(CStr(ctl.key & ""))
    On Error GoTo 0
End Function

Private Function SAPProbe_ParseCount(ByVal rawText As String, ByVal patternText As String, ByRef resultCount As Long) As Boolean
    Dim p As Long, tailText As String, digits As String, i As Long, ch As String
    p = InStr(1, rawText, patternText, vbTextCompare)
    If p = 0 Then Exit Function
    tailText = Mid$(rawText, p + Len(patternText))
    For i = 1 To Len(tailText)
        ch = Mid$(tailText, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch ElseIf digits <> "" Then Exit For
    Next i
    If digits = "" Then Exit Function
    resultCount = CLng(digits)
    SAPProbe_ParseCount = True
End Function

Private Function SAPProbe_RunID() As String
    SAPProbe_RunID = "PROBE_" & Format$(Now, "yyyymmdd_hhnnss")
End Function

Private Sub SAPProbe_Log(ByVal runID As String, ByVal probeName As String, ByVal actionName As String, _
                         ByVal expectedKey As String, ByVal ob As String, ByVal doc As String, ByVal wt As String, _
                         ByVal expectedText As String, ByVal liveText As String, ByVal targetText As String, _
                         ByVal resultText As String, ByVal errorText As String, ByVal findMs As Double, ByVal readMs As Double)
    Dim ws As Worksheet, r As Long, totalMs As Double
    Set ws = SAPProbe_EnsureLogSheet()
    r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    If r < 2 Then r = 2
    totalMs = findMs + readMs
    ws.Cells(r, 1).Value2 = Now
    ws.Cells(r, 2).Value2 = runID
    ws.Cells(r, 3).Value2 = probeName
    ws.Cells(r, 4).Value2 = actionName
    ws.Cells(r, 5).Value2 = SAPProbe_SessionIDText()
    ws.Cells(r, 6).Value2 = expectedKey
    ws.Cells(r, 7).Value2 = ob
    ws.Cells(r, 8).Value2 = doc
    ws.Cells(r, 9).Value2 = wt
    ws.Cells(r, 10).Value2 = expectedText
    ws.Cells(r, 11).Value2 = liveText
    ws.Cells(r, 12).Value2 = targetText
    ws.Cells(r, 13).Value2 = findMs
    ws.Cells(r, 14).Value2 = readMs
    ws.Cells(r, 15).Value2 = resultText
    ws.Cells(r, 16).Value2 = errorText
    ws.Cells(r, 17).Value2 = totalMs
    RT_Log "SAP_PROBE_" & probeName, ob, doc, resultText, actionName & "; " & expectedKey & "; " & targetText & "; " & errorText
End Sub

Private Function SAPProbe_EnsureLogSheet() As Worksheet
    On Error Resume Next
    Set SAPProbe_EnsureLogSheet = ThisWorkbook.Worksheets(SP_LOG)
    On Error GoTo 0
    If SAPProbe_EnsureLogSheet Is Nothing Then
        Set SAPProbe_EnsureLogSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        SAPProbe_EnsureLogSheet.Name = SP_LOG
    End If
    If SAPProbe_Text(SAPProbe_EnsureLogSheet.Cells(1, 1).Value2) <> "Time" Then
        SAPProbe_EnsureLogSheet.Range("A1:Q1").Value = Array("Time", "RunID", "Probe", "Action", "SessionID", "ExpectedKey", "OB", "Document", "WT", "Expected", "Live", "Target", "FindMs", "ReadMs", "Result", "ErrorStage", "TotalMs")
        SAPProbe_EnsureLogSheet.rows(1).Font.Bold = True
    End If
End Function

Private Function SAPProbe_SessionIDText() As String
    SAPProbe_SessionIDText = "EXISTING_ONLY"
End Function

Private Function SAPProbe_SourceKey(ByVal material As String, ByVal binText As String, ByVal batchText As String) As String
    SAPProbe_SourceKey = Trim$(material) & "|" & Trim$(binText) & "|" & Trim$(batchText)
End Function

Private Function SAPProbe_KeyPart(ByVal keyText As String, ByVal partNo As Long) As String
    Dim p As Variant
    p = Split(keyText, "|")
    If partNo >= 1 And partNo <= UBound(p) + 1 Then SAPProbe_KeyPart = Trim$(CStr(p(partNo - 1)))
End Function

Private Function SAPProbe_Text(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    SAPProbe_Text = Trim$(CStr(valueIn & ""))
End Function

Private Function SAPProbe_Number(ByVal valueIn As Variant) As Double
    Dim s As String
    If IsNumeric(valueIn) Then SAPProbe_Number = CDbl(valueIn): Exit Function
    s = Replace(Replace(Replace(SAPProbe_Text(valueIn), ",", ""), ".", ""), " ", "")
    If IsNumeric(s) Then SAPProbe_Number = CDbl(s)
End Function

Private Function SAPProbe_Elapsed(ByVal started As Double) As Double
    If Timer >= started Then
        SAPProbe_Elapsed = Timer - started
    Else
        SAPProbe_Elapsed = 86400# - started + Timer
    End If
End Function
