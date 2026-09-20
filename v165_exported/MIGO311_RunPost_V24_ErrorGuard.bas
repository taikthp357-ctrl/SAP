Attribute VB_Name = "MIGO311_RunPost_V24_ErrorGuard"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub ATT_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub ATT_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

Public LastMigoRunSucceeded As Boolean
Public LastMigoProcessedCount As Long
Public LastMigoPostedCount As Long
Public LastMigoRunMessage As String
Public LastMigoTotalSeconds As Double
Private mToolBusy As Boolean
Private mStop As Boolean
Private mSourceFolder As String
Public Const ATT_FAST_ENGINE_VERSION As String = "1.0.0-OFFLINE_READY"
Private mFastInput As Variant
Private mFastDocs As Object
Private mFastFiles As Collection
Private mFastFileIndex As Object
Private mFastHistory As Object
Private mFastRunId As String
Private mFastLogRow As Long
Private mFastSapUser As String
Private mFastGosProfile As String
Private mFastScanSeconds As Double
Private mFastExcelSeconds As Double
Private mFastSapSeconds As Double
Private mLastGosMethod As String
Private mLastGosControlID As String
Private mLastGosControlType As String
Private mLastGosResolved As Boolean
Private mLastGosViewSelector As String
Private mLastGosCreateSelector As String
Private mLastGosResult As Object
Private mLastGosIndicatorID As String
Private mGosDiagnosticRunId As String
Private Const SH_INPUT As String = "INPUT"
Private Const SH_SCAN As String = ""
Private Const SH_CONFIG As String = "CONFIG"
Private Const SH_HISTORY As String = "ATT_HISTORY"
Private Const SH_LOG As String = "LOG"
Private Const SH_RAW As String = "OCR_RAW"
Private Const FIRST_ROW As Long = 7
Private Const C_MAT As Long = 1
Private Const C_DOC As Long = 2
Private Const C_PO As Long = 3
Private Const C_TICKET As Long = 4
Private Const C_BBBG As Long = 5
Private Const C_VAT As Long = 6
Private Const C_OCR_PO As Long = 7
Private Const C_OCR_MAT_BBBG As Long = 8
Private Const C_OCR_TICKET As Long = 9
Private Const C_OCR_MAT_VAT As Long = 10
Private Const C_CHECK As Long = 11
Private Const C_HEAD As Long = 12
Private Const C_ATT_BBBG As Long = 13
Private Const C_ATT_VAT As Long = 14
Private Const C_ATT As Long = 15
Private Const C_NOTE As Long = 16

Public Sub CAPTURE_GOS_POPUP()
    Dim ses As Object, popup As Object, titles As Object, errText As String, doc As String, t0 As Double
    If Not ToolTryBegin("CAPTURE GOS POPUP") Then Exit Sub
    t0 = Timer: mFastRunId = "GOS_CAPTURE_" & Format$(Now, "yyyymmdd_hhnnss")
    On Error GoTo EH
    Set ses = ModuleSAP.ATT_GetMigoSession()
    If ses Is Nothing Then Err.Raise vbObjectError + 7931, , "Khong tim thay SAP MIGO session."
    FAST_LoadGosProfile ses: doc = GOS_CurrentDoc(ses)
    Set popup = ATT_WaitSapWindow(ses, 1, 1200)
    If popup Is Nothing Then
        GOS_LogWindowState ses, doc, "CAPTURE_MANUAL_POPUP", "MANUAL", "", "POPUP_NOT_OPEN"
        GOS_LogKnownControls ses, doc, "CAPTURE_MANUAL_POPUP", "MANUAL", "", "POPUP_NOT_OPEN"
        MsgBox "Khong thay Attachment popup. Hay bam Display Attachment bang tay trong SAP, giu popup mo, sau do bam CAPTURE GOS POPUP.", vbExclamation
        GoTo SafeExit
    End If
    GOS_LogSapMap ses, popup, doc, "CAPTURE_MANUAL_POPUP", "MANUAL", "wnd[1]", "MANUAL_POPUP_OPEN"
    Set titles = GOS_ReadOpenPopupTitles(popup, errText)
    If errText <> "" Then
        GOS_LogSapMap ses, popup, doc, "CAPTURE_MANUAL_POPUP", "MANUAL", "wnd[1]", "MANUAL_POPUP_LIST_FAIL"
        GOS_LogEvent "GOS_MANUAL_CAPTURE", "FAIL", doc, "CAPTURE_MANUAL_POPUP", "MANUAL", "wnd[1]", "wnd[1]", "", "LIST_NOT_FOUND", errText, Timer - t0, ses
        MsgBox "Da capture popup nhung chua doc duoc list. Xem LOG GOS_SAP_MAP.", vbExclamation
    Else
        GOS_LogEvent "GOS_MANUAL_CAPTURE", "SUCCESS", doc, "CAPTURE_MANUAL_POPUP", "MANUAL", "wnd[1]", "wnd[1]", GOS_TitleSummary(titles), "", "", Timer - t0, ses
        MsgBox "CAPTURE PASS: doc duoc " & CStr(titles.count) & " attachment. Khong dong popup va khong attach file.", vbInformation
    End If
SafeExit:
    ToolEnd
    Exit Sub
EH:
    GOS_LogEvent "GOS_MANUAL_CAPTURE", "FAIL", doc, "CAPTURE_MANUAL_POPUP", "MANUAL", "", "", "", "CAPTURE_EXCEPTION", Err.description, Timer - t0, ses
    MsgBox "CAPTURE GOS POPUP loi: " & Err.description, vbCritical
    Resume SafeExit
End Sub

Private Function GOS_ReadOpenPopupTitles(ByVal popup As Object, ByRef errText As String) As Object
    Dim d As Object, grid As Object, tbl As Object, r As Long, title As String, listStage As String, probe As String
    Set d = CreateObject("Scripting.Dictionary"): errText = ""
    If Not GOS_WaitForAttachmentList(popup, grid, tbl, listStage, probe) Then
        errText = "LIST_NOT_FOUND: " & probe
        Set GOS_ReadOpenPopupTitles = d
        Exit Function
    End If
    If Not grid Is Nothing Then
        For r = 0 To CLng(grid.rowCount) - 1
            title = ATT_GridTitle(grid, r)
            If title <> "" Then d(ATT_NormTitle(title)) = title
        Next r
    ElseIf Not tbl Is Nothing Then
        For r = 0 To CLng(tbl.rowCount) - 1
            title = ATT_TableTitle(tbl, r)
            If title <> "" Then d(ATT_NormTitle(title)) = title
        Next r
    End If
    Set GOS_ReadOpenPopupTitles = d
End Function

Private Function GOS_TryDirectDisplayButton(ByVal ses As Object, ByRef resultText As String) As Boolean
    Dim roots As Variant, rootId As Variant, root As Object, hit As Object, detail As String, buttonId As String
    Dim errNo As Long, errDesc As String
    roots = Array("wnd[0]/titl", "wnd[0]/titl/shellcont", "wnd[0]/titl/shellcont/shell", "wnd[0]/tbar[1]", "wnd[0]/tbar[0]")
    For Each rootId In roots
        Set root = Nothing
        On Error Resume Next: Set root = ses.findById(CStr(rootId), False): On Error GoTo 0
        If Not root Is Nothing Then
            detail = "": Set hit = GOS_FindDisplayAttachmentNode(root, 0, detail)
            If Not hit Is Nothing Then
                buttonId = GOS_DetailValue(detail, "ID=")
                On Error Resume Next
                Err.Clear
                If buttonId <> "" Then
                    CallByName hit, "PressButton", VbMethod, buttonId
                Else
                    CallByName hit, "Press", VbMethod
                End If
                errNo = Err.Number: errDesc = Err.description
                Err.Clear: On Error GoTo 0
                If errNo = 0 Then
                    resultText = "DIRECT_PRESS_OK|ROOT=" & CStr(rootId) & "|" & detail
                    GOS_TryDirectDisplayButton = True
                Else
                    resultText = "DIRECT_PRESS_FAIL=" & CStr(errNo) & ":" & errDesc & "|ROOT=" & CStr(rootId) & "|" & detail
                End If
                Exit Function
            End If
        End If
    Next rootId
    resultText = "DIRECT_BUTTON_NOT_FOUND"
End Function

Private Function GOS_DetailValue(ByVal detail As String, ByVal token As String) As String
    Dim p As Long, e As Long
    p = InStr(1, detail, token, vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(token): e = InStr(p, detail, "|")
    If e = 0 Then e = Len(detail) + 1
    GOS_DetailValue = Trim$(Mid$(detail, p, e - p))
End Function

Private Sub GOS_LogKnownControls(ByVal ses As Object, ByVal doc As String, ByVal action As String, ByVal method As String, ByVal controlId As String, ByVal mapStage As String)
    Dim ids As Variant, oneId As Variant, obj As Object, output As String
    ids = Array("wnd[0]", "wnd[0]/titl", "wnd[0]/titl/shellcont", "wnd[0]/titl/shellcont/shell", "wnd[0]/tbar[0]", "wnd[0]/tbar[1]", "wnd[0]/usr", "wnd[0]/sbar")
    For Each oneId In ids
        Set obj = Nothing
        On Error Resume Next: Set obj = ses.findById(CStr(oneId), False): On Error GoTo 0
        If Not obj Is Nothing Then output = output & "|ROOT=" & CStr(oneId) & ";" & GOS_ControlMap(obj)
    Next oneId
    GOS_LogEvent "GOS_KNOWN_MAP", "INFO", doc, action, method, controlId, "", "MAP_STAGE=" & mapStage & "|" & GOS_LogSafe(output, 18000), "", "", 0, ses
End Sub

Private Function GOS_WaitForPopupQuick(ByVal ses As Object, ByRef popup As Object, ByVal timeoutMs As Long, ByRef resultText As String) As Boolean
    Dim startedAt As Double, elapsedMs As Double
    startedAt = Timer
    Do
        Set popup = Nothing
        On Error Resume Next: Set popup = ses.findById("wnd[1]", False): On Error GoTo 0
        If Not popup Is Nothing Then
            resultText = "POPUP_WAIT_MS=" & Format$(ATT_Elapsed(startedAt) * 1000#, "0")
            GOS_WaitForPopupQuick = True
            Exit Function
        End If
        elapsedMs = ATT_Elapsed(startedAt) * 1000#
        If elapsedMs >= timeoutMs Then Exit Do
        DoEvents
        If elapsedMs < 1000# Then ATT_Sleep 50 Else ATT_Sleep 100
    Loop
    resultText = "NO_POPUP_WAIT_MS=" & Format$(ATT_Elapsed(startedAt) * 1000#, "0")
End Function

Private Sub GOS_LogWindowState(ByVal ses As Object, ByVal doc As String, ByVal action As String, ByVal method As String, ByVal controlId As String, ByVal stateStage As String)
    Dim i As Long, w As Object, output As String
    For i = 0 To 4
        Set w = Nothing
        On Error Resume Next: Set w = ses.findById("wnd[" & CStr(i) & "]", False): On Error GoTo 0
        If Not w Is Nothing Then output = output & "|" & GOS_WindowDescriptor(w)
    Next i
    output = output & "|" & GOS_StatusBarDescriptor(ses)
    GOS_LogEvent "GOS_WINDOW_STATE", "INFO", doc, action, method, controlId, "", "STAGE=" & stateStage & "|" & GOS_LogSafe(output, 8000), "", "", 0, ses
End Sub

Private Function GOS_WindowDescriptor(ByVal w As Object) As String
    Dim objId As String, typ As String, nm As String, txt As String, modal As String
    On Error Resume Next
    objId = CStr(w.id): typ = CStr(w.Type): nm = CStr(w.Name): txt = CStr(w.text): modal = CStr(w.IsModal)
    On Error GoTo 0
    GOS_WindowDescriptor = "WINDOW{ID=" & objId & ";TYPE=" & typ & ";NAME=" & nm & ";MODAL=" & modal & ";TEXT=" & txt & "}"
End Function

Private Function GOS_StatusBarDescriptor(ByVal ses As Object) As String
    Dim sbar As Object, txt As String, msgType As String
    On Error Resume Next
    Set sbar = ses.findById("wnd[0]/sbar", False)
    If Not sbar Is Nothing Then txt = CStr(sbar.text): msgType = CStr(sbar.messageType)
    On Error GoTo 0
    GOS_StatusBarDescriptor = "STATUSBAR{TYPE=" & msgType & ";TEXT=" & txt & "}"
End Function
Private Sub GOS_LogWindowMap(ByVal ses As Object, ByVal doc As String, ByVal action As String, ByVal method As String, ByVal controlId As String, ByVal mapStage As String)
    GOS_LogKnownControls ses, doc, action, method, controlId, "WINDOW_" & mapStage
End Sub

Private Sub GOS_CommitProfile(ByVal ses As Object, ByVal doc As String, ByVal action As String, ByVal method As String, ByVal controlId As String, ByVal selector As String)
    Dim profileKey As String, errNo As Long, errDesc As String
    profileKey = FAST_GosProfileKey(ses)
    On Error Resume Next
    Err.Clear
    ThisWorkbook.Save
    errNo = Err.Number: errDesc = Err.description
    Err.Clear
    On Error GoTo 0
    If errNo = 0 Then
        GOS_LogEvent "GOS_PROFILE_COMMIT", "DONE", doc, action, method, controlId, "", "PROFILE=" & profileKey & "|SELECTOR=" & selector, "", "", 0, ses
    Else
        GOS_LogEvent "GOS_PROFILE_COMMIT", "WARN", doc, action, method, controlId, "", "PROFILE=" & profileKey & "|SELECTOR=" & selector, "WORKBOOK_SAVE_FAIL", CStr(errNo) & ":" & errDesc, 0, ses
    End If
End Sub

Private Function GOS_WaitForAttachmentList(ByVal popup As Object, ByRef grid As Object, ByRef tbl As Object, ByRef listStage As String, ByRef probe As String) As Boolean
    Dim startedAt As Double, elapsedMs As Double
    startedAt = Timer: probe = ""
    Do
        Set grid = ATT_FindAttachmentGrid(popup)
        If Not grid Is Nothing Then
            listStage = "GRID_READY"
            probe = "WAIT_MS=" & Format$(ATT_Elapsed(startedAt) * 1000#, "0") & "|" & GOS_ControlDescriptor(grid)
            GOS_WaitForAttachmentList = True
            Exit Function
        End If
        Set tbl = ATT_FindType(popup, "GuiTableControl")
        If Not tbl Is Nothing Then
            listStage = "TABLE_READY"
            probe = "WAIT_MS=" & Format$(ATT_Elapsed(startedAt) * 1000#, "0") & "|" & GOS_ControlDescriptor(tbl)
            GOS_WaitForAttachmentList = True
            Exit Function
        End If
        elapsedMs = ATT_Elapsed(startedAt) * 1000#
        If elapsedMs >= 3500# Then Exit Do
        DoEvents
        If elapsedMs < 1000# Then ATT_Sleep 50 Else ATT_Sleep 100
    Loop
    listStage = "LIST_CONTROL_NOT_FOUND"
    probe = "WAIT_MS=" & Format$(ATT_Elapsed(startedAt) * 1000#, "0") & "|NO_GRID_OR_TABLE"
End Function

Private Function GOS_ControlDescriptor(ByVal obj As Object) As String
    Dim objId As String, typ As String, subTyp As String, rows As Long
    On Error Resume Next
    objId = CStr(obj.id): typ = CStr(obj.Type): subTyp = CStr(obj.SubType)
    Err.Clear: rows = CLng(CallByName(obj, "RowCount", VbGet)): If Err.Number <> 0 Then rows = -1
    On Error GoTo 0
    GOS_ControlDescriptor = "LIST_ID=" & objId & ";TYPE=" & typ & ";SUB=" & subTyp & ";ROWCOUNT=" & rows
End Function
Private Function GOS_SelectActionWithFallback(ByVal ses As Object, ByVal shell As Object, ByVal action As String, ByVal doc As String, ByVal method As String, ByVal controlId As String, ByRef popup As Object, ByRef selectorUsed As String, ByRef detail As String) As Boolean
    Dim candidates As Collection, selector As Variant, candidateLog As String, oneResult As String, idx As Long, waitResult As String
    Set candidates = New Collection
    idx = 0
    If UCase$(action) = "VIEW_ATTACHMENTS" Then
        idx = 1: oneResult = "": waitResult = "": Set popup = Nothing
        If GOS_TryDirectDisplayButton(ses, oneResult) Then
            GOS_LogEvent "MENU_ATTEMPT", "SELECTED", doc, action, method, controlId, "", "TRY=" & idx & "|SELECTOR=DIRECT_BUTTON:Display Attachment|RESULT=" & oneResult, "", "", 0, ses
            If GOS_WaitForPopupQuick(ses, popup, 3500, waitResult) Then
                selectorUsed = "DIRECT_BUTTON:Display Attachment"
                GOS_LogEvent "MENU_ATTEMPT", "SUCCESS", doc, action, method, controlId, "wnd[1]", "TRY=" & idx & "|SELECTOR=" & selectorUsed & "|" & waitResult, "", "", 0, ses
                GOS_SelectActionWithFallback = True
                Exit Function
            End If
            oneResult = oneResult & ";" & waitResult
            GOS_LogWindowState ses, doc, action, method, controlId, "NO_POPUP_DIRECT_BUTTON"
            GOS_LogEvent "MENU_ATTEMPT", "NO_POPUP", doc, action, method, controlId, "", "TRY=" & idx & "|SELECTOR=DIRECT_BUTTON:Display Attachment|RESULT=" & oneResult, "POPUP_NOT_FOUND", oneResult, 0, ses
        Else
            GOS_LogEvent "MENU_ATTEMPT", "SKIP", doc, action, method, controlId, "", "TRY=" & idx & "|SELECTOR=DIRECT_BUTTON:Display Attachment|RESULT=" & oneResult, "DIRECT_BUTTON_NOT_FOUND", oneResult, 0, ses
        End If
        If candidateLog <> "" Then candidateLog = candidateLog & " || "
        candidateLog = candidateLog & "TRY" & idx & "=DIRECT_BUTTON:" & oneResult
        GOS_AddActionCandidate candidates, mLastGosViewSelector
        GOS_AddActionCandidate candidates, "BUTTON:%GOS_VIEW_ATTA"
        GOS_AddActionCandidate candidates, "TOOLBAR:%GOS_VIEW_ATTA"
        GOS_AddActionCandidate candidates, "ID:%GOS_VIEW_ATTA"
        GOS_AddActionCandidate candidates, "TEXT:Display Attachment"
        GOS_AddActionCandidate candidates, "TEXT:Display Attachments"
        GOS_AddActionCandidate candidates, "TEXT:Attachment list"
    Else
        GOS_AddActionCandidate candidates, mLastGosCreateSelector
        GOS_AddActionCandidate candidates, "ID:%GOS_PCATTA_CREA"
        GOS_AddActionCandidate candidates, "BUTTON:%GOS_PCATTA_CREA"
        GOS_AddActionCandidate candidates, "TOOLBAR:%GOS_PCATTA_CREA"
        GOS_AddActionCandidate candidates, "TEXT:Create Attachment"
        GOS_AddActionCandidate candidates, "TEXT:Create attachment"
    End If
    For Each selector In candidates
        idx = idx + 1: oneResult = "": waitResult = "": Set popup = Nothing
        If GOS_TryActionSelector(shell, CStr(selector), oneResult) Then
            GOS_LogEvent "MENU_ATTEMPT", "SELECTED", doc, action, method, controlId, "", "TRY=" & idx & "|SELECTOR=" & CStr(selector) & "|RESULT=" & oneResult, "", "", 0, ses
            If GOS_WaitForPopupQuick(ses, popup, 3500, waitResult) Then
                selectorUsed = CStr(selector)
                GOS_LogEvent "MENU_ATTEMPT", "SUCCESS", doc, action, method, controlId, "wnd[1]", "TRY=" & idx & "|SELECTOR=" & selectorUsed & "|" & waitResult, "", "", 0, ses
                GOS_SelectActionWithFallback = True
                Exit Function
            End If
            oneResult = oneResult & ";" & waitResult
            GOS_LogWindowState ses, doc, action, method, controlId, "NO_POPUP_TRY_" & idx
            GOS_LogEvent "MENU_ATTEMPT", "NO_POPUP", doc, action, method, controlId, "", "TRY=" & idx & "|SELECTOR=" & CStr(selector) & "|RESULT=" & oneResult, "POPUP_NOT_FOUND", oneResult, 0, ses
        Else
            GOS_LogEvent "MENU_ATTEMPT", "FAIL", doc, action, method, controlId, "", "TRY=" & idx & "|SELECTOR=" & CStr(selector) & "|RESULT=" & oneResult, "MENU_SELECTOR_FAIL", oneResult, 0, ses
        End If
        If candidateLog <> "" Then candidateLog = candidateLog & " || "
        candidateLog = candidateLog & "TRY" & idx & "=" & CStr(selector) & ":" & oneResult
    Next selector
    detail = GOS_LogSafe(candidateLog, 7000)
End Function

Private Sub GOS_AddActionCandidate(ByVal candidates As Collection, ByVal selector As String)
    Dim i As Long
    selector = Trim$(selector)
    If selector = "" Then Exit Sub
    For i = 1 To candidates.count
        If UCase$(CStr(candidates(i))) = UCase$(selector) Then Exit Sub
    Next i
    candidates.Add selector
End Sub
Private Function GOS_TryActionSelector(ByVal shell As Object, ByVal selector As String, ByRef resultText As String) As Boolean
    Dim payload As String, errNo As Long, errDesc As String, prefix As String
    prefix = UCase$(Left$(selector, InStr(selector, ":")))
    payload = Mid$(selector, InStr(selector, ":") + 1)
    On Error Resume Next
    Err.Clear
    Select Case prefix
        Case "ID:"
            shell.pressContextButton "%GOS_TOOLBOX"
            If Err.Number = 0 Then
                ATT_Sleep 100
                shell.selectContextMenuItem payload
            End If
        Case "TEXT:"
            shell.pressContextButton "%GOS_TOOLBOX"
            If Err.Number = 0 Then
                ATT_Sleep 100
                CallByName shell, "selectContextMenuItemByText", VbMethod, payload
            End If
        Case "BUTTON:"
            CallByName shell, "pressButton", VbMethod, payload
        Case "TOOLBAR:"
            CallByName shell, "pressToolbarButton", VbMethod, payload
        Case Else
            resultText = "INVALID_SELECTOR=" & selector
            On Error GoTo 0
            Exit Function
    End Select
    errNo = Err.Number: errDesc = Err.description
    Err.Clear: On Error GoTo 0
    If errNo = 0 Then
        resultText = "SELECT_OK|METHOD=" & Left$(prefix, Len(prefix) - 1)
        GOS_TryActionSelector = True
    Else
        resultText = "SELECT_FAIL=" & CStr(errNo) & ":" & errDesc & "|METHOD=" & Left$(prefix, Len(prefix) - 1)
    End If
End Function

Private Sub GOS_SaveActionProfile(ByVal ses As Object, ByVal action As String, ByVal selector As String)
    Dim profileKey As String, cfgKey As String, sharedKey As String
    profileKey = FAST_GosProfileKey(ses)
    If UCase$(action) = "VIEW_ATTACHMENTS" Then
        cfgKey = "ATT_GOS_VIEW_SELECTOR_" & profileKey
        sharedKey = "ATT_GOS_VIEW_SELECTOR_SHARED"
        mLastGosViewSelector = selector
    Else
        cfgKey = "ATT_GOS_CREATE_SELECTOR_" & profileKey
        sharedKey = "ATT_GOS_CREATE_SELECTOR_SHARED"
        mLastGosCreateSelector = selector
    End If
    ATT_SetCfg cfgKey, selector, "Auto-saved GOS action selector for SAP system/client/user"
    ' --- Also refresh SHARED key so new users benefit from this session ---
    ATT_SetCfg sharedKey, selector, "Shared GOS action selector (auto-refreshed from last successful run)"
    FAST_Log "GOS_ACTION_PROFILE", "SAVED", "", "PROFILE=" & profileKey & ";ACTION=" & action & ";SELECTOR=" & selector
End Sub

Private Sub GOS_LogSapMap(ByVal ses As Object, ByVal shell As Object, ByVal doc As String, ByVal action As String, ByVal method As String, ByVal controlId As String, ByVal mapStage As String)
    Dim mapText As String
    mapText = GOS_ControlMap(shell)
    GOS_LogEvent "GOS_SAP_MAP", "INFO", doc, action, method, controlId, "", "MAP_STAGE=" & mapStage & "|" & mapText, "", "", 0, ses
End Sub
Private Function GOS_ControlMap(ByVal root As Object) As String
    Dim output As String
    If root Is Nothing Then GOS_ControlMap = "MAP_ROOT_NOT_FOUND": Exit Function
    GOS_MapNode root, 0, output
    GOS_ControlMap = GOS_LogSafe(output, 18000)
End Function
Private Sub GOS_MapNode(ByVal obj As Object, ByVal depth As Long, ByRef output As String)
    Dim objId As String, typ As String, subTyp As String, nm As String, txt As String, tip As String, acc As String
    Dim count As Long, rowCount As Long, colCount As Long, i As Long, child As Object, col As Object, bid As String, btip As String, btxt As String
    If obj Is Nothing Or depth > 7 Or Len(output) > 16500 Then Exit Sub
    On Error Resume Next
    objId = CStr(obj.id): typ = CStr(obj.Type): subTyp = CStr(obj.SubType): nm = CStr(obj.Name)
    txt = CStr(obj.text): tip = CStr(obj.Tooltip): acc = CStr(obj.AccTooltip)
    Err.Clear: rowCount = CLng(CallByName(obj, "RowCount", VbGet)): If Err.Number <> 0 Then rowCount = -1
    Err.Clear: colCount = CLng(CallByName(obj, "ColumnCount", VbGet)): If Err.Number <> 0 Then colCount = -1
    Err.Clear
    On Error GoTo 0
    output = output & "|NODE" & depth & "{ID=" & objId & ";TYPE=" & typ & ";SUB=" & subTyp & ";NAME=" & nm & ";ROWCOUNT=" & rowCount & ";COLCOUNT=" & colCount & ";TEXT=" & txt & ";TIP=" & tip & ";ACC=" & acc & "}"
    On Error Resume Next
    Err.Clear: count = CLng(CallByName(obj, "ButtonCount", VbGet))
    If Err.Number <> 0 Then count = 0
    Err.Clear: On Error GoTo 0
    If count > 0 And count < 64 Then
        For i = 0 To count - 1
            bid = "": btip = "": btxt = ""
            On Error Resume Next
            bid = CStr(CallByName(obj, "GetButtonId", VbMethod, i))
            btip = CStr(CallByName(obj, "GetButtonTooltip", VbMethod, i))
            btxt = CStr(CallByName(obj, "GetButtonText", VbMethod, i))
            On Error GoTo 0
            output = output & "|BUTTON" & i & "{ID=" & bid & ";TIP=" & btip & ";TEXT=" & btxt & "}"
        Next i
    End If
    Set col = Nothing
    On Error Resume Next
    Set col = obj.Children
    If Not col Is Nothing Then count = CLng(col.count) Else count = 0
    On Error GoTo 0
    If count > 48 Then count = 48
    For i = 0 To count - 1
        Set child = Nothing
        On Error Resume Next: Set child = col.item(i): On Error GoTo 0
        If Not child Is Nothing Then GOS_MapNode child, depth + 1, output
    Next i
End Sub

Private Function GOS_DocFieldMap(ByVal ses As Object) As String
    Dim root As Object, output As String
    On Error Resume Next: Set root = ses.findById("wnd[0]", False): On Error GoTo 0
    GOS_DocMapNode root, 0, output
    GOS_DocFieldMap = GOS_LogSafe(output, 8000)
End Function

Private Sub GOS_DocMapNode(ByVal obj As Object, ByVal depth As Long, ByRef output As String)
    Dim objId As String, typ As String, nm As String, txt As String, count As Long, i As Long, col As Object, child As Object
    If obj Is Nothing Or depth > 5 Or Len(output) > 7000 Then Exit Sub
    On Error Resume Next
    objId = CStr(obj.id): typ = CStr(obj.Type): nm = CStr(obj.Name): txt = CStr(obj.text)
    On Error GoTo 0
    If InStr(1, UCase$(objId & " " & nm), "DOC", vbTextCompare) > 0 Or InStr(1, UCase$(objId & " " & nm), "MAT", vbTextCompare) > 0 Then
        output = output & "|FIELD{ID=" & objId & ";TYPE=" & typ & ";NAME=" & nm & ";TEXT=" & txt & "}"
    End If
    Set col = Nothing
    On Error Resume Next: Set col = obj.Children: If Not col Is Nothing Then count = CLng(col.count) Else count = 0: On Error GoTo 0
    If count > 48 Then count = 48
    For i = 0 To count - 1
        Set child = Nothing
        On Error Resume Next: Set child = col.item(i): On Error GoTo 0
        If Not child Is Nothing Then GOS_DocMapNode child, depth + 1, output
    Next i
End Sub

Private Function GOS_LogSafe(ByVal sourceText As String, ByVal maxLen As Long) As String
    sourceText = Replace(Replace(Replace(CStr(sourceText), "|", "/"), vbCr, " "), vbLf, " ")
    If Len(sourceText) > maxLen Then sourceText = Left$(sourceText, maxLen) & "...[TRUNCATED]"
    GOS_LogSafe = sourceText
End Function

Private Function FAST_GosProfileKey(ByVal ses As Object) As String
    Dim sapSystem As String, sapClient As String, sapUser As String
    On Error Resume Next
    sapSystem = UCase$(Trim$(CStr(ses.info.SystemName)))
    sapClient = UCase$(Trim$(CStr(ses.info.client)))
    sapUser = UCase$(Trim$(CStr(ses.info.User)))
    On Error GoTo 0
    If sapUser = "" Then sapUser = UCase$(Trim$(Application.UserName))
    If sapSystem = "" Then sapSystem = "SAP"
    If sapClient = "" Then sapClient = "DEFAULT"
    FAST_GosProfileKey = FAST_ProfileToken(sapSystem & "_" & sapClient & "_" & sapUser)
End Function

Private Function FAST_ProfileToken(ByVal sourceText As String) As String
    Dim i As Long, ch As String, output As String
    sourceText = UCase$(sourceText)
    For i = 1 To Len(sourceText)
        ch = Mid$(sourceText, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then output = output & ch Else output = output & "_"
    Next i
    Do While InStr(output, "__") > 0: output = Replace(output, "__", "_"): Loop
    FAST_ProfileToken = Trim$(output)
End Function

Private Sub GOS_SaveUserProfile(ByVal ses As Object, ByVal method As String, ByVal controlId As String, ByVal controlType As String)
    Dim profileKey As String
    profileKey = FAST_GosProfileKey(ses)
    ATT_SetCfg "ATT_GOS_METHOD_" & profileKey, method, "Auto-saved GOS method for SAP system/client/user"
    ATT_SetCfg "ATT_GOS_CONTROL_" & profileKey, controlId, "Auto-saved GOS control for SAP system/client/user"
    ATT_SetCfg "ATT_GOS_TYPE_" & profileKey, controlType, "Auto-saved GOS control type for SAP system/client/user"
    ' --- Also refresh SHARED keys so new users benefit from this session ---
    ATT_SetCfg "ATT_GOS_CONTROL_SHARED", controlId, "Shared GOS control (auto-refreshed from last successful run)"
    ATT_SetCfg "ATT_GOS_METHOD_SHARED", method, "Shared GOS method (auto-refreshed from last successful run)"
    ATT_SetCfg "ATT_GOS_TYPE_SHARED", controlType, "Shared GOS control type (auto-refreshed from last successful run)"
    FAST_Log "GOS_PROFILE", "SAVED", "", "PROFILE=" & profileKey & ";METHOD=" & method & ";CONTROL=" & controlId
End Sub

Private Sub ATT_SetCfg(ByVal key As String, ByVal cfgValue As String, ByVal description As String)
    Dim ws As Worksheet, f As Range, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    Set f = ws.Columns(1).Find(What:=key, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then
        r = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 1
        ws.Cells(r, 1).value = key: ws.Cells(r, 2).value = cfgValue: ws.Cells(r, 3).value = description
    Else
        ws.Cells(f.Row, 2).value = cfgValue
        If Trim$(CStr(ws.Cells(f.Row, 3).value)) = "" Then ws.Cells(f.Row, 3).value = description
    End If
End Sub

Private Function FAST_FileMatchesType(ByVal fileName As String, ByVal typ As String) As Boolean
    Dim s As String
    s = ATT_NormText(fileName)
    Select Case UCase$(typ)
        Case "VAT"
            FAST_FileMatchesType = ATT_HasKeyword(fileName, ATT_Cfg("VAT_KEYWORDS", "VAT,HOADON,HOA DON,INVOICE")) Or _
                                  InStr(s, "HOA DON GIA TRI GIA TANG") > 0 Or InStr(s, "VAT INVOICE") > 0
        Case "BBBG"
            FAST_FileMatchesType = ATT_HasKeyword(fileName, ATT_Cfg("BBBG_KEYWORDS", "BBBG,BBGH,BIENBAN,BIEN BAN,BAN GIAO"))
    End Select
End Function

Private Function GOS_TitleHasType(ByVal title As String, ByVal typ As String) As Boolean
    Dim s As String
    s = ATT_NormText(title)
    Select Case UCase$(Trim$(typ))
        Case "VAT"
            GOS_TitleHasType = GOS_HasSapToken(s, "VAT") Or GOS_HasSapToken(s, "HOADON") Or GOS_HasSapToken(s, "INVOICE")
        Case "BBBG"
            GOS_TitleHasType = GOS_HasSapToken(s, "BB") Or GOS_HasSapToken(s, "BBBG") Or GOS_HasSapToken(s, "BBGH") Or _
                                  InStr(1, s, "BIEN BAN", vbTextCompare) > 0 Or InStr(1, s, "BAN GIAO", vbTextCompare) > 0
    End Select
End Function

Private Sub FAST_AddFileIndexItem(ByVal idx As Object, ByVal key As String, ByVal fullPath As String, ByVal fileName As String, ByVal rawTicket As String, ByVal typ As String, ByVal fileSize As Double, ByVal lastModified As String)
    Dim bucket As Collection, item As Object
    If Not idx.Exists(key) Then Set bucket = New Collection: idx.Add key, bucket
    Set bucket = idx(key)
    Set item = CreateObject("Scripting.Dictionary"): item.CompareMode = vbTextCompare
    item("FullPath") = fullPath: item("FileName") = fileName: item("Ticket") = rawTicket
    item("Type") = typ: item("LastModified") = lastModified: item("Size") = fileSize
    bucket.Add item
End Sub

Private Function FAST_DocHasMissingType(ByVal d As Object, ByVal titles As Object, ByVal typ As String) As Boolean
    Dim k As Variant, g As Object, sapPresent As Boolean, expected As Boolean
    typ = UCase$(Trim$(typ))
    If typ = "VAT" And Not ATT_BoolCfg("VAT_REQUIRED", True) Then Exit Function
    sapPresent = GOS_HasType(titles, typ)
    For Each k In d("GROUPS").keys
        Set g = d("GROUPS")(k)
        expected = (Trim$(CStr(g(typ))) <> "")
        If Not expected And Not sapPresent Then
            FAST_DocHasMissingType = True
            Exit Function
        End If
    Next k
End Function

Private Sub ATT_HistoryOnce(ByVal docNo As String, ByVal ticket As String, ByVal typ As String, ByVal evidence As String, ByVal result As String, ByVal note As String)
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets(SH_HISTORY)
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    For r = lastRow To 2 Step -1
        If CStr(ws.Cells(r, 2).value) = docNo And UCase$(CStr(ws.Cells(r, 4).value)) = UCase$(typ) And _
           CStr(ws.Cells(r, 5).value) = evidence And UCase$(CStr(ws.Cells(r, 6).value)) = UCase$(result) Then Exit Sub
    Next r
    ATT_History docNo, ticket, typ, evidence, result, note
End Sub

Private Sub FAST_RecordSapAttachmentEvidence(ByVal docNo As String, ByVal d As Object, ByVal titles As Object)
    Dim k As Variant, title As String, typ As Variant, ticket As String
    If titles Is Nothing Then Exit Sub
    ticket = FAST_FirstGroupTicket(d)
    For Each k In titles.keys
        title = CStr(titles(k))
        For Each typ In Array("BBBG", "VAT")
            If GOS_TitleHasType(title, typ) Then ATT_HistoryOnce docNo, ticket, typ, "SAP_TITLE:" & title, "SAP_EXISTING", "DISPLAY_ATTACHMENT"
        Next typ
    Next k
End Sub

Private Sub FAST_EnsureInputUI(ByVal ws As Worksheet)
    ws.Cells(6, C_OCR_PO).value = "OCR PO"
    ws.Cells(6, C_OCR_MAT_BBBG).value = "OCR Material BBBG"
    ws.Cells(6, C_OCR_TICKET).value = "OCR So phieu"
    ws.Cells(6, C_OCR_MAT_VAT).value = "OCR Material VAT"
    ws.Cells(6, C_CHECK).value = "Check Status"
    ws.Cells(6, C_HEAD).value = "HeadText Status"
    ws.Cells(6, C_ATT_BBBG).value = "ATT BBBG"
    ws.Cells(6, C_ATT_VAT).value = "ATT VAT"
    ws.Cells(6, C_ATT).value = "ATT Status"
    ws.Cells(6, C_NOTE).value = "Note"
    With ws.Range(ws.Cells(6, 1), ws.Cells(6, C_NOTE))
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Columns("R").ColumnWidth = 22
    ws.Columns("S").ColumnWidth = 16
    ws.Cells(1, 18).value = "FAST RUN SUMMARY"
    With ws.Range(ws.Cells(1, 18), ws.Cells(1, 19))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With
End Sub

Private Sub FAST_UpdateDashboard(ByVal ws As Worksheet, ByVal queueCount As Long, ByVal okCount As Long, ByVal errCount As Long, ByVal skipCount As Long, ByVal runState As String, ByVal elapsedSeconds As Double)
    FAST_EnsureInputUI ws
    ws.Cells(2, 18).value = "Queue DOC": ws.Cells(2, 19).value = queueCount
    ws.Cells(3, 18).value = "Success": ws.Cells(3, 19).value = okCount
    ws.Cells(4, 18).value = "Error": ws.Cells(4, 19).value = errCount
    ws.Cells(5, 18).value = "Skipped": ws.Cells(5, 19).value = skipCount
    ws.Cells(6, 18).value = "Elapsed sec": ws.Cells(6, 19).value = Round(elapsedSeconds, 1)
    ws.Cells(7, 18).value = "Run status": ws.Cells(7, 19).value = runState
    ws.Cells(8, 18).value = "Updated": ws.Cells(8, 19).value = Now
    ws.Cells(8, 19).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Range(ws.Cells(2, 18), ws.Cells(8, 19)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(2, 18), ws.Cells(8, 18)).Font.Bold = True
    Select Case UCase$(runState)
        Case "DONE": ws.Cells(7, 19).Interior.Color = RGB(198, 239, 206)
        Case "FATAL": ws.Cells(7, 19).Interior.Color = RGB(255, 199, 206)
        Case Else: ws.Cells(7, 19).Interior.Color = RGB(255, 235, 156)
    End Select
End Sub

Private Sub FAST_RefreshStatusStyles(ByVal ws As Worksheet)
    Dim r As Long, lastRow As Long
    lastRow = ATT_LastRow(ws, C_DOC)
    For r = FIRST_ROW To lastRow
        FAST_ColorStatusCell ws.Cells(r, C_CHECK), CStr(ws.Cells(r, C_CHECK).value)
        FAST_ColorStatusCell ws.Cells(r, C_ATT_BBBG), CStr(ws.Cells(r, C_ATT_BBBG).value)
        FAST_ColorStatusCell ws.Cells(r, C_ATT_VAT), CStr(ws.Cells(r, C_ATT_VAT).value)
        FAST_ColorStatusCell ws.Cells(r, C_ATT), CStr(ws.Cells(r, C_ATT).value)
    Next r
End Sub

Private Sub FAST_ColorStatusCell(ByVal target As Range, ByVal statusText As String)
    Dim s As String
    s = UCase$(Trim$(statusText))
    target.Font.Bold = (s <> "")
    target.Interior.pattern = xlNone
    If s = "" Then Exit Sub
    If InStr(s, "ERROR") > 0 Or InStr(s, "MISSING") > 0 Or InStr(s, "NOT_FOUND") > 0 Then
        target.Interior.Color = RGB(255, 199, 206)
    ElseIf InStr(s, "PARTIAL") > 0 Or InStr(s, "SKIP") > 0 Or InStr(s, "UNVERIFIED") > 0 Then
        target.Interior.Color = RGB(255, 235, 156)
    ElseIf InStr(s, "DONE") > 0 Or s = "CHECK_OK" Then
        target.Interior.Color = RGB(198, 239, 206)
    End If
End Sub

Public Sub FAST_ReconcileExistingStatuses()
    Dim ws As Worksheet, r As Long, lastRow As Long, b As String, v As String, docNo As String, docs As Object
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    Set docs = CreateObject("Scripting.Dictionary"): docs.CompareMode = vbTextCompare
    lastRow = ATT_LastRow(ws, C_DOC)
    For r = FIRST_ROW To lastRow
        docNo = Trim$(CStr(ws.Cells(r, C_DOC).value))
        If docNo <> "" Then
            If Not docs.Exists(docNo) Then docs.Add docNo, True
            b = UCase$(Trim$(CStr(ws.Cells(r, C_ATT_BBBG).value)))
            v = UCase$(Trim$(CStr(ws.Cells(r, C_ATT_VAT).value)))
            If b = "DONE" And v = "DONE" Then
                ws.Cells(r, C_ATT).value = "ATT_DONE"
            ElseIf b = "DONE" Then
                ws.Cells(r, C_ATT).value = "ATT_PARTIAL_NO_VAT"
            ElseIf v = "DONE" Then
                ws.Cells(r, C_ATT).value = "ATT_PARTIAL_NO_BBBG"
            Else
                ws.Cells(r, C_ATT).value = "ATT_PARTIAL_MISSING_BBBG_VAT"
            End If
        End If
    Next r
    FAST_RefreshStatusStyles ws
    FAST_UpdateDashboard ws, docs.count, 0, 0, 0, "READY", 0
End Sub

Private Function ATT_RunTesseractPass(ByVal imagePath As String, ByVal psm As String, ByRef status As String) As String
    Dim exe As String, base As String, cmd As String, rc As Long, txt As String
    exe = ATT_TesseractPath()
    If exe = "" Then status = "TESSERACT_MISSING": Exit Function
    base = ATT_TempFolder("OCR") & "\result_" & psm
    cmd = ATT_Q(exe) & " " & ATT_Q(imagePath) & " " & ATT_Q(base) & _
          " -l " & ATT_Cfg("OCR_LANGUAGE", "eng+vie") & " --psm " & psm & _
          " --dpi 300 -c preserve_interword_spaces=1"
    rc = ATT_RunWait(cmd)
    If rc <> 0 Then status = "OCR_FAIL_" & rc: Exit Function
    txt = base & ".txt"
    If Dir$(txt) = "" Then status = "OCR_NO_OUTPUT": Exit Function
    ATT_RunTesseractPass = ATT_ReadText(txt): status = "OCR_OK"
End Function

Private Function ATT_RunTesseractDigits(ByVal imagePath As String, ByVal psm As String) As String
    Dim exe As String, base As String, cmd As String, txt As String
    exe = ATT_TesseractPath()
    If exe = "" Then Exit Function
    base = ATT_TempFolder("OCRNUM") & "\digits_" & psm
    cmd = ATT_Q(exe) & " " & ATT_Q(imagePath) & " " & ATT_Q(base) & _
          " -l eng --oem 1 --psm " & psm & " --dpi 300" & _
          " -c thresholding_method=2 -c thresholding_kfactor=0.34" & _
          " -c tessedit_char_whitelist=0123456789 -c preserve_interword_spaces=1"
    If ATT_RunWait(cmd) <> 0 Then Exit Function
    txt = base & ".txt"
    If Dir$(txt) <> "" Then ATT_RunTesseractDigits = ATT_ReadText(txt)
End Function

Private Function ATT_TicketNearMatch(ByVal a As String, ByVal b As String) As Boolean
    Dim ca As String, cb As String, i As Long, diff As Long
    ca = ATT_TicketCode(a): cb = ATT_TicketCode(b)
    If ca = "" Or cb = "" Or Len(ca) <> Len(cb) Then Exit Function
    If Len(ca) < 4 Then Exit Function
    For i = 1 To Len(ca)
        If Mid$(ca, i, 1) <> Mid$(cb, i, 1) Then diff = diff + 1
        If diff > 1 Then Exit Function
    Next i
    ATT_TicketNearMatch = (diff = 1)
End Function

Private Function ATT_AppendNote(ByVal currentText As String, ByVal newText As String) As String
    If Trim$(currentText) = "" Then ATT_AppendNote = newText Else ATT_AppendNote = currentText & " " & newText
End Function


Private Function ATT_ExtractPO(ByVal raw As String, ByVal groups As Object) As Object
    Dim found As Object, expected As Object, g As Object, poSet As Object
    Dim groupKey As Variant, poKey As Variant, re As Object
    Dim code As String, pattern As String, i As Long
    Set found = CreateObject("Scripting.Dictionary")
    Set expected = CreateObject("Scripting.Dictionary")

    ' Chi tim cac SA/PO dang co trong cot dau vao. Khong bien O/I/L thanh 0/1
    ' vi cach sua rong nay co the tao ra mot PO khong ton tai tren chung tu.
    For Each groupKey In groups.keys
        Set g = groups(groupKey)
        Set poSet = g("PO")
        For Each poKey In poSet.keys
            code = UCase$(ATT_Clean(CStr(poKey)))
            If Len(code) = 10 And Left$(code, 1) = "5" And IsNumeric(code) Then expected(code) = True
        Next poKey
    Next groupKey

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    For Each poKey In expected.keys
        code = CStr(poKey)
        pattern = "(^|[^0-9])"
        For i = 1 To Len(code)
            If i > 1 Then pattern = pattern & "[ -]?"
            pattern = pattern & Mid$(code, i, 1)
        Next i
        pattern = pattern & "([^0-9]|$)"
        re.pattern = pattern
        If re.Test(raw) Then found(code) = True
    Next poKey

    Set ATT_ExtractPO = found
End Function


Private Function ATT_RunTesseract(ByVal imagePath As String, ByRef status As String) As String
    Dim psm1 As String, psm2 As String, s1 As String, s2 As String, st1 As String, st2 As String
    psm1 = ATT_Cfg("OCR_PSM", "6")
    psm2 = ATT_Cfg("OCR_PSM_SECONDARY", "11")
    s1 = ATT_RunTesseractPass(imagePath, psm1, st1)
    If psm2 <> "" And psm2 <> psm1 Then s2 = ATT_RunTesseractPass(imagePath, psm2, st2)
    ATT_RunTesseract = s1
    If Trim$(s2) <> "" Then ATT_RunTesseract = ATT_RunTesseract & vbCrLf & "--- OCR PASS " & psm2 & " ---" & vbCrLf & s2
    If Trim$(ATT_RunTesseract) = "" Then status = st1 & "/" & st2 Else status = "OCR_MULTI_" & psm1 & "_" & psm2
End Function

Private Function ATT_ReadPreparedImage(ByVal imagePath As String, ByRef status As String) As String
    Dim prepared As String, out As String, stMain As String, stSecond As String, digits As String
    ' WIA la thanh phan anh co san cua Windows, chay trong Excel; khong goi PowerShell.
    prepared = ATT_PreprocessImageWIA(imagePath)
    If prepared = "" Then prepared = imagePath
    out = ATT_RunTesseractPass(prepared, ATT_Cfg("OCR_PSM", "6"), stMain)
    out = out & vbCrLf & "--- IMAGE SPARSE PASS ---" & vbCrLf & _
          ATT_RunTesseractPass(prepared, ATT_Cfg("OCR_PSM_SECONDARY", "11"), stSecond)
    digits = ATT_RunTesseractDigits(prepared, ATT_Cfg("OCR_IMAGE_NUMERIC_PSM", "6"))
    If Trim$(digits) <> "" Then out = out & vbCrLf & "--- IMAGE DIGITS ONLY ---" & vbCrLf & digits
    If prepared = imagePath Then status = "IMAGE_WIA_FALLBACK" Else status = "IMAGE_WIA_PREPROCESSED"
    ATT_ReadPreparedImage = out
End Function

Private Function ATT_PreprocessImageWIA(ByVal imagePath As String) As String
    Const WIA_SCALE As String = "{4EBB0166-C18B-4065-9332-109015741711}"
    Const WIA_CONVERT As String = "{42A6E907-1D2F-4B38-AC50-31ADBE2AB3C2}"
    Const WIA_PNG As String = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
    Dim img As Object, proc As Object, result As Object, outFolder As String, outPath As String
    On Error GoTo EH
    outFolder = ATT_TempFolder("IMG_WIA"): outPath = outFolder & "\image_scaled.png"
    Set img = CreateObject("WIA.ImageFile"): img.LoadFile imagePath
    Set proc = CreateObject("WIA.ImageProcess")
    proc.Filters.Add WIA_SCALE
    proc.Filters(1).Properties("MaximumWidth").value = 3600
    proc.Filters(1).Properties("MaximumHeight").value = 3600
    proc.Filters(1).Properties("PreserveAspectRatio").value = True
    proc.Filters.Add WIA_CONVERT
    proc.Filters(2).Properties("FormatID").value = WIA_PNG
    Set result = proc.Apply(img): result.SaveFile outPath
    If Dir$(outPath) <> "" Then ATT_PreprocessImageWIA = outPath
    Exit Function
EH:
    ATT_PreprocessImageWIA = ""
End Function


Private Sub ATT_MatchGroups(ByVal groups As Object, ByVal infos As Collection)
    Dim k As Variant, g As Object, info As Object, b As Collection, v As Collection
    Dim matOk As Boolean, poOk As Boolean, i As Long, bInfo As Object, ticketRead As String
    Dim ticketFuzzy As Boolean
    For Each k In groups.keys
        Set g = groups(k): Set b = New Collection: Set v = New Collection
        If Trim$(g("TICKET")) = "" Then
            g("STATUS") = "NEED_SOPHIEU": g("NOTE") = "Nhap So phieu tu MB51 roi CHECK lai.": GoTo ContinueGroup
        End If
        For i = 1 To infos.count
            Set info = infos(i)
            If info("TYPE") <> "VAT" Then
                poOk = ATT_ContainsAll(info("PO"), g("PO"))
                matOk = ATT_Overlap(info("MAT"), g("MAT")) > 0
                ticketRead = ATT_Clean(info("TICKET"))
                If poOk And matOk Then
                    If ticketRead = "" Or ATT_TicketMatches(ticketRead, g("TICKET")) Or ATT_TicketNearMatch(ticketRead, g("TICKET")) Then b.Add info
                End If
            End If
        Next i
        If b.count = 0 Then
            g("STATUS") = "NO_BBBG_MATCH": g("NOTE") = "Khong co BBBG khop tat ca PO/SA + Material + So phieu."
        ElseIf b.count > 1 Then
            g("STATUS") = "MULTI_BBBG_MATCH": g("NOTE") = b.count & " file BBBG cung khop."
        Else
            Set bInfo = b(1): g("BBBG") = bInfo("PATH")
            g("OCR_PO") = ATT_JoinKeys(bInfo("PO"), ";")
            g("OCR_MAT_B") = ATT_JoinKeys(bInfo("MAT"), ";")
            ticketRead = ATT_Clean(bInfo("TICKET"))
            ticketFuzzy = (ticketRead <> "" And Not ATT_TicketMatches(ticketRead, g("TICKET")) And ATT_TicketNearMatch(ticketRead, g("TICKET")))
            If ticketRead = "" Then
                g("OCR_TICKET") = "NOT_READ"
            ElseIf ticketFuzzy Then
                g("OCR_TICKET") = g("TICKET")
                g("NOTE") = "OCR doc " & ticketRead & "; da doi chieu PO + Material va sua thanh " & g("TICKET") & "."
            Else
                g("OCR_TICKET") = ticketRead
            End If
            For i = 1 To infos.count
                Set info = infos(i)
                If info("TYPE") = "VAT" And ATT_FileTicketMatch(info("NAME"), g("TICKET")) Then
                    matOk = (ATT_Overlap(info("MAT"), g("MAT")) > 0 And ATT_Overlap(info("MAT"), bInfo("MAT")) > 0)
                    If matOk Then v.Add info
                End If
            Next i
            If v.count = 0 Then
                If ATT_BoolCfg("VAT_REQUIRED", True) Then
                    g("STATUS") = "NO_VAT_MATCH": g("NOTE") = ATT_AppendNote(CStr(g("NOTE")), "VAT can ten file khop So phieu va ma BIW khop BBBG.")
                Else
                    g("STATUS") = "CHECK_OK": g("NOTE") = ATT_AppendNote(CStr(g("NOTE")), "Khong co VAT; CONFIG cho phep BBBG-only.")
                End If
            ElseIf v.count > 1 Then
                g("STATUS") = "MULTI_VAT_MATCH": g("NOTE") = ATT_AppendNote(CStr(g("NOTE")), v.count & " file VAT cung khop.")
            Else
                Set info = v(1): g("VAT") = info("PATH")
                g("OCR_MAT_V") = ATT_JoinKeys(info("MAT"), ";")
                g("STATUS") = "CHECK_OK"
                If g("OCR_TICKET") = "NOT_READ" Then g("NOTE") = ATT_AppendNote(CStr(g("NOTE")), "BBBG khong doc duoc So phieu; da xac nhan bang tat ca PO/SA + Material.")
            End If
        End If
ContinueGroup:
    Next k
End Sub
Private Function ATT_ValidateMaterialDocs(ByVal ws As Worksheet) As Boolean
    Dim r As Long, lr As Long, missingCount As Long
    Dim mat As String, doc As String, po As String, ticket As String
    lr = Application.Max(ATT_LastRow(ws, C_MAT), ATT_LastRow(ws, C_DOC), ATT_LastRow(ws, C_PO), ATT_LastRow(ws, C_TICKET))
    ATT_ValidateMaterialDocs = True
    For r = FIRST_ROW To lr
        mat = ATT_Clean(ws.Cells(r, C_MAT).value)
        doc = ATT_Clean(ws.Cells(r, C_DOC).value)
        po = ATT_Clean(ws.Cells(r, C_PO).value)
        ticket = ATT_Clean(ws.Cells(r, C_TICKET).value)
        If doc = "" And (mat <> "" Or po <> "" Or ticket <> "") Then
            ws.Cells(r, C_CHECK).value = "INPUT_MISSING_DOC"
            ws.Cells(r, C_NOTE).value = "Thieu Material DOC cot B; hay dien DOC de MIGO co the chay."
            missingCount = missingCount + 1
        End If
    Next r
    If missingCount > 0 Then
        ATT_ValidateMaterialDocs = False
        MsgBox missingCount & " dong dang thieu Material DOC o cot B." & vbCrLf & _
               "Hay dien DOC, sau do bam CHECK lai de tool mo MIGO.", _
               vbExclamation, "CAN DIEN MATERIAL DOC"
    End If
End Function

Public Sub OCR_CheckFiles()
    Dim ws As Worksheet, groups As Object, files As Collection
    Dim folder As String, t0 As Double
    If Not ToolTryBegin("CHECK TEN FILE") Then Exit Sub
    t0 = Timer: On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    mStop = False: ATT_ClearCheckOutput ws
    If Not ATT_ValidateMaterialDocs(ws) Then GoTo SafeExit
    Set groups = ATT_BuildGroups(ws)
    If groups.count = 0 Then Err.Raise vbObjectError + 7001, , "Khong co dong hop le tu dong " & FIRST_ROW & "."
    folder = Trim$(mSourceFolder)
    If folder = "" Then folder = Trim$(ATT_Cfg("SOURCE_FOLDER", ""))
    If folder = "" Or Dir$(folder, vbDirectory) = "" Then
        folder = ATT_PickFolder()
        If folder = "" Then GoTo SafeExit
        mSourceFolder = folder
    End If
    Set files = New Collection
    ATT_ListFiles folder, ATT_BoolCfg("SCAN_SUBFOLDERS", True), files
    If files.count = 0 Then Err.Raise vbObjectError + 7003, , "Thu muc khong co PDF/JPG/JPEG/PNG/TIF/TIFF."
    ATT_ListFilesNoOcr files
    ATT_MatchFileNamesToInput groups, files
    ATT_WriteGroups ws, groups
    ATT_Log "CHECK", "DONE_FILENAME", "", "Da doi chieu ten " & files.count & " file voi " & groups.count & " nhom INPUT."
SafeExit:
    LastMigoTotalSeconds = ATT_Elapsed(t0)
    ToolEnd
    Application.StatusBar = False
    Exit Sub
EH:
    ATT_Log "CHECK", "FAIL", "", CStr(Err.Number) & " - " & Err.description
    MsgBox "CHECK TEN FILE loi " & Err.Number & ": " & Err.description, vbCritical, "TOOL BBBG/VAT"
    Resume SafeExit
End Sub
Public Sub CheckFiles_NoOCR()
    OCR_CheckFiles
End Sub
Private Sub ATT_ListFilesNoOcr(ByVal files As Collection)
    Dim p As Variant, typ As String, ticket As String
    For Each p In files
        typ = ATT_Classify(ATT_FileName(CStr(p)), "")
        ticket = ATT_TicketFromFileName(ATT_FileName(CStr(p)))
        FAST_Log "FILE_SCAN", typ, "", "FILE=" & ATT_FileName(CStr(p)) & ";TICKET=" & ticket
    Next p
End Sub
Private Sub ATT_MatchFileNamesToInput(ByVal groups As Object, ByVal files As Collection)
    Dim k As Variant, g As Object, p As Variant, inputTicket As String, fileTicket As String
    Dim b As Collection, v As Collection, u As Collection, typ As String, pathList As String
    For Each k In groups.keys
        Set g = groups(k): Set b = New Collection: Set v = New Collection: Set u = New Collection
        inputTicket = ATT_Clean(CStr(g("TICKET")))
        g("HEAD_STATUS") = "CHUA_KIEM_TRA_MIGO"
        If inputTicket = "" Then
            g("STATUS") = "NEED_SOPHIEU": g("NOTE") = "Can nhap So phieu trong INPUT de doi chieu ten file."
            GoTo ContinueGroupFilename
        End If
        Application.StatusBar = "Doi chieu ten file: " & inputTicket
        For Each p In files
            fileTicket = ATT_TicketFromFileName(ATT_FileName(CStr(p)))
            If ATT_TicketMatches(fileTicket, inputTicket) Or ATT_FileTicketMatch(ATT_FileName(CStr(p)), inputTicket) Then
                typ = ATT_Classify(ATT_FileName(CStr(p)), "")
                Select Case typ
                    Case "VAT": v.Add CStr(p)
                    Case "BBBG": b.Add CStr(p)
                    Case Else: u.Add CStr(p)
                End Select
            End If
        Next p
        If u.count > 0 Then
            g("STATUS") = "UNKNOWN_FILE_TYPE"
            g("NOTE") = u.count & " file dung So phieu nhung chua ro BBBG/VAT: " & inputTicket
            GoTo ContinueGroupFilename
        End If
        pathList = ""
        For Each p In b
            If pathList <> "" Then pathList = pathList & "|"
            pathList = pathList & CStr(p)
        Next p
        g("BBBG") = pathList
        If v.count > 0 Then g("VAT") = CStr(v(1)) Else g("VAT") = ""
        If b.count = 0 And v.count = 0 Then
            g("STATUS") = "NO_BBBG_VAT_MATCH": g("NOTE") = "Khong co BBBG/VAT co So phieu khop INPUT: " & inputTicket
        ElseIf b.count = 0 Then
            g("STATUS") = "PARTIAL_NO_BBBG": g("NOTE") = "Thieu BBBG; se ATT VAT neu HeadText MIGO dung."
        ElseIf v.count = 0 Then
            g("STATUS") = "PARTIAL_NO_VAT": g("NOTE") = b.count & " BBBG khop INPUT; thieu VAT, se ATT BBBG neu HeadText MIGO dung."
        ElseIf v.count > 1 Then
            g("STATUS") = "PARTIAL_MULTI_VAT": g("NOTE") = b.count & " BBBG khop; co nhieu VAT, bo qua VAT de tranh nham."
        Else
            g("STATUS") = "CHECK_OK": g("NOTE") = b.count & " BBBG + 1 VAT khop INPUT."
        End If
        g("OCR_TICKET") = inputTicket
ContinueGroupFilename:
        DoEvents
    Next k
End Sub

Private Function ATT_TicketFromFileName(ByVal fileName As String) As String
    Dim re As Object, ms As Object, s As String
    s = UCase$(ATT_BaseName(fileName))
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    re.pattern = "\bIN[ ]*[-_. ]*[ ]*([0-9]{2,24})\b"
    If re.Test(s) Then
        Set ms = re.Execute(s): ATT_TicketFromFileName = "IN - " & CStr(ms(0).SubMatches(0)): Exit Function
    End If
    re.pattern = "(^|[^0-9])([0-9]{2,24})([^0-9]|$)"
    If re.Test(s) Then
        Set ms = re.Execute(s): ATT_TicketFromFileName = "IN - " & CStr(ms(0).SubMatches(1))
    End If
End Function
Public Sub SAP_FillHeadTextOnly()
    Dim ws As Worksheet, docs As Object, ses As Object, k As Variant, d As Object
    Dim t0 As Double, errText As String, currentText As String, n As Long
    If Not ToolTryBegin("DIEN HEADTEXT") Then Exit Sub
    t0 = Timer: On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    Set docs = ATT_BuildDocTickets(ws)
    If docs.count = 0 Then Err.Raise vbObjectError + 7010, , "Khong co DOC + So phieu de xu ly."
    Set ses = ModuleSAP.ATT_GetMigoSession()
    If ses Is Nothing Then Err.Raise vbObjectError + 7011, , "Khong tim thay SAP session. Hay dang nhap SAP truoc."
    For Each k In docs.keys
        Set d = docs(k)
        If d("TICKETS").count = 0 Then
            ATT_SetDocStatus ws, CStr(k), C_HEAD, "NEED_SOPHIEU", "Can nhap So phieu tu MB51."
        ElseIf d("TICKETS").count > 1 Then
            ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_MULTI_SOPHIEU", "DOC co nhieu So phieu; khong tu ghi de HeadText."
        Else
            Application.StatusBar = "HeadText DOC " & CStr(k)
            errText = "": currentText = ""
            If ATT_OpenMigoDoc(ses, CStr(k), errText) Then
                If ATT_ReadHeadText(ses, currentText, errText) Then
                    If Trim$(currentText) = "" Then
                        If ATT_WriteHeadText(ses, ATT_FirstKey(d("TICKETS")), errText) Then
                            ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_DONE", ""
                            n = n + 1
                        Else
                            ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_FAIL", errText
                        End If
                    ElseIf ATT_SameToken(currentText, ATT_FirstKey(d("TICKETS"))) Then
                        ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_OK", "Da co dung So phieu."
                    Else
                        ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_EXISTS_DIFFERENT", "SAP=" & currentText
                    End If
                Else
                    ATT_SetDocStatus ws, CStr(k), C_HEAD, "HEADTEXT_READ_FAIL", errText
                End If
            Else
                ATT_SetDocStatus ws, CStr(k), C_HEAD, "OPEN_DOC_FAIL", errText
            End If
        End If
        DoEvents
    Next k
      MsgBox "Da xu ly HeadText cho " & n & " DOC.", vbInformation, "DIEN HEADTEXT"
SafeExit:
    LastMigoTotalSeconds = ATT_Elapsed(t0): ToolEnd: Application.StatusBar = False
    Exit Sub
EH:
    ATT_Log "HEADTEXT", "FAIL", "", CStr(Err.Number) & " - " & Err.description
      MsgBox "DIEN HEADTEXT loi " & Err.Number & ": " & Err.description, vbCritical
    Resume SafeExit
End Sub
Private Function ATT_WasPreviouslyAttached(ByVal docNo As String, ByVal attachmentType As String, ByVal filePath As String) As Boolean
    Dim ws As Worksheet, lr As Long, r As Long
    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_HISTORY)
    lr = ws.Cells(ws.rows.count, 2).End(xlUp).Row
    For r = lr To 2 Step -1
        If StrComp(Trim$(CStr(ws.Cells(r, 2).value)), Trim$(docNo), vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(ws.Cells(r, 4).value)), Trim$(attachmentType), vbTextCompare) = 0 And _
           StrComp(Trim$(CStr(ws.Cells(r, 5).value)), Trim$(filePath), vbTextCompare) = 0 And _
           UCase$(Trim$(CStr(ws.Cells(r, 6).value))) = "DONE" Then
            ATT_WasPreviouslyAttached = True
            Exit Function
        End If
    Next r
EH:
End Function

Public Sub SAP_AttachCheckedFiles()
    Dim ws As Worksheet, docs As Object, groups As Object, files As Collection, ses As Object, docKey As Variant, docInfo As Object
    Dim titles As Object, grpKey As Variant, g As Object, errText As String, gosCtx As Object
    Dim expected As Object, titleB As String, titleV As String, doneDocs As Long, t0 As Double, docAttachFailed As Boolean
    If Not ToolTryBegin("ATT SAP") Then Exit Sub
    t0 = Timer: On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    If Not ATT_ValidateMaterialDocs(ws) Then GoTo SafeExit
    Set ses = ModuleSAP.ATT_GetMigoSession()
    If ses Is Nothing Then Err.Raise vbObjectError + 7021, , "Khong tim thay SAP session."
    ATT_ActivateMigoSession ses
    If Trim$(mSourceFolder) = "" Then mSourceFolder = Trim$(ATT_Cfg("SOURCE_FOLDER", ""))
    If Trim$(mSourceFolder) = "" Or Dir$(mSourceFolder, vbDirectory) = "" Then
        mSourceFolder = ATT_PickFolder()
        If mSourceFolder = "" Then GoTo SafeExit
    End If
    Set files = New Collection
    ATT_ListFiles mSourceFolder, ATT_BoolCfg("SCAN_SUBFOLDERS", True), files
    If files.count = 0 Then Err.Raise vbObjectError + 7022, , "Thu muc khong co file de ATT."
    Set groups = ATT_BuildGroups(ws)
    If groups.count = 0 Then Err.Raise vbObjectError + 7023, , "Khong co DOC + So phieu hop le de ATT."
    ATT_ListFilesNoOcr files
    ATT_MatchFileNamesToInput groups, files
    ATT_WriteGroups ws, groups
    Set docs = ATT_BuildAttachDocs(ws)
    If docs.count = 0 Then Err.Raise vbObjectError + 7020, , "Khong co nhom HeadText + file hop le de ATT."
    For Each docKey In docs.keys
        Set docInfo = docs(docKey)
        Application.StatusBar = "ATT DOC " & CStr(docKey)
        errText = ""
        docAttachFailed = False
        If Not ATT_OpenMigoDoc(ses, CStr(docKey), errText) Then
            ATT_SetDocStatus ws, CStr(docKey), C_HEAD, "MIGO_DOC_OPEN_FAIL", errText
            ATT_SetDocStatus ws, CStr(docKey), C_ATT, "OPEN_DOC_FAIL", errText
            Exit For
        End If
        Dim liveHeadText As String, expectedHeadText As String
        liveHeadText = "": expectedHeadText = ""
        For Each grpKey In docInfo("GROUPS").keys
            Set g = docInfo("GROUPS")(grpKey): expectedHeadText = CStr(g("TICKET")): Exit For
        Next grpKey
        If Not ATT_ReadHeadText(ses, liveHeadText, errText) Then
            ATT_SetDocStatus ws, CStr(docKey), C_HEAD, "HEADTEXT_READ_FAIL", errText
            ATT_SetDocStatus ws, CStr(docKey), C_ATT, "HEADTEXT_READ_FAIL", errText
            Exit For
        End If
        If Not ATT_TicketMatches(liveHeadText, expectedHeadText) Then
            ATT_SetDocStatus ws, CStr(docKey), C_HEAD, "HEADTEXT_MISMATCH", "SAP=" & liveHeadText & "; INPUT=" & expectedHeadText
            ATT_SetDocStatus ws, CStr(docKey), C_ATT, "HEADTEXT_CHANGED", "SAP=" & liveHeadText & "; CHECK=" & expectedHeadText
            Exit For
        End If
        Set gosCtx = GOS_CreateContext(ses, CStr(docKey))
        Set titles = GOS_ReadAttachments(ses, gosCtx, errText)
        ATT_SetDocStatus ws, CStr(docKey), C_HEAD, "HEADTEXT_OK", ""
        If titles Is Nothing Then
            ATT_Log "ATT", "LIST_UNREADABLE_CONTINUE", CStr(docKey), errText
            Set titles = CreateObject("Scripting.Dictionary")
        End If
        Set expected = CreateObject("Scripting.Dictionary")
        For Each grpKey In docInfo("GROUPS").keys
            Set g = docInfo("GROUPS")(grpKey)
            titleB = ATT_FileName(CStr(g("BBBG")))
            titleV = ATT_FileName(CStr(g("VAT")))
            If Trim$(CStr(g("BBBG"))) <> "" Then expected(ATT_NormTitle(titleB)) = titleB
            If Trim$(CStr(g("VAT"))) <> "" Then expected(ATT_NormTitle(titleV)) = titleV
            Dim bParts As Variant, bIdx As Long, bPath As String, bAllOk As Boolean
            If Trim$(CStr(g("BBBG"))) = "" Then
                ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "NOT_FOUND"
            ElseIf GOS_HasType(titles, "BBBG") Or ATT_TitleExists(titles, titleB) Or ATT_WasPreviouslyAttached(CStr(docKey), "BBBG", CStr(g("BBBG"))) Then
                ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "ALREADY_ATTACHED"
            Else
                bParts = Split(CStr(g("BBBG")), "|"): bAllOk = True
                For bIdx = LBound(bParts) To UBound(bParts)
                    bPath = Trim$(CStr(bParts(bIdx)))
                    If bPath <> "" Then
                        If Not ATT_AttachOne(ses, bPath, titleB, errText, gosCtx) Then
                            bAllOk = False: ATT_AddNoteRows ws, g("ROWS"), "BBBG: " & errText
                        End If
                    End If
                Next bIdx
                If bAllOk Then
                    ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "ATT_SENT"
                Else
                    ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "ATT_PARTIAL_FAIL"
                    docAttachFailed = True
                    Exit For
                End If
            End If
            If Trim$(CStr(g("VAT"))) = "" Then
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "NOT_FOUND"
            ElseIf GOS_HasType(titles, "VAT") Or ATT_TitleExists(titles, titleV) Or ATT_WasPreviouslyAttached(CStr(docKey), "VAT", CStr(g("VAT"))) Then
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "ALREADY_ATTACHED"
            ElseIf ATT_AttachOne(ses, CStr(g("VAT")), titleV, errText, gosCtx) Then
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "ATT_SENT"
            Else
                ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "ATT_FAIL"
                ATT_AddNoteRows ws, g("ROWS"), "VAT: " & errText
                docAttachFailed = True
                Exit For
            End If
        Next grpKey
        If docAttachFailed Then Exit For
        If ATT_DocHasMissingVat(docInfo, titles) Then
            ATT_SetDocStatus ws, CStr(docKey), C_ATT, "ATT_PARTIAL_NO_VAT", "Da ATT BBBG; chua co file VAT."
        Else
            ATT_SetDocStatus ws, CStr(docKey), C_ATT, "ATT_DONE", ""
            doneDocs = doneDocs + 1
        End If
        ATT_MarkVerified ws, docInfo, titles
        If Not ATT_ReturnToMigoEntry(ses, errText) Then
            ATT_SetDocStatus ws, CStr(docKey), C_ATT, "ATT_DONE", "ATT da thanh cong, nhung chua ve man hinh nhap DOC: " & errText
            Exit For
        End If
NextDoc:
        DoEvents
    Next docKey
    ' ATT completion dialog suppressed for current-MIGO run
SafeExit:
    LastMigoTotalSeconds = ATT_Elapsed(t0): ToolEnd: Application.StatusBar = False
    Exit Sub
EH:
    ATT_Log "ATT", "FAIL", "", CStr(Err.Number) & " - " & Err.description
    ' ATT error dialog suppressed for current-MIGO run
    Resume SafeExit
End Sub

'====================== CHECK / GROUPING ======================
Private Function ATT_BuildGroups(ByVal ws As Worksheet) As Object
    Dim groups As Object, g As Object, r As Long, lr As Long, key As String
    Dim mat As String, doc As String, po As String, ticket As String
    Set groups = CreateObject("Scripting.Dictionary")
    lr = ATT_LastRow(ws, C_DOC)
    For r = FIRST_ROW To lr
        mat = ATT_Clean(ws.Cells(r, C_MAT).value)
        doc = ATT_Clean(ws.Cells(r, C_DOC).value)
        po = ATT_Clean(ws.Cells(r, C_PO).value)
        ticket = ATT_Clean(ws.Cells(r, C_TICKET).value)
        If mat & doc & po & ticket <> "" Then
            If doc = "" Or po = "" Or mat = "" Then
                ws.Cells(r, C_CHECK).value = "INPUT_MISSING"
                ws.Cells(r, C_NOTE).value = "Can Material + DOC + PO/SA."
            Else
                key = UCase$(doc & "|" & ticket)
                If Not groups.Exists(key) Then
                    Set g = CreateObject("Scripting.Dictionary")
                    g("DOC") = doc: g("TICKET") = ticket
                    Set g("PO") = CreateObject("Scripting.Dictionary")
                    Set g("MAT") = CreateObject("Scripting.Dictionary")
                    Set g("ROWS") = New Collection
                    g("BBBG") = "": g("VAT") = "": g("STATUS") = "": g("NOTE") = ""
                    groups.Add key, g
                End If
                Set g = groups(key)
                g("PO")(UCase$(po)) = True
                g("MAT")(ATT_NormCode(mat)) = True
                g("ROWS").Add r
            End If
        End If
    Next r
    Set ATT_BuildGroups = groups
End Function

Private Function ATT_MarkGroupsNeedTicket(ByVal groups As Object) As Long
    Dim k As Variant, g As Object
    For Each k In groups.keys
        Set g = groups(k)
        If Trim$(CStr(g("TICKET"))) = "" Then
            g("STATUS") = "NEED_SOPHIEU"
            g("NOTE") = "Nhap So phieu o cot D theo dang NCC IN - 4288 roi CHECK lai."
            ATT_MarkGroupsNeedTicket = ATT_MarkGroupsNeedTicket + 1
        End If
    Next k
End Function
Private Function ATT_ScanAndOcr(ByVal files As Collection, ByVal groups As Object) As Collection
    Dim infos As New Collection, p As Variant, info As Object, raw As String, stat As String, rows As Long
    rows = 1
    For Each p In files
        If mStop Then Exit For
        Application.StatusBar = "OCR " & rows & "/" & files.count & ": " & ATT_FileName(CStr(p))
        stat = "": raw = ATT_ReadDocument(CStr(p), stat)
        Set info = CreateObject("Scripting.Dictionary")
        info("PATH") = CStr(p): info("NAME") = ATT_FileName(CStr(p))
        info("RAW") = raw: info("OCR_STATUS") = stat
        info("TYPE") = ATT_Classify(info("NAME"), raw)
        Set info("PO") = ATT_ExtractPO(raw, groups)
        Set info("MAT") = ATT_ExtractMaterials(raw)
        info("TICKET") = ATT_ExtractTicket(raw)
        infos.Add info
        FAST_Log "OCR_SCAN", CStr(info("TYPE")), "", "FILE=" & CStr(info("NAME")) & ";STATUS=" & stat & ";TICKET=" & CStr(info("TICKET"))
        rows = rows + 1
        DoEvents
    Next p
    Set ATT_ScanAndOcr = infos
End Function

Private Sub ATT_WriteGroups(ByVal ws As Worksheet, ByVal groups As Object)
    Dim k As Variant, g As Object, r As Variant
    For Each k In groups.keys
        Set g = groups(k)
        For Each r In g("ROWS")
            ws.Cells(r, C_BBBG).value = g("BBBG")
            ws.Cells(r, C_VAT).value = g("VAT")
            If g.Exists("OCR_PO") Then ws.Cells(r, C_OCR_PO).value = g("OCR_PO")
            If g.Exists("OCR_MAT_B") Then ws.Cells(r, C_OCR_MAT_BBBG).value = g("OCR_MAT_B")
            If g.Exists("OCR_TICKET") Then ws.Cells(r, C_OCR_TICKET).value = g("OCR_TICKET")
            If g.Exists("OCR_MAT_V") Then ws.Cells(r, C_OCR_MAT_VAT).value = g("OCR_MAT_V")
            If g.Exists("HEAD_STATUS") Then ws.Cells(r, C_HEAD).value = g("HEAD_STATUS")
            ws.Cells(r, C_CHECK).value = g("STATUS")
            ws.Cells(r, C_NOTE).value = g("NOTE")
        Next r
    Next k
End Sub

'====================== OCR / FILE ======================
Private Function ATT_ReadDocument(ByVal path As String, ByRef status As String) As String
    Dim ext As String, tmp As String, txtExe As String, ps1 As String, f As String, out As String
    ext = LCase$(ATT_Ext(path)): status = ""
    If ext <> "pdf" Then
        ATT_ReadDocument = ATT_ReadPreparedImage(path, status): Exit Function
    End If
    tmp = ATT_TempFolder("PDF")
    txtExe = ATT_Cfg("PDFTOTEXT_EXE", "")
    If txtExe <> "" And Dir$(txtExe) <> "" Then
        f = tmp & "\pdf.txt"
        If ATT_RunWait(ATT_Q(txtExe) & " -layout " & ATT_Q(path) & " " & ATT_Q(f)) = 0 Then
            out = ATT_ReadText(f)
            If Len(Trim$(out)) >= CLng(Val(ATT_Cfg("PDF_TEXT_MIN_CHARS", "30"))) Then
                status = "PDF_TEXT": ATT_ReadDocument = out: Exit Function
            End If
        End If
    End If
    ps1 = ThisWorkbook.path & "\helpers\RenderPdf.ps1"
    If Dir$(ps1) = "" Then status = "PDF_RENDERER_MISSING": Exit Function
    If ATT_RunWait("powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & ATT_Q(ps1) & _
                   " -PdfPath " & ATT_Q(path) & " -OutputDir " & ATT_Q(tmp) & _
                   " -MaxPages " & ATT_Cfg("PDF_MAX_PAGES", "5")) <> 0 Then
        status = "PDF_RENDER_FAIL": Exit Function
    End If
    f = Dir$(tmp & "\page_*.png")
    Do While f <> ""
        out = out & vbCrLf & ATT_RunTesseract(tmp & "\" & f, status)
        f = Dir$()
    Loop
    If Len(Trim$(out)) = 0 Then status = "PDF_OCR_EMPTY" Else status = "PDF_OCR"
    ATT_ReadDocument = out
End Function

Private Function ATT_OcrPreflight(ByRef msg As String) As Boolean
    Dim p As String
    p = ATT_TesseractPath()
    If p = "" Then
        msg = "Khong tim thay tesseract.exe. Cap nhat CONFIG/TESSERACT_EXE.": Exit Function
    End If
    ATT_OcrPreflight = True
End Function

Private Function ATT_TesseractPath() As String
    Dim p As String
    p = ThisWorkbook.path & "\runtime\Tesseract-OCR\tesseract.exe"
    If Dir$(p) <> "" Then ATT_TesseractPath = p: Exit Function
    p = ATT_Cfg("TESSERACT_EXE", "")
    If p <> "" And Dir$(p) <> "" Then ATT_TesseractPath = p: Exit Function
    p = Environ$("LOCALAPPDATA") & "\Tesseract-OCR\tesseract.exe"
    If Dir$(p) <> "" Then ATT_TesseractPath = p: Exit Function
    p = Environ$("ProgramFiles") & "\Tesseract-OCR\tesseract.exe"
    If Dir$(p) <> "" Then ATT_TesseractPath = p
End Function
Private Function ATT_Classify(ByVal fileName As String, ByVal raw As String) As String
    Dim s As String
    s = ATT_NormText(fileName & " " & raw)
    If ATT_HasKeyword(fileName, ATT_Cfg("VAT_KEYWORDS", "VAT,HOADON,HOA DON,INVOICE")) Or _
       InStr(s, "HOA DON GIA TRI GIA TANG") > 0 Or InStr(s, "VAT INVOICE") > 0 Then
        ATT_Classify = "VAT"
    ElseIf ATT_HasKeyword(fileName, ATT_Cfg("BBBG_KEYWORDS", "BBBG,BBGH,BIENBAN,BIEN BAN,BAN GIAO")) Then
        ATT_Classify = "BBBG"
    Else
        ATT_Classify = "UNKNOWN"
    End If
End Function

Private Function ATT_ExtractTicket(ByVal raw As String) As String
    Dim re As Object, ms As Object, s As String
    s = UCase$(raw)
    Set re = CreateObject("VBScript.RegExp"): re.Global = False: re.IgnoreCase = True
    re.pattern = "(SO[ ]*PHIEU|PHIEU[ ]*SO|BBBG[ ]*(NO|NUMBER)?)[ ]*[:#-]?[ ]*([A-Z0-9_-]{2,24})"
    If re.Test(s) Then
        Set ms = re.Execute(s)
        ATT_ExtractTicket = CStr(ms(0).SubMatches(ms(0).SubMatches.count - 1))
    End If
End Function

Private Function ATT_ExtractMaterials(ByVal raw As String) As Object
    Dim d As Object, lines As Variant, x As Variant, compact As String, m As Object, i As Long, v As String
    Set d = CreateObject("Scripting.Dictionary")
    lines = Split(Replace(raw, vbCr, vbLf), vbLf)
    Set m = CreateObject("VBScript.RegExp"): m.Global = True: m.IgnoreCase = True
    m.pattern = "\bB[I1]W[0-9A-Z]{6,24}(AA|AB)\b"
    For Each x In lines
        compact = UCase$(CStr(x))
        compact = Replace(compact, " ", ""): compact = Replace(compact, "-", ""): compact = Replace(compact, "_", "")
        If m.Test(compact) Then
            For i = 0 To m.Execute(compact).count - 1
                v = UCase$(m.Execute(compact)(i).value)
                v = Replace(v, "B1W", "BIW")
                d(v) = True
            Next i
        End If
    Next x
    Set ATT_ExtractMaterials = d
End Function

Private Function ATT_RegexDict(ByVal text As String, ByVal pattern As String) As Object
    Dim d As Object, re As Object, ms As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp"): re.Global = True: re.IgnoreCase = True: re.pattern = pattern
    If re.Test(text) Then
        Set ms = re.Execute(text)
        For i = 0 To ms.count - 1: d(UCase$(ms(i).value)) = True: Next i
    End If
    Set ATT_RegexDict = d
End Function

Private Sub ATT_ListFiles(ByVal folder As String, ByVal recursive As Boolean, ByRef out As Collection)
    Dim fso As Object, fld As Object, fil As Object, subf As Object, ext As String
    Set fso = CreateObject("Scripting.FileSystemObject"): Set fld = fso.GetFolder(folder)
    For Each fil In fld.files
        ext = LCase$(fso.GetExtensionName(fil.path))
        Select Case ext
            Case "pdf", "jpg", "jpeg", "png", "tif", "tiff"
                out.Add fil.path
        End Select
    Next fil
    If recursive Then For Each subf In fld.SubFolders: ATT_ListFiles subf.path, True, out: Next subf
End Sub
Private Sub ATT_ActivateMigoSession(ByVal ses As Object)
    On Error Resume Next
    ses.findById("wnd[0]").maximize
    AppActivate "MIGO"
    On Error GoTo 0
End Sub
Private Function ATT_WaitForMigoEntryFast(ByVal ses As Object, ByVal fieldId As String, ByVal timeoutMs As Long, ByRef fieldObj As Object, ByRef errText As String) As Boolean
    Dim startedAt As Double, popup As Object, root As Object, candidate As Object
    Dim busy As Boolean, fallbackDone As Boolean, elapsedMs As Double
    Set fieldObj = Nothing
    startedAt = Timer
    Do
        Set popup = Nothing
        Set candidate = Nothing
        busy = False
        On Error Resume Next
        Set popup = ses.findById("wnd[1]", False)
        Set candidate = ses.findById(fieldId, False)
        busy = CBool(ses.busy)
        On Error GoTo 0
        If Not popup Is Nothing Then
            errText = "MIGO_EXIT_CONFIRMATION_REQUIRED"
            Exit Function
        End If
        If ATT_IsMigoDocField(candidate, fieldId) Then Set fieldObj = candidate
        If fieldObj Is Nothing And Not fallbackDone Then
            elapsedMs = ATT_Elapsed(startedAt) * 1000#
            If elapsedMs >= 250# Then
                fallbackDone = True
                On Error Resume Next
                Set root = ses.findById("wnd[0]/usr", False)
                On Error GoTo 0
                Set fieldObj = ATT_FindMigoDocFieldInTree(root, fieldId)
            End If
        End If
        If Not fieldObj Is Nothing And Not busy Then
            ATT_WaitForMigoEntryFast = True
            Exit Function
        End If
        elapsedMs = ATT_Elapsed(startedAt) * 1000#
        If elapsedMs >= CDbl(timeoutMs) Then
            errText = "MIGO_DOC_FIELD_NOT_READY"
            Exit Function
        End If
        DoEvents
        If busy Then ATT_Sleep 25 Else ATT_Sleep 10
    Loop
End Function
Private Function ATT_ReturnToMigoEntry(ByVal ses As Object, ByRef errText As String) As Boolean
    Dim o As Object, fieldId As String, t0 As Double
    On Error GoTo EH
    t0 = Timer
    fieldId = ATT_Cfg("MIGO_DOC_FIELD_ID", "wnd[0]/usr/ctxtGODYNPRO-MAT_DOC")
    ses.findById("wnd[0]").sendVKey 12
    If Not ATT_WaitForMigoEntryFast(ses, fieldId, 10000, o, errText) Then Exit Function
    FAST_Log "MIGO_READY", "DONE", GOS_CurrentDoc(ses), "ELAPSED_SEC=" & Format$(ATT_Elapsed(t0), "0.000")
    ATT_ReturnToMigoEntry = True
    Exit Function
EH:
    errText = "RETURN_TO_MIGO_ENTRY: " & Err.description
End Function
Private Function ATT_WaitForSapIdle(ByVal ses As Object, ByVal timeoutMs As Long, ByRef errText As String) As Boolean
    Dim startedAt As Double
    On Error GoTo EH
    startedAt = Timer
    Do While ses.busy
        ATT_Wait 100
        If ATT_Elapsed(startedAt) * 1000# >= timeoutMs Then
            errText = "SAP_UI_BUSY_TIMEOUT"
            Exit Function
        End If
    Loop
    ATT_WaitForSapIdle = True
    Exit Function
EH:
    errText = "SAP_UI_IDLE_CHECK: " & Err.description
End Function
Private Function ATT_FindChildByIdContains(ByVal obj As Object, ByVal needle As String) As Object
    Dim child As Object, found As Object
    On Error Resume Next
    If obj Is Nothing Then Exit Function
    If InStr(1, CStr(obj.id), needle, vbTextCompare) > 0 Then
        Set ATT_FindChildByIdContains = obj
        Exit Function
    End If
    For Each child In obj.Children
        Set found = ATT_FindChildByIdContains(child, needle)
        If Not found Is Nothing Then
            Set ATT_FindChildByIdContains = found
            Exit Function
        End If
    Next child
    On Error GoTo 0
End Function
Private Function ATT_FindChildByNameExact(ByVal obj As Object, ByVal targetName As String) As Object
    Dim child As Object, found As Object
    On Error Resume Next
    If obj Is Nothing Then Exit Function
    If StrComp(Trim$(CStr(obj.Name)), targetName, vbTextCompare) = 0 Then
        Set ATT_FindChildByNameExact = obj
        Exit Function
    End If
    For Each child In obj.Children
        Set found = ATT_FindChildByNameExact(child, targetName)
        If Not found Is Nothing Then
            Set ATT_FindChildByNameExact = found
            Exit Function
        End If
    Next child
    On Error GoTo 0
End Function
Private Function ATT_IsMigoDocField(ByVal o As Object, ByVal fieldId As String) As Boolean
    Dim typ As String, objId As String, nm As String
    On Error Resume Next
    If o Is Nothing Then Exit Function
    typ = CStr(o.Type)
    objId = CStr(o.id)
    nm = CStr(o.Name)
    On Error GoTo 0
    If InStr(1, UCase$(typ), "MENU", vbTextCompare) > 0 Then Exit Function
    If InStr(1, UCase$(typ), "TEXTFIELD", vbTextCompare) = 0 And InStr(1, UCase$(typ), "INPUTFIELD", vbTextCompare) = 0 Then Exit Function
    If InStr(1, objId, "GODYNPRO-MAT_DOC", vbTextCompare) = 0 Then Exit Function
    If InStr(1, objId, "/usr/", vbTextCompare) = 0 Then Exit Function
    If nm <> "" And StrComp(nm, "GODYNPRO-MAT_DOC", vbTextCompare) <> 0 Then Exit Function
    ATT_IsMigoDocField = True
End Function

Private Function ATT_FindMigoDocFieldInTree(ByVal obj As Object, ByVal fieldId As String) As Object
    Dim child As Object, found As Object
    On Error Resume Next
    If obj Is Nothing Then Exit Function
    If ATT_IsMigoDocField(obj, fieldId) Then
        Set ATT_FindMigoDocFieldInTree = obj
        Exit Function
    End If
    For Each child In obj.Children
        Set found = ATT_FindMigoDocFieldInTree(child, fieldId)
        If Not found Is Nothing Then
            Set ATT_FindMigoDocFieldInTree = found
            Exit Function
        End If
    Next child
    On Error GoTo 0
End Function

Private Function ATT_FindMigoDocField(ByVal ses As Object, ByVal fieldId As String) As Object
    Dim o As Object, root As Object
    On Error Resume Next
    Set o = ses.findById(fieldId, False)
    On Error GoTo 0
    If ATT_IsMigoDocField(o, fieldId) Then
        Set ATT_FindMigoDocField = o
        Exit Function
    End If
    Set o = Nothing
    On Error Resume Next
    Set root = ses.findById("wnd[0]/usr", False)
    On Error GoTo 0
    Set o = ATT_FindMigoDocFieldInTree(root, fieldId)
    If o Is Nothing Then
        On Error Resume Next
        Set root = ses.findById("wnd[0]", False)
        On Error GoTo 0
        Set o = ATT_FindMigoDocFieldInTree(root, fieldId)
    End If
    Set ATT_FindMigoDocField = o
End Function
Private Function ATT_WaitMigoDocLoaded(ByVal ses As Object, ByVal docNo As String, ByVal timeoutMs As Long, ByRef errText As String) As Boolean
    Dim startedAt As Double, caption As String, sType As String, sText As String
    Dim fieldId As String, docField As Object, headField As Object, fieldText As String, busy As Boolean
    Dim elapsedMs As Double, sleepMs As Long
    On Error GoTo EH
    startedAt = Timer
    fieldId = ATT_Cfg("MIGO_DOC_FIELD_ID", "wnd[0]/usr/ctxtGODYNPRO-MAT_DOC")
    Do
        caption = "": sType = "": sText = "": fieldText = "": busy = False
        Set docField = Nothing
        Set headField = Nothing
        On Error Resume Next
        busy = CBool(ses.busy)
        caption = CStr(ses.findById("wnd[0]").text)
        sType = CStr(ses.findById("wnd[0]/sbar").messageType)
        sText = CStr(ses.findById("wnd[0]/sbar").text)
        Set docField = ATT_FindMigoDocField(ses, fieldId)
        If Not docField Is Nothing Then fieldText = CStr(docField.text)
        Set headField = ses.findById("wnd[0]").findByName("GOHEAD-BKTXT", "GuiTextField")
        On Error GoTo EH
        If sType = "E" Or sType = "A" Or sType = "X" Then
            errText = "MIGO_LOAD_ERROR: " & sText
            Exit Function
        End If
        If Not busy Then
            If InStr(1, ATT_NormCode(caption), ATT_NormCode(docNo), vbTextCompare) > 0 Then
                ATT_WaitMigoDocLoaded = True
                Exit Function
            End If
            If Not docField Is Nothing Then
                If ATT_SameToken(fieldText, docNo) And Not headField Is Nothing Then
                    ATT_WaitMigoDocLoaded = True
                    Exit Function
                End If
            End If
        End If
        elapsedMs = ATT_Elapsed(startedAt) * 1000#
        If elapsedMs >= CDbl(timeoutMs) Then
            errText = "MIGO_DOC_LOAD_TIMEOUT: expected=" & docNo & "; field=" & fieldText & "; title=" & caption & "; sbar=" & sText
            Exit Function
        End If
        DoEvents
        If elapsedMs < 1000# Then sleepMs = 10 Else sleepMs = 25
        ATT_Sleep sleepMs
    Loop
EH:
    errText = "MIGO_DOC_LOAD_CHECK: " & Err.description
End Function
Private Function ATT_OpenMigoDoc(ByVal ses As Object, ByVal docNo As String, ByRef errText As String) As Boolean
    Dim root As Object, o As Object, fieldId As String, caption As String
    Dim requiredKey As String, currentKey As String, t0 As Double
    Dim stage As String, eNo As Long, eDesc As String
    On Error GoTo EH
    t0 = Timer
    fieldId = ATT_Cfg("MIGO_DOC_FIELD_ID", "wnd[0]/usr/ctxtGODYNPRO-MAT_DOC")

    stage = "READ_WINDOW"
    On Error Resume Next
    Err.Clear
    Set root = ses.findById("wnd[0]")
    eNo = Err.Number: eDesc = Err.description: Err.Clear
    On Error GoTo EH
    If eNo <> 0 Or root Is Nothing Then
        errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & IIf(eDesc <> "", eDesc, "WINDOW_NOT_FOUND")
        Exit Function
    End If

    stage = "READ_CAPTION"
    On Error Resume Next
    Err.Clear
    caption = CStr(root.text)
    eNo = Err.Number: eDesc = Err.description: Err.Clear
    On Error GoTo EH
    If eNo <> 0 Then
        errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & eDesc
        Exit Function
    End If

    If InStr(1, caption, "Material Document", vbTextCompare) > 0 Then
        stage = "EXIT_LOADED_DOC"
        On Error Resume Next
        Err.Clear
        root.sendVKey 12
        eNo = Err.Number: eDesc = Err.description: Err.Clear
        On Error GoTo EH
        If eNo <> 0 Then
            errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & eDesc
            Exit Function
        End If
        If Not ATT_WaitForMigoEntryFast(ses, fieldId, 10000, o, errText) Then Exit Function
        Set root = Nothing
        On Error Resume Next
        Set root = ses.findById("wnd[0]")
        On Error GoTo EH
        If root Is Nothing Then
            errText = "OPEN_MIGO_DOC_STAGE=RELOAD_WINDOW: WINDOW_NOT_FOUND"
            Exit Function
        End If
    End If

    stage = "ACTION_CONTROL"
    Set o = Nothing
    On Error Resume Next
    Err.Clear
    Set o = root.findByName("GODYNPRO-ACTION", "GuiComboBox")
    eNo = Err.Number: eDesc = Err.description: Err.Clear
    On Error GoTo EH
    If eNo <> 0 Then
        errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & eDesc
        Exit Function
    End If
    requiredKey = ATT_Cfg("MIGO_ACTION_KEY", "A12")
    If Not o Is Nothing Then
        stage = "ACTION_KEY"
        On Error Resume Next
        Err.Clear
        currentKey = CStr(o.key)
        If Err.Number = 0 And currentKey <> requiredKey Then o.key = requiredKey
        eNo = Err.Number: eDesc = Err.description: Err.Clear
        On Error GoTo EH
        If eNo <> 0 Then
            errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & eDesc
            Exit Function
        End If
        If currentKey <> requiredKey Then
            If Not ATT_WaitForSapIdle(ses, 10000, errText) Then Exit Function
        End If
    End If

    stage = "REFDOC_CONTROL"
    Set o = Nothing
    On Error Resume Next
    Err.Clear
    Set o = root.findByName("GODYNPRO-REFDOC", "GuiComboBox")
    eNo = Err.Number: eDesc = Err.description: Err.Clear
    On Error GoTo EH
    If eNo <> 0 Then
        errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & eDesc
        Exit Function
    End If
    requiredKey = ATT_Cfg("MIGO_REFDOC_KEY", "R02")
    If Not o Is Nothing Then
        stage = "REFDOC_KEY"
        On Error Resume Next
        Err.Clear
        currentKey = CStr(o.key)
        If Err.Number = 0 And currentKey <> requiredKey Then o.key = requiredKey
        eNo = Err.Number: eDesc = Err.description: Err.Clear
        On Error GoTo EH
        If eNo <> 0 Then
            errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & eDesc
            Exit Function
        End If
        If currentKey <> requiredKey Then
            If Not ATT_WaitForSapIdle(ses, 10000, errText) Then Exit Function
        End If
    End If

    stage = "REFRESH_WINDOW"
    On Error Resume Next
    Err.Clear
    Set root = ses.findById("wnd[0]")
    eNo = Err.Number: eDesc = Err.description: Err.Clear
    On Error GoTo EH
    If eNo <> 0 Or root Is Nothing Then
        errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & IIf(eDesc <> "", eDesc, "WINDOW_NOT_FOUND")
        Exit Function
    End If
    stage = "FIELD_LOOKUP"
    Set o = ATT_FindMigoDocField(ses, fieldId)
    If o Is Nothing Then
        errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": MIGO_DOC_FIELD_NOT_FOUND: " & fieldId
        Exit Function
    End If

    stage = "FIELD_FOCUS"
    On Error Resume Next
    Err.Clear
    o.SetFocus
    eNo = Err.Number: eDesc = Err.description: Err.Clear
    On Error GoTo EH
    If eNo <> 0 Then
        errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & eDesc
        Exit Function
    End If

    ' Ghi một lần. Việc xóa rồi ghi liên tiếp đã gây invalid argument trên một số trạng thái MIGO.
    stage = "FIELD_CLEAR"
    On Error Resume Next
    Err.Clear
    o.text = ""
    eNo = Err.Number: eDesc = Err.description: Err.Clear
    On Error GoTo EH
    If eNo <> 0 Then
        errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & eDesc
        Exit Function
    End If
    ATT_Wait 50

    stage = "FIELD_WRITE"
    On Error Resume Next
    Err.Clear
    o.text = CStr(docNo)
    eNo = Err.Number: eDesc = Err.description: Err.Clear
    On Error GoTo EH
    If eNo <> 0 Then
        errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & eDesc
        Exit Function
    End If
    If Not ATT_SameToken(CStr(o.text), CStr(docNo)) Then
        ATT_Wait 50
        On Error Resume Next
        Err.Clear
        o.text = CStr(docNo)
        eNo = Err.Number: eDesc = Err.description: Err.Clear
        On Error GoTo EH
        If eNo <> 0 Then
            errText = "OPEN_MIGO_DOC_STAGE=FIELD_REWRITE: " & eDesc
            Exit Function
        End If
        If Not ATT_SameToken(CStr(o.text), CStr(docNo)) Then
            errText = "MIGO_DOC_INPUT_MISMATCH: expected=" & CStr(docNo) & "; actual=" & CStr(o.text)
            Exit Function
        End If
    End If
    On Error Resume Next
    o.caretPosition = Len(CStr(docNo))
    Err.Clear
    On Error GoTo EH
    stage = "SUBMIT_REFRESH"
    On Error Resume Next
    Err.Clear
    Set root = ses.findById("wnd[0]")
    eNo = Err.Number: eDesc = Err.description: Err.Clear
    On Error GoTo EH
    If eNo <> 0 Or root Is Nothing Then
        errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & IIf(eDesc <> "", eDesc, "WINDOW_NOT_FOUND")
        Exit Function
    End If
    stage = "SUBMIT"
    On Error Resume Next
    Err.Clear
    root.sendVKey 0
    eNo = Err.Number: eDesc = Err.description: Err.Clear
    On Error GoTo EH
    If eNo <> 0 Then
        errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & eDesc
        Exit Function
    End If

    stage = "LOAD_WAIT"
    If Not ATT_WaitMigoDocLoaded(ses, docNo, 20000, errText) Then Exit Function
    FAST_Log "DOC_OPEN_READY", "DONE", docNo, "ELAPSED_SEC=" & Format$(ATT_Elapsed(t0), "0.000")
    ATT_OpenMigoDoc = True
    Exit Function
EH:
    errText = "OPEN_MIGO_DOC_STAGE=" & stage & ": " & Err.description
End Function

Private Function ATT_MigoHasLoadedDocument(ByVal ses As Object) As Boolean
    Dim o As Object, v As String
    On Error Resume Next
    Set o = ses.findById("wnd[0]/usr", False).findByName("GOHEAD-BKTXT", "GuiTextField")
    If Not o Is Nothing Then
        v = Trim$(CStr(o.text))
        ATT_MigoHasLoadedDocument = (v <> "")
    End If
    On Error GoTo 0
End Function
Private Function ATT_ReadHeadText(ByVal ses As Object, ByRef value As String, ByRef errText As String) As Boolean
    Dim o As Object, root As Object, startedAt As Double
    On Error GoTo EH
    value = ""
    On Error Resume Next
    Set o = ses.findById("wnd[0]").findByName("GOHEAD-BKTXT", "GuiTextField")
    On Error GoTo EH
    If o Is Nothing Then
        On Error Resume Next
        ses.SendCommand "MIGO_OK_HEADER_OPEN"
        On Error GoTo EH
        startedAt = Timer
        Do
            Set o = Nothing
            On Error Resume Next
            Set o = ses.findById("wnd[0]").findByName("GOHEAD-BKTXT", "GuiTextField")
            If o Is Nothing Then
                Set root = ses.findById("wnd[0]", False)
                Set o = ATT_FindChildByNameExact(root, "GOHEAD-BKTXT")
            End If
            On Error GoTo EH
            If Not o Is Nothing Then Exit Do
            If ATT_Elapsed(startedAt) * 1000# >= 1000# Then Exit Do
            DoEvents: ATT_Sleep 25
        Loop
    End If
    If o Is Nothing Then errText = "READ_HEADTEXT: GOHEAD-BKTXT_NOT_FOUND": Exit Function
    value = CStr(o.text)
    ATT_ReadHeadText = True
    Exit Function
EH:
    errText = "READ_HEADTEXT: " & Err.description
End Function
Private Function ATT_WriteHeadText(ByVal ses As Object, ByVal value As String, ByRef errText As String) As Boolean
    Dim o As Object, t As String
    On Error GoTo EH
    Set o = ses.findById("wnd[0]").findByName("GOHEAD-BKTXT", "GuiTextField")
    If o Is Nothing Then errText = "WRITE_HEADTEXT: GOHEAD-BKTXT_NOT_FOUND": Exit Function
    o.text = value
    ses.findById("wnd[0]").sendVKey 11
    If Not ATT_WaitForSapIdle(ses, 5000, errText) Then Exit Function
    t = CStr(ses.findById("wnd[0]/sbar").messageType)
    If t = "E" Or t = "A" Or t = "X" Then errText = CStr(ses.findById("wnd[0]/sbar").text): Exit Function
    ATT_WriteHeadText = True
    Exit Function
EH:
    errText = "WRITE_HEADTEXT: " & Err.description
End Function
Private Function ATT_FindAttachmentGrid(ByVal root As Object) As Object
    Dim child As Object, found As Object, typeName As String, subTypeName As String
    On Error Resume Next
    If root Is Nothing Then Exit Function
    typeName = CStr(root.Type)
    subTypeName = CStr(root.SubType)
    If typeName = "GuiGridView" Or (typeName = "GuiShell" And InStr(1, subTypeName, "Grid", vbTextCompare) > 0) Then
        Set ATT_FindAttachmentGrid = root
        Exit Function
    End If
    For Each child In root.Children
        Set found = ATT_FindAttachmentGrid(child)
        If Not found Is Nothing Then
            Set ATT_FindAttachmentGrid = found
            Exit Function
        End If
    Next child
    On Error GoTo 0
End Function
Private Function ATT_GetAttachmentTitles(ByVal ses As Object, ByRef errText As String) As Object
    Dim ctx As Object
    Set ctx = GOS_CreateContext(ses, GOS_CurrentDoc(ses))
    Set ATT_GetAttachmentTitles = GOS_ReadAttachments(ses, ctx, errText)
End Function
Private Function ATT_WaitSapWindow(ByVal ses As Object, ByVal windowIndex As Long, ByVal timeoutMs As Long) As Object
    Dim started As Double, elapsed As Double, w As Object, sleepMs As Long
    started = Timer
    Do
        Set w = Nothing
        On Error Resume Next
        Set w = ses.findById("wnd[" & CStr(windowIndex) & "]", False)
        On Error GoTo 0
        If Not w Is Nothing Then
            Set ATT_WaitSapWindow = w
            Exit Function
        End If
        elapsed = ATT_Elapsed(started) * 1000#
        If elapsed < 1000# Then sleepMs = 25 Else sleepMs = 50
        DoEvents
        ATT_Sleep sleepMs
        If elapsed >= CDbl(timeoutMs) Then Exit Do
    Loop
End Function
Private Function ATT_StatusIsAttachmentSuccess(ByVal messageType As String, ByVal messageText As String) As Boolean
    Dim t As String, hasObject As Boolean, hasSuccess As Boolean
    If UCase$(Trim$(messageType)) <> "S" Then Exit Function
    t = UCase$(Trim$(messageText))
    hasObject = (InStr(1, t, "ATTACH", vbTextCompare) > 0 Or InStr(1, t, "DOCUMENT", vbTextCompare) > 0 Or InStr(1, t, "FILE", vbTextCompare) > 0)
    hasSuccess = (InStr(1, t, "CREATED", vbTextCompare) > 0 Or InStr(1, t, "SAVED", vbTextCompare) > 0 Or InStr(1, t, "STORED", vbTextCompare) > 0 Or InStr(1, t, "SUCCESS", vbTextCompare) > 0 Or InStr(1, t, "THANH CONG", vbTextCompare) > 0)
    ATT_StatusIsAttachmentSuccess = (hasObject And hasSuccess)
End Function
Private Function ATT_WaitForAttachment(ByVal ses As Object, ByVal fileName As String, ByVal statusBefore As String, ByRef errText As String) As Boolean
    Dim startedAt As Double, sType As String, sText As String, statusNow As String
    On Error GoTo EH
    startedAt = Timer
    Do
        If Not ATT_WaitForSapIdle(ses, 10000, errText) Then Exit Function
        sType = CStr(ses.findById("wnd[0]/sbar").messageType): sText = CStr(ses.findById("wnd[0]/sbar").text)
        statusNow = sType & "|" & sText
        If (sType = "E" Or sType = "A" Or sType = "X") And statusNow <> statusBefore Then errText = "ATT_SAP_ERROR: " & sText: Exit Function
        If ATT_StatusIsAttachmentSuccess(sType, sText) Then
            If Not GOS_WaitPopupGone(ses, 5000, errText) Then Exit Function
            ATT_WaitForAttachment = True: Exit Function
        End If
        If ATT_Elapsed(startedAt) >= 30# Then errText = "ATT_SUCCESS_MESSAGE_TIMEOUT: file=" & fileName & "; status=" & statusNow: Exit Function
        ATT_Sleep 100
        DoEvents
    Loop
EH:
    errText = "ATT_WAIT_CONFIRM: " & Err.description
End Function
Private Function ATT_AttachOne(ByVal ses As Object, ByVal src As String, ByVal title As String, ByRef errText As String, Optional ByVal gosCtx As Object) As Boolean
    Dim w As Object, statusBefore As String, t0 As Double, stage As String, doc As String, method As String, controlId As String
    t0 = Timer: doc = GOS_CurrentDoc(ses): method = "DIRECT_TITLE_SHELL": controlId = "wnd[0]/titl/shellcont/shell"
    On Error GoTo EH
    stage = "FILE_NOT_FOUND"
    If Len(Dir$(src)) = 0 Then errText = "Khong tim thay file: " & src: GoTo Fail
    statusBefore = CStr(ses.findById("wnd[0]/sbar").messageType) & "|" & CStr(ses.findById("wnd[0]/sbar").text)
    stage = "CREATE_ACTION"
    If Not GOS_OpenAction(ses, "CREATE_ATTACHMENT", gosCtx, w, errText) Then GoTo Fail
    stage = "FILE_DIALOG_FAIL"
    If ATT_SetSapFileDialog(w, src) Then
        GOS_LogEvent "PATH_FILLED", "DONE", doc, "CREATE_ATTACHMENT", method, controlId, "wnd[1]", "FILE=" & ATT_FileName(src), "", "", Timer - t0, ses
    Else
        On Error Resume Next: w.Close: On Error GoTo EH
        stage = "WINDOWS_NATIVE_FALLBACK"
        GOS_LogEvent "FILE_DIALOG_FALLBACK", "START", doc, "CREATE_ATTACHMENT", "WINDOWS_NATIVE", controlId, "WINDOWS_DIALOG", "FILE=" & ATT_FileName(src), "", "", Timer - t0, ses
        If Not WinImportFile(src, 15) Then errText = "FILE_DIALOG_CONTROL_NOT_FOUND": GoTo Fail
        GOS_LogEvent "PATH_FILLED", "DONE", doc, "CREATE_ATTACHMENT", "WINDOWS_NATIVE", controlId, "WINDOWS_DIALOG", "FILE=" & ATT_FileName(src), "", "", Timer - t0, ses
    End If
    stage = "ATTACHMENT_CONFIRM_TIMEOUT"
    If Not ATT_WaitForAttachment(ses, ATT_FileName(src), statusBefore, errText) Then GoTo Fail
    GOS_LogEvent "UPLOAD_CONFIRMED", "SUCCESS", doc, "CREATE_ATTACHMENT", method, controlId, "wnd[1]", "FILE=" & ATT_FileName(src), "", "", Timer - t0, ses
    ATT_AttachOne = True
    Exit Function
Fail:
    GOS_LogEvent "CREATE_END", "FAIL", doc, "CREATE_ATTACHMENT", method, controlId, "wnd[1]", "FILE=" & ATT_FileName(src), stage, errText, Timer - t0, ses
    Exit Function
EH:
    errText = "ATT_FILE: " & Err.description
    GOS_LogEvent "CREATE_END", "FAIL", doc, "CREATE_ATTACHMENT", method, controlId, "wnd[1]", "FILE=" & ATT_FileName(src), stage, errText, Timer - t0, ses
End Function

Private Function ATT_SetSapFileDialog(ByVal w As Object, ByVal fullPath As String) As Boolean
    Dim fso As Object, p As String, n As String, o As Object
    Set fso = CreateObject("Scripting.FileSystemObject"): p = fso.GetParentFolderName(fullPath): n = fso.GetFileName(fullPath)
    On Error Resume Next
    Set o = Nothing
    Set o = w.findById("usr/ctxtDY_PATH")
    If Not o Is Nothing Then o.text = p
    Set o = Nothing
    Set o = w.findById("usr/ctxtDY_FILENAME"): If Not o Is Nothing Then o.text = n
    If Not o Is Nothing Then w.findById("tbar[0]/btn[0]").press: ATT_SetSapFileDialog = (Err.Number = 0): Exit Function
    Err.Clear: Set o = w.findById("usr/ctxtRLGRAP-FILENAME")
    If Not o Is Nothing Then o.text = fullPath: w.findById("tbar[0]/btn[0]").press: ATT_SetSapFileDialog = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function ATT_FindType(ByVal root As Object, ByVal typeName As String) As Object
    Dim i As Long, c As Object, x As Object
    On Error Resume Next
    If CStr(root.Type) = typeName Then Set ATT_FindType = root: Exit Function
    For i = 0 To root.Children.count - 1
        Set c = root.Children(i)
        Set x = ATT_FindType(c, typeName)
        If Not x Is Nothing Then Set ATT_FindType = x: Exit Function
    Next i
End Function
Private Function ATT_GridTitle(ByVal grid As Object, ByVal r As Long) As String
    Dim a As Variant, x As Variant, v As String, cols As Variant, header As String
    On Error Resume Next
    cols = grid.ColumnOrder
    For Each x In cols
        header = CStr(grid.GetColumnTitle(CStr(x)))
        If InStr(1, header, "TITLE", vbTextCompare) > 0 Or InStr(1, header, "DESCRIPTION", vbTextCompare) > 0 Then
            v = CStr(grid.GetCellValue(r, CStr(x)))
            If Trim$(v) <> "" Then ATT_GridTitle = v: Exit Function
        End If
    Next x
    On Error GoTo 0
    a = Array("BITM_DESCR", "OBJDES", "TITLE", "DESCRIPT", "FILENAME")
    For Each x In a
        On Error Resume Next: v = CStr(grid.GetCellValue(r, CStr(x))): On Error GoTo 0
        If Trim$(v) <> "" Then ATT_GridTitle = v: Exit Function
    Next x
End Function

'====================== DOC / ATT GROUPS ======================
Private Function ATT_BuildDocTickets(ByVal ws As Worksheet) As Object
    Dim d As Object, x As Object, r As Long, lr As Long, doc As String, t As String
    Set d = CreateObject("Scripting.Dictionary"): lr = ATT_LastRow(ws, C_DOC)
    For r = FIRST_ROW To lr
        doc = ATT_Clean(ws.Cells(r, C_DOC).value): t = ATT_Clean(ws.Cells(r, C_TICKET).value)
        If doc <> "" Then
            If Not d.Exists(doc) Then Set x = CreateObject("Scripting.Dictionary"): Set x("TICKETS") = CreateObject("Scripting.Dictionary"): d.Add doc, x
            Set x = d(doc): If t <> "" Then x("TICKETS")(UCase$(t)) = True
        End If
    Next r
    Set ATT_BuildDocTickets = d
End Function
Private Function ATT_BuildAttachDocs(ByVal ws As Worksheet) As Object
    Dim docs As Object, di As Object, groups As Object, g As Object, r As Long, lr As Long, key As String
    Dim doc As String, t As String, st As String
    Set docs = CreateObject("Scripting.Dictionary"): lr = ATT_LastRow(ws, C_DOC)
    For r = FIRST_ROW To lr
        st = UCase$(ATT_Clean(ws.Cells(r, C_CHECK).value))
        If st = "CHECK_OK" Or Left$(st, 8) = "PARTIAL_" Then
            doc = ATT_Clean(ws.Cells(r, C_DOC).value): t = ATT_Clean(ws.Cells(r, C_TICKET).value): key = UCase$(doc & "|" & t)
            If Not docs.Exists(doc) Then Set di = CreateObject("Scripting.Dictionary"): Set di("GROUPS") = CreateObject("Scripting.Dictionary"): docs.Add doc, di
            Set di = docs(doc): Set groups = di("GROUPS")
            If Not groups.Exists(key) Then
                Set g = CreateObject("Scripting.Dictionary"): g("TICKET") = t
                g("BBBG") = CStr(ws.Cells(r, C_BBBG).value): g("VAT") = CStr(ws.Cells(r, C_VAT).value)
                Set g("ROWS") = New Collection: groups.Add key, g
            End If
            Set g = groups(key): g("ROWS").Add r
        End If
    Next r
    Set ATT_BuildAttachDocs = docs
End Function
Private Function ATT_DocHasMissingVat(ByVal docInfo As Object, Optional ByVal titles As Object) As Boolean
    Dim k As Variant, g As Object, sapVAT As Boolean
    sapVAT = GOS_HasType(titles, "VAT")
    For Each k In docInfo("GROUPS").keys
        Set g = docInfo("GROUPS")(k)
        If Trim$(CStr(g("VAT"))) = "" And Not sapVAT Then ATT_DocHasMissingVat = True: Exit Function
    Next k
End Function
Private Sub ATT_MarkVerified(ByVal ws As Worksheet, ByVal docInfo As Object, Optional ByVal titles As Object)
    Dim k As Variant, g As Object, docNo As String, sapBBBG As Boolean, sapVAT As Boolean
    Dim localBBBG As Boolean, localVAT As Boolean, finalBBBG As Boolean, finalVAT As Boolean
    sapBBBG = GOS_HasType(titles, "BBBG"): sapVAT = GOS_HasType(titles, "VAT")
    For Each k In docInfo("GROUPS").keys
        Set g = docInfo("GROUPS")(k): docNo = CStr(ws.Cells(g("ROWS")(1), C_DOC).value)
        localBBBG = (Trim$(CStr(g("BBBG"))) <> ""): localVAT = (Trim$(CStr(g("VAT"))) <> "")
        finalBBBG = localBBBG Or sapBBBG: finalVAT = localVAT Or sapVAT
        If finalBBBG Then
            ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "DONE"
            If sapBBBG And Not localBBBG Then ATT_AddNoteRows ws, g("ROWS"), "EXISTS_SAP_BBBG"
            If localBBBG Then ATT_History docNo, CStr(g("TICKET")), "BBBG", CStr(g("BBBG")), "DONE"
        Else
            ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "NOT_FOUND"
        End If
        If finalVAT Then
            ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "DONE"
            If sapVAT And Not localVAT Then ATT_AddNoteRows ws, g("ROWS"), "EXISTS_SAP_VAT"
            If localVAT Then ATT_History docNo, CStr(g("TICKET")), "VAT", CStr(g("VAT")), "DONE"
        Else
            ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "MISSING_VAT"
        End If
        If finalBBBG And finalVAT Then
            ATT_SetRows ws, g("ROWS"), C_CHECK, "CHECK_OK"
        ElseIf finalVAT Then
            ATT_SetRows ws, g("ROWS"), C_CHECK, "PARTIAL_NO_BBBG"
        ElseIf finalBBBG Then
            ATT_SetRows ws, g("ROWS"), C_CHECK, "PARTIAL_NO_VAT"
        Else
            ATT_SetRows ws, g("ROWS"), C_CHECK, "NO_BBBG_VAT_MATCH"
        End If
    Next k
End Sub

'====================== GENERIC HELPERS ======================
Private Function ATT_Cfg(ByVal key As String, ByVal defaultValue As String) As String
    Dim ws As Worksheet, f As Range
    On Error GoTo x
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    Set f = ws.Columns(1).Find(What:=key, LookAt:=xlWhole, MatchCase:=False)
    If Not f Is Nothing Then ATT_Cfg = CStr(ws.Cells(f.Row, 2).value) Else ATT_Cfg = defaultValue
    Exit Function
x:  ATT_Cfg = defaultValue
End Function
Private Function ATT_BoolCfg(ByVal key As String, ByVal def As Boolean) As Boolean
    Dim s As String: s = UCase$(ATT_Cfg(key, IIf(def, "TRUE", "FALSE")))
    ATT_BoolCfg = (s = "TRUE" Or s = "1" Or s = "YES")
End Function
Private Function ATT_PickFolder() As String
    Dim fd As Object: Set fd = Application.FileDialog(4)
    fd.title = "Chon thu muc chua BBBG/VAT": fd.AllowMultiSelect = False
    If fd.Show = -1 Then ATT_PickFolder = fd.SelectedItems(1)
End Function
Private Function ATT_LastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    ATT_LastRow = ws.Cells(ws.rows.count, col).End(xlUp).Row
    If ATT_LastRow < FIRST_ROW Then ATT_LastRow = FIRST_ROW
End Function
Private Function ATT_Clean(ByVal v As Variant) As String
    ATT_Clean = Trim$(Replace(CStr(v), ChrW$(160), " "))
End Function
Private Function ATT_NormText(ByVal s As String) As String
    s = UCase$(s): s = Replace(s, "_", " "): s = Replace(s, "-", " "): s = Replace(s, ".", " ")
    s = Replace(s, vbTab, " "): s = Replace(s, vbCr, " "): s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    ATT_NormText = Trim$(s)
End Function
Private Function ATT_NormCode(ByVal s As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp"): re.Global = True: re.pattern = "[^A-Za-z0-9]"
    s = UCase$(re.Replace(s, "")): ATT_NormCode = Replace(s, "B1W", "BIW")
End Function
Private Function ATT_TicketCode(ByVal ticket As String) As String
    Dim re As Object, ms As Object, s As String
    s = ATT_NormText(ticket)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    re.pattern = "\bIN[ ]*([0-9]{2,24})\b"
    If re.Test(s) Then
        Set ms = re.Execute(s)
        ATT_TicketCode = "IN" & CStr(ms(0).SubMatches(0))
    Else
        ATT_TicketCode = ATT_NormCode(ticket)
    End If
End Function

Private Function ATT_TicketMatches(ByVal a As String, ByVal b As String) As Boolean
    Dim ca As String, cb As String
    ca = ATT_TicketCode(a): cb = ATT_TicketCode(b)
    If ca = "" Or cb = "" Then Exit Function
    ATT_TicketMatches = (ca = cb)
End Function
Private Function ATT_FileTicketMatch(ByVal fileName As String, ByVal ticket As String) As Boolean
    Dim re As Object, ms As Object, s As String, digits As String, compact As String
    s = UCase$(ATT_BaseName(fileName))
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    re.pattern = "\bIN[ ]*[-_. ]*[ ]*([0-9]{2,24})\b"
    If re.Test(UCase$(ticket)) Then
        Set ms = re.Execute(UCase$(ticket))
        digits = CStr(ms(0).SubMatches(0))
        re.pattern = "(^|[^0-9])IN[ ]*[-_. ]*[ ]*0*" & digits & "([^0-9]|$)"
        If re.Test(s) Then ATT_FileTicketMatch = True: Exit Function
        re.pattern = "(^|[^0-9])0*" & digits & "([^0-9]|$)"
        ATT_FileTicketMatch = re.Test(s)
        Exit Function
    End If
    compact = ATT_NormCode(s)
    ATT_FileTicketMatch = (InStr(1, compact, ATT_TicketCode(ticket), vbTextCompare) > 0)
End Function
Private Function ATT_TextHasToken(ByVal text As String, ByVal ticket As String) As Boolean
    ATT_TextHasToken = ATT_TokenSequence(ATT_NormText(text), ATT_NormText(ticket))
End Function
Private Function ATT_TokenSequence(ByVal hay As String, ByVal needle As String) As Boolean
    Dim h As Variant, n As Variant, i As Long, j As Long, ok As Boolean
    If needle = "" Then Exit Function
    h = Split(hay, " "): n = Split(needle, " ")
    For i = LBound(h) To UBound(h) - (UBound(n) - LBound(n))
        ok = True
        For j = LBound(n) To UBound(n)
            If ATT_NormCode(CStr(h(i + j))) <> ATT_NormCode(CStr(n(j))) Then ok = False: Exit For
        Next j
        If ok Then ATT_TokenSequence = True: Exit Function
    Next i
End Function
Private Function ATT_SameToken(ByVal a As String, ByVal b As String) As Boolean
    ATT_SameToken = (ATT_NormCode(a) = ATT_NormCode(b))
End Function
Private Function ATT_ContainsAll(ByVal have As Object, ByVal need As Object) As Boolean
    Dim k As Variant
    ATT_ContainsAll = True
    For Each k In need.keys
        If Not have.Exists(UCase$(CStr(k))) Then
            ATT_ContainsAll = False
            Exit Function
        End If
    Next k
End Function
Private Function ATT_Overlap(ByVal a As Object, ByVal b As Object) As Long
    Dim k As Variant
    For Each k In b.keys
        If a.Exists(UCase$(CStr(k))) Then ATT_Overlap = ATT_Overlap + 1
    Next k
End Function
Private Function ATT_JoinKeys(ByVal d As Object, ByVal SEP As String) As String
    Dim k As Variant
    For Each k In d.keys
        If ATT_JoinKeys <> "" Then ATT_JoinKeys = ATT_JoinKeys & SEP
        ATT_JoinKeys = ATT_JoinKeys & CStr(k)
    Next k
End Function
Private Function ATT_HasKeyword(ByVal s As String, ByVal csv As String) As Boolean
    Dim x As Variant
    For Each x In Split(csv, ",")
        If InStr(1, UCase$(s), Trim$(UCase$(CStr(x))), vbTextCompare) > 0 Then
            ATT_HasKeyword = True
            Exit Function
        End If
    Next x
End Function
Private Function ATT_Title(ByVal ticket As String, ByVal kind As String) As String
    Dim fmt As String
    If UCase$(kind) = "VAT" Then fmt = ATT_Cfg("VAT_TITLE_FORMAT", "{SOPHIEU} VAT") Else fmt = ATT_Cfg("BBBG_TITLE_FORMAT", "{SOPHIEU} BBBG")
    ATT_Title = Trim$(Replace(fmt, "{SOPHIEU}", ticket, 1, -1, vbTextCompare))
End Function
Private Function ATT_NormTitle(ByVal s As String) As String
    s = ATT_BaseName(s): ATT_NormTitle = ATT_NormCode(s)
End Function
Private Function ATT_TitleExists(ByVal titles As Object, ByVal expected As String) As Boolean
    Dim k As Variant, e As String: e = ATT_NormTitle(expected)
    For Each k In titles.keys: If InStr(1, CStr(k), e, vbTextCompare) > 0 Or InStr(1, e, CStr(k), vbTextCompare) > 0 Then ATT_TitleExists = True: Exit Function
    Next k
End Function
Private Function ATT_AllTitlesExist(ByVal titles As Object, ByVal expected As Object) As Boolean
    Dim k As Variant
    ATT_AllTitlesExist = True
    For Each k In expected.keys
        If Not ATT_TitleExists(titles, expected(k)) Then
            ATT_AllTitlesExist = False
            Exit Function
        End If
    Next k
End Function
Private Function ATT_FirstKey(ByVal d As Object) As String
    Dim a As Variant
    If d Is Nothing Then Exit Function
    If d.count = 0 Then Exit Function
    a = d.keys
    ATT_FirstKey = CStr(a(0))
End Function
Private Function ATT_FileName(ByVal p As String) As String
    ATT_FileName = CreateObject("Scripting.FileSystemObject").GetFileName(p)
End Function
Private Function ATT_BaseName(ByVal p As String) As String
    ATT_BaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(p)
End Function
Private Function ATT_Ext(ByVal p As String) As String
    ATT_Ext = CreateObject("Scripting.FileSystemObject").GetExtensionName(p)
End Function
Private Function ATT_SafeName(ByVal s As String) As String
    Dim x As Variant: For Each x In Array("\", "/", ":", "*", "?", """", "<", ">", "|"): s = Replace(s, CStr(x), "_"): Next x
    ATT_SafeName = s
End Function
Private Function ATT_Q(ByVal s As String) As String
    ATT_Q = Chr$(34) & s & Chr$(34)
End Function
Private Function ATT_RunWait(ByVal cmd As String) As Long
    ATT_RunWait = CreateObject("WScript.Shell").Run(cmd, 0, True)
End Function
Private Function ATT_ReadText(ByVal p As String) As String
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open: st.LoadFromFile p: ATT_ReadText = st.ReadText: st.Close
End Function
Private Function ATT_TempFolder(ByVal prefix As String) As String
    Dim p As String: p = Environ$("TEMP") & "\" & prefix & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Timer * 100), "00000000")
    MkDir p: ATT_TempFolder = p
End Function
Private Sub ATT_Wait(ByVal ms As Long)
    DoEvents
    ATT_Sleep ms
    DoEvents
End Sub
Private Function ATT_Elapsed(ByVal t0 As Double) As Double
    ATT_Elapsed = Timer - t0
    If ATT_Elapsed < 0 Then ATT_Elapsed = ATT_Elapsed + 86400#
End Function

Private Sub ATT_ClearCheckOutput(ByVal ws As Worksheet)
    Dim lr As Long: lr = ATT_LastRow(ws, C_DOC)
    If lr < FIRST_ROW Then Exit Sub
    ws.Range(ws.Cells(FIRST_ROW, C_BBBG), ws.Cells(lr, C_NOTE)).ClearContents
End Sub
Private Sub ATT_SetRows(ByVal ws As Worksheet, ByVal rows As Collection, ByVal col As Long, ByVal value As String)
    Dim r As Variant
    For Each r In rows
        ws.Cells(r, col).value = value
    Next r
End Sub
Private Sub ATT_AddNoteRows(ByVal ws As Worksheet, ByVal rows As Collection, ByVal msg As String)
    Dim r As Variant
    For Each r In rows
        If ws.Cells(r, C_NOTE).value <> "" Then
            ws.Cells(r, C_NOTE).value = ws.Cells(r, C_NOTE).value & " | "
        End If
        ws.Cells(r, C_NOTE).value = ws.Cells(r, C_NOTE).value & msg
    Next r
End Sub
Private Sub ATT_SetDocStatus(ByVal ws As Worksheet, ByVal doc As String, ByVal col As Long, ByVal status As String, ByVal note As String)
    Dim r As Long, lr As Long
    lr = ATT_LastRow(ws, C_DOC)
    For r = FIRST_ROW To lr
        If ATT_Clean(ws.Cells(r, C_DOC).value) = doc Then
            ws.Cells(r, col).value = status
            If note <> "" Then ATT_AddNoteRows ws, ATT_OneRow(r), note
        End If
    Next r
End Sub
Private Function ATT_OneRow(ByVal r As Long) As Collection
    Dim c As New Collection
    c.Add r
    Set ATT_OneRow = c
End Function
Private Sub ATT_Log(ByVal stage As String, ByVal status As String, ByVal doc As String, ByVal msg As String)
    Dim ws As Worksheet, r As Long: Set ws = ThisWorkbook.Worksheets(SH_LOG)
    r = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).value = Now: ws.Cells(r, 2).value = stage: ws.Cells(r, 3).value = status: ws.Cells(r, 4).value = doc: ws.Cells(r, 5).value = msg
End Sub
Private Sub ATT_History(ByVal doc As String, ByVal ticket As String, ByVal typ As String, ByVal path As String, ByVal result As String, Optional ByVal note As String = "")
    Dim ws As Worksheet, r As Long, sapUser As String: Set ws = ThisWorkbook.Worksheets(SH_HISTORY)
    r = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 1
    sapUser = mFastSapUser: If sapUser = "" Then sapUser = Application.UserName
    ws.Cells(r, 1).value = Now: ws.Cells(r, 2).value = doc: ws.Cells(r, 3).value = ticket: ws.Cells(r, 4).value = typ: ws.Cells(r, 5).value = path: ws.Cells(r, 6).value = result: ws.Cells(r, 7).value = sapUser
    If note <> "" Then ws.Cells(r, 8).value = note
End Sub

'====================== COMPATIBILITY STUBS ======================
Public Function ToolTryBegin(ByVal actionName As String) As Boolean
    If mToolBusy Then MsgBox "Tool dang chay. Hay doi luot hien tai ket thuc.", vbExclamation: Exit Function
    mToolBusy = True: ToolTryBegin = True
End Function
Public Sub ToolEnd()
    mToolBusy = False
End Sub
Public Function ToolRunIsActive() As Boolean
    ToolRunIsActive = mToolBusy
End Function
Public Function IsRunInProgress() As Boolean
    IsRunInProgress = mToolBusy
End Function
Public Sub ResetTransientState()
    mToolBusy = False
    mStop = False
End Sub
Public Sub GetExistingOBState(ByRef readyCount As Long, ByRef postedCount As Long, ByRef staleCount As Long)
    readyCount = 0
    postedCount = 0
    staleCount = 0
End Sub
Public Sub Run_MIGO311_Post_FirstReadyOB()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_Post_AllReadyOB()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_Post_FirstReadyOB_AllowYellow()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_Post_AllReadyOB_AllowYellow()
    ATT_DisabledOldMacro
End Sub
Public Sub Post_MIGO_311_FromOBGroups_V3()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_OneClick()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_Post_FirstReadyOB_UncheckRedOK()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_Post_AllReadyOB_UncheckRedOK()
    ATT_DisabledOldMacro
End Sub
Public Sub Run_MIGO311_Post_AllReadyOB_PartialRescue()
    ATT_DisabledOldMacro
End Sub
Private Sub ATT_DisabledOldMacro()
    MsgBox "Macro MIGO 311 cu khong dung trong tool ATT nay.", vbInformation
End Sub
Public Sub SelectSourceFolder_NoOCR()
    Dim p As String
    p = ATT_PickFolder()
    If p <> "" Then
        mSourceFolder = p
        MsgBox "Da chon thu muc BBBG/VAT:" & vbCrLf & p, vbInformation, "BBBG/VAT - NO OCR"
    End If
End Sub
Private Function ATT_TableTitle(ByVal tbl As Object, ByVal rowIndex As Long) As String
    On Error Resume Next
    ATT_TableTitle = CStr(tbl.GetCell(rowIndex, 1).text)
    If ATT_TableTitle = "" Then ATT_TableTitle = CStr(tbl.GetCell(rowIndex, 0).text)
    On Error GoTo 0
End Function
Public Sub RUN_FAST_ALL()
    Dim ws As Worksheet, ses As Object, docs As Object, d As Object
    Dim folder As String, k As Variant, t0 As Double, okCount As Long, errCount As Long, skipCount As Long
    Dim errText As String, state As String
    If Not ToolTryBegin("RUN FAST ALL") Then Exit Sub
    t0 = Timer: mFastRunId = Format$(Now, "yyyymmdd_hhnnss") & "_FAST"
    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    FAST_EnsureInputUI ws
    If Not ATT_ValidateMaterialDocs(ws) Then GoTo SafeExit
    Set docs = FAST_BuildDocQueue(ws)
    If docs.count = 0 Then Err.Raise vbObjectError + 7801, , "Khong co DOC hop le trong INPUT."
    FAST_UpdateDashboard ws, docs.count, 0, 0, 0, "PREPARING", 0
    mFastExcelSeconds = ATT_Elapsed(t0)
    FAST_Log "BUILD_QUEUE", "DONE", "", "DOC=" & docs.count
    folder = Trim$(mSourceFolder)
    If folder = "" Then folder = Trim$(ATT_Cfg("SOURCE_FOLDER", ""))
    If folder = "" Or Dir$(folder, vbDirectory) = "" Then
        folder = ATT_PickFolder()
        If folder = "" Then GoTo SafeExit
        mSourceFolder = folder
    End If
    Set mFastFiles = New Collection
    ATT_ListFiles folder, ATT_BoolCfg("SCAN_SUBFOLDERS", True), mFastFiles
    If mFastFiles.count = 0 Then Err.Raise vbObjectError + 7803, , "SOURCE_FOLDER khong co file BBBG/VAT."
    Set mFastFileIndex = FAST_BuildFileIndex(mFastFiles)
    FAST_MatchQueue docs
    FAST_WriteQueue ws, docs
    FAST_RefreshStatusStyles ws
    mFastScanSeconds = ATT_Elapsed(t0) - mFastExcelSeconds
    FAST_Log "SCAN_FOLDER", "DONE", "", "FILES=" & mFastFiles.count
    Set ses = ModuleSAP.ATT_GetMigoSession()
    If ses Is Nothing Then Err.Raise vbObjectError + 7804, , "Khong tim thay SAP session."
    ATT_ActivateMigoSession ses
    FAST_LoadHistory ws.Parent
    FAST_LoadGosProfile ses
    FAST_Log "SAP_CONNECT", "DONE", "", "SESSION_REUSED=TRUE"
    FAST_UpdateDashboard ws, docs.count, okCount, errCount, skipCount, "RUNNING", ATT_Elapsed(t0)
    For Each k In docs.keys
        Set d = docs(k): state = CStr(d("STATE"))
        If state = "DONE" Then
            FAST_Log "DOC", "SKIP_DONE", CStr(k), ""
        ElseIf Not FAST_DocReady(d) Then
            skipCount = skipCount + 1
            ATT_SetDocStatus ws, CStr(k), C_ATT, "SKIP_NOT_READY", "Khong co ticket hoac khong co file BBBG/VAT match."
            FAST_Log "DOC", "SKIP_NOT_READY", CStr(k), ""
        ElseIf Not FAST_DocSingleTicket(d) Then
            skipCount = skipCount + 1
            ATT_SetDocStatus ws, CStr(k), C_ATT, "SKIP_MULTI_SOPHIEU", "DOC co nhieu So phieu; khong tu dong ghi de HeadText."
            FAST_Log "DOC", "SKIP_MULTI_SOPHIEU", CStr(k), ""
        Else
            errText = ""
            If FAST_ProcessDoc(ses, CStr(k), d, errText) Then
                okCount = okCount + 1
            Else
                errCount = errCount + 1
                If FAST_IsSystemFatal(errText) Then Err.Raise vbObjectError + 7805, , errText
            End If
        End If
        DoEvents
    Next k
    mFastSapSeconds = ATT_Elapsed(t0) - mFastExcelSeconds - mFastScanSeconds
    FAST_RefreshStatusStyles ws
    FAST_UpdateDashboard ws, docs.count, okCount, errCount, skipCount, "DONE", ATT_Elapsed(t0)
    FAST_Log "RUN", "DONE", "", "OK=" & okCount & ";ERROR=" & errCount & ";SKIP=" & skipCount & ";TOTAL_SEC=" & Format$(ATT_Elapsed(t0), "0.00")
    MsgBox "RUN FAST ALL da ket thuc." & vbCrLf & _
           "DOC thanh cong: " & okCount & vbCrLf & "DOC loi: " & errCount & vbCrLf & _
           "DOC bo qua: " & skipCount, vbInformation, "ATT BBBG/VAT"
SafeExit:
    If Not ws Is Nothing Then FAST_RefreshStatusStyles ws
    LastMigoTotalSeconds = ATT_Elapsed(t0)
ToolEnd:     Application.StatusBar = False
    Exit Sub
EH:
    FAST_Log "RUN", "FATAL", "", CStr(Err.Number) & " - " & Err.description
    If Not ws Is Nothing Then
        If docs Is Nothing Then
            FAST_UpdateDashboard ws, 0, okCount, errCount, skipCount, "FATAL", ATT_Elapsed(t0)
        Else
            FAST_UpdateDashboard ws, docs.count, okCount, errCount, skipCount, "FATAL", ATT_Elapsed(t0)
        End If
    End If
    MsgBox "RUN FAST ALL loi: " & Err.description, vbCritical, "ATT BBBG/VAT"
    Resume SafeExit
End Sub

Public Sub OFFLINE_VALIDATE()
    Dim ws As Worksheet, docs As Object, folder As String, files As Collection
    Dim checks As Long, passed As Long, failures As String
    If Not ToolTryBegin("OFFLINE VALIDATE") Then Exit Sub
    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    checks = checks + 1
    If ATT_LastRow(ws, C_DOC) >= FIRST_ROW Then passed = passed + 1 Else failures = failures & "INPUT khong co du lieu; "
    checks = checks + 1
    Set docs = FAST_BuildDocQueue(ws)
    If docs.count > 0 Then passed = passed + 1 Else failures = failures & "Khong build duoc DocMap; "
    checks = checks + 1
    If FAST_RequiredSheetsOk(failures) Then passed = passed + 1
    checks = checks + 1
    If FAST_RequiredConfigOk(failures) Then passed = passed + 1
    checks = checks + 1
    If FAST_ButtonsOk(ws, failures) Then passed = passed + 1
    folder = Trim$(mSourceFolder)
    If folder = "" Then folder = Trim$(ATT_Cfg("SOURCE_FOLDER", ""))
    If folder <> "" And Dir$(folder, vbDirectory) <> "" Then
        Set files = New Collection
        ATT_ListFiles folder, ATT_BoolCfg("SCAN_SUBFOLDERS", True), files
        checks = checks + 1
        If files.count >= 0 Then passed = passed + 1 Else failures = failures & "File scan loi; "
        Set mFastFileIndex = FAST_BuildFileIndex(files)
        checks = checks + 1
        If Not mFastFileIndex Is Nothing Then passed = passed + 1 Else failures = failures & "FileIndex loi; "
    End If
    FAST_Log "OFFLINE_VALIDATE", IIf(failures = "", "PASS", "FAIL"), "", "PASS=" & passed & "/" & checks & "; " & failures
    If failures = "" Then
        MsgBox "OFFLINE VALIDATE PASS: " & passed & "/" & checks & " checks." & vbCrLf & _
               "Chua ket noi SAP va chua upload attachment.", vbInformation, "ATT BBBG/VAT"
    Else
        MsgBox "OFFLINE VALIDATE FAIL: " & passed & "/" & checks & vbCrLf & failures, vbExclamation, "ATT BBBG/VAT"
    End If
SafeExit:
ToolEnd:     Application.StatusBar = False
    Exit Sub
EH:
    FAST_Log "OFFLINE_VALIDATE", "FAIL", "", CStr(Err.Number) & " - " & Err.description
    MsgBox "OFFLINE VALIDATE loi: " & Err.description, vbCritical, "ATT BBBG/VAT"
    Resume SafeExit
End Sub

Private Function FAST_BuildDocQueue(ByVal ws As Worksheet) As Object
    Dim docs As Object, d As Object, groups As Object, g As Object
    Dim src As Variant, lr As Long, i As Long, rowNo As Long, doc As String, ticket As String, key As String, k As Variant
    Dim mat As String, po As String
    Set docs = CreateObject("Scripting.Dictionary"): docs.CompareMode = vbTextCompare
    lr = ATT_LastRow(ws, C_DOC)
    If lr < FIRST_ROW Then Set FAST_BuildDocQueue = docs: Exit Function
    src = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lr, 16)).Value2
    For i = 1 To UBound(src, 1)
        rowNo = FIRST_ROW + i - 1
        mat = FAST_ValueText(src(i, C_MAT)): doc = FAST_ValueText(src(i, C_DOC))
        po = FAST_ValueText(src(i, C_PO)): ticket = FAST_ValueText(src(i, C_TICKET))
        If mat <> "" And doc <> "" And po <> "" Then
            If Not docs.Exists(doc) Then
                Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
                d("DOC") = doc: Set d("GROUPS") = CreateObject("Scripting.Dictionary")
                Set d("ROWS") = New Collection: d("OPEN_COUNT") = 0: d("STATE") = "PENDING"
                docs.Add doc, d
            End If
            Set d = docs(doc): d("ROWS").Add rowNo
            key = UCase$(doc & "|" & FAST_TicketKey(ticket)): Set groups = d("GROUPS")
            If Not groups.Exists(key) Then
                Set g = CreateObject("Scripting.Dictionary"): g.CompareMode = vbTextCompare
                g("TICKET") = ticket: g("BBBG") = FAST_ValueText(src(i, C_BBBG))
                g("VAT") = FAST_ValueText(src(i, C_VAT)): Set g("ROWS") = New Collection
                groups.Add key, g
            End If
            Set g = groups(key): g("ROWS").Add rowNo
        End If
    Next i
    For Each k In docs.keys
        Set d = docs(k): d("STATE") = FAST_DocStateFromRows(ws, d("ROWS"))
    Next k
    Set FAST_BuildDocQueue = docs
End Function
Private Function FAST_DocStateFromRows(ByVal ws As Worksheet, ByVal rows As Collection) As String
    Dim r As Variant, s As String, allDone As Boolean
    allDone = True
    For Each r In rows
        s = UCase$(FAST_ValueText(ws.Cells(CLng(r), C_ATT).Value2))
        If s <> "DONE" And s <> "ALREADY_ATTACHED" And s <> "ATT_DONE" Then allDone = False
        If s = "ERROR" Or InStr(1, s, "FAIL", vbTextCompare) > 0 Then FAST_DocStateFromRows = "ERROR": Exit Function
    Next r
    If allDone Then FAST_DocStateFromRows = "DONE" Else FAST_DocStateFromRows = "PENDING"
End Function
Private Function FAST_BuildFileIndex(ByVal files As Collection) As Object
    Dim idx As Object, p As Variant, key As String, fn As String, rawTicket As String, typ As String
    Dim fileSize As Double, lastModified As String, hasBBBG As Boolean, hasVAT As Boolean
    Set idx = CreateObject("Scripting.Dictionary"): idx.CompareMode = vbTextCompare
    For Each p In files
        fn = ATT_FileName(CStr(p)): rawTicket = ATT_TicketFromFileName(fn): key = FAST_TicketKey(rawTicket)
        hasBBBG = FAST_FileMatchesType(fn, "BBBG")
        hasVAT = FAST_FileMatchesType(fn, "VAT")
        If hasBBBG And hasVAT Then
            typ = "BBBG+VAT"
        ElseIf hasBBBG Then
            typ = "BBBG"
        ElseIf hasVAT Then
            typ = "VAT"
        Else
            typ = "UNKNOWN"
        End If
        FAST_LogFileMatch fn, typ, rawTicket, key
        FAST_SafeFileMetadata CStr(p), fileSize, lastModified
        If key <> "" Then
            If hasBBBG Then FAST_AddFileIndexItem idx, key, CStr(p), fn, rawTicket, "BBBG", fileSize, lastModified
            If hasVAT Then FAST_AddFileIndexItem idx, key, CStr(p), fn, rawTicket, "VAT", fileSize, lastModified
        End If
    Next p
    Set FAST_BuildFileIndex = idx
End Function

Private Sub FAST_MatchQueue(ByVal docs As Object)
    Dim dk As Variant, gk As Variant, d As Object, g As Object, bucket As Collection, item As Object
    Dim key As String, typ As String, b As String, v As String
    For Each dk In docs.keys
        Set d = docs(dk)
        For Each gk In d("GROUPS").keys
            Set g = d("GROUPS")(gk): key = FAST_TicketKey(CStr(g("TICKET"))): b = "": v = ""
            If key = "" Then g("STATUS") = "NEED_SOPHIEU": GoTo NextGroup
            If mFastFileIndex.Exists(key) Then
                Set bucket = mFastFileIndex(key)
                For Each item In bucket
                    typ = UCase$(CStr(item("Type")))
                    If typ = "BBBG" Then
                        If b <> "" Then b = b & "|"
                        b = b & CStr(item("FullPath"))
                    ElseIf typ = "VAT" And v = "" Then
                        v = CStr(item("FullPath"))
                    End If
                Next item
            End If
            g("BBBG") = b: g("VAT") = v
            If b = "" And v = "" Then g("STATUS") = "NO_BBBG_VAT_MATCH"
            If b <> "" And v = "" Then g("STATUS") = "PARTIAL_NO_VAT"
            If b = "" And v <> "" Then g("STATUS") = "PARTIAL_NO_BBBG"
            If b <> "" And v <> "" Then g("STATUS") = "CHECK_OK"
NextGroup:
        Next gk
    Next dk
End Sub

Private Sub FAST_WriteQueue(ByVal ws As Worksheet, ByVal docs As Object)
    Dim dk As Variant, gk As Variant, d As Object, g As Object, r As Variant
    For Each dk In docs.keys
        Set d = docs(dk)
        For Each gk In d("GROUPS").keys
            Set g = d("GROUPS")(gk)
            For Each r In g("ROWS")
                ws.Cells(CLng(r), C_BBBG).value = g("BBBG")
                ws.Cells(CLng(r), C_VAT).value = g("VAT")
                ws.Cells(CLng(r), C_CHECK).value = g("STATUS")
            Next r
        Next gk
    Next dk
End Sub
Private Function FAST_ProcessDoc(ByVal ses As Object, ByVal docNo As String, ByVal d As Object, ByRef errText As String) As Boolean
    Dim liveHead As String, expectedHead As String, titles As Object, gk As Variant, g As Object
    Dim paths As Variant, p As Variant, title As String, ws As Worksheet, gosCtx As Object
    Dim indicatorState As String, indicatorDetail As String, gosMethod As String, gosControl As String
    Dim docOpened As Boolean, missingBBBG As Boolean, missingVAT As Boolean
    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    d("STATE") = "RUNNING": d("OPEN_COUNT") = CLng(d("OPEN_COUNT")) + 1
    FAST_Log "DOC_START", "RUNNING", docNo, "OPEN_COUNT=" & d("OPEN_COUNT")
    docOpened = FAST_OpenDocWithRetry(ses, docNo, errText)
    If Not docOpened Then GoTo Failed
    liveHead = "": If Not ATT_ReadHeadText(ses, liveHead, errText) Then GoTo Failed
    expectedHead = FAST_FirstGroupTicket(d)
    If expectedHead = "" Then errText = "NEED_SOPHIEU": GoTo Failed
    If Trim$(liveHead) = "" Then
        If Not ATT_WriteHeadText(ses, expectedHead, errText) Then GoTo Failed
    ElseIf Not FAST_TicketMatches(liveHead, expectedHead) Then
        errText = "HEADTEXT_DIFFERENT: SAP=" & liveHead & "; INPUT=" & expectedHead: GoTo Failed
    End If
    Set gosCtx = GOS_CreateContext(ses, docNo)
    If gosCtx Is Nothing Then errText = "GOS_CONTEXT_FAIL": GoTo Failed
    errText = ""
    If FAST_HistoryConfirmsComplete(docNo, d) Then
        Set titles = CreateObject("Scripting.Dictionary")
        FAST_Log "GOS_VIEW_SKIP", "HISTORY_CONFIRMED", docNo, "VIEW_SKIPPED=TRUE"
    Else
        gosMethod = "DIRECT_TITLE_SHELL": gosControl = "wnd[0]/titl/shellcont/shell"
        On Error Resume Next
        If gosCtx.Exists("METHOD") Then gosMethod = CStr(gosCtx("METHOD"))
        If gosCtx.Exists("CONTROL_ID") Then gosControl = CStr(gosCtx("CONTROL_ID"))
        On Error GoTo EH
        indicatorDetail = ""
        indicatorState = GOS_DisplayAttachmentState(ses, indicatorDetail)
        GOS_LogEvent "GOS_DISPLAY_BUTTON", IIf(indicatorState = "PRESENT", "PRESENT", "NONE"), docNo, "VIEW_ATTACHMENTS", gosMethod, gosControl, "", "STATE=" & indicatorState & "|DETAIL=" & indicatorDetail, "", "", 0, ses
        If indicatorState = "PRESENT" Then
            Set titles = GOS_ReadAttachments(ses, gosCtx, errText)
            If titles Is Nothing Or errText <> "" Then GoTo Failed
            FAST_RecordSapAttachmentEvidence docNo, d, titles
        Else
            Set titles = CreateObject("Scripting.Dictionary")
            FAST_Log "GOS_VIEW_SKIP", "NO_DISPLAY_ATTACHMENT_BUTTON", docNo, "VIEW_SKIPPED=TRUE|DETAIL=" & indicatorDetail
        End If
    End If
    For Each gk In d("GROUPS").keys
        Set g = d("GROUPS")(gk)
        If Trim$(CStr(g("TICKET"))) = "" Then GoTo NextGroup
        If Trim$(CStr(g("BBBG"))) <> "" Then
            paths = Split(CStr(g("BBBG")), "|")
            For Each p In paths
                If Trim$(CStr(p)) <> "" Then
                    title = ATT_FileName(CStr(p))
                    If Not GOS_HasType(titles, "BBBG") And Not ATT_TitleExists(titles, title) And Not FAST_HistoryHas(docNo, "BBBG", CStr(p)) Then
                        If Not FAST_AttachWithRetry(ses, CStr(p), title, errText, gosCtx) Then GoTo Failed
                        titles(ATT_NormTitle(title)) = title
                    End If
                End If
            Next p
        End If
        If Trim$(CStr(g("VAT"))) <> "" Then
            title = ATT_FileName(CStr(g("VAT")))
            If Not GOS_HasType(titles, "VAT") And Not ATT_TitleExists(titles, title) And Not FAST_HistoryHas(docNo, "VAT", CStr(g("VAT"))) Then
                If Not FAST_AttachWithRetry(ses, CStr(g("VAT")), title, errText, gosCtx) Then GoTo Failed
                titles(ATT_NormTitle(title)) = title
            End If
        End If
NextGroup:
    Next gk
    missingBBBG = FAST_DocHasMissingType(d, titles, "BBBG")
    missingVAT = FAST_DocHasMissingType(d, titles, "VAT")
    If missingBBBG And missingVAT Then
        ATT_SetDocStatus ws, docNo, C_ATT, "ATT_PARTIAL_MISSING_BBBG_VAT", "Missing BBBG and VAT source files."
        d("STATE") = "PARTIAL"
    ElseIf missingBBBG Then
        ATT_SetDocStatus ws, docNo, C_ATT, "ATT_PARTIAL_NO_BBBG", "VAT processed; BBBG source is missing."
        d("STATE") = "PARTIAL"
    ElseIf missingVAT Then
        ATT_SetDocStatus ws, docNo, C_ATT, "ATT_PARTIAL_NO_VAT", "BBBG processed; VAT source is missing."
        d("STATE") = "PARTIAL"
    Else
        ATT_SetDocStatus ws, docNo, C_ATT, "ATT_DONE", "FAST_ONE_DOC_PASS"
        d("STATE") = "DONE"
    End If
    FAST_MarkVerified ws, d, titles
    If Not ATT_ReturnToMigoEntry(ses, errText) Then GoTo Failed
    FAST_Log "DOC_END", CStr(d("STATE")), docNo, "OPEN_COUNT=" & d("OPEN_COUNT")
    FAST_ProcessDoc = True: Exit Function
Failed:
    d("STATE") = "ERROR": ATT_SetDocStatus ws, docNo, C_ATT, "ERROR", errText
    FAST_Log "DOC_END", "ERROR", docNo, errText
    If docOpened Then
        On Error Resume Next: ATT_ReturnToMigoEntry ses, errText: On Error GoTo 0
    Else
        FAST_Log "MIGO_READY", "SKIP", docNo, "REASON=DOC_OPEN_FAILED|NEXT_DOC_RESETS"
    End If
    Exit Function
EH:
    errText = "DOC_ERROR: " & Err.description: Resume Failed
End Function
Private Function FAST_IsDocOpenRetryable(ByVal errText As String) As Boolean
    Dim v As String
    v = UCase$(CStr(errText))
    FAST_IsDocOpenRetryable = (InStr(1, v, "MIGO_DOC_FIELD_NOT_READY", vbTextCompare) > 0 _
        Or InStr(1, v, "SAP_UI_BUSY_TIMEOUT", vbTextCompare) > 0 _
        Or InStr(1, v, "MIGO_DOC_LOAD_TIMEOUT", vbTextCompare) > 0 _
        Or InStr(1, v, "MIGO_DOC_LOAD_CHECK", vbTextCompare) > 0)
End Function

Private Function FAST_OpenDocWithRetry(ByVal ses As Object, ByVal docNo As String, ByRef errText As String) As Boolean
    Dim firstErr As String
    If ATT_OpenMigoDoc(ses, docNo, errText) Then
        FAST_OpenDocWithRetry = True
        Exit Function
    End If
    firstErr = errText
    If Not FAST_IsDocOpenRetryable(firstErr) Then
        FAST_Log "DOC_OPEN_RETRY", "SKIP", docNo, "REASON=NONRETRYABLE|ERROR=" & firstErr
        Exit Function
    End If
    ATT_Wait 100
    If ATT_OpenMigoDoc(ses, docNo, errText) Then FAST_OpenDocWithRetry = True
End Function

Private Function FAST_AttachWithRetry(ByVal ses As Object, ByVal path As String, ByVal title As String, ByRef errText As String, Optional ByVal gosCtx As Object) As Boolean
    If ATT_AttachOne(ses, path, title, errText, gosCtx) Then FAST_AttachWithRetry = True: Exit Function
    FAST_Log "GOS_RETRY", "FALLBACK", GOS_CurrentDoc(ses), "FILE=" & ATT_FileName(path) & "|METHOD=RETRY_CONTEXT|ERROR=" & errText
    If ATT_AttachOne(ses, path, title, errText, gosCtx) Then FAST_AttachWithRetry = True
End Function

Private Function FAST_FirstGroupTicket(ByVal d As Object) As String
    Dim k As Variant
    For Each k In d("GROUPS").keys: FAST_FirstGroupTicket = CStr(d("GROUPS")(k)("TICKET")): Exit Function: Next k
End Function

Private Sub FAST_LoadHistory(ByVal wb As Workbook)
    Dim ws As Worksheet, src As Variant, lr As Long, i As Long, key As String
    Set mFastHistory = CreateObject("Scripting.Dictionary"): mFastHistory.CompareMode = vbTextCompare
    Set ws = wb.Worksheets(SH_HISTORY): lr = ws.Cells(ws.rows.count, 2).End(xlUp).Row
    If lr < 2 Then Exit Sub
    src = ws.Range("B2:F" & lr).Value2
    For i = 1 To UBound(src, 1)
        If UCase$(FAST_ValueText(src(i, 5))) = "DONE" Then
            key = FAST_HistoryKey(FAST_ValueText(src(i, 1)), FAST_ValueText(src(i, 3)), FAST_ValueText(src(i, 4)))
            mFastHistory(key) = True
        End If
    Next i
End Sub

Private Function FAST_HistoryKey(ByVal docNo As String, ByVal typ As String, ByVal path As String) As String
    FAST_HistoryKey = UCase$(Trim$(docNo) & "|" & Trim$(typ) & "|" & Trim$(path))
End Function

Private Function FAST_HistoryHas(ByVal docNo As String, ByVal typ As String, ByVal path As String) As Boolean
    If mFastHistory Is Nothing Then Exit Function
    FAST_HistoryHas = mFastHistory.Exists(FAST_HistoryKey(docNo, typ, path))
End Function
Private Sub FAST_LoadGosProfile(ByVal ses As Object)
    Dim profileKey As String
    On Error Resume Next
    mFastSapUser = CStr(ses.info.User): If mFastSapUser = "" Then mFastSapUser = Application.UserName
    profileKey = FAST_GosProfileKey(ses)
    mFastGosProfile = ATT_Cfg("ATT_GOS_METHOD_" & profileKey, "AUTO")
    mLastGosControlID = ATT_Cfg("ATT_GOS_CONTROL_" & profileKey, "")
    mLastGosMethod = ATT_Cfg("ATT_GOS_METHOD_" & profileKey, "DIRECT_TITLE_SHELL")
    mLastGosControlType = ATT_Cfg("ATT_GOS_TYPE_" & profileKey, "GuiShell")
    mLastGosViewSelector = ATT_Cfg("ATT_GOS_VIEW_SELECTOR_" & profileKey, "")
    mLastGosCreateSelector = ATT_Cfg("ATT_GOS_CREATE_SELECTOR_" & profileKey, "")
    ' --- SHARED FALLBACK: new user with no user-specific profile ---
    If mLastGosControlID = "" Then
        mLastGosControlID = ATT_Cfg("ATT_GOS_CONTROL_SHARED", "")
        mLastGosMethod = ATT_Cfg("ATT_GOS_METHOD_SHARED", "DIRECT_TITLE_SHELL")
        mLastGosControlType = ATT_Cfg("ATT_GOS_TYPE_SHARED", "GuiShell")
    End If
    If mLastGosViewSelector = "" Then
        mLastGosViewSelector = ATT_Cfg("ATT_GOS_VIEW_SELECTOR_SHARED", "")
    End If
    If mLastGosCreateSelector = "" Then
        mLastGosCreateSelector = ATT_Cfg("ATT_GOS_CREATE_SELECTOR_SHARED", "")
    End If
    mLastGosResolved = (mLastGosControlID <> "")
    On Error GoTo 0
End Sub

Private Function FAST_IsSystemFatal(ByVal s As String) As Boolean
    FAST_IsSystemFatal = (InStr(1, s, "SAP session", vbTextCompare) > 0 Or InStr(1, s, "scripting", vbTextCompare) > 0 Or _
                          InStr(1, s, "disconnected", vbTextCompare) > 0 Or InStr(1, s, "invalid object", vbTextCompare) > 0)
End Function
Private Function FAST_RequiredSheetsOk(ByRef failures As String) As Boolean
    Dim x As Variant, ws As Worksheet
    For Each x In Array(SH_INPUT, SH_CONFIG, SH_HISTORY, SH_LOG)
        On Error Resume Next: Set ws = ThisWorkbook.Worksheets(CStr(x)): On Error GoTo 0
        If ws Is Nothing Then failures = failures & "Thieu sheet " & CStr(x) & "; "
        Set ws = Nothing
    Next x
    FAST_RequiredSheetsOk = (InStr(failures, "Thieu sheet") = 0)
End Function

Private Function FAST_RequiredConfigOk(ByRef failures As String) As Boolean
    Dim k As Variant, v As String
    For Each k In Array("MIGO_DOC_FIELD_ID", "MIGO_ACTION_KEY", "MIGO_REFDOC_KEY")
        v = ATT_Cfg(CStr(k), "")
        If v = "" Then failures = failures & "Thieu CONFIG " & CStr(k) & "; ": Exit Function
    Next k
    FAST_RequiredConfigOk = True
End Function

Private Function FAST_ButtonsOk(ByVal ws As Worksheet, ByRef failures As String) As Boolean
    Dim s As Shape, seen1 As Boolean, seen2 As Boolean
    For Each s In ws.Shapes
        If s.Name = "btnRunFastAll" And s.OnAction = "RUN_FAST_ALL" Then seen1 = True
        If s.Name = "btnOfflineValidate" And s.OnAction = "OFFLINE_VALIDATE" Then seen2 = True
    Next s
    If Not seen1 Then failures = failures & "Thieu button RUN_FAST_ALL; "
    If Not seen2 Then failures = failures & "Thieu button OFFLINE_VALIDATE; "
    FAST_ButtonsOk = seen1 And seen2
End Function

Private Function FAST_ValueText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    FAST_ValueText = Trim$(Replace(CStr(v), ChrW$(160), " "))
End Function

Private Sub FAST_Log(ByVal stage As String, ByVal status As String, ByVal doc As String, ByVal msg As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG): r = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).value = Now: ws.Cells(r, 2).value = stage: ws.Cells(r, 3).value = status
    ws.Cells(r, 4).value = doc: ws.Cells(r, 5).value = msg: ws.Cells(r, 7).value = mFastRunId
    On Error GoTo 0
End Sub


Public Function OFFLINE_VALIDATE_SILENT() As String
    Dim ws As Worksheet, docs As Object, failures As String, folder As String
    Dim files As Collection, passed As Long, checks As Long
    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SH_INPUT)
    checks = checks + 1: If ATT_LastRow(ws, C_DOC) >= FIRST_ROW Then passed = passed + 1 Else failures = failures & "INPUT;"
    checks = checks + 1: Set docs = FAST_BuildDocQueue(ws): If docs.count > 0 Then passed = passed + 1 Else failures = failures & "DOC_QUEUE;"
    checks = checks + 1: If FAST_RequiredSheetsOk(failures) Then passed = passed + 1
    checks = checks + 1: If FAST_RequiredConfigOk(failures) Then passed = passed + 1
    checks = checks + 1: If FAST_ButtonsOk(ws, failures) Then passed = passed + 1
    folder = Trim$(mSourceFolder): If folder = "" Then folder = Trim$(ATT_Cfg("SOURCE_FOLDER", ""))
    If folder <> "" And Dir$(folder, vbDirectory) <> "" Then
        Set files = New Collection: ATT_ListFiles folder, ATT_BoolCfg("SCAN_SUBFOLDERS", True), files
        checks = checks + 1: Set mFastFileIndex = FAST_BuildFileIndex(files): If Not mFastFileIndex Is Nothing Then passed = passed + 1 Else failures = failures & "FILE_INDEX;"
    End If
    OFFLINE_VALIDATE_SILENT = "PASS=" & passed & "/" & checks & ";FAILURES=" & failures & ";SAP=NOT_CALLED"
    Exit Function
EH:
    OFFLINE_VALIDATE_SILENT = "ERROR=" & Err.Number & ";" & Err.description
End Function
Private Function FAST_TicketKey(ByVal ticket As String) As String
    Dim re As Object, ms As Object, s As String, code As String
    s = UCase$(ATT_NormText(ticket))
    Set re = CreateObject("VBScript.RegExp"): re.Global = False: re.IgnoreCase = True
    re.pattern = "\bIN[ ]*[-_. ]*[ ]*([0-9]{2,24})\b"
    If re.Test(s) Then
        Set ms = re.Execute(s): code = CStr(ms(0).SubMatches(0))
    Else
        re.pattern = "(^|[^0-9])0*([0-9]{2,24})([^0-9]|$)"
        If re.Test(s) Then
            Set ms = re.Execute(s): code = CStr(ms(0).SubMatches(1))
        Else
            code = ATT_NormCode(s)
        End If
    End If
    Do While Len(code) > 1 And Left$(code, 1) = "0": code = Mid$(code, 2): Loop
    FAST_TicketKey = code
End Function

Private Function FAST_TicketMatches(ByVal a As String, ByVal b As String) As Boolean
    Dim ca As String, cb As String
    ca = FAST_TicketKey(a): cb = FAST_TicketKey(b)
    FAST_TicketMatches = (ca <> "" And cb <> "" And ca = cb)
End Function

Private Function FAST_TicketNormalizationOk() As Boolean
    FAST_TicketNormalizationOk = (FAST_TicketKey("4467") = "4467" And _
        FAST_TicketKey("IN - 4467") = "4467" And _
        FAST_TicketKey("Automech IN - 4467") = "4467" And _
        FAST_TicketMatches("Automech IN - 4467", "4467"))
End Function

Private Function FAST_DocReady(ByVal d As Object) As Boolean
    Dim k As Variant, g As Object, hasFile As Boolean, hasTicket As Boolean
    For Each k In d("GROUPS").keys
        Set g = d("GROUPS")(k)
        If Trim$(CStr(g("TICKET"))) <> "" Then hasTicket = True
        If Trim$(CStr(g("BBBG"))) <> "" Or Trim$(CStr(g("VAT"))) <> "" Then hasFile = True
    Next k
    FAST_DocReady = (hasTicket And hasFile)
End Function
Private Function FAST_DocHasMissingVat(ByVal d As Object, Optional ByVal titles As Object) As Boolean
    FAST_DocHasMissingVat = FAST_DocHasMissingType(d, titles, "VAT")
End Function



Private Function FAST_DocSingleTicket(ByVal d As Object) As Boolean
    Dim k As Variant, g As Object, ticket As String
    For Each k In d("GROUPS").keys
        Set g = d("GROUPS")(k)
        If Trim$(CStr(g("TICKET"))) = "" Then Exit Function
        If ticket = "" Then
            ticket = CStr(g("TICKET"))
        ElseIf Not ATT_TicketMatches(ticket, CStr(g("TICKET"))) Then
            Exit Function
        End If
    Next k
    FAST_DocSingleTicket = (ticket <> "")
End Function

Private Function FAST_MarkVerified(ByVal ws As Worksheet, ByVal d As Object, Optional ByVal titles As Object) As Boolean
    Dim k As Variant, g As Object, p As Variant, parts As Variant, docNo As String
    Dim sapBBBG As Boolean, sapVAT As Boolean, localBBBG As Boolean, localVAT As Boolean
    Dim finalBBBG As Boolean, finalVAT As Boolean
    sapBBBG = GOS_HasType(titles, "BBBG"): sapVAT = GOS_HasType(titles, "VAT")
    For Each k In d("GROUPS").keys
        Set g = d("GROUPS")(k): docNo = CStr(ws.Cells(g("ROWS")(1), C_DOC).value)
        localBBBG = (Trim$(CStr(g("BBBG"))) <> "")
        localVAT = (Trim$(CStr(g("VAT"))) <> "")
        finalBBBG = localBBBG Or sapBBBG
        finalVAT = localVAT Or sapVAT
        If finalBBBG Then
            ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "DONE"
            If sapBBBG And Not localBBBG Then ATT_AddNoteRows ws, g("ROWS"), "EXISTS_SAP_BBBG"
            If localBBBG Then
                parts = Split(CStr(g("BBBG")), "|")
                For Each p In parts
                    If Trim$(CStr(p)) <> "" Then ATT_History docNo, CStr(g("TICKET")), "BBBG", Trim$(CStr(p)), "DONE"
                Next p
            End If
        Else
            ATT_SetRows ws, g("ROWS"), C_ATT_BBBG, "NOT_FOUND"
        End If
        If finalVAT Then
            ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "DONE"
            If sapVAT And Not localVAT Then ATT_AddNoteRows ws, g("ROWS"), "EXISTS_SAP_VAT"
            If localVAT Then ATT_History docNo, CStr(g("TICKET")), "VAT", CStr(g("VAT")), "DONE"
        Else
            ATT_SetRows ws, g("ROWS"), C_ATT_VAT, "MISSING_VAT"
        End If
        If finalBBBG And finalVAT Then
            ATT_SetRows ws, g("ROWS"), C_CHECK, "CHECK_OK"
        ElseIf finalVAT Then
            ATT_SetRows ws, g("ROWS"), C_CHECK, "PARTIAL_NO_BBBG"
        ElseIf finalBBBG Then
            ATT_SetRows ws, g("ROWS"), C_CHECK, "PARTIAL_NO_VAT"
        Else
            ATT_SetRows ws, g("ROWS"), C_CHECK, "NO_BBBG_VAT_MATCH"
        End If
    Next k
    FAST_MarkVerified = True
End Function

Private Sub GOS_LogEvent(ByVal stage As String, ByVal status As String, ByVal doc As String, ByVal action As String, ByVal method As String, ByVal controlId As String, ByVal popup As String, ByVal result As String, ByVal errorStage As String, ByVal errText As String, ByVal elapsed As Double, ByVal ses As Object)
    Dim ws As Worksheet, r As Long, msg As String, runId As String
    Dim sapSystem As String, client As String, sapUser As String, sessId As String, tcode As String
    On Error Resume Next
    If Not ses Is Nothing Then
        sapSystem = CStr(ses.info.SystemName)
        client = CStr(ses.info.client)
        sapUser = CStr(ses.info.User)
        sessId = CStr(ses.id)
        tcode = CStr(ses.info.Transaction)
    End If
    On Error GoTo 0
    If sapUser = "" Then sapUser = Application.UserName
    runId = mFastRunId
    If runId = "" Then runId = "GOS_" & Format$(Now, "yyyymmdd_hhnnss")
    errText = Replace(Replace(Replace(CStr(errText), "|", "/"), vbCr, " "), vbLf, " ")
    msg = "ACTION=" & action & "|METHOD=" & method & "|CONTROL_ID=" & controlId & "|POPUP=" & popup & "|RESULT=" & result & "|ERROR_STAGE=" & errorStage & "|ERROR=" & errText & "|ELAPSED_SEC=" & Format$(elapsed, "0.000") & "|SAP_SYSTEM=" & sapSystem & "|CLIENT=" & client & "|SAP_USER=" & sapUser & "|WINDOWS_USER=" & Environ$("USERNAME") & "|COMPUTER=" & Environ$("COMPUTERNAME") & "|SESSION_ID=" & sessId & "|TCODE=" & tcode
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.rows.count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).value = Now
    ws.Cells(r, 2).value = stage
    ws.Cells(r, 3).value = status
    ws.Cells(r, 4).value = doc
    ws.Cells(r, 5).value = msg
    ws.Cells(r, 7).value = runId
    On Error GoTo 0
End Sub
Private Function GOS_CurrentDoc(ByVal ses As Object) As String
    Dim o As Object, fieldId As String
    On Error Resume Next
    fieldId = ATT_Cfg("ATT_DOC_FIELD_ID", ATT_Cfg("MIGO_DOC_FIELD_ID", "wnd[0]/usr/ctxtGODYNPRO-MAT_DOC"))
    Set o = ATT_FindMigoDocField(ses, fieldId)
    If Not o Is Nothing Then GOS_CurrentDoc = Trim$(CStr(o.text))
    On Error GoTo 0
End Function
Public Sub TEST_GOS()
    Dim ses As Object, titles As Object, errText As String, doc As String, t0 As Double, docMap As String, failStage As String
    If Not ToolTryBegin("TEST GOS") Then Exit Sub
    t0 = Timer: mFastRunId = "GOS_" & Format$(Now, "yyyymmdd_hhnnss")
    On Error GoTo EH
    Set ses = ModuleSAP.ATT_GetMigoSession()
    If ses Is Nothing Then
        GOS_LogEvent "GOS_TEST", "FAIL", "", "VIEW", "TEST_EXISTING_LIST", "", "", "", "SAP_SESSION", "Khong tim thay phien SAP MIGO.", Timer - t0, ses
        MsgBox "TEST GOS: khong tim thay phien SAP MIGO. Xem LOG.", vbExclamation
        GoTo SafeExit
    End If
    FAST_LoadGosProfile ses
    doc = GOS_CurrentDoc(ses)
    If doc = "" Then
        docMap = GOS_DocFieldMap(ses)
        GOS_LogEvent "MIGO_DOC_MAP", "INFO", "", "VIEW", "TEST_EXISTING_LIST", "", "", docMap, "DOC_NOT_FOUND", "Khong doc duoc so chung tu MIGO hien tai.", Timer - t0, ses
        GOS_LogEvent "GOS_TEST", "FAIL", "", "VIEW", "TEST_EXISTING_LIST", "", "", "PROFILE=" & FAST_GosProfileKey(ses), "DOC_NOT_FOUND", "Khong doc duoc so chung tu MIGO hien tai.", Timer - t0, ses
        MsgBox "TEST GOS: hay mo MIGO va nap DOC truoc. LOG da co MIGO_DOC_MAP.", vbExclamation
        GoTo SafeExit
    End If
    GOS_LogEvent "GOS_TEST", "START", doc, "VIEW", "TEST_MULTI_METHOD", mLastGosControlID, "", "PROFILE=" & FAST_GosProfileKey(ses) & "|CACHED_VIEW=" & mLastGosViewSelector, "", "", Timer - t0, ses
    Set titles = ATT_GetAttachmentTitles(ses, errText)
    If titles Is Nothing Or errText <> "" Then
        failStage = "LIST_NOT_FOUND"
        On Error Resume Next
        If Not mLastGosResult Is Nothing Then
            If mLastGosResult.Exists("ErrorStage") Then failStage = CStr(mLastGosResult("ErrorStage"))
        End If
        On Error GoTo EH
        If InStr(1, UCase$(errText), "MENU_ACTION_OR_POPUP_ALL_FAILED", vbTextCompare) > 0 Then failStage = "VIEW_COMMAND_NOT_EXECUTED"
        GOS_LogEvent "GOS_TEST", "FAIL", doc, "VIEW", "TEST_MULTI_METHOD", mLastGosControlID, "", "CACHED_VIEW=" & mLastGosViewSelector, failStage, errText, Timer - t0, ses
        If failStage = "VIEW_COMMAND_NOT_EXECUTED" Then
            GOS_LogEvent "GOS_MANUAL_NEXT_STEP", "INFO", doc, "VIEW", "MANUAL_CAPTURE_REQUIRED", mLastGosControlID, "", "NEXT=OPEN_DISPLAY_ATTACHMENT_MANUALLY_THEN_CLICK_CAPTURE_POPUP_MANUAL", failStage, "Menu scripting returned without an attachment popup.", Timer - t0, ses
            MsgBox "TEST GOS: SAP nhan lenh menu nhung khong mo Attachment popup." & vbCrLf & vbCrLf & "Buoc tiep theo: mo Display Attachment bang tay trong SAP, giu popup mo, sau do bam CAPTURE POPUP (MANUAL) tren sheet INPUT.", vbExclamation
        Else
            MsgBox "TEST GOS: da mo popup nhung khong doc duoc danh sach. Xem LOG: GOS_SAP_MAP.", vbExclamation
        End If
        GoTo SafeExit
    End If
    GOS_LogEvent "GOS_TEST", "SUCCESS", doc, "VIEW", mLastGosMethod, mLastGosControlID, "wnd[1]", "COUNT=" & CStr(titles.count) & "|VIEW_SELECTOR=" & mLastGosViewSelector & "|PROFILE=" & FAST_GosProfileKey(ses), "", "", Timer - t0, ses
    MsgBox "TEST GOS PASS: da luu launcher va VIEW selector cho SAP user hien tai. Attachment: " & CStr(titles.count), vbInformation
SafeExit:
    ToolEnd
    Exit Sub
EH:
    GOS_LogEvent "GOS_TEST", "FAIL", doc, "VIEW", "TEST_MULTI_METHOD", mLastGosControlID, "", "", "TEST_EXCEPTION", Err.description, Timer - t0, ses
    MsgBox "TEST GOS loi: " & Err.description & ". Xem LOG.", vbCritical
    Resume SafeExit
End Sub
Private Function GOS_TypeOfTitle(ByVal title As String) As String
    If GOS_TitleHasType(title, "BBBG") And GOS_TitleHasType(title, "VAT") Then
        GOS_TypeOfTitle = "BBBG+VAT"
    ElseIf GOS_TitleHasType(title, "VAT") Then
        GOS_TypeOfTitle = "VAT"
    ElseIf GOS_TitleHasType(title, "BBBG") Then
        GOS_TypeOfTitle = "BBBG"
    Else
        GOS_TypeOfTitle = "UNKNOWN"
    End If
End Function
Private Function GOS_HasType(ByVal titles As Object, ByVal typ As String) As Boolean
    Dim k As Variant, title As String
    If titles Is Nothing Then Exit Function
    On Error GoTo Done
    For Each k In titles.keys
        title = CStr(titles(k))
        If GOS_TitleHasType(title, typ) Then GOS_HasType = True: Exit Function
    Next k
Done:
End Function

Private Function GOS_TitleSummary(ByVal titles As Object) As String
    Dim k As Variant, title As String, typ As String, s As String, n As Long
    If titles Is Nothing Then GOS_TitleSummary = "COUNT=0": Exit Function
    On Error GoTo Done
    s = "COUNT=" & CStr(titles.count)
    For Each k In titles.keys
        title = Replace(Replace(Replace(CStr(titles(k)), "|", "/"), vbCr, " "), vbLf, " ")
        typ = GOS_TypeOfTitle(title)
        n = n + 1
        s = s & "|TITLE" & CStr(n) & "=" & title & "|TYPE" & CStr(n) & "=" & typ
    Next k
Done:
    GOS_TitleSummary = s
End Function

Private Sub FAST_LogFileMatch(ByVal fileName As String, ByVal typ As String, ByVal rawTicket As String, ByVal normalizedTicket As String)
    Dim result As String, reason As String, msg As String
    result = "MATCH": reason = ""
    If typ <> "BBBG" And typ <> "VAT" Then result = "NO_MATCH": reason = "CLASS_UNKNOWN"
    If normalizedTicket = "" Then result = "NO_MATCH": reason = "TICKET_NOT_FOUND"
    msg = "FILE=" & Replace(fileName, "|", "/") & "|CLASS=" & typ & "|TICKET_RAW=" & rawTicket & "|TICKET_NORMALIZED=" & normalizedTicket & "|MATCH=" & normalizedTicket & "|RESULT=" & result
    If reason <> "" Then msg = msg & "|REASON=" & reason
    FAST_Log "FILE_MATCH", result, "", msg
End Sub
Private Function GOS_CreateContext(ByVal ses As Object, ByVal doc As String) As Object
    Dim ctx As Object, candidate As Object, method As String, controlId As String, controlType As String
    Set ctx = CreateObject("Scripting.Dictionary"): ctx.CompareMode = vbTextCompare
    method = "DIRECT_TITLE_SHELL": controlId = "wnd[0]/titl/shellcont/shell": controlType = "GuiShell"
    On Error Resume Next
    If mLastGosResolved And mLastGosControlID <> "" Then
        Set candidate = ses.findById(mLastGosControlID, False)
        If Not candidate Is Nothing Then
            controlId = mLastGosControlID: method = mLastGosMethod: controlType = mLastGosControlType
        Else
            mLastGosResolved = False
        End If
    End If
    If candidate Is Nothing Then Set candidate = ses.findById(controlId, False)
    If Not candidate Is Nothing Then
        mLastGosResolved = True: mLastGosControlID = controlId: mLastGosMethod = method: mLastGosControlType = controlType
    End If
    On Error GoTo 0
    ctx("DOC") = doc: ctx("METHOD") = method: ctx("CONTROL_ID") = controlId: ctx("CONTROL_TYPE") = controlType: ctx("RESOLVED") = Not candidate Is Nothing
    If Not CBool(ctx("RESOLVED")) Then Set GOS_CreateContext = Nothing Else Set GOS_CreateContext = ctx
End Function
Private Function GOS_OpenAction(ByVal ses As Object, ByVal action As String, ByVal ctx As Object, ByRef popup As Object, ByRef errText As String) As Boolean
    Dim shell As Object, controlId As String, method As String, controlType As String, stage As String, doc As String, t0 As Double, result As Object
    Dim selectorUsed As String, selectionDetail As String
    t0 = Timer: doc = GOS_CurrentDoc(ses): method = "DIRECT_TITLE_SHELL": controlId = "wnd[0]/titl/shellcont/shell": controlType = "GuiShell"
    If Not ctx Is Nothing Then If ctx.Exists("CONTROL_ID") Then controlId = CStr(ctx("CONTROL_ID")): method = CStr(ctx("METHOD")): controlType = CStr(ctx("CONTROL_TYPE"))
    Set result = CreateObject("Scripting.Dictionary"): result.CompareMode = vbTextCompare
    result("Success") = False: result("MethodUsed") = method: result("ControlID") = controlId: result("ControlType") = controlType
    result("PopupWindow") = "": result("PopupType") = "": result("ErrorStage") = "": result("ErrorCode") = "": result("ElapsedMs") = 0: result("Action") = action
    Set mLastGosResult = result
    GOS_LogEvent IIf(UCase$(action) = "VIEW_ATTACHMENTS", "GOS_VIEW_START", "CREATE_START"), "START", doc, action, method, controlId, "", "", "", "", 0, ses
    On Error GoTo EH
    stage = "CONTROL_NOT_FOUND": Set shell = ses.findById(controlId, False)
    If shell Is Nothing Then errText = "CONTROL_NOT_FOUND: " & controlId: GoTo Fail
    GOS_LogEvent "LAUNCHER_FOUND", "DONE", doc, action, method, controlId, "", "", "", "", Timer - t0, ses
    GOS_LogSapMap ses, shell, doc, action, method, controlId, "PRE_SELECT"
    GOS_LogWindowState ses, doc, action, method, controlId, "BEFORE_SELECT"
    stage = "MENU_ACTION_OR_POPUP_ALL_FAILED"
    selectionDetail = "": selectorUsed = ""
    If Not GOS_SelectActionWithFallback(ses, shell, action, doc, method, controlId, popup, selectorUsed, selectionDetail) Then
        errText = "MENU_ACTION_OR_POPUP_ALL_FAILED: " & selectionDetail
        GoTo Fail
    End If
    result("MethodUsed") = method & ";" & selectorUsed
    GOS_LogEvent IIf(UCase$(action) = "VIEW_ATTACHMENTS", "VIEW_MENU_SELECTED", "CREATE_SELECTED"), "DONE", doc, action, method, controlId, "", "SELECTOR=" & selectorUsed, "", "", Timer - t0, ses
    result("Success") = True: result("PopupWindow") = "wnd[1]": result("PopupType") = CStr(popup.Type): result("ElapsedMs") = CLng((Timer - t0) * 1000)
    mLastGosResolved = True: mLastGosControlID = controlId: mLastGosMethod = method: mLastGosControlType = controlType
    GOS_SaveUserProfile ses, method, controlId, controlType
    GOS_SaveActionProfile ses, action, selectorUsed
    GOS_CommitProfile ses, doc, action, method, controlId, selectorUsed
    GOS_LogEvent IIf(UCase$(action) = "VIEW_ATTACHMENTS", "POPUP_DETECTED", "FILE_POPUP_DETECTED"), "DONE", doc, action, method, controlId, "wnd[1]", "SELECTOR=" & selectorUsed, "", "", Timer - t0, ses
    GOS_OpenAction = True: Exit Function
Fail:
    result("ErrorStage") = stage: result("ErrorCode") = GOS_ErrorCode(stage, errText): result("ElapsedMs") = CLng((Timer - t0) * 1000)
    mLastGosResolved = False
    GOS_LogWindowState ses, doc, action, method, controlId, "FAILED_" & stage
    GOS_LogWindowMap ses, doc, action, method, controlId, "FAILED_" & stage
    GOS_LogEvent IIf(UCase$(action) = "VIEW_ATTACHMENTS", "GOS_VIEW_END", "CREATE_END"), "FAIL", doc, action, method, controlId, "", "SELECTOR=" & selectorUsed & "|" & selectionDetail, stage, errText, Timer - t0, ses
    Exit Function
EH:
    errText = "GOS_OPEN: " & Err.description: result("ErrorStage") = stage: result("ErrorCode") = GOS_ErrorCode(stage, errText): result("ElapsedMs") = CLng((Timer - t0) * 1000)
    mLastGosResolved = False
    GOS_LogWindowState ses, doc, action, method, controlId, "EXCEPTION_" & stage
    GOS_LogWindowMap ses, doc, action, method, controlId, "EXCEPTION_" & stage
    GOS_LogEvent IIf(UCase$(action) = "VIEW_ATTACHMENTS", "GOS_VIEW_END", "CREATE_END"), "FAIL", doc, action, method, controlId, "", "SELECTOR=" & selectorUsed & "|" & selectionDetail, stage, errText, Timer - t0, ses
End Function
Private Function GOS_ReadAttachments(ByVal ses As Object, ByVal ctx As Object, ByRef errText As String) As Object
    Dim d As Object, grid As Object, tbl As Object, w As Object, r As Long, title As String, t0 As Double, doc As String
    Dim closeErr As String, listStage As String, probe As String, listInfo As String
    Set d = CreateObject("Scripting.Dictionary")
    errText = "": t0 = Timer: doc = GOS_CurrentDoc(ses)
    If Not GOS_OpenAction(ses, "VIEW_ATTACHMENTS", ctx, w, errText) Then
        Set GOS_ReadAttachments = Nothing
        Exit Function
    End If
    GOS_LogSapMap ses, w, doc, "VIEW_ATTACHMENTS", "POPUP_DISPATCH", "wnd[1]", "POPUP_OPEN"
    If GOS_WaitForAttachmentList(w, grid, tbl, listStage, probe) Then
        If Not grid Is Nothing Then
            listStage = "GRID_LIST"
            For r = 0 To CLng(grid.rowCount) - 1
                title = ATT_GridTitle(grid, r)
                If title <> "" Then d(ATT_NormTitle(title)) = title
            Next r
        ElseIf Not tbl Is Nothing Then
            listStage = "TABLE_LIST"
            For r = 0 To CLng(tbl.rowCount) - 1
                title = ATT_TableTitle(tbl, r)
                If title <> "" Then d(ATT_NormTitle(title)) = title
            Next r
        Else
            errText = "LIST_DISPATCH_ERROR: No supported control after wait."
        End If
    Else
        errText = "LIST_NOT_FOUND: Attachment list did not render within wait window. " & probe
        GOS_LogSapMap ses, w, doc, "VIEW_ATTACHMENTS", "POPUP_DISPATCH", "wnd[1]", "POPUP_LIST_TIMEOUT"
    End If
    listInfo = "STAGE=" & listStage & "|" & probe & "|" & GOS_TitleSummary(d)
    GOS_LogEvent "TITLES_READ", IIf(errText = "", "DONE", "FAIL"), doc, "VIEW_ATTACHMENTS", "POPUP_DISPATCH", "wnd[1]", "wnd[1]", listInfo, IIf(errText = "", "", "LIST_NOT_FOUND"), errText, Timer - t0, ses
    If errText <> "" Then GOS_LogSapMap ses, w, doc, "VIEW_ATTACHMENTS", "POPUP_DISPATCH", "wnd[1]", "POPUP_LIST_FAIL"
    On Error Resume Next
    w.Close
    On Error GoTo 0
    closeErr = ""
    If Not GOS_WaitPopupGone(ses, 2500, closeErr) Then
        GOS_LogEvent "POPUP_CLOSED", "WARN", doc, "VIEW_ATTACHMENTS", "POPUP_DISPATCH", "wnd[1]", "wnd[1]", "", "POPUP_CLOSE_TIMEOUT", closeErr, Timer - t0, ses
        If errText = "" Then errText = closeErr
    Else
        GOS_LogEvent "POPUP_CLOSED", "DONE", doc, "VIEW_ATTACHMENTS", "POPUP_DISPATCH", "wnd[1]", "", "", "", "", Timer - t0, ses
    End If
    If errText = "" Then
        GOS_LogEvent "GOS_VIEW_END", "SUCCESS", doc, "VIEW_ATTACHMENTS", "POPUP_DISPATCH", "wnd[0]/titl/shellcont/shell", "wnd[1]", GOS_TitleSummary(d), "", "", Timer - t0, ses
    End If
    Set GOS_ReadAttachments = d
End Function
Private Function GOS_WaitPopupGone(ByVal ses As Object, ByVal timeoutMs As Long, ByRef errText As String) As Boolean
    Dim t0 As Double, w As Object, busy As Boolean, settleStart As Double
    t0 = Timer: settleStart = 0
    Do
        Set w = Nothing
        busy = False
        On Error Resume Next
        Set w = ses.findById("wnd[1]", False)
        busy = CBool(ses.busy)
        On Error GoTo 0
        If w Is Nothing Then
            If Not busy Then
                GOS_WaitPopupGone = True
                Exit Function
            End If
            If settleStart = 0 Then settleStart = Timer
            If ATT_Elapsed(settleStart) * 1000# >= 250# Then
                GOS_WaitPopupGone = True
                Exit Function
            End If
        Else
            settleStart = 0
        End If
        DoEvents
        ATT_Sleep 25
        If ATT_Elapsed(t0) * 1000# >= CDbl(timeoutMs) Then
            errText = "POPUP_CLOSE_TIMEOUT"
            Exit Function
        End If
    Loop
End Function

Private Function GOS_ErrorCode(ByVal stage As String, ByVal errText As String) As String
    If InStr(1, errText, "CONTROL_NOT_FOUND", vbTextCompare) > 0 Then GOS_ErrorCode = "CONTROL_NOT_FOUND" Else GOS_ErrorCode = stage
End Function

Private Function GOS_LastActionResult() As Object
    Set GOS_LastActionResult = mLastGosResult
End Function

Private Function FAST_HistoryConfirmsComplete(ByVal docNo As String, ByVal d As Object) As Boolean
    Dim k As Variant, g As Object, hasExpected As Boolean
    For Each k In d("GROUPS").keys
        Set g = d("GROUPS")(k)
        If Trim$(CStr(g("BBBG"))) <> "" Then hasExpected = True: If Not FAST_HistoryHas(docNo, "BBBG", CStr(g("BBBG"))) Then Exit Function
        If Trim$(CStr(g("VAT"))) <> "" Then hasExpected = True: If Not FAST_HistoryHas(docNo, "VAT", CStr(g("VAT"))) Then Exit Function
    Next k
    FAST_HistoryConfirmsComplete = hasExpected
End Function


Private Sub FAST_SafeFileMetadata(ByVal filePath As String, ByRef fileSize As Double, ByRef lastModified As String)
    Dim fso As Object, f As Object
    fileSize = 0: lastModified = ""
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filePath)
    If Err.Number <> 0 Or f Is Nothing Then
        FAST_Log "FILE_METADATA", "WARN", "", "FILE=" & Replace(filePath, "|", "/") & "|RESULT=METADATA_FAIL|ERROR=" & CStr(Err.Number) & " - " & Err.description
        Err.Clear
    Else
        fileSize = CDbl(f.Size)
        lastModified = CStr(f.DateLastModified)
    End If
    On Error GoTo 0
End Sub
Private Sub GOS_IndicatorDiagnostic(ByVal ses As Object, ByVal doc As String)
    Dim runKey As String, ids As Variant, i As Long, id As String, obj As Object, detail As String
    runKey = mFastRunId
    If runKey = "" Then runKey = "GOS_DIAG_ONCE"
    If mGosDiagnosticRunId = runKey Then Exit Sub
    mGosDiagnosticRunId = runKey
    ids = Array("wnd[0]/titl/shellcont", "wnd[0]/titl/shellcont/shell", "wnd[0]/titl")
    detail = "RUN=" & runKey & "|"
    For i = LBound(ids) To UBound(ids)
        id = CStr(ids(i))
        Set obj = Nothing
        On Error Resume Next
        Set obj = ses.findById(id, False)
        On Error GoTo 0
        If obj Is Nothing Then
            detail = detail & id & "=MISSING|"
        Else
            detail = detail & id & "=" & GOS_DiagMeta(obj) & "|" & GOS_ToolbarDiagnostic(obj) & "|" & GOS_ChildrenDiagnostic(obj) & ";"
        End If
        If Len(detail) > 6000 Then Exit For
    Next i
    GOS_LogEvent "GOS_INDICATOR_PROBE", "DONE", doc, "VIEW_ATTACHMENTS", "DIAGNOSTIC_ONCE", "wnd[0]/titl/shellcont", "", Left$(detail, 6000), "", "", 0, ses
End Sub

Private Function GOS_DiagMeta(ByVal obj As Object) As String
    Dim s As String, v As String
    If obj Is Nothing Then Exit Function
    On Error Resume Next
    s = "TYPE=" & GOS_DiagValue(obj.Type) & ",SUBTYPE=" & GOS_DiagValue(obj.SubType) & ",ID=" & GOS_DiagValue(obj.id) & ",NAME=" & GOS_DiagValue(obj.Name) & ",ICON=" & GOS_DiagValue(obj.IconName) & ",TIP=" & GOS_DiagValue(obj.Tooltip) & ",ACC=" & GOS_DiagValue(obj.AccTooltip)
    On Error GoTo 0
    GOS_DiagMeta = s
End Function

Private Function GOS_DiagValue(ByVal value As Variant) As String
    Dim s As String
    On Error Resume Next
    s = CStr(value)
    s = Replace(s, "|", "/")
    s = Replace(s, ";", ",")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    GOS_DiagValue = Left$(s, 160)
End Function

Private Function GOS_ToolbarDiagnostic(ByVal obj As Object) As String
    Dim count As Long, i As Long, id As String, icon As String, tip As String, txt As String, s As String
    If obj Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear: count = CLng(CallByName(obj, "ButtonCount", VbGet))
    If Err.Number <> 0 Then
        Err.Clear
        GOS_ToolbarDiagnostic = "BCOUNT=UNSUPPORTED"
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    s = "BCOUNT=" & CStr(count)
    If count > 32 Then count = 32
    For i = 0 To count - 1
        id = "": icon = "": tip = "": txt = ""
        On Error Resume Next
        Err.Clear: id = CStr(CallByName(obj, "GetButtonId", VbMethod, i))
        Err.Clear: icon = CStr(CallByName(obj, "GetButtonIcon", VbMethod, i))
        Err.Clear: tip = CStr(CallByName(obj, "GetButtonTooltip", VbMethod, i))
        Err.Clear: txt = CStr(CallByName(obj, "GetButtonText", VbMethod, i))
        On Error GoTo 0
        If id <> "" Or icon <> "" Or tip <> "" Or txt <> "" Then
            s = s & "|B" & CStr(i) & "=" & GOS_DiagValue(id) & "," & GOS_DiagValue(icon) & "," & GOS_DiagValue(tip) & "," & GOS_DiagValue(txt)
        End If
        If Len(s) > 2500 Then Exit For
    Next i
    GOS_ToolbarDiagnostic = s
End Function

Private Function GOS_ChildrenDiagnostic(ByVal obj As Object) As String
    Dim col As Object, child As Object, count As Long, i As Long, s As String
    If obj Is Nothing Then Exit Function
    Set col = Nothing
    On Error Resume Next
    Set col = obj.Children
    Err.Clear: count = CLng(col.count)
    If Err.Number <> 0 Or col Is Nothing Then
        On Error GoTo 0
        GOS_ChildrenDiagnostic = "CHILDREN=UNSUPPORTED"
        Exit Function
    End If
    On Error GoTo 0
    s = "CHILDREN=" & CStr(count)
    If count > 16 Then count = 16
    For i = 0 To count - 1
        Set child = Nothing
        On Error Resume Next
        Set child = col.item(i)
        On Error GoTo 0
        If Not child Is Nothing Then
            s = s & "|C" & CStr(i) & "=" & GOS_DiagMeta(child)
        End If
        If Len(s) > 2500 Then Exit For
    Next i
    GOS_ChildrenDiagnostic = s
End Function
Private Function GOS_DisplayAttachmentText(ByVal value As String) As Boolean
    Dim s As String
    s = UCase$(Trim$(CStr(value)))
    s = Replace(s, " ", "")
    s = Replace(s, "_", "")
    s = Replace(s, "-", "")
    s = Replace(s, ".", "")
    GOS_DisplayAttachmentText = (InStr(1, s, "DISPLAYATTACHMENT", vbTextCompare) > 0)
End Function

Private Function GOS_FindDisplayAttachmentNode(ByVal obj As Object, ByVal depth As Long, ByRef detail As String) As Object
    Dim typ As String, objId As String, nm As String, txt As String, tip As String, acc As String, defTip As String
    Dim count As Long, i As Long, btip As String, btxt As String, bAcc As String, bid As String
    Dim col As Object, child As Object, hit As Object, d As String
    If obj Is Nothing Or depth > 8 Then Exit Function
    On Error Resume Next
    typ = CStr(obj.Type): objId = CStr(obj.id): nm = CStr(obj.Name)
    txt = CStr(obj.text): tip = CStr(obj.Tooltip): acc = CStr(obj.AccTooltip): defTip = CStr(obj.DefaultTooltip)
    On Error GoTo 0
    If GOS_DisplayAttachmentText(txt) Or GOS_DisplayAttachmentText(tip) Or GOS_DisplayAttachmentText(acc) Or GOS_DisplayAttachmentText(defTip) Then
        detail = "NODE_ID=" & objId & "|TYPE=" & typ & "|NAME=" & nm & "|TEXT=" & txt & "|TIP=" & tip & "|ACC=" & acc
        Set GOS_FindDisplayAttachmentNode = obj
        Exit Function
    End If
    On Error Resume Next
    Err.Clear: count = CLng(CallByName(obj, "ButtonCount", VbGet))
    If Err.Number <> 0 Then count = 0
    On Error GoTo 0
    If count > 0 And count < 128 Then
        For i = 0 To count - 1
            btip = "": btxt = "": bAcc = "": bid = ""
            On Error Resume Next
            Err.Clear: bid = CStr(CallByName(obj, "GetButtonId", VbMethod, i))
            Err.Clear: btip = CStr(CallByName(obj, "GetButtonTooltip", VbMethod, i))
            Err.Clear: btxt = CStr(CallByName(obj, "GetButtonText", VbMethod, i))
            Err.Clear: bAcc = CStr(CallByName(obj, "GetButtonAccTooltip", VbMethod, i))
            On Error GoTo 0
            If GOS_DisplayAttachmentText(btip) Or GOS_DisplayAttachmentText(btxt) Or GOS_DisplayAttachmentText(bAcc) Then
                detail = "BUTTON=" & CStr(i) & "|ID=" & bid & "|TIP=" & btip & "|TEXT=" & btxt & "|ACC=" & bAcc
                Set GOS_FindDisplayAttachmentNode = obj
                Exit Function
            End If
        Next i
    End If
    Set col = Nothing
    On Error Resume Next
    Set col = obj.Children
    Err.Clear: count = CLng(col.count)
    If Err.Number <> 0 Or col Is Nothing Then count = 0
    On Error GoTo 0
    If count > 64 Then count = 64
    For i = 0 To count - 1
        Set child = Nothing
        On Error Resume Next
        Set child = col.item(i)
        On Error GoTo 0
        If Not child Is Nothing Then
            d = "": Set hit = GOS_FindDisplayAttachmentNode(child, depth + 1, d)
            If Not hit Is Nothing Then
                detail = d
                Set GOS_FindDisplayAttachmentNode = hit
                Exit Function
            End If
        End If
    Next i
End Function

Private Function GOS_DisplayAttachmentState(ByVal ses As Object, ByRef detail As String) As String
    Dim roots As Variant, i As Long, root As Object, hit As Object, d As String, rootId As String
    detail = "": GOS_DisplayAttachmentState = "NONE"
    roots = Array("wnd[0]/titl", "wnd[0]/tbar[1]", "wnd[0]/tbar[0]")
    For i = LBound(roots) To UBound(roots)
        rootId = CStr(roots(i))
        Set root = Nothing
        On Error Resume Next
        Set root = ses.findById(rootId, False)
        On Error GoTo 0
        If Not root Is Nothing Then
            d = "": Set hit = GOS_FindDisplayAttachmentNode(root, 0, d)
            If Not hit Is Nothing Then
                detail = "ROOT=" & rootId & "|" & d
                GOS_DisplayAttachmentState = "PRESENT"
                Exit Function
            End If
        End If
    Next i
    detail = "NO_DISPLAY_ATTACHMENT_BUTTON"
End Function
Private Function GOS_AttachmentIndicatorState(ByVal ses As Object, ByVal ctx As Object, ByRef detail As String) As String
    Dim ids As Variant, i As Long, id As String, obj As Object, state As String, d As String
    detail = "": GOS_AttachmentIndicatorState = "UNKNOWN"
    ids = Array("wnd[0]/titl/shellcont", "wnd[0]/titl/shellcont/shell", "wnd[0]/titl")
    If Trim$(mLastGosIndicatorID) <> "" Then
        ids = Array(mLastGosIndicatorID, "wnd[0]/titl/shellcont", "wnd[0]/titl/shellcont/shell", "wnd[0]/titl")
    End If
    For i = LBound(ids) To UBound(ids)
        id = CStr(ids(i))
        Set obj = Nothing
        On Error Resume Next
        Set obj = ses.findById(id, False)
        On Error GoTo 0
        If Not obj Is Nothing Then
            state = GOS_IndicatorObjectState(obj, d)
            If state <> "UNKNOWN" Then
                detail = id & "|" & d
                GOS_AttachmentIndicatorState = state
                mLastGosIndicatorID = id
                Exit Function
            End If
            state = GOS_IndicatorChildrenState(obj, d)
            If state <> "UNKNOWN" Then
                detail = id & "|CHILD|" & d
                GOS_AttachmentIndicatorState = state
                mLastGosIndicatorID = id
                Exit Function
            End If
        End If
    Next i
    detail = "UNRESOLVED|CONTROL_ID=wnd[0]/titl/shellcont|SHELL=wnd[0]/titl/shellcont/shell"
End Function
Private Function GOS_IndicatorObjectState(ByVal obj As Object, ByRef detail As String) As String
    Dim icon As String, tip As String, accTip As String, accText As String, txt As String, defTip As String
    Dim typeName As String, subTypeName As String, signature As String
    Dim count As Long, i As Long, buttonIcon As String, buttonTip As String, buttonText As String
    Dim buttonSignature As String
    detail = "": GOS_IndicatorObjectState = "UNKNOWN"
    If obj Is Nothing Then Exit Function
    icon = "": tip = "": accTip = "": accText = "": txt = "": defTip = "": typeName = "": subTypeName = ""
    On Error Resume Next
    icon = CStr(obj.IconName)
    tip = CStr(obj.Tooltip)
    accTip = CStr(obj.AccTooltip)
    accText = CStr(obj.accText)
    txt = CStr(obj.text)
    defTip = CStr(obj.DefaultTooltip)
    typeName = CStr(obj.Type)
    subTypeName = CStr(obj.SubType)
    On Error GoTo 0
    signature = UCase$(icon)
    If InStr(1, signature, "ICON_GOS_SERVICES_ATT_REL", vbTextCompare) > 0 Or _
       InStr(1, signature, "ICON_GOS_SERVICES_ATTACHMENT", vbTextCompare) > 0 Then
        detail = "TYPE=" & typeName & "|SUBTYPE=" & subTypeName & "|ICON=" & icon & "|TIP=" & tip
        GOS_IndicatorObjectState = "PRESENT"
        Exit Function
    End If
    If InStr(1, signature, "ICON_GOS_SERVICES_RELATIONS", vbTextCompare) > 0 Or _
       StrComp(signature, "ICON_GOS_SERVICES", vbTextCompare) = 0 Then
        detail = "TYPE=" & typeName & "|SUBTYPE=" & subTypeName & "|ICON=" & icon & "|TIP=" & tip
        GOS_IndicatorObjectState = "NONE"
        Exit Function
    End If
    On Error Resume Next
    count = CLng(CallByName(obj, "ButtonCount", VbGet))
    On Error GoTo 0
    If count > 0 And count < 256 Then
        For i = 0 To count - 1
            buttonIcon = "": buttonTip = "": buttonText = ""
            On Error Resume Next
            buttonIcon = CStr(CallByName(obj, "GetButtonIcon", VbMethod, i))
            buttonTip = CStr(CallByName(obj, "GetButtonTooltip", VbMethod, i))
            buttonText = CStr(CallByName(obj, "GetButtonText", VbMethod, i))
            On Error GoTo 0
            buttonSignature = UCase$(buttonIcon)
            If InStr(1, buttonSignature, "ICON_GOS_SERVICES_ATT_REL", vbTextCompare) > 0 Or _
               InStr(1, buttonSignature, "ICON_GOS_SERVICES_ATTACHMENT", vbTextCompare) > 0 Then
                detail = "TYPE=" & typeName & "|BUTTON=" & CStr(i) & "|ICON=" & buttonIcon & "|TIP=" & buttonTip & "|TEXT=" & buttonText
                GOS_IndicatorObjectState = "PRESENT"
                Exit Function
            End If
            If InStr(1, buttonSignature, "ICON_GOS_SERVICES_RELATIONS", vbTextCompare) > 0 Or _
               StrComp(buttonSignature, "ICON_GOS_SERVICES", vbTextCompare) = 0 Then
                detail = "TYPE=" & typeName & "|BUTTON=" & CStr(i) & "|ICON=" & buttonIcon & "|TIP=" & buttonTip & "|TEXT=" & buttonText
                GOS_IndicatorObjectState = "NONE"
            End If
        Next i
        If GOS_IndicatorObjectState = "NONE" Then Exit Function
    End If
    detail = "TYPE=" & typeName & "|SUBTYPE=" & subTypeName & "|ICON=" & icon & "|TIP=" & tip & "|ACCTIP=" & accTip & "|ACCTEXT=" & accText & "|TEXT=" & txt & "|DEFAULT=" & defTip
End Function
Private Function GOS_IndicatorChildrenState(ByVal obj As Object, ByRef detail As String) As String
    Dim child As Object, state As String, d As String, childId As String
    detail = "": GOS_IndicatorChildrenState = "UNKNOWN"
    If obj Is Nothing Then Exit Function
    On Error Resume Next
    For Each child In obj.Children
        childId = "": childId = CStr(child.id)
        state = GOS_IndicatorObjectState(child, d)
        If state <> "UNKNOWN" Then
            detail = "CHILD_ID=" & childId & "|" & d
            GOS_IndicatorChildrenState = state
            Exit Function
        End If
    Next child
    On Error GoTo 0
End Function



Private Function GOS_HasSapToken(ByVal sourceText As String, ByVal token As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False: re.IgnoreCase = True
    re.pattern = "(^|[^A-Z0-9])" & token & "([^A-Z0-9]|$)"
    GOS_HasSapToken = re.Test(sourceText)
End Function

