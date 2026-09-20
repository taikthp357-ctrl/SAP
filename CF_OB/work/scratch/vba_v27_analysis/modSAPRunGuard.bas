Attribute VB_Name = "modSAPRunGuard"
Option Explicit

' Central contract for SAP MAP_MODE / RUN_MODE separation.
'
' MAP_MODE is the only mode that may discover controls or use a mapper.
' RUN_MODE is fail-closed: it may use SAP GUI Scripting on exact mapped IDs,
' but it never falls back to captions, coordinates, SendKeys, or guessed IDs.
' Probe mode is read-only even when the caller is otherwise in RUN_MODE.

Public Function SAPRun_Mode() As String
    Dim modeText As String
    ' Fail closed: a missing or unknown mode must never become RUN.
    modeText = UCase$(Trim$(ModuleHelper.ToolConfig("SAP_MODE", "MAP")))
    Select Case modeText
        Case "MAP", "PROBE", "LIVE_PILOT", "RUN"
        Case Else: modeText = "INVALID"
    End Select
    SAPRun_Mode = modeText
End Function

Public Function SAPRun_IsMapMode() As Boolean
    SAPRun_IsMapMode = (SAPRun_Mode() = "MAP")
End Function

Public Function SAPRun_IsRunMode() As Boolean
    ' Exact-only safety mode used by shared SAP helpers.  PROBE, LIVE_PILOT,
    ' and INVALID are not RUN; they must be explicitly handled by their gate.
    SAPRun_IsRunMode = (SAPRun_Mode() = "RUN")
End Function

Public Function SAPRun_IsProbeMode() As Boolean
    SAPRun_IsProbeMode = (SAPRun_Mode() = "PROBE" Or _
                          ModuleHelper.ToolConfigBool("SAP_PROBE_MODE", False))
End Function

Public Function SAPRun_IsLivePilotMode() As Boolean
    SAPRun_IsLivePilotMode = (SAPRun_Mode() = "LIVE_PILOT")
End Function

Public Function SAPRun_ScreenControlAllowed() As Boolean
    If SAPRun_IsMapMode() Then
        SAPRun_ScreenControlAllowed = ModuleHelper.ToolConfigBool("MAP_ALLOW_SCREEN_CONTROL", True)
    Else
        ' Production default is hard-coded safe even if the CONFIG row is
        ' missing or malformed.
        SAPRun_ScreenControlAllowed = ModuleHelper.ToolConfigBool("RUN_ALLOW_SCREEN_CONTROL", False)
    End If
End Function

Public Function SAPRun_FallbackAllowed() As Boolean
    If SAPRun_IsMapMode() Then
        SAPRun_FallbackAllowed = ModuleHelper.ToolConfigBool("MAP_ALLOW_FALLBACK", False)
    Else
        SAPRun_FallbackAllowed = ModuleHelper.ToolConfigBool("RUN_ALLOW_FALLBACK", False)
    End If
End Function

Public Function SAPRun_RequireProbeRuntime(ByVal stageName As String, _
                                           ByVal requiredKeysCsv As String, _
                                           ByRef messageText As String, _
                                           Optional ByVal requireLockedInRun As Boolean = True) As Boolean
    Dim missing As String
    messageText = ""

    If SAPRun_Mode() <> "PROBE" And Not (SAPRun_Mode() = "RUN" And SAPRun_IsProbeMode()) Then
        messageText = "STOP_PROBE_MODE_REQUIRED: set SAP_MODE=PROBE (pre-lock) or SAP_MODE=RUN plus SAP_PROBE_MODE=TRUE (post-lock) before " & stageName & "."
        Exit Function
    End If

    ' MAP probe is used only to validate a candidate map before it is locked.
    ' RUN probe is the later proof that the locked map has no screen/fallback
    ' dependency. Neither branch creates a session or performs a write.
    If SAPRun_Mode() = "RUN" Then
        If Not SAPRun_RequireRunMode(stageName, messageText) Then Exit Function
        If requireLockedInRun Then
            If Not SAPRun_RequireLockedMap(stageName, requiredKeysCsv, messageText) Then Exit Function
        Else
            missing = SAPRun_MissingKeys(requiredKeysCsv)
            If missing <> "" Then
                messageText = "STOP_CONTROL_NOT_FOUND: RUN probe missing CONFIG technical IDs for " & _
                              stageName & ": " & missing
                Exit Function
            End If
        End If
    ElseIf SAPRun_Mode() = "PROBE" Then
        missing = SAPRun_MissingKeys(requiredKeysCsv)
        If missing <> "" Then
            messageText = "STOP_CONTROL_NOT_FOUND: PROBE missing CONFIG technical IDs for " & _
                          stageName & ": " & missing
            Exit Function
        End If
    Else
        messageText = "STOP_PROBE_MODE_REQUIRED: set SAP_MODE=PROBE for pre-lock read-only evidence, or SAP_MODE=RUN after LOCK."
        Exit Function
    End If
    SAPRun_RequireProbeRuntime = True
End Function

Public Function SAPRun_RequireMapMode(ByVal operationName As String, _
                                      ByRef messageText As String) As Boolean
    messageText = ""
    If Not SAPRun_IsMapMode() Then
        messageText = "STOP_MAP_MODE_REQUIRED: " & operationName & _
                      " is available only when CONFIG SAP_MODE=MAP."
        Exit Function
    End If
    If Not SAPRun_ScreenControlAllowed() Then
        messageText = "STOP_MAP_SCREEN_CONTROL_DISABLED: enable MAP_ALLOW_SCREEN_CONTROL only for mapping."
        Exit Function
    End If
    SAPRun_RequireMapMode = True
End Function

Public Function SAPRun_RequireRunMode(ByVal operationName As String, _
                                      ByRef messageText As String) As Boolean
    messageText = ""
    If SAPRun_Mode() <> "RUN" Then
        messageText = "STOP_RUN_MODE_REQUIRED: " & operationName & _
                      " is an operational entry point; set CONFIG SAP_MODE=RUN."
        Exit Function
    End If
    If SAPRun_ScreenControlAllowed() Then
        messageText = "STOP_RUN_SCREEN_CONTROL_ENABLED: RUN_ALLOW_SCREEN_CONTROL must remain FALSE."
        Exit Function
    End If
    If SAPRun_FallbackAllowed() Then
        messageText = "STOP_RUN_FALLBACK_ENABLED: RUN_ALLOW_FALLBACK must remain FALSE."
        Exit Function
    End If
    If ModuleHelper.ToolConfigBool("RUN_ALLOW_SENDKEYS", False) Then
        messageText = "STOP_RUN_SENDKEYS_ENABLED: RUN_ALLOW_SENDKEYS must remain FALSE."
        Exit Function
    End If
    If ModuleHelper.ToolConfigBool("RUN_ALLOW_COORDINATE", False) Then
        messageText = "STOP_RUN_COORDINATE_ENABLED: RUN_ALLOW_COORDINATE must remain FALSE."
        Exit Function
    End If
    If ModuleHelper.ToolConfigBool("RUN_ALLOW_FUZZY_FALLBACK", False) Then
        messageText = "STOP_RUN_FUZZY_FALLBACK_ENABLED: RUN_ALLOW_FUZZY_FALLBACK must remain FALSE."
        Exit Function
    End If
    SAPRun_RequireRunMode = True
End Function

Public Function SAPRun_RequireRuntime(ByVal stageName As String, _
                                      ByVal requiredKeysCsv As String, _
                                      ByRef messageText As String) As Boolean
    ModuleHelper.ResetToolConfigCache
    If Not SAPRun_RequireRunMode(stageName, messageText) Then Exit Function
    If Not SAPRun_RequireLockedMap(stageName, requiredKeysCsv, messageText) Then Exit Function
    SAPRun_RequireRuntime = True
End Function

Public Function SAPRun_RequireLockedMap(ByVal stageName As String, _
                                        ByVal requiredKeysCsv As String, _
                                        ByRef messageText As String) As Boolean
    Dim ws As Worksheet, mapState As String
    Dim keys As Variant, oneKey As Variant, keyText As String
    Dim missing As String

    messageText = ""
    If SAPRun_IsMapMode() Then
        SAPRun_RequireLockedMap = True
        Exit Function
    End If

    If Not ModuleHelper.ToolConfigBool("RUN_REQUIRE_LOCKED_MAP", True) Then
        SAPRun_RequireLockedMap = True
        Exit Function
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("SAP_MAP")
    On Error GoTo 0
    If ws Is Nothing Then
        messageText = "STOP_MAP_NOT_LOCKED: SAP_MAP sheet is missing for " & stageName & "."
        Exit Function
    End If

    mapState = UCase$(Trim$(CStr(ws.Range("D6").Value2 & "")))
    If mapState <> "LOCKED" Then
        messageText = "STOP_MAP_NOT_LOCKED: SAP_MAP state is '" & mapState & _
                      "'; lock the reviewed technical IDs before RUN stage " & stageName & "."
        Exit Function
    End If

    missing = SAPRun_MissingKeys(requiredKeysCsv)
    If missing <> "" Then
        messageText = "STOP_MAP_NOT_LOCKED: missing CONFIG technical IDs for " & _
                      stageName & ": " & missing
        Exit Function
    End If

    If Not SAPRun_VerifyLockedFingerprint(messageText) Then Exit Function

    SAPRun_RequireLockedMap = True
End Function

Public Function SAPRun_MissingKeys(ByVal requiredKeysCsv As String) As String
    Dim keys As Variant, oneKey As Variant, keyText As String
    Dim missing As String
    keys = Split(requiredKeysCsv, ",")
    For Each oneKey In keys
        keyText = UCase$(Trim$(CStr(oneKey & "")))
        If keyText <> "" Then
            If Trim$(ModuleHelper.ToolConfig(keyText, "")) = "" Then
                If missing <> "" Then missing = missing & ","
                missing = missing & keyText
            End If
        End If
    Next oneKey
    SAPRun_MissingKeys = missing
End Function

Public Function SAPRun_RequiredV1Keys() As String
    SAPRun_RequiredV1Keys = _
        "MON_TREE_ID,MON_OUTBOUND_NODE_KEY,MON_RESULT_GRID_ID,MON_WT_GRID_ID," & _
        "MON_OB_INPUT_ID,MON_EXECUTE_ID,MON_WT_ACTION_ID,MON_UPPER_OB_COL,MON_UPPER_DOC_COL," & _
        "MON_UPPER_WA_COL,MON_UPPER_PICKING_COL,MON_UPPER_GI_COL,MON_UPPER_PICKING_PLAN_COL," & _
        "MON_WT_DOC_COL,MON_WT_WHO_COL,MON_WT_COL,MON_WT_ITEM_COL,MON_WT_MATERIAL_COL," & _
        "MON_WT_BATCH_COL,MON_WT_BIN_COL,MON_WT_QTY_COL,MON_WT_STATUS_COL," & _
        "POST_GRID_ID,POST_PRODUCT_COL,POST_AVAILABLE_QTY_COL,POST_BATCH_COL,POST_SOURCE_BIN_COL," & _
        "POST_TST_COL,POST_TOTAL_STOCK_COL,POST_DEST_BIN_COL,POST_QTY_CHG_COL,POST_DEST_BATCH_COL," & _
        "POST_RESULT_COUNT_ID,POST_RESULT_COUNT_PATTERN,POST_ADVANCED_TOGGLE_ID,POST_ADVANCED_RESET_ID," & _
        "POST_ADVANCED_START_ID,POST_PRODUCT_MULTI_ID,POST_OWNER_ID,POST_PRODUCT_POPUP_CLOSE_ID," & _
        "POST_PRODUCT_POPUP_CLEAR_ID,POST_PRODUCT_PASTE_ID,POST_PRODUCT_POPUP_OK_ID,POST_SAVE_ID," & _
        "RELEASE_MON_MATERIAL_ID,RELEASE_MON_BIN_ID,RELEASE_MON_BATCH_ID,RELEASE_MON_EXECUTE_ID," & _
        "RELEASE_MON_RESULT_GRID_ID,RELEASE_MON_RESULT_COUNT_ID,RELEASE_MON_RESULT_COUNT_PATTERN," & _
        "RELEASE_MON_OB_COL,RELEASE_MON_WH_COL,RELEASE_MON_WT_COL,RELEASE_MON_MATERIAL_COL," & _
        "RELEASE_MON_BATCH_COL,RELEASE_MON_BIN_COL,RELEASE_MON_QTY_COL,RELEASE_MON_STATUS_COL"
End Function

Public Function SAPRun_VerifyLockedFingerprint(ByRef messageText As String) As Boolean
    Dim lockedFingerprint As String, actualFingerprint As String
    messageText = ""
    ModuleHelper.ResetToolConfigCache
    ' The CF pilot extends SAP_MAP with CF-only action keys. Keep the
    ' original locked V1 fingerprint for read-only WH/POST runtime checks so
    ' adding a reviewed CF extension cannot break TAI WH.
    lockedFingerprint = Trim$(ModuleHelper.ToolConfig("SAP_MAP_BASE_FINGERPRINT", ""))
    If lockedFingerprint = "" Then lockedFingerprint = Trim$(ModuleHelper.ToolConfig("SAP_MAP_FINGERPRINT", ""))
    If lockedFingerprint = "" Then
        messageText = "STOP_MAP_CHANGED_AFTER_LOCK: SAP_MAP_FINGERPRINT is blank."
        Exit Function
    End If
    actualFingerprint = SAPRun_CalculateFingerprint(SAPRun_RequiredV1Keys())
    If StrComp(lockedFingerprint, actualFingerprint, vbTextCompare) <> 0 Then
        messageText = "STOP_MAP_CHANGED_AFTER_LOCK: one or more locked technical IDs or SAP identity values changed."
        Exit Function
    End If
    SAPRun_VerifyLockedFingerprint = True
End Function

Public Function SAPRun_CalculateFingerprint(ByVal requiredKeysCsv As String) As String
    Dim oneKey As Variant, keyText As String, canonical As String
    Dim i As Long, h As Double, codePoint As Long
    h = 104729#
    canonical = "SYSTEM=" & UCase$(Trim$(ModuleHelper.ToolConfig("SAP_SYSTEM", ""))) & vbLf & _
                "CLIENT=" & UCase$(Trim$(ModuleHelper.ToolConfig("SAP_CLIENT", ""))) & vbLf & _
                "USER=" & UCase$(Trim$(ModuleHelper.ToolConfig("SAP_USER", ""))) & vbLf
    For Each oneKey In Split(requiredKeysCsv, ",")
        keyText = UCase$(Trim$(CStr(oneKey & "")))
        If keyText <> "" Then
            canonical = canonical & keyText & "=" & Trim$(ModuleHelper.ToolConfig(keyText, "")) & vbLf
        End If
    Next oneKey
    For i = 1 To Len(canonical)
        codePoint = AscW(Mid$(canonical, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536
        h = h * 65599# + CDbl(codePoint)
        h = h - Int(h / 2147483647#) * 2147483647#
    Next i
    SAPRun_CalculateFingerprint = "MAPFP1-" & Hex$(CLng(h)) & "-" & CStr(Len(canonical))
End Function

Public Sub CF_LockSAPMap()
    Dim requiredKeys As String
    Dim ws As Worksheet, missing As String, messageText As String
    Dim reviewed As Long, ambiguous As Long, mapVersion As String, fingerprint As String
    On Error GoTo Failed
    If Not SAPRun_RequireMapMode("CF_LockSAPMap", messageText) Then
        MsgBox messageText, vbExclamation, "SAP MAP"
        Exit Sub
    End If
    requiredKeys = SAPRun_RequiredV1Keys()
    missing = SAPRun_MissingKeys(requiredKeys)
    If missing <> "" Then
        messageText = "STOP_MAP_NOT_READY: cannot LOCK SAP_MAP; missing CONFIG keys: " & missing
        MsgBox messageText, vbExclamation, "SAP MAP"
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets("SAP_MAP")
    If Not SAPRun_ReviewMap(requiredKeys, reviewed, ambiguous, messageText) Then
        MsgBox messageText, vbExclamation, "SAP MAP"
        Exit Sub
    End If
    If Not SAPRun_RequiredProbePass(requiredKeys, messageText) Then
        MsgBox messageText, vbExclamation, "SAP MAP"
        Exit Sub
    End If
    mapVersion = Format$(Now, "yyyymmdd_hhnnss")
    fingerprint = SAPRun_CalculateFingerprint(requiredKeys)
    SAPRun_SetConfigValue "SAP_MAP_VERSION", mapVersion
    SAPRun_SetConfigValue "SAP_MAP_LOCKED_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    SAPRun_SetConfigValue "SAP_MAP_FINGERPRINT", fingerprint
    SAPRun_SetConfigValue "SAP_MAP_LOCKED", "TRUE"
    SAPRun_MarkMapLocked ws, requiredKeys
    ws.Range("D6").Value2 = "LOCKED"
    ws.Range("B6").Value2 = Now
    ws.Range("F7").Value2 = reviewed
    ws.Range("H7").Value2 = reviewed
    ws.Range("J7").Value2 = ambiguous
    ModuleLog.WriteLog "SAP_MAP", "", "", "LOCKED", "MapVersion=" & mapVersion & "; Fingerprint=" & fingerprint & "; reviewed exact IDs/probes passed; RUN exact-only contract enabled."
    MsgBox "SAP_MAP is LOCKED. RUN will use exact SAP GUI Scripting IDs only.", vbInformation, "SAP MAP"
    Exit Sub
Failed:
    MsgBox "CF_LockSAPMap failed: Err " & CStr(Err.Number) & ": " & Err.Description, vbExclamation, "SAP MAP"
End Sub

Public Sub CF_ReviewSAPMap()
    Dim reviewed As Long, ambiguous As Long, messageText As String
    If Not SAPRun_RequireMapMode("CF_ReviewSAPMap", messageText) Then
        MsgBox messageText, vbExclamation, "SAP MAP REVIEW"
        Exit Sub
    End If
    If Not SAPRun_ReviewMap(SAPRun_RequiredV1Keys(), reviewed, ambiguous, messageText) Then
        MsgBox messageText, vbExclamation, "SAP MAP REVIEW"
    Else
        MsgBox messageText, vbInformation, "SAP MAP REVIEW"
    End If
End Sub

Public Function SAPRun_ReviewMap(ByVal requiredKeysCsv As String, ByRef reviewedCount As Long, _
                                 ByRef ambiguousCount As Long, ByRef messageText As String) As Boolean
    Dim ws As Worksheet, oneKey As Variant, keyText As String
    Dim rowNo As Long, foundRows As Long, idText As String, typeText As String, stageText As String
    Dim needControl As Boolean, needColumn As Boolean, needPattern As Boolean, missing As String, ambiguous As String
    On Error GoTo Failed
    messageText = "": reviewedCount = 0: ambiguousCount = 0
    Set ws = ThisWorkbook.Worksheets("SAP_MAP")
    For Each oneKey In Split(requiredKeysCsv, ",")
        keyText = UCase$(Trim$(CStr(oneKey & "")))
        If keyText <> "" Then
            idText = Trim$(ModuleHelper.ToolConfig(keyText, ""))
            If idText = "" Then
                missing = SAPRun_AppendCsv(missing, keyText)
            Else
                foundRows = 0: rowNo = SAPRun_FindMapRow(ws, keyText, foundRows)
                If foundRows <> 1 Then
                    ambiguousCount = ambiguousCount + 1
                    ambiguous = SAPRun_AppendCsv(ambiguous, keyText)
                Else
                    typeText = UCase$(Trim$(CStr(ws.Cells(rowNo, 7).Value2 & "")))
                    stageText = UCase$(Trim$(CStr(ws.Cells(rowNo, 3).Value2 & "")))
                    needColumn = (InStr(keyText, "_COL") > 0)
                    needPattern = (InStr(keyText, "_PATTERN") > 0)
                    needControl = Not needColumn And Not needPattern
                    If needColumn And typeText <> "GRID_COLUMN" Then
                        ambiguousCount = ambiguousCount + 1
                        ambiguous = SAPRun_AppendCsv(ambiguous, keyText & "(TYPE)")
                    ElseIf needControl And InStr(typeText, "GRID") = 0 And InStr(typeText, "BUTTON") = 0 And _
                           InStr(typeText, "TEXTFIELD") = 0 And InStr(typeText, "CTEXTFIELD") = 0 And _
                           InStr(typeText, "COMBO") = 0 And InStr(typeText, "MENU") = 0 And _
                           InStr(typeText, "TREE") = 0 And InStr(typeText, "STATUS") = 0 Then
                        ambiguousCount = ambiguousCount + 1
                        ambiguous = SAPRun_AppendCsv(ambiguous, keyText & "(TYPE)")
                    ElseIf Trim$(CStr(ws.Cells(rowNo, 6).Value2 & "")) <> idText Then
                        ambiguousCount = ambiguousCount + 1
                        ambiguous = SAPRun_AppendCsv(ambiguous, keyText & "(ID)")
                    Else
                        ws.Cells(rowNo, 15).Value2 = Now
                        ws.Cells(rowNo, 16).Value2 = "REVIEWED " & stageText
                        reviewedCount = reviewedCount + 1
                    End If
                End If
            End If
        End If
    Next oneKey
    If missing <> "" Or ambiguous <> "" Then
        messageText = "STOP_MAP_REVIEW_FAILED: missing=" & missing & "; ambiguous/type mismatch=" & ambiguous
        Exit Function
    End If
    ThisWorkbook.Worksheets("SAP_MAP").Range("D6").Value2 = "REVIEWED"
    ThisWorkbook.Worksheets("SAP_MAP").Range("F7").Value2 = reviewedCount
    ThisWorkbook.Worksheets("SAP_MAP").Range("J7").Value2 = ambiguousCount
    messageText = "SAP_MAP_REVIEW_PASS: reviewed=" & CStr(reviewedCount) & "; ambiguous=0. Run live read-only probes before LOCK."
    SAPRun_ReviewMap = True
    Exit Function
Failed:
    messageText = "STOP_MAP_REVIEW_FAILED: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub SAPRun_MarkMapLocked(ByVal ws As Worksheet, ByVal requiredKeysCsv As String)
    Dim oneKey As Variant, rowNo As Long, foundRows As Long
    For Each oneKey In Split(requiredKeysCsv, ",")
        foundRows = 0
        rowNo = SAPRun_FindMapRow(ws, UCase$(Trim$(CStr(oneKey & ""))), foundRows)
        If rowNo > 0 And foundRows = 1 Then
            ws.Cells(rowNo, 13).Value2 = "YES"
            ws.Cells(rowNo, 15).Value2 = Now
            ws.Cells(rowNo, 16).Value2 = "LOCKED exact ID; fingerprint protected."
        End If
    Next oneKey
End Sub

Private Function SAPRun_FindMapRow(ByVal ws As Worksheet, ByVal configKey As String, ByRef matchCount As Long) As Long
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).row
    For r = 11 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 2).Value2 & "")), configKey, vbTextCompare) = 0 Then
            matchCount = matchCount + 1
            SAPRun_FindMapRow = r
        End If
    Next r
End Function

Private Function SAPRun_RequiredProbePass(ByVal requiredKeysCsv As String, ByRef messageText As String) As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long, keyText As String
    Dim probePass As Boolean, requiredKey As Variant
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("SAP_PROBE_LOG")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For Each requiredKey In Array("POST", "MON", "RELEASE")
        probePass = False
        For r = 2 To lastRow
            keyText = UCase$(Trim$(CStr(ws.Cells(r, 3).Value2 & "")))
            If keyText = CStr(requiredKey) And UCase$(Trim$(CStr(ws.Cells(r, 15).Value2 & ""))) = "PASS" Then
                probePass = True
                Exit For
            End If
        Next r
        If Not probePass Then
            messageText = "STOP_MAP_LIVE_PROBE_REQUIRED: no PASS evidence for " & CStr(requiredKey) & " read-only probe."
            Exit Function
        End If
    Next requiredKey
    SAPRun_RequiredProbePass = True
    Exit Function
Failed:
    messageText = "STOP_MAP_LIVE_PROBE_REQUIRED: SAP_PROBE_LOG is missing or unreadable."
End Function

Private Function SAPRun_AppendCsv(ByVal currentText As String, ByVal nextText As String) As String
    If Trim$(nextText) = "" Then
        SAPRun_AppendCsv = currentText
    ElseIf Trim$(currentText) = "" Then
        SAPRun_AppendCsv = nextText
    Else
        SAPRun_AppendCsv = currentText & "," & nextText
    End If
End Function

Private Sub SAPRun_SetConfigValue(ByVal keyName As String, ByVal valueText As String)
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("CONFIG")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value2 & "")), keyName, vbTextCompare) = 0 Then
            ws.Cells(r, 2).Value2 = valueText
            ModuleHelper.ResetToolConfigCache
            Exit Sub
        End If
    Next r
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).Value2 = keyName
    ws.Cells(lastRow, 2).Value2 = valueText
    ws.Cells(lastRow, 3).Value2 = "TEXT"
    ws.Cells(lastRow, 4).Value2 = "YES"
    ws.Cells(lastRow, 5).Value2 = "SYSTEM"
    ws.Cells(lastRow, 6).Value2 = "Generated only after reviewed read-only probes pass."
    ModuleHelper.ResetToolConfigCache
End Sub

Public Function SAPRun_FindExact(ByVal ses As Object, ByVal configKey As String, _
                                 ByRef messageText As String) As Object
    Dim idText As String
    messageText = ""
    If ses Is Nothing Then
        messageText = "STOP_SESSION_NOT_FOUND: exact lookup " & configKey & " has no SAP session."
        Exit Function
    End If
    idText = Trim$(ModuleHelper.ToolConfig(configKey, ""))
    If idText = "" Then
        messageText = "STOP_CONTROL_NOT_FOUND: CONFIG " & configKey & " is blank."
        Exit Function
    End If
    On Error Resume Next
    Set SAPRun_FindExact = ses.FindById(idText, False)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If SAPRun_FindExact Is Nothing Then
        messageText = "STOP_CONTROL_NOT_FOUND: " & configKey & "=" & idText
    End If
End Function

Public Function SAPRun_ExactID(ByVal configKey As String, _
                               ByRef messageText As String) As String
    SAPRun_ExactID = Trim$(ModuleHelper.ToolConfig(configKey, ""))
    messageText = ""
    If SAPRun_ExactID = "" Then
        messageText = "STOP_CONTROL_NOT_FOUND: CONFIG " & configKey & " is blank."
    End If
End Function

Public Function SAPRun_RequireReadOnly(ByVal actionName As String, _
                                       ByRef messageText As String) As Boolean
    Dim actionText As String, keyName As String
    messageText = ""
    actionText = UCase$(Replace(Trim$(actionName), " ", "_"))
    If Not SAPRun_IsProbeMode() Then
        SAPRun_RequireReadOnly = True
        Exit Function
    End If

    Select Case actionText
        Case "CANCEL": keyName = "PROBE_ALLOW_CANCEL"
        Case "SAVE": keyName = "PROBE_ALLOW_SAVE"
        Case "CREATE": keyName = "PROBE_ALLOW_CREATE"
        Case "CF", "CONFIRM": keyName = "PROBE_ALLOW_CF"
        Case "WRITE", "SAP_WRITE", "POST": keyName = "PROBE_ALLOW_WRITE"
        Case Else
            SAPRun_RequireReadOnly = True
            Exit Function
    End Select

    If ModuleHelper.ToolConfigBool(keyName, False) Then
        SAPRun_RequireReadOnly = True
    Else
        messageText = "STOP_PROBE_SIDE_EFFECT_BLOCKED: " & actionName & _
                      " is disabled while SAP_PROBE_MODE=TRUE."
    End If
End Function

Public Sub SAPRun_CompileProbe()
    Dim msg As String
    If SAPRun_Mode() <> "MAP" And SAPRun_Mode() <> "RUN" Then _
        Err.Raise vbObjectError + 9981, , "SAP_RUN_GUARD_MODE_PROBE_FAIL"
    If SAPRun_RequireReadOnly("READ", msg) = False Then _
        Err.Raise vbObjectError + 9982, , msg
End Sub
