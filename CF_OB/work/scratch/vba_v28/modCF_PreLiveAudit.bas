Attribute VB_Name = "modCF_PreLiveAudit"
Option Explicit

' Read-only readiness gate for the first real SAP Confirm (CF) pilot.
' This module never presses a SAP control. It validates local state, exact
' map/action contracts, CF_ONLY classification, whitelist and time ceilings.

Private Const AUDIT_READY As String = "LIVE_PILOT_READY"
Private Const AUDIT_NOT_READY As String = "NOT_READY"
Private Const EPS As Double = 0.000001
Private Const CF_MAP_KEYS As String = _
    "MON_TREE_ID,MON_OUTBOUND_NODE_KEY,MON_RESULT_GRID_ID,MON_WT_GRID_ID," & _
    "MON_OB_INPUT_ID,MON_EXECUTE_ID,MON_WT_ACTION_ID,MON_UPPER_OB_COL," & _
    "MON_UPPER_DOC_COL,MON_UPPER_WA_COL,MON_UPPER_PICKING_COL,MON_UPPER_GI_COL," & _
    "MON_UPPER_PICKING_PLAN_COL,MON_WT_DOC_COL,MON_WT_WHO_COL,MON_WT_ITEM_COL," & _
    "MON_WT_MATERIAL_COL,MON_WT_BATCH_COL,MON_WT_BIN_COL,MON_WT_QTY_COL," & _
    "MON_WT_STATUS_COL,CF_GRID_ID,CF_BUTTON_ID,CF_CONFIRM_ID"

Public Sub CF_PreLiveAudit()
    Dim reportText As String
    If CF_PreLiveAudit_Run(reportText) Then
        MsgBox reportText, vbInformation, "PRE-LIVE AUDIT"
    Else
        MsgBox reportText, vbExclamation, "PRE-LIVE AUDIT"
    End If
End Sub

' Non-UI wrapper used by automated regression checks; it still performs no SAP write.
Public Sub CF_PreLiveAudit_ExecuteNoUI()
    Dim reportText As String
    Call CF_PreLiveAudit_Run(reportText)
End Sub

Public Function CF_PreLiveAudit_Run(ByRef reportText As String) As Boolean
    Dim blockers As Collection, runID As String, mapFP As String, resultText As String
    Dim integrityOK As Boolean, compileOK As Boolean, snapshotOK As Boolean
    Dim plannerOK As Boolean, mapOK As Boolean, actionMapOK As Boolean
    Dim configOK As Boolean, guardsOK As Boolean, timeoutOK As Boolean, loggingOK As Boolean, executorOK As Boolean
    Dim msg As String, started As Double

    On Error GoTo Failed
    started = Timer
    Set blockers = New Collection
    runID = Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng((Timer - Int(Timer)) * 1000#), "000")

    configOK = AuditExecutionConfig(blockers)
    integrityOK = AuditWorkbookStructure(blockers)
    compileOK = AuditCompile(blockers)
    msg = ""
    snapshotOK = CF_CheckPlanner_ValidateSnapshot(msg)
    If Not snapshotOK Then AddBlocker blockers, msg
    plannerOK = AuditPlanner(blockers)
    mapOK = AuditSAPMap(blockers, mapFP)
    actionMapOK = AuditActionMap(blockers)
    guardsOK = AuditWriteGuards(blockers)
    executorOK = AuditExecutor(blockers)
    timeoutOK = AuditTimeouts(blockers)
    loggingOK = AuditLogging(blockers)

    CF_PreLiveEnsureActionLogSchema
    If blockers.count = 0 Then resultText = AUDIT_READY Else resultText = AUDIT_NOT_READY

    AuditSetConfig "PRELIVE_AUDIT_RUN_ID", runID
    AuditSetConfig "PRELIVE_AUDIT_RESULT", resultText
    AuditSetConfig "LIVE_READY", IIf(resultText = AUDIT_READY, "TRUE", "FALSE")
    AuditSetConfig "PRELIVE_AUDIT_MAP_FINGERPRINT", mapFP
    AuditSetConfig "PRELIVE_AUDIT_AT", Format$(Now, "yyyy-mm-dd hh:nn:ss")

    reportText = AuditSummary(configOK, integrityOK, compileOK, snapshotOK, plannerOK, mapOK, _
                               actionMapOK, guardsOK, timeoutOK, loggingOK, executorOK, _
                               resultText, runID, blockers, Timer - started)
    On Error Resume Next
    ModuleLog.WriteLog "PRELIVE_AUDIT", reportText, runID, resultText
    On Error GoTo 0
    CF_PreLiveAudit_Run = (resultText = AUDIT_READY)
    Exit Function

Failed:
    If blockers Is Nothing Then Set blockers = New Collection
    AddBlocker blockers, "STOP_PRELIVE_AUDIT_ERROR: Err " & CStr(Err.Number) & ": " & Err.Description
    AuditSetConfig "PRELIVE_AUDIT_RESULT", AUDIT_NOT_READY
    AuditSetConfig "PRELIVE_AUDIT_RUN_ID", runID
    reportText = AuditSummary(False, False, False, False, False, False, False, False, False, False, False, _
                              AUDIT_NOT_READY, runID, blockers, Timer - started)
    CF_PreLiveAudit_Run = False
End Function

' Every future CF executor must call this guard before any SAP side effect.
Public Function CF_PreLiveAudit_Require(ByRef messageText As String) As Boolean
    Dim savedResult As String, savedFP As String, currentFP As String
    On Error GoTo Failed
    messageText = ""
    ModuleHelper.ResetToolConfigCache
    savedResult = UCase$(Trim$(ModuleHelper.ToolConfig("PRELIVE_AUDIT_RESULT", "")))
    If savedResult <> AUDIT_READY Then
        messageText = "STOP_PRELIVE_AUDIT_REQUIRED: latest result is '" & savedResult & "'."
        Exit Function
    End If
    savedFP = Trim$(ModuleHelper.ToolConfig("PRELIVE_AUDIT_MAP_FINGERPRINT", ""))
    currentFP = AuditMapFingerprint()
    If savedFP = "" Or StrComp(savedFP, currentFP, vbTextCompare) <> 0 Then
        messageText = "STOP_MAP_CHANGED_AFTER_LOCK: PRELIVE audit map fingerprint differs from current map."
        Exit Function
    End If
    If Not WriteWhitelist("CF", messageText) Then Exit Function
    CF_PreLiveAudit_Require = True
    Exit Function
Failed:
    messageText = "STOP_PRELIVE_AUDIT_REQUIRED: guard error. Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_PreLiveActionAllowed(ByVal actionName As String, ByRef messageText As String) As Boolean
    Dim actionText As String
    actionText = UCase$(Replace(Trim$(actionName), " ", "_"))
    If actionText <> "CF" And actionText <> "CONFIRM" Then
        messageText = "STOP_ACTION_NOT_ENABLED: only CF is whitelisted in the first pilot."
        Exit Function
    End If
    CF_PreLiveActionAllowed = CF_PreLiveAudit_Require(messageText)
End Function

Private Function AuditExecutionConfig(ByVal blockers As Collection) As Boolean
    Dim initialCount As Long, levelText As String, modeText As String
    Dim liveEnabled As Boolean, keyText As Variant

    initialCount = blockers.count
    levelText = UCase$(Trim$(ModuleHelper.ToolConfig("CF_RUN_LEVEL", "")))
    modeText = SAPRun_Mode()
    liveEnabled = ModuleHelper.ToolConfigBool("CF_LIVE_ENABLED", False)

    Select Case levelText
        Case "MAP", "OFFLINE_AUDIT", "DRY_RUN", "PILOT_LIVE", "RUN_LIVE"
        Case Else: AddBlocker blockers, "STOP_CF_RUN_LEVEL_INVALID: CF_RUN_LEVEL must be MAP, OFFLINE_AUDIT, DRY_RUN, PILOT_LIVE or RUN_LIVE."
    End Select

    If modeText = "INVALID" Then AddBlocker blockers, "STOP_SAP_MODE_INVALID: SAP_MODE is missing or unknown; fail-closed mode is required."
    If levelText = "MAP" And modeText <> "MAP" Then AddBlocker blockers, "STOP_MAP_MODE_REQUIRED: CF_RUN_LEVEL=MAP requires SAP_MODE=MAP."
    If levelText = "OFFLINE_AUDIT" And modeText <> "MAP" And modeText <> "PROBE" Then _
        AddBlocker blockers, "STOP_OFFLINE_MODE_REQUIRED: OFFLINE_AUDIT permits only SAP_MODE=MAP or PROBE."
    If levelText = "DRY_RUN" And Not SAPRun_IsProbeMode() Then _
        AddBlocker blockers, "STOP_PROBE_MODE_REQUIRED: DRY_RUN requires explicit read-only probe mode."
    If (levelText = "PILOT_LIVE" Or levelText = "RUN_LIVE") And _
       (modeText <> "RUN" Or Not liveEnabled) Then _
        AddBlocker blockers, "STOP_LIVE_MODE_NOT_ARMED: live level requires SAP_MODE=RUN and CF_LIVE_ENABLED=TRUE."
    If (levelText <> "PILOT_LIVE" And levelText <> "RUN_LIVE") And liveEnabled Then _
        AddBlocker blockers, "STOP_CF_LIVE_MODE_MISMATCH: CF_LIVE_ENABLED must remain FALSE below PILOT_LIVE."

    For Each keyText In Array("CORPORATE_SAFE_MODE", "LOCAL_ONLY", "AUDIT_FAIL_CLOSED", _
                              "AUDIT_REQUIRE_EXECUTOR", "AUDIT_REQUIRE_CALL_PATH", _
                              "CF_REQUIRE_EXECUTOR", "CF_REQUIRE_PRE_VERIFY", _
                              "CF_REQUIRE_POST_VERIFY", "CF_REQUIRE_STATUS_C_PROOF", _
                              "CF_REQUIRE_SINGLE_EXECUTOR", "DO_NOT_KILL_USER_EXCEL", _
                              "HEARTBEAT_REQUIRED", "LOG_LOCAL_ONLY")
        AuditRequireBooleanValue CStr(keyText), "TRUE", blockers
    Next keyText

    For Each keyText In Array("ALLOW_WEB", "ALLOW_HTTP", "ALLOW_EXTERNAL_API", _
                              "ALLOW_FILE_UPLOAD", "ALLOW_LOG_UPLOAD", "ALLOW_SCREENSHOT_UPLOAD", _
                              "ALLOW_GITHUB_PUSH", "ALLOW_GIT_FETCH", "ALLOW_PACKAGE_DOWNLOAD", _
                              "ALLOW_INFINITE_LOOP", "ALLOW_UNBOUNDED_RETRY", _
                              "CF_ALLOW_UNPROVEN_SUCCESS", "CF_ALLOW_VERIFY_FALLBACK", _
                              "CF_ALLOW_WRITE_FALLBACK", "RUN_ALLOW_SCREEN_CONTROL", _
                              "RUN_ALLOW_FALLBACK", "RUN_ALLOW_SENDKEYS", "RUN_ALLOW_COORDINATE", _
                              "RUN_ALLOW_FUZZY_FALLBACK", "PROBE_ALLOW_CANCEL", "PROBE_ALLOW_SAVE", _
                              "PROBE_ALLOW_CREATE", "PROBE_ALLOW_CF", "PROBE_ALLOW_WRITE", _
                              "ALLOW_QTY_WRITE", "ALLOW_POST", "ALLOW_CANCEL", "ALLOW_CREATE", _
                              "ALLOW_DELETE", "ALLOW_MIGO")
        AuditRequireBooleanValue CStr(keyText), "FALSE", blockers
    Next keyText

    If UCase$(Trim$(ModuleHelper.ToolConfig("NETWORK_REQUIRED_POLICY", ""))) <> "STOP" Then _
        AddBlocker blockers, "STOP_NETWORK_POLICY: NETWORK_REQUIRED_POLICY must be STOP."
    If UCase$(Trim$(ModuleHelper.ToolConfig("LOG_UPLOAD", ""))) <> "FALSE" Then _
        AddBlocker blockers, "STOP_LOG_UPLOAD_ENABLED: local-only logging requires LOG_UPLOAD=FALSE."

    AuditRequireNumericRange "CODEX_COMMAND_TIMEOUT_SEC", 1#, 3600#, blockers
    AuditRequireNumericRange "CODEX_PROBE_TIMEOUT_SEC", 1#, 3600#, blockers
    AuditRequireNumericRange "EXCEL_OPEN_TIMEOUT_SEC", 1#, 3600#, blockers
    AuditRequireNumericRange "EXCEL_COM_TIMEOUT_SEC", 1#, 3600#, blockers
    AuditRequireNumericRange "SAP_CONTROL_TIMEOUT_SEC", 1#, 3600#, blockers
    AuditRequireNumericRange "SAP_MAP_TIMEOUT_SEC", 1#, 3600#, blockers
    AuditRequireNumericRange "SAP_TRANSACTION_TIMEOUT_SEC", 1#, 3600#, blockers
    AuditRequireNumericRange "HEARTBEAT_TIMEOUT_SEC", 1#, 3600#, blockers
    AuditRequireNumericRange "CF_MAX_WRITE_ATTEMPTS", 1#, 1#, blockers
    AuditRequireNumericRange "CF_MAX_CONFIRM_ATTEMPTS", 1#, 1#, blockers

    If blockers.count = initialCount Then AuditExecutionConfig = True
End Function

Private Function AuditRequireBooleanValue(ByVal keyName As String, ByVal expectedValue As String, _
                                          ByVal blockers As Collection) As Boolean
    Dim raw As String, expectedText As String
    raw = UCase$(Trim$(ModuleHelper.ToolConfig(keyName, "")))
    expectedText = UCase$(Trim$(expectedValue))
    If raw <> "TRUE" And raw <> "FALSE" Then
        AddBlocker blockers, "STOP_CONFIG_BOOLEAN: " & keyName & " must be explicitly TRUE/FALSE."
        Exit Function
    End If
    If raw <> expectedText Then
        AddBlocker blockers, "STOP_CONFIG_SAFETY: " & keyName & " must be " & expectedText & "."
        Exit Function
    End If
    AuditRequireBooleanValue = True
End Function

Private Function AuditRequireNumericRange(ByVal keyName As String, ByVal minValue As Double, _
                                          ByVal maxValue As Double, ByVal blockers As Collection) As Boolean
    Dim raw As String, valueNumber As Double
    raw = Trim$(ModuleHelper.ToolConfig(keyName, ""))
    If raw = "" Or Not IsNumeric(raw) Then
        AddBlocker blockers, "STOP_CONFIG_NUMBER: " & keyName & " must be numeric."
        Exit Function
    End If
    valueNumber = CDbl(raw)
    If valueNumber < minValue Or valueNumber > maxValue Then
        AddBlocker blockers, "STOP_CONFIG_RANGE: " & keyName & " must be " & CStr(minValue) & ".." & CStr(maxValue) & "."
        Exit Function
    End If
    AuditRequireNumericRange = True
End Function

Private Function AuditWorkbookStructure(ByVal blockers As Collection) As Boolean
    Dim names As Variant, oneName As Variant, ws As Worksheet
    Dim headers As Variant, i As Long, missing As String
    On Error GoTo Failed
    names = Array("INPUT", "DATA_CORE", "OB_QUEUE", "OB_WH_CACHE", "PLAN", _
                  "PLAN_SOURCE_LEGS", "PLAN_RELEASE", "POSTING_CACHE", "POSTING_SUMMARY", _
                  "RUN_LOG", "SYSTEM", "CONFIG", "SAP_ACTION_MAP")
    For Each oneName In names
        Set ws = GetSheet(CStr(oneName))
        If ws Is Nothing Then
            If missing <> "" Then missing = missing & ","
            missing = missing & CStr(oneName)
        End If
        Set ws = Nothing
    Next oneName
    If missing <> "" Then
        AddBlocker blockers, "STOP_WORKBOOK_STRUCTURE: missing sheet(s)=" & missing
        Exit Function
    End If
    If UCase$(Trim$(CStr(ThisWorkbook.Worksheets("SAP_ACTION_MAP").Cells(1, 1).Value2 & ""))) <> "ACTIONKEY" Then
        AddBlocker blockers, "STOP_ACTION_MAP_CONTRACT: SAP_ACTION_MAP!A1 is not ActionKey."
        Exit Function
    End If
    headers = Array("TIME", "RUN ID", "STAGE", "OB", "DOCUMENT", "ITEM", "WH", "WT", _
                    "MATERIAL", "ACTION", "STATUS", "MESSAGE", "SAP SEC", "EXCEL SEC", _
                    "SESSION", "TCODE", "TITLE")
    For i = 0 To UBound(headers)
        If UCase$(Trim$(CStr(ThisWorkbook.Worksheets("RUN_LOG").Cells(8, i + 1).Value2 & ""))) <> headers(i) Then
            AddBlocker blockers, "STOP_LOG_SCHEMA: RUN_LOG row 8 column " & CStr(i + 1) & " is not " & headers(i) & "."
            Exit Function
        End If
    Next i
    AuditWorkbookStructure = True
    Exit Function
Failed:
    AddBlocker blockers, "STOP_WORKBOOK_STRUCTURE: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function AuditCompile(ByVal blockers As Collection) As Boolean
    Dim probes As Variant, p As Variant, errText As String
    probes = Array("SAPActionMap_CompileProbe", "SAPRun_CompileProbe", "SAPProbe_CompileProbe", _
                   "WT_StateEngine_CompileProbe", "AP_CompileProbe", "CF_InputLayout_CompileProbe", _
                   "CF_CheckPlanner_CompileProbe", "CF_ExecutePlanToCFHandoff_StaticCompileProbe", _
                   "CF_ReleaseHolderReader_CompileProbe", "MIGO311_CompileProbe")
    For Each p In probes
        errText = ""
        If Not RunProbe(CStr(p), errText) Then
            AddBlocker blockers, "STOP_VBA_COMPILE: " & CStr(p) & ": " & errText
            Exit Function
        End If
    Next p
    AuditCompile = True
End Function

Private Function RunProbe(ByVal probeName As String, ByRef errorText As String) As Boolean
    On Error GoTo Failed
    Select Case probeName
        Case "SAPActionMap_CompileProbe": SAPActionMap_CompileProbe
        Case "SAPRun_CompileProbe": SAPRun_CompileProbe
        Case "SAPProbe_CompileProbe": SAPProbe_CompileProbe
        Case "WT_StateEngine_CompileProbe": WT_StateEngine_CompileProbe
        Case "AP_CompileProbe": AP_CompileProbe
        Case "CF_InputLayout_CompileProbe": CF_InputLayout_CompileProbe
        Case "CF_CheckPlanner_CompileProbe": CF_CheckPlanner_CompileProbe
        Case "CF_ExecutePlanToCFHandoff_StaticCompileProbe": CF_ExecutePlanToCFHandoff_StaticCompileProbe
        Case "CF_ReleaseHolderReader_CompileProbe": CF_ReleaseHolderReader_CompileProbe
        Case "MIGO311_CompileProbe": MIGO311_CompileProbe
        Case Else: errorText = "unknown probe": Exit Function
    End Select
    RunProbe = True
    Exit Function
Failed:
    errorText = "Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function AuditPlanner(ByVal blockers As Collection) As Boolean
    Dim wsPlan As Worksheet, wsWH As Worksheet, runID As String, planRunID As String
    Dim lastRow As Long, r As Long, whRow As Long, decision As String, reasonText As String, cfRows As Long
    On Error GoTo Failed
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
    runID = Trim$(CStr(wsWH.Range("W2").Value2 & ""))
    planRunID = Trim$(CStr(wsPlan.Range("AB2").Value2 & ""))
    If runID = "" Then AddBlocker blockers, "STOP_PLAN_RUNID: WH_ACTUAL!W2 is blank."
    If planRunID = "" Then
        AddBlocker blockers, "STOP_PLAN_NOT_READY: PLAN has no PLAN Run ID; run CHECK after the latest WH snapshot."
    ElseIf StrComp(planRunID, runID, vbTextCompare) <> 0 Then
        AddBlocker blockers, "STOP_PLAN_STALE: PLAN Run ID differs from WH_ACTUAL Run ID."
    End If
    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    If lastRow < 2 Then
        AddBlocker blockers, "STOP_PLAN_NOT_READY: PLAN has no data rows."
        Exit Function
    End If
    For r = 2 To lastRow
        decision = UCase$(Trim$(CStr(wsPlan.Cells(r, 20).Value2 & "")))
        If decision = "CF_ONLY" Then
            whRow = CLng(val(wsPlan.Cells(r, 27).Value2))
            reasonText = ""
            If AuditCFOnlyRow(wsPlan, wsWH, r, whRow, reasonText) Then
                cfRows = cfRows + 1
            Else
                AddBlocker blockers, "STOP_CF_ONLY_NOT_PROVEN: PLAN row " & CStr(r) & ": " & reasonText
            End If
        End If
    Next r
    If cfRows = 0 Then AddBlocker blockers, "STOP_CF_ONLY_NOT_AVAILABLE: no deterministic CF_ONLY row is ready for pilot."
    AuditPlanner = (cfRows > 0)
    Exit Function
Failed:
    AddBlocker blockers, "STOP_PLAN_AUDIT: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function AuditCFOnlyRow(ByVal wsPlan As Worksheet, ByVal wsWH As Worksheet, _
                                ByVal planRow As Long, ByVal whRow As Long, ByRef reasonText As String) As Boolean
    Dim obState As String, wtState As String, rawStatus As String
    Dim sysQty As Double, targetQty As Double, actualQty As Double
    Dim sysBin As String, targetBin As String, actualBin As String
    Dim sysBatch As String, targetBatch As String, actualBatch As String
    If whRow < 5 Then reasonText = "WH Actual Row is missing.": Exit Function
    obState = UCase$(Trim$(CStr(wsWH.Cells(whRow, 25).Value2 & "")))
    wtState = UCase$(Trim$(CStr(wsWH.Cells(whRow, 26).Value2 & "")))
    rawStatus = UCase$(Trim$(CStr(wsWH.Cells(whRow, 10).Value2 & "")))
    If obState <> "OB_OPEN_OR_PARTIAL" Then reasonText = "OB state is not OB_OPEN_OR_PARTIAL.": Exit Function
    If wtState <> "WT_ACTIVE" Then reasonText = "WT state is not WT_ACTIVE.": Exit Function
    If rawStatus <> "" Then reasonText = "raw WT status is not blank.": Exit Function
    If Not IsNumeric(wsWH.Cells(whRow, 12).Value2) Then reasonText = "Actual Qty is not numeric.": Exit Function
    actualQty = CDbl(wsWH.Cells(whRow, 12).Value2)
    If actualQty <= 0# Then reasonText = "Actual Qty is not > 0.": Exit Function
    If Not IsNumeric(wsPlan.Cells(planRow, 7).Value2) Or Not IsNumeric(wsPlan.Cells(planRow, 8).Value2) Then reasonText = "Sys/Target Qty is not numeric.": Exit Function
    sysQty = CDbl(wsPlan.Cells(planRow, 7).Value2): targetQty = CDbl(wsPlan.Cells(planRow, 8).Value2)
    sysBin = Trim$(CStr(wsPlan.Cells(planRow, 9).Value2 & "")): targetBin = Trim$(CStr(wsPlan.Cells(planRow, 12).Value2 & ""))
    sysBatch = Trim$(CStr(wsPlan.Cells(planRow, 11).Value2 & "")): targetBatch = Trim$(CStr(wsPlan.Cells(planRow, 14).Value2 & ""))
    actualBin = Trim$(CStr(wsWH.Cells(whRow, 14).Value2 & "")): actualBatch = Trim$(CStr(wsWH.Cells(whRow, 22).Value2 & ""))
    ' Actual Batch is optional user input.  When the proven actual bin is the
    ' system/target bin, a blank batch means no batch change; the planner has
    ' already resolved the target batch to the proven system batch.
    If actualBatch = "" And StrComp(actualBin, targetBin, vbTextCompare) = 0 And _
       StrComp(sysBin, targetBin, vbTextCompare) = 0 Then actualBatch = sysBatch
    If Abs(sysQty - targetQty) > EPS Or Abs(actualQty - targetQty) > EPS Then reasonText = "Qty does not match target.": Exit Function
    If StrComp(sysBin, targetBin, vbTextCompare) <> 0 Or StrComp(actualBin, targetBin, vbTextCompare) <> 0 Then reasonText = "Bin does not match target.": Exit Function
    If StrComp(sysBatch, targetBatch, vbTextCompare) <> 0 Or StrComp(actualBatch, targetBatch, vbTextCompare) <> 0 Then reasonText = "Batch does not match target.": Exit Function
    If UCase$(Trim$(CStr(wsPlan.Cells(planRow, 21).Value2 & ""))) = "TRUE" Or _
       UCase$(Trim$(CStr(wsPlan.Cells(planRow, 22).Value2 & ""))) = "TRUE" Or _
       UCase$(Trim$(CStr(wsPlan.Cells(planRow, 23).Value2 & ""))) = "TRUE" Then reasonText = "planner change flag is TRUE.": Exit Function
    If CDbl(val(wsPlan.Cells(planRow, 17).Value2)) > EPS Then reasonText = "Need Release is nonzero.": Exit Function
    AuditCFOnlyRow = True
End Function

Private Function AuditSAPMap(ByVal blockers As Collection, ByRef fingerprintText As String) As Boolean
    Dim ws As Worksheet, mapState As String, missing As String, lockedFP As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("SYSTEM")
    mapState = UCase$(Trim$(CStr(ws.Range("D6").Value2 & "")))
    If mapState <> "LOCKED" Or Not ModuleHelper.ToolConfigBool("SAP_MAP_LOCKED", False) Then AddBlocker blockers, "STOP_MAP_NOT_LOCKED: SAP_MAP must be LOCKED before CF pilot."
    missing = SAPRun_MissingKeys(CF_MAP_KEYS)
    If missing <> "" Then AddBlocker blockers, "STOP_CONTROL_NOT_FOUND: CF required CONFIG IDs are missing=" & missing
    lockedFP = Trim$(ModuleHelper.ToolConfig("SAP_MAP_FINGERPRINT", ""))
    fingerprintText = AuditMapFingerprint()
    If lockedFP = "" Then
        AddBlocker blockers, "STOP_MAP_CHANGED_AFTER_LOCK: SAP_MAP_FINGERPRINT is blank."
    ElseIf StrComp(lockedFP, fingerprintText, vbTextCompare) <> 0 Then
        AddBlocker blockers, "STOP_MAP_CHANGED_AFTER_LOCK: current CF map fingerprint differs from locked value."
    End If
    AuditSAPMap = (mapState = "LOCKED" And missing = "" And lockedFP <> "" And StrComp(lockedFP, fingerprintText, vbTextCompare) = 0)
    Exit Function
Failed:
    AddBlocker blockers, "STOP_MAP_AUDIT: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function AuditActionMap(ByVal blockers As Collection) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long, keyText As String, idText As String, statusText As String
    Dim wtOK As Boolean, cfOK As Boolean
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("SAP_ACTION_MAP")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        keyText = UCase$(Trim$(CStr(ws.Cells(r, 1).Value2 & "")))
        idText = Trim$(CStr(ws.Cells(r, 5).Value2 & ""))
        statusText = UCase$(Trim$(CStr(ws.Cells(r, 6).Value2 & "")))
        If (keyText = "MON_WT_GRID_ID" Or keyText = "MON_WT_ACTION_ID") And statusText = "VERIFIED" And idText <> "" Then wtOK = True
        If (keyText = "MON_CONFIRM_FG_COMMAND_ID" Or keyText = "MON_CONFIRM_BG_COMMAND_ID") And statusText = "VERIFIED" And idText <> "" Then cfOK = True
    Next r
    If Not wtOK Then AddBlocker blockers, "CF_WT_NAV_NOT_VERIFIED: exact WT grid/action is not VERIFIED in SAP_ACTION_MAP."
    If Not cfOK Then AddBlocker blockers, "CF_CONFIRM_ID_NOT_VERIFIED: no VERIFIED Confirm command/control exists in SAP_ACTION_MAP."
    AuditActionMap = (wtOK And cfOK)
    Exit Function
Failed:
    AddBlocker blockers, "STOP_ACTION_MAP_AUDIT: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function AuditWriteGuards(ByVal blockers As Collection) As Boolean
    Dim msg As String, levelText As String, initialCount As Long
    Dim writeEnabled As Boolean

    initialCount = blockers.count
    levelText = UCase$(Trim$(ModuleHelper.ToolConfig("CF_RUN_LEVEL", "")))

    If levelText = "PILOT_LIVE" Or levelText = "RUN_LIVE" Then
        If Not SAPRun_RequireRunMode("CF pilot", msg) Then AddBlocker blockers, msg
        If Not WriteWhitelist("CF", msg) Then AddBlocker blockers, msg
        If ModuleHelper.ToolConfigBool("SAP_PROBE_MODE", False) Then _
            AddBlocker blockers, "STOP_PROBE_MODE_REQUIRED: SAP_PROBE_MODE must be FALSE for live pilot."
    Else
        writeEnabled = ModuleHelper.ToolConfigBool("ALLOW_CF", False) Or _
                       ModuleHelper.ToolConfigBool("ALLOW_QTY_WRITE", False) Or _
                       ModuleHelper.ToolConfigBool("ALLOW_POST", False) Or _
                       ModuleHelper.ToolConfigBool("ALLOW_CANCEL", False) Or _
                       ModuleHelper.ToolConfigBool("ALLOW_CREATE", False) Or _
                       ModuleHelper.ToolConfigBool("ALLOW_DELETE", False) Or _
                       ModuleHelper.ToolConfigBool("ALLOW_MIGO", False)
        If writeEnabled Then AddBlocker blockers, "STOP_WRITE_DISABLED: no SAP write whitelist may be enabled below PILOT_LIVE."
        If SAPRun_IsRunMode() Then AddBlocker blockers, "STOP_OFFLINE_MODE_REQUIRED: non-live audit cannot run with SAP_MODE=RUN."
    End If

    AuditWriteGuards = (blockers.count = initialCount)
End Function

Private Function WriteWhitelist(ByVal actionName As String, ByRef messageText As String) As Boolean
    Dim actionText As String
    actionText = UCase$(Replace(Trim$(actionName), " ", "_"))
    If actionText <> "CF" And actionText <> "CONFIRM" Then
        messageText = "STOP_ACTION_NOT_ENABLED: action=" & actionName & " is outside the CF-only whitelist."
        Exit Function
    End If
    If Not ModuleHelper.ToolConfigBool("ALLOW_CF", False) Then
        messageText = "STOP_ACTION_NOT_ENABLED: ALLOW_CF must be TRUE only for the CF pilot."
        Exit Function
    End If
    If ModuleHelper.ToolConfigBool("ALLOW_QTY_WRITE", False) Or ModuleHelper.ToolConfigBool("ALLOW_POST", False) Or _
       ModuleHelper.ToolConfigBool("ALLOW_CANCEL", False) Or ModuleHelper.ToolConfigBool("ALLOW_CREATE", False) Or _
       ModuleHelper.ToolConfigBool("ALLOW_DELETE", False) Or ModuleHelper.ToolConfigBool("ALLOW_MIGO", False) Then
        messageText = "STOP_WRITE_WHITELIST: CF is not the only enabled write action."
        Exit Function
    End If
    WriteWhitelist = True
End Function

Private Function AuditExecutor(ByVal blockers As Collection) As Boolean
    Dim bodyText As String, executorBody As String, reworkBody As String, ok As Boolean
    ok = True

    If Not FindPublicProcedureBody("CF_RunSinglePilot", executorBody) Then
        AddBlocker blockers, "CF_EXECUTOR_NOT_IMPLEMENTED: public CF_RunSinglePilot procedure is missing."
        ok = False
    End If
    If Not FindPublicProcedureBody("CF_PreVerifyTarget", bodyText) Then
        AddBlocker blockers, "CF_PREVERIFY_NOT_IMPLEMENTED: public CF_PreVerifyTarget procedure is missing."
        ok = False
    End If
    If Not FindPublicProcedureBody("CF_ExecuteConfirm", bodyText) Then
        AddBlocker blockers, "CF_CONFIRM_NOT_IMPLEMENTED: public CF_ExecuteConfirm procedure is missing."
        ok = False
    End If
    If Not FindPublicProcedureBody("CF_PostVerifyTarget", bodyText) Then
        AddBlocker blockers, "CF_POSTVERIFY_NOT_IMPLEMENTED: public CF_PostVerifyTarget procedure is missing."
        ok = False
    End If
    If Not FindPublicProcedureBody("CF_VerifyStatusC", bodyText) Then
        AddBlocker blockers, "CF_STATE_C_PROOF_NOT_IMPLEMENTED: public CF_VerifyStatusC procedure is missing."
        ok = False
    ElseIf InStr(1, UCase$(bodyText), """C""", vbTextCompare) = 0 Then
        AddBlocker blockers, "CF_STATE_C_PROOF_NOT_IMPLEMENTED: CF_VerifyStatusC has no explicit status C comparison."
        ok = False
    End If

    If ok Then
        If Not ProcedureCalls(executorBody, "CF_PreVerifyTarget") Then AddBlocker blockers, "CF_CALL_PATH_MISSING: CF_RunSinglePilot does not call CF_PreVerifyTarget.": ok = False
        If Not ProcedureCalls(executorBody, "CF_ExecuteConfirm") Then AddBlocker blockers, "CF_CALL_PATH_MISSING: CF_RunSinglePilot does not call CF_ExecuteConfirm.": ok = False
        If Not ProcedureCalls(executorBody, "CF_PostVerifyTarget") Then AddBlocker blockers, "CF_CALL_PATH_MISSING: CF_RunSinglePilot does not call CF_PostVerifyTarget.": ok = False
        If Not ProcedureCalls(executorBody, "CF_VerifyStatusC") Then AddBlocker blockers, "CF_CALL_PATH_MISSING: CF_RunSinglePilot does not call CF_VerifyStatusC.": ok = False
    End If

    If Not FindPublicProcedureBody("CWR_ExecutePlanToCFHandoffCore", reworkBody) Then
        AddBlocker blockers, "REWORK_EXECUTOR_NOT_IMPLEMENTED: CWR_ExecutePlanToCFHandoffCore is missing."
        ok = False
    Else
        If Not ProcedureCalls(reworkBody, "CWR_LoadExecutionContext") Then AddBlocker blockers, "REWORK_CALL_PATH_MISSING: executor does not load exact PLAN context.": ok = False
        If Not ProcedureCalls(reworkBody, "CWR_AssertCFHandoff") Then AddBlocker blockers, "REWORK_CF_HANDOFF_LOCK_MISSING": ok = False
        If Not ProcedureCalls(reworkBody, "CWR_AssertReworkGate") Then AddBlocker blockers, "REWORK_WRITE_GATE_MISSING": ok = False
        If Not ProcedureCalls(reworkBody, "Codex_CancelOldWTAndRecord") Then AddBlocker blockers, "REWORK_CANCEL_STAGE_MISSING": ok = False
        If Not ProcedureCalls(reworkBody, "Codex_CreateOutboundWTFromExactCandidate") Then AddBlocker blockers, "REWORK_CREATE_STAGE_MISSING": ok = False
        If ProcedureCalls(reworkBody, "CF_ExecuteConfirm") Or ProcedureCalls(reworkBody, "CF_RunSinglePilot") Then
            AddBlocker blockers, "REWORK_CF_VIOLATION: rework executor must stop before CF."
            ok = False
        End If
    End If
    If Not CodeHasToken("REWORK_ACTION_WHITELIST") Or Not CodeHasToken("REWORK_NEVER_CF") Then
        AddBlocker blockers, "REWORK_GUARD_CONFIG_MISSING"
        ok = False
    End If
    AuditExecutor = ok
End Function

Private Function FindPublicProcedureBody(ByVal procedureName As String, ByRef bodyText As String) As Boolean
    Dim component As Object, codeModule As Object
    Dim lineNo As Long, endLine As Long, lineText As String, normalized As String
    bodyText = ""
    On Error GoTo ExitHere
    For Each component In ThisWorkbook.VBProject.VBComponents
        If component.Name <> "modCF_PreLiveAudit" Then
            Set codeModule = component.codeModule
            For lineNo = 1 To codeModule.CountOfLines
                lineText = codeModule.lines(lineNo, 1)
                If IsProcedureDeclaration(lineText, procedureName) Then
                    endLine = lineNo
                    Do While endLine < codeModule.CountOfLines
                        endLine = endLine + 1
                        normalized = UCase$(Trim$(StripVBAComment(codeModule.lines(endLine, 1))))
                        If normalized = "END SUB" Or normalized = "END FUNCTION" Or _
                           normalized = "END PROPERTY" Or Left$(normalized, 10) = "END PROPERTY" Then Exit Do
                    Loop
                    bodyText = codeModule.lines(lineNo, endLine - lineNo + 1)
                    FindPublicProcedureBody = True
                    Exit Function
                End If
            Next lineNo
        End If
    Next component
ExitHere:
End Function

Private Function IsProcedureDeclaration(ByVal lineText As String, ByVal procedureName As String) As Boolean
    Dim normalized As String, prefix As Variant, nameText As String, cutAt As Long
    normalized = UCase$(Trim$(StripVBAComment(lineText)))
    If normalized = "" Then Exit Function
    For Each prefix In Array("PUBLIC SUB ", "PRIVATE SUB ", "FRIEND SUB ", "SUB ", _
                             "PUBLIC FUNCTION ", "PRIVATE FUNCTION ", "FRIEND FUNCTION ", "FUNCTION ")
        If Left$(normalized, Len(CStr(prefix))) = CStr(prefix) Then
            nameText = Trim$(Mid$(normalized, Len(CStr(prefix)) + 1))
            cutAt = InStr(1, nameText, "(", vbBinaryCompare)
            If cutAt > 0 Then nameText = Left$(nameText, cutAt - 1)
            cutAt = InStr(1, nameText, " ", vbBinaryCompare)
            If cutAt > 0 Then nameText = Left$(nameText, cutAt - 1)
            IsProcedureDeclaration = (StrComp(nameText, UCase$(procedureName), vbTextCompare) = 0)
            Exit Function
        End If
    Next prefix
End Function

Private Function ProcedureCalls(ByVal bodyText As String, ByVal procedureName As String) As Boolean
    Dim lines As Variant, oneLine As Variant, normalized As String, nameText As String
    nameText = UCase$(procedureName)
    lines = Split(bodyText, vbCrLf)
    For Each oneLine In lines
        normalized = UCase$(Trim$(StripVBAComment(CStr(oneLine))))
        If InStr(1, normalized, "CALL " & nameText, vbTextCompare) > 0 Or _
           InStr(1, normalized, nameText & "(", vbTextCompare) > 0 Then
            ProcedureCalls = True
            Exit Function
        End If
    Next oneLine
End Function

Private Function StripVBAComment(ByVal lineText As String) As String
    Dim commentAt As Long
    commentAt = InStr(1, lineText, "'", vbBinaryCompare)
    If commentAt > 0 Then
        StripVBAComment = Left$(lineText, commentAt - 1)
    Else
        StripVBAComment = lineText
    End If
End Function

Private Function AuditTimeouts(ByVal blockers As Collection) As Boolean
    Dim keys As Variant, oneKey As Variant, raw As String, missing As String
    keys = Array("CF_NAV_TIMEOUT_SEC", "CF_ACTION_TIMEOUT_SEC", "CF_POSTVERIFY_TIMEOUT_SEC", _
                 "MAP_SCAN_TIMEOUT_SEC", "POST_CAPTURE_TIMEOUT_SEC", "RELEASE_QUERY_TIMEOUT_SEC")
    For Each oneKey In keys
        raw = Trim$(ModuleHelper.ToolConfig(CStr(oneKey), ""))
        If raw = "" Or Not IsNumeric(raw) Or CDbl(raw) <= 0# Then
            If missing <> "" Then missing = missing & ","
            missing = missing & CStr(oneKey)
        End If
    Next oneKey
    If missing <> "" Then AddBlocker blockers, "STOP_TIMEOUT_SETTING: missing/non-positive timeout setting(s)=" & missing
    If Not CodeHasToken("CF_NAV_TIMEOUT_SEC") Or Not CodeHasToken("CF_ACTION_TIMEOUT_SEC") Or Not CodeHasToken("CF_POSTVERIFY_TIMEOUT_SEC") Then AddBlocker blockers, "STOP_CF_TIMEOUT_CODE: CF action path has no bounded timeout constants."
    AuditTimeouts = (missing = "") And CodeHasToken("CF_NAV_TIMEOUT_SEC") And CodeHasToken("CF_ACTION_TIMEOUT_SEC") And CodeHasToken("CF_POSTVERIFY_TIMEOUT_SEC")
End Function

Private Function AuditLogging(ByVal blockers As Collection) As Boolean
    Dim stages As Variant, oneStage As Variant, missing As String
    stages = Array("ACTION_START", "PRE_VERIFY", "ACTION_SENT", "POST_VERIFY", "DONE")
    For Each oneStage In stages
        If Not CodeHasToken(CStr(oneStage)) Then
            If missing <> "" Then missing = missing & ","
            missing = missing & CStr(oneStage)
        End If
    Next oneStage
    If missing <> "" Then AddBlocker blockers, "STOP_LOGGING_CONTRACT: missing CF action stages=" & missing
    AuditLogging = (missing = "")
End Function

Private Function CodeHasToken(ByVal tokenText As String) As Boolean
    Dim c As Object, cm As Object, codeText As String
    On Error GoTo ExitHere
    For Each c In ThisWorkbook.VBProject.VBComponents
        If c.Name <> "modCF_PreLiveAudit" Then
            Set cm = c.codeModule
            If cm.CountOfLines > 0 Then
                codeText = cm.lines(1, cm.CountOfLines)
                If InStr(1, codeText, tokenText, vbTextCompare) > 0 Then CodeHasToken = True: Exit Function
            End If
        End If
    Next c
ExitHere:
End Function

Private Function AuditMapFingerprint() As String
    Dim canonical As String, oneKey As Variant, keyText As String
    Dim h As Double, i As Long, codePoint As Long
    On Error GoTo Failed
    canonical = "SYSTEM=" & UCase$(Trim$(ModuleHelper.ToolConfig("SAP_SYSTEM", ""))) & vbLf & _
                "CLIENT=" & UCase$(Trim$(ModuleHelper.ToolConfig("SAP_CLIENT", ""))) & vbLf & _
                "USER=" & UCase$(Trim$(ModuleHelper.ToolConfig("SAP_USER", ""))) & vbLf
    For Each oneKey In Split(CF_MAP_KEYS, ",")
        keyText = UCase$(Trim$(CStr(oneKey & "")))
        If keyText <> "" Then canonical = canonical & keyText & "=" & Trim$(ModuleHelper.ToolConfig(keyText, "")) & vbLf
    Next oneKey
    h = 104729#
    For i = 1 To Len(canonical)
        codePoint = AscW(Mid$(canonical, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536
        h = h * 65599# + CDbl(codePoint)
        h = h - Int(h / 2147483647#) * 2147483647#
    Next i
    If h < 0# Then h = 0#
    If h >= 2147483647# Then h = 2147483646#
    AuditMapFingerprint = "MAPFP1-" & Hex$(CLng(h)) & "-" & CStr(Len(canonical))
    Exit Function
Failed:
    AuditMapFingerprint = ""
End Function

Private Function AuditSummary(ByVal configOK As Boolean, ByVal integrityOK As Boolean, ByVal compileOK As Boolean, ByVal snapshotOK As Boolean, _
                              ByVal plannerOK As Boolean, ByVal mapOK As Boolean, ByVal actionMapOK As Boolean, _
                              ByVal guardsOK As Boolean, ByVal timeoutOK As Boolean, ByVal loggingOK As Boolean, _
                              ByVal executorOK As Boolean, ByVal resultText As String, ByVal runID As String, _
                              ByVal blockers As Collection, ByVal elapsedSeconds As Double) As String
    Dim s As String, i As Long
    s = "PRE-LIVE AUDIT" & vbCrLf & _
        "Config Contract... " & PassText(configOK) & vbCrLf & _
        "Integrity........ " & PassText(integrityOK) & vbCrLf & _
        "Compile.......... " & PassText(compileOK) & vbCrLf & _
        "Snapshot......... " & PassText(snapshotOK) & vbCrLf & _
        "Planner.......... " & PassText(plannerOK) & vbCrLf & _
        "SAP Map.......... " & PassText(mapOK) & vbCrLf & _
        "Action Map....... " & PassText(actionMapOK) & vbCrLf & _
        "Write Guards..... " & PassText(guardsOK) & vbCrLf & _
        "Timeouts......... " & PassText(timeoutOK) & vbCrLf & _
        "Logging.......... " & PassText(loggingOK) & vbCrLf & _
        "CF Executor...... " & PassText(executorOK) & vbCrLf & _
        "CF Pilot......... " & IIf(resultText = AUDIT_READY, "READY", "FAIL") & vbCrLf & _
        "RESULT = " & resultText & vbCrLf & "RunID = " & runID & vbCrLf & _
        "Elapsed = " & Format$(elapsedSeconds, "0.000") & " s"
    If blockers.count > 0 Then
        s = s & vbCrLf & "BLOCKERS:"
        For i = 1 To blockers.count
            s = s & vbCrLf & "- " & CStr(blockers(i))
        Next i
    End If
    AuditSummary = s
End Function

Private Function PassText(ByVal ok As Boolean) As String
    If ok Then PassText = "PASS" Else PassText = "FAIL"
End Function

Private Sub AddBlocker(ByVal blockers As Collection, ByVal text As String)
    If blockers Is Nothing Then Exit Sub
    If Trim$(text) <> "" Then blockers.Add text
End Sub

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Sub AuditSetConfig(ByVal keyName As String, ByVal valueText As String)
    Dim ws As Worksheet, lastRow As Long, r As Long, found As Boolean
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("CONFIG")
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 3 To lastRow
        If StrComp(UCase$(Trim$(CStr(ws.Cells(r, 1).Value2 & ""))), UCase$(keyName), vbTextCompare) = 0 Then
            ws.Cells(r, 2).Value2 = valueText
            found = True
            Exit For
        End If
    Next r
    If Not found Then ws.Cells(lastRow + 1, 1).Value2 = keyName: ws.Cells(lastRow + 1, 2).Value2 = valueText
    ModuleHelper.ResetToolConfigCache
ExitHere:
End Sub

Public Sub CF_PreLiveEnsureActionLogSchema()
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("RUN_LOG")
    ws.Range("R8:X8").Value = Array("Expected Before", "Live Before", "Target", "Live After", "Result", "Error Stage", "Total MS")
    ws.Range("R8:X8").Font.Bold = True
ExitHere:
End Sub

Public Sub CF_PreLiveLogActionStage(ByVal runID As String, ByVal planID As String, ByVal stageName As String, _
                                    ByVal ob As String, ByVal wt As String, ByVal actionName As String, _
                                    ByVal expectedBefore As String, ByVal liveBefore As String, ByVal targetText As String, _
                                    ByVal liveAfter As String, ByVal resultText As String, ByVal errorStage As String, _
                                    ByVal totalMs As Double)
    Dim ws As Worksheet, r As Long
    On Error GoTo ExitHere
    CF_PreLiveEnsureActionLogSchema
    Set ws = ThisWorkbook.Worksheets("RUN_LOG")
    r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    If r < 9 Then r = 9
    ws.Cells(r, 1).Value2 = Now: ws.Cells(r, 2).Value2 = runID: ws.Cells(r, 3).Value2 = stageName
    ws.Cells(r, 4).Value2 = ob: ws.Cells(r, 8).Value2 = wt: ws.Cells(r, 10).Value2 = actionName
    ws.Cells(r, 11).Value2 = resultText: ws.Cells(r, 12).Value2 = "PlanID=" & planID & "; ErrorStage=" & errorStage
    ws.Cells(r, 18).Value2 = expectedBefore: ws.Cells(r, 19).Value2 = liveBefore: ws.Cells(r, 20).Value2 = targetText
    ws.Cells(r, 21).Value2 = liveAfter: ws.Cells(r, 22).Value2 = resultText: ws.Cells(r, 23).Value2 = errorStage
    ws.Cells(r, 24).Value2 = totalMs
ExitHere:
End Sub
