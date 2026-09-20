Attribute VB_Name = "modCF_CheckPlanner"
Option Explicit

' User-facing CHECK planner.  V2.3 scheduler-contract build: CHECK writes the
' execution/group metadata consumed by the locked outbound-WT scheduler.
' This module is intentionally read-only against SAP: it may load /SCWM/POST
' only through the existing read-only loader, then builds PLAN / source legs /
' release candidates. No Cancel, Save, Create, CF, MIGO, or SAP write exists.

Private Const CP_PLAN As String = "PLAN"
Private Const CP_PLAN_LEGS As String = "PLAN_SOURCE_LEGS"
Private Const CP_PLAN_RELEASE As String = "PLAN_RELEASE"
Private Const CP_WH As String = "DATA_CORE"
Private Const CP_POST_CACHE As String = "POSTING_CACHE"
Private Const CP_FIRST_WH_ROW As Long = 5
Private Const CP_EPS As Double = 0.000001
Private Const CP_COL_QTY_MODE As Long = 28
Private Const CP_COL_EXEC_MODE As Long = 29
Private Const CP_COL_EXEC_STATE As Long = 30
Private Const CP_COL_DOC_GROUP As Long = 31
Private Const CP_COL_LOCATION_GROUP As Long = 32
Private Const CP_COL_CREATE_GROUP As Long = 33
Private Const CP_COL_PLAN_FINGERPRINT As Long = 34
Private Const CP_COL_QTY_DELTA As Long = 35
' Group contract is deliberately outside the legacy A:AT lock metadata.  This
' keeps the existing lock cells stable while CHECK adds an explicit aggregate
' quantity contract consumed by Button #4.
Private Const CP_COL_QTY_GROUP_ID As Long = 47
Private Const CP_COL_GROUP_CURRENT As Long = 48
Private Const CP_COL_GROUP_TARGET As Long = 49
Private Const CP_COL_GROUP_DELTA As Long = 50
Private Const CP_COL_GROUP_MEMBERS As Long = 51
Private Const CP_COL_GROUP_NEED_QTY As Long = 52

Public gCP_LastCheckResult As String

Public Sub CF_Check()
    Dim wsWH As Worksheet, wsPlan As Worksheet
    Dim exceptions As Object, fastRows As Long, checkedRows As Long
    Dim messageText As String, started As Double, oldSilent As Boolean
    Dim checkResult As String, readyRows As Long, stopRows As Long
    Dim postingRefreshRequired As Boolean
    Dim failureNumber As Long, failureDescription As String
    Dim checkHistoryRunID As String

    On Error GoTo Failed
    oldSilent = gRT_Silent
    gCP_LastCheckResult = ""
    started = RT_StartTimer()
    RT_SetStatus "STAGE", "CHECK PREPARE"
    Set wsWH = ThisWorkbook.Worksheets(CP_WH)
    Set wsPlan = CP_RequireSheet(CP_PLAN, False)
    CF_HistoryEnsureLedger
    checkHistoryRunID = CF_HistoryNewRunID("CHECK", CP_Text(wsWH.Range("W2").Value2))
    CP_ResetOutputs wsPlan
    Set exceptions = CreateObject("Scripting.Dictionary")
    exceptions.CompareMode = vbTextCompare

    If Not CP_ValidateSnapshot(wsWH, messageText) Then Err.Raise vbObjectError + 9810, , messageText
    If Not AP_NormalizeAndValidateAll(wsWH, messageText) Then Err.Raise vbObjectError + 9811, , messageText
    If Not CP_ValidateUserInputGate(wsWH, messageText) Then Err.Raise vbObjectError + 9812, , messageText
    CP_Preclassify wsWH, wsPlan, exceptions, fastRows, checkedRows

    If exceptions.count > 0 Then
        ' CHECK is a deterministic local decision step.  It must never start,
        ' navigate, or reuse SAP /SCWM/POST implicitly.  A user-triggered
        ' TAI POSTING refreshes the authoritative cache; otherwise the plan
        ' remains stopped with the precise cache/capture reason.
        If Not CP_ValidatePostingCache(wsWH, messageText) Then
            CP_MarkPostingUnavailable wsPlan, messageText
            postingRefreshRequired = CP_IsPostingRefreshRequired(messageText)
        Else
            CP_ResolveExceptions wsWH, wsPlan
        End If
    End If

    ' Resolve any batch inferred from the authoritative Posting snapshot before
    ' deriving group keys.  The executor must never receive a source group with
    ' a blank batch or a per-row quantity contract when the ODO item is shared.
    CP_RefreshPlanGroupKeys wsPlan
    CP_BuildQtyGroupContract wsPlan

    CP_BuildReleasePlan wsWH, wsPlan
    CP_PlanOutcome wsPlan, readyRows, stopRows, checkResult
    If postingRefreshRequired Then checkResult = "PLAN_NEEDS_TAI_POSTING"
    gCP_LastCheckResult = checkResult
    CP_WriteSummary wsPlan, fastRows, checkedRows, exceptions.count, started, checkResult, readyRows, stopRows
    wsPlan.Range("AM1").Value2 = "CHECK HISTORY RUN ID"
    wsPlan.Range("AM2").Value2 = checkHistoryRunID
    CP_RecordCheckHistory wsPlan, checkHistoryRunID
    CF_HistoryAppendSummary checkHistoryRunID, "CHECK", checkResult, _
        CP_Text(ThisWorkbook.Worksheets("CF_CONTROL").Range("A7").Value2), checkedRows, readyRows, stopRows, _
        0, 0, 0, checkResult, "Posting materials=" & CStr(exceptions.count) & "; SAP_WRITE=0"
    RT_SetStatus "PLAN", checkResult
    RT_SetStatus "TIME", Format$(RT_Elapsed(started), "0.000") & " s"
    If postingRefreshRequired Then
        RT_SetStatus "STAGE", "CHECK: TAI POSTING REQUIRED"
    Else
        RT_SetStatus "STAGE", "CHECK DONE"
    End If
    RT_Log "CHECK", "", "", "PASS", checkResult & "; fast=" & CStr(fastRows) & _
           "; postingMaterial=" & CStr(exceptions.count) & _
           IIf(postingRefreshRequired, "; NEXT=TAI_POSTING_THEN_CHECK", "") & _
           "; SAP_WRITE=0; CANCEL=0; SAVE=0; CREATE=0; CF=0"
    If Not gRT_Silent Then
        MsgBox "CHECK hoàn tất: " & checkResult & vbCrLf & _
               "Fast path: " & CStr(fastRows) & vbCrLf & _
               "Posting materials: " & CStr(exceptions.count) & vbCrLf & _
               "Xem Decision / Reason tại WH_ACTUAL và sheet PLAN.", vbInformation, "CHECK"
    End If
    CF_RefreshInputWorklist
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    If failureDescription = "" Then failureDescription = "Unexpected CHECK failure without VBA error detail."
    On Error Resume Next
    gCP_LastCheckResult = "STOP|" & failureDescription
    If checkHistoryRunID = "" Then checkHistoryRunID = CF_HistoryNewRunID("CHECK_FAIL")
    CF_HistoryAppend checkHistoryRunID, "CHECK", "CHECK_FAIL", "", "", "", "", "", "", _
        "CHECK_PLAN", "FAIL", "", "", "", "", failureDescription
    CF_HistoryAppendSummary checkHistoryRunID, "CHECK", "CHECK_FAIL", "", 0, 0, 1, 0, 0, 0, _
        "FAIL", failureDescription
    gRT_Silent = oldSilent
    Set gMPOST_CheckMaterialFilter = Nothing
    RT_SetStatus "STAGE", "CHECK FAIL"
    RT_SetStatus "TIME", Format$(RT_Elapsed(started), "0.000") & " s"
    RT_Log "CHECK", "", "", "FAIL", "Err " & CStr(failureNumber) & ": " & failureDescription & _
           "; SAP_WRITE=0; CANCEL=0; SAVE=0; CREATE=0; CF=0"
    If Not gRT_Silent Then MsgBox "CHECK failed" & vbCrLf & failureDescription, vbExclamation, "CHECK"
End Sub

' Script-safe CHECK entry point.  It keeps the same local planning behavior
' as the user button but suppresses only Excel message boxes; it does not
' load SAP, send a write, Cancel, Create, or CF.
Public Sub CF_Check_Silent()
    Dim oldSilent As Boolean
    On Error GoTo Failed
    oldSilent = gRT_Silent
    gRT_Silent = True
    CF_Check
CleanExit:
    gRT_Silent = oldSilent
    Exit Sub
Failed:
    gRT_Silent = oldSilent
    Err.Raise Err.Number, , Err.Description
End Sub

Public Sub CF_CheckPlanner_CompileProbe()
    Dim decision As String, pipeline As String, reasonText As String
    CP_Classify "OB_DONE", "WT_ACTIVE", 10#, 10#, "BIN-A", "BIN-A", "B1", "B1", _
                decision, pipeline, reasonText
    CP_Assert decision = "DONE", "OB_DONE"
    CP_Classify "OB_OPEN_OR_PARTIAL", "WT_ACTIVE", 0#, 10#, "BIN-A", "BIN-A", "B1", "B1", _
                decision, pipeline, reasonText
    CP_Assert decision = "CANCEL_TO_ZERO", "CANCEL_TO_ZERO"
    CP_Classify "OB_OPEN_OR_PARTIAL", "WT_ACTIVE", 10#, 10#, "BIN-A", "BIN-A", "B1", "B1", _
                decision, pipeline, reasonText
    CP_Assert decision = "CF_ONLY", "CF_ONLY"
    CP_Classify "OB_OPEN_OR_PARTIAL", "WT_ACTIVE", 6#, 10#, "BIN-A", "BIN-A", "B1", "B1", _
                decision, pipeline, reasonText
    CP_Assert decision = "CHANGE_QTY", "CHANGE_QTY"
    CP_Classify "OB_OPEN_OR_PARTIAL", "WT_ACTIVE", 12#, 10#, "BIN-A", "BIN-A", "B1", "B1", _
                decision, pipeline, reasonText
    CP_Assert decision = "POSTING_REQUIRED", "QTY_INCREASE"
    CP_Classify "OB_OPEN_OR_PARTIAL", "WT_ACTIVE", 10#, 10#, "BIN-B", "BIN-A", "B1", "B1", _
                decision, pipeline, reasonText
    CP_Assert decision = "POSTING_REQUIRED", "BIN_CHANGE"
    CP_Classify "OB_OPEN_OR_PARTIAL", "WT_CANCELLED_ZERO", 10#, 10#, "BIN-A", "BIN-A", "B1", "B1", _
                decision, pipeline, reasonText
    CP_Assert decision = "STOP_OB_QTY_ZERO", "A_ZERO"
    CP_Assert Abs(CP_ProjectAvailableAfterOwnCancel(100#, 600#, 500#) - 600#) <= CP_EPS, "SELF_HELD_PROJECTED"
    CP_Assert Abs(CP_ProjectAvailableAfterOwnCancel(0#, 293#, 293#) - 293#) <= CP_EPS, "SELF_HELD_SHORTAGE_REMAINS"
    CP_Assert CP_IsPostingRefreshRequired("STOP_STALE_POSTING_CACHE: cache RunID differs."), "POSTING_REFRESH_REQUIRED"
    CP_Assert Not CP_IsPostingRefreshRequired("STOP_POSTING_NO_STOCK: full capture returned no stock."), "POSTING_NO_STOCK_IS_RESULT"
    CP_Assert CP_DeriveQtyMode(150#, 100#, "BIN-A", "BIN-A", "B1", "B1") = "DELTA_ADD", "DELTA_ADD_MODE"
    CP_Assert CP_DeriveQtyMode(80#, 100#, "BIN-A", "BIN-A", "B1", "B1") = "FINAL_TOTAL", "FINAL_TOTAL_MODE"
End Sub

Public Sub CF_CheckPlanner_LocalSnapshotProbe()
    Dim wsWH As Worksheet, wsPlan As Worksheet, exceptions As Object
    Dim msg As String, fastRows As Long, checkedRows As Long
    Set wsWH = ThisWorkbook.Worksheets(CP_WH)
    If Not CP_ValidateSnapshot(wsWH, msg) Then Err.Raise vbObjectError + 9821, , msg
    Set wsPlan = CP_RequireSheet(CP_PLAN, False)
    CP_ResetOutputs wsPlan
    Set exceptions = CreateObject("Scripting.Dictionary")
    exceptions.CompareMode = vbTextCompare
    CP_Preclassify wsWH, wsPlan, exceptions, fastRows, checkedRows
    CP_Assert checkedRows > 0, "LOCAL_ROWS"
End Sub

Public Function CF_CheckPlanner_ValidateSnapshot(ByRef messageText As String) As Boolean
    Dim wsWH As Worksheet
    On Error GoTo Failed
    Set wsWH = ThisWorkbook.Worksheets(CP_WH)
    CF_CheckPlanner_ValidateSnapshot = CP_ValidateSnapshot(wsWH, messageText)
    Exit Function
Failed:
    messageText = "STOP_STALE_WH_SNAPSHOT: validation failed: " & Err.Description
End Function

Private Function CP_ValidateSnapshot(ByVal ws As Worksheet, ByRef messageText As String) As Boolean
    Dim runID As String, lastRow As Long, gateText As String
    Dim wsQ As Worksheet, wsCache As Worksheet, wsInput As Worksheet
    Dim qLast As Long, cLast As Long, r As Long, qRunID As String, cacheRunID As String
    Dim inputRows As Object, queueRows As Object, cacheKeys As Object, actualKeys As Object
    Dim obText As String, returnedOB As String, sourceKey As String, inputOB As String
    Dim inputRowNo As Long, queueCount As Long, keyItem As Variant
    If ws Is Nothing Then
        messageText = "STOP_SNAPSHOT_MISSING: WH_ACTUAL is missing."
        Exit Function
    End If
    runID = CP_Text(ws.Range("W2").Value2)
    lastRow = CP_LastRow(ws, 1, CP_FIRST_WH_ROW)
    gateText = CP_Text(ws.Range("X2").Value2)
    If Left$(gateText, Len("STOP_STALE_WH_SNAPSHOT")) = "STOP_STALE_WH_SNAPSHOT" Then
        messageText = gateText
        Exit Function
    End If
    If runID = "" Or lastRow < CP_FIRST_WH_ROW Then
        messageText = "STOP_SNAPSHOT_NOT_READY: run TẢI WH before CHECK."
        Exit Function
    End If
    On Error GoTo Failed
    Set wsQ = ThisWorkbook.Worksheets("OB_QUEUE")
    Set wsCache = ThisWorkbook.Worksheets("OB_WH_CACHE")
    Set wsInput = ThisWorkbook.Worksheets("CF_CONTROL")
    Set inputRows = CreateObject("Scripting.Dictionary"): inputRows.CompareMode = vbTextCompare
    Set queueRows = CreateObject("Scripting.Dictionary"): queueRows.CompareMode = vbTextCompare
    Set cacheKeys = CreateObject("Scripting.Dictionary"): cacheKeys.CompareMode = vbTextCompare
    Set actualKeys = CreateObject("Scripting.Dictionary"): actualKeys.CompareMode = vbTextCompare
    For r = 7 To 506
        inputOB = CP_Text(wsInput.Cells(r, 1).Value2)
        If CP_IsOBInput(inputOB) Then inputRows(CStr(r)) = inputOB
    Next r
    If inputRows.count = 0 Then
        messageText = "STOP_STALE_WH_SNAPSHOT: INPUT has no current OB list."
        Exit Function
    End If
    qLast = CP_LastRow(wsQ, 1, 2)
    For r = 2 To qLast
        qRunID = CP_Text(wsQ.Cells(r, 18).Value2)
        obText = CP_Text(wsQ.Cells(r, 3).Value2)
        returnedOB = CP_Text(wsQ.Cells(r, 4).Value2)
        If StrComp(qRunID, runID, vbTextCompare) = 0 Then
            If UCase$(CP_Text(wsQ.Cells(r, 7).Value2)) <> "PASS" Then
                messageText = "STOP_STALE_WH_SNAPSHOT: latest WH RunID contains a non-PASS OB."
                Exit Function
            End If
            If obText = "" Or StrComp(obText, returnedOB, vbTextCompare) <> 0 Then
                messageText = "STOP_STALE_WH_SNAPSHOT: OB_QUEUE Requested/Returned mismatch."
                Exit Function
            End If
            inputRowNo = CLng(val(wsQ.Cells(r, 2).Value2))
            If inputRowNo < 7 Or Not inputRows.Exists(CStr(inputRowNo)) Then
                messageText = "STOP_STALE_WH_SNAPSHOT: OB_QUEUE input row is not in current INPUT."
                Exit Function
            End If
            If StrComp(CStr(inputRows(CStr(inputRowNo))), obText, vbTextCompare) <> 0 Then
                messageText = "STOP_STALE_WH_SNAPSHOT: INPUT OB differs from latest OB_QUEUE batch."
                Exit Function
            End If
            If queueRows.Exists(CStr(inputRowNo)) Then
                messageText = "STOP_STALE_WH_SNAPSHOT: duplicate OB_QUEUE input row in latest batch."
                Exit Function
            End If
            queueRows(CStr(inputRowNo)) = True
            queueCount = queueCount + 1
        End If
    Next r
    If queueCount <> inputRows.count Then
        messageText = "STOP_STALE_WH_SNAPSHOT: INPUT OB count does not equal latest PASS batch."
        Exit Function
    End If
    For Each keyItem In inputRows.keys
        If Not queueRows.Exists(CStr(keyItem)) Then
            messageText = "STOP_STALE_WH_SNAPSHOT: INPUT row is absent from latest PASS batch: " & CStr(keyItem)
            Exit Function
        End If
    Next keyItem
    cLast = CP_LastRow(wsCache, 1, 2)
    For r = 2 To cLast
        cacheRunID = CP_Text(wsCache.Cells(r, 20).Value2)
        If StrComp(cacheRunID, runID, vbTextCompare) = 0 Then
            If UCase$(CP_Text(wsCache.Cells(r, 16).Value2)) <> "PASS" Then
                messageText = "STOP_STALE_WH_SNAPSHOT: OB_WH_CACHE has a non-PASS row."
                Exit Function
            End If
            sourceKey = CP_Text(wsCache.Cells(r, 18).Value2)
            If sourceKey = "" Or cacheKeys.Exists(sourceKey) Then
                messageText = "STOP_STALE_WH_SNAPSHOT: OB_WH_CACHE SourceKey is empty or duplicate."
                Exit Function
            End If
            cacheKeys(sourceKey) = True
        End If
    Next r
    If cacheKeys.count = 0 Then
        messageText = "STOP_STALE_WH_SNAPSHOT: no PASS cache row matches WH_ACTUAL RunID."
        Exit Function
    End If
    For r = CP_FIRST_WH_ROW To lastRow
        sourceKey = CP_Text(ws.Cells(r, 20).Value2)
        If sourceKey = "" Or actualKeys.Exists(sourceKey) Then
            messageText = "STOP_STALE_WH_SNAPSHOT: WH_ACTUAL SourceKey is empty or duplicate."
            Exit Function
        End If
        actualKeys(sourceKey) = True
    Next r
    If actualKeys.count <> cacheKeys.count Then
        messageText = "STOP_STALE_WH_SNAPSHOT: WH_ACTUAL row count does not equal cache row count."
        Exit Function
    End If
    For Each keyItem In cacheKeys.keys
        If Not actualKeys.Exists(CStr(keyItem)) Then
            messageText = "STOP_STALE_WH_SNAPSHOT: WH_ACTUAL SourceKey differs from cache."
            Exit Function
        End If
    Next keyItem
    CP_ValidateSnapshot = True
    Exit Function
Failed:
    messageText = "STOP_STALE_WH_SNAPSHOT: validation failed: " & Err.Description
End Function

' CHECK is only allowed after the user has supplied the picking intent.
' Qty and bin are mandatory for every live row.  Actual Batch remains optional:
' when it is empty the planner retains the WH batch as the requirement and
' resolves eligible stock at the user-selected source bin from Posting data.
Private Function CP_ValidateUserInputGate(ByVal ws As Worksheet, ByRef messageText As String) As Boolean
    Dim lastRow As Long, r As Long
    Dim material As String, obState As String, wtState As String
    Dim actualQty As String, actualBin As String
    On Error GoTo Failed

    lastRow = CP_LastRow(ws, 1, CP_FIRST_WH_ROW)
    For r = CP_FIRST_WH_ROW To lastRow
        material = CP_Text(ws.Cells(r, 6).Value2)
        obState = UCase$(CP_Text(ws.Cells(r, 25).Value2))
        wtState = UCase$(CP_Text(ws.Cells(r, 26).Value2))
        If material <> "" And obState = "OB_OPEN_OR_PARTIAL" And _
           (wtState = "WT_ACTIVE" Or wtState = "WT_CANCELLED_REMAINING") Then
            actualQty = CP_Text(ws.Cells(r, 12).Value2)
            actualBin = CP_Text(ws.Cells(r, 14).Value2)
            If actualQty = "" Then
                messageText = "STOP_USER_INPUT_REQUIRED: row " & CStr(r) & _
                    " needs Actual Qty before CHECK."
                Exit Function
            End If
            If actualBin = "" Then
                messageText = "STOP_USER_INPUT_REQUIRED: row " & CStr(r) & _
                    " needs Area/Bin before CHECK."
                Exit Function
            End If
        End If
    Next r
    CP_ValidateUserInputGate = True
    Exit Function
Failed:
    messageText = "STOP_USER_INPUT_REQUIRED: input-gate validation failed: " & Err.Description
End Function

Private Function CP_IsOBInput(ByVal valueText As String) As Boolean
    Dim s As String
    s = Replace(Trim$(valueText), " ", "")
    If Len(s) < 6 Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    CP_IsOBInput = True
End Function

Private Sub CP_Preclassify(ByVal wsWH As Worksheet, ByVal wsPlan As Worksheet, _
                           ByVal exceptions As Object, ByRef fastRows As Long, _
                           ByRef checkedRows As Long)
    Dim lastRow As Long, r As Long, outRow As Long
    Dim obState As String, wtState As String, material As String
    Dim sysQty As Double, targetQty As Double, sysBin As String, targetBin As String
    Dim sysBatch As String, targetBatch As String
    Dim decision As String, pipeline As String, reasonText As String, gateText As String

    lastRow = CP_LastRow(wsWH, 1, CP_FIRST_WH_ROW)
    outRow = 2
    For r = CP_FIRST_WH_ROW To lastRow
        material = CP_Text(wsWH.Cells(r, 6).Value2)
        If material = "" Then GoTo NextRow
        checkedRows = checkedRows + 1
        obState = UCase$(CP_Text(wsWH.Cells(r, 25).Value2))
        wtState = UCase$(CP_Text(wsWH.Cells(r, 26).Value2))
        sysQty = CP_Number(wsWH.Cells(r, 7).Value2)
        targetQty = CP_Number(wsWH.Cells(r, 12).Value2)
        sysBin = CP_Text(wsWH.Cells(r, 8).Value2)
        targetBin = CP_Text(wsWH.Cells(r, 14).Value2)
        sysBatch = CP_Text(wsWH.Cells(r, 9).Value2)
        targetBatch = CP_Text(wsWH.Cells(r, 22).Value2)
        If targetBin = "" Then targetBin = sysBin
        ' Same-bin rows retain their proven System Batch.  For a changed bin,
        ' a blank Actual Batch means the planner must resolve source legs from
        ' active WH at that bin instead of guessing a batch from Posting stock.
        If targetBatch = "" And StrComp(targetBin, sysBin, vbTextCompare) = 0 Then targetBatch = sysBatch

        If CP_Text(wsWH.Cells(r, 12).Value2) = "" And obState = "OB_OPEN_OR_PARTIAL" And _
           (wtState = "WT_ACTIVE" Or wtState = "WT_CANCELLED_REMAINING") Then
            decision = "STOP_STATE_NOT_PROVEN"
            pipeline = ""
            reasonText = "Actual Qty is required; use 0 only when the user explicitly skips this WT."
        Else
            CP_Classify obState, wtState, targetQty, sysQty, targetBin, sysBin, targetBatch, sysBatch, _
                        decision, pipeline, reasonText
        End If
        gateText = "PRECHECK"
        If decision = "CANCEL_TO_ZERO" Then gateText = "PASS"
        CP_WritePlanRow wsPlan, outRow, wsWH, r, targetQty, targetBin, targetBatch, _
                        decision, gateText, reasonText, pipeline
        CP_WriteExecutionContract wsPlan, outRow, decision
        CP_WriteWHDecision wsWH, r, decision, reasonText, pipeline
        ' Both a replacement and a quantity-only rework need an exact
        ' current source before Button #4 can receive a locked execution row.
        ' Previously CHANGE_QTY never entered the source resolver, leaving the
        ' executor to discover that gap late.
        If decision = "POSTING_REQUIRED" Or decision = "CHANGE_QTY" Then
            exceptions(material) = True
        Else
            fastRows = fastRows + 1
        End If
        outRow = outRow + 1
NextRow:
    Next r
End Sub

Private Sub CP_Classify(ByVal obState As String, ByVal wtState As String, _
                        ByVal targetQty As Double, ByVal sysQty As Double, _
                        ByVal targetBin As String, ByVal sysBin As String, _
                        ByVal targetBatch As String, ByVal sysBatch As String, _
                        ByRef decision As String, ByRef pipeline As String, _
                        ByRef reasonText As String)
    Dim needQty As Boolean, needBin As Boolean, needBatch As Boolean
    decision = "STOP_STATE_NOT_PROVEN"
    pipeline = ""
    reasonText = "OB state is not proven."

    If obState = "OB_DONE" Then
        decision = "DONE": reasonText = "OB_DONE: no Posting or user action."
        Exit Sub
    End If
    If obState <> "OB_OPEN_OR_PARTIAL" Then Exit Sub
    If wtState = "WT_DONE" Then
        decision = "DONE": reasonText = "WT_DONE: confirmed WT."
        Exit Sub
    End If
    If wtState = "WT_CANCELLED_ZERO" Then
        decision = "STOP_OB_QTY_ZERO": reasonText = "Cancelled WT has Remaining Qty=0."
        Exit Sub
    End If
    If wtState = "WT_CANCELLED_PENDING_REMAINING" Then
        ' Status A is a resumable, already-cancelled member.  CHECK must still
        ' lock it into the same Document/Item QtyGroup so Button #4 can read the
        ' aggregate PRDO total (for example 246 -> 256) before any write.
        If targetQty < 0# Or targetQty <= CP_EPS Then
            If targetQty <= CP_EPS Then
                decision = "CANCEL_TO_ZERO": pipeline = "RECONCILE_CANCELLED_WT -> SET_PRDO_QTY_ZERO -> USER_CF_HANDOFF"
                reasonText = "WT status A; aggregate PRDO zero must be proven before handoff."
            Else
                decision = "STOP_STATE_NOT_PROVEN": reasonText = "Actual Qty cannot be negative."
            End If
            Exit Sub
        End If
        If sysQty <= 0# Or targetBin = "" Then
            decision = "STOP_STATE_NOT_PROVEN": reasonText = "Missing Sys Qty or Target Bin."
            Exit Sub
        End If
        needQty = Abs(targetQty - sysQty) > CP_EPS
        needBin = StrComp(targetBin, sysBin, vbTextCompare) <> 0
        needBatch = StrComp(targetBatch, sysBatch, vbTextCompare) <> 0
        If Not needQty And Not needBin And Not needBatch Then
            decision = "CF_ONLY": pipeline = "CF_USER_HANDOFF"
            reasonText = "Cancelled WT status A reconciled; no target change remains for SAP."
        Else
            decision = "POSTING_REQUIRED"
            pipeline = "RECONCILE_CANCELLED_WT -> "
            If needQty Then pipeline = pipeline & "CHANGE_QTY -> "
            If needBin Or needBatch Then pipeline = pipeline & "POST_CHANGE_BIN_BATCH -> "
            pipeline = pipeline & "CREATE_OR_REACQUIRE_WT -> CF_USER_HANDOFF"
            reasonText = "WT status A is accepted only as a group-resumable row; PRDO aggregate readback is mandatory."
        End If
        Exit Sub
    End If
    If wtState <> "WT_ACTIVE" And wtState <> "WT_CANCELLED_REMAINING" Then
        decision = "STOP_STATE_NOT_PROVEN": reasonText = "WT state is not eligible."
        Exit Sub
    End If
    If targetQty < 0# Then
        decision = "STOP_STATE_NOT_PROVEN": reasonText = "Actual Qty cannot be negative."
        Exit Sub
    End If
    If targetQty <= CP_EPS Then
        decision = "CANCEL_TO_ZERO"
        pipeline = "CANCEL_OLD_WT -> SET_PRDO_QTY_ZERO -> USER_CF_HANDOFF"
        reasonText = "Actual Qty=0: cancel the exact active WT and set the exact PRDO item Qty to 0; no new WT is created."
        Exit Sub
    End If
    If sysQty <= 0# Or targetBin = "" Then
        decision = "STOP_STATE_NOT_PROVEN": reasonText = "Missing Sys Qty or Target Bin."
        Exit Sub
    End If

    needQty = Abs(targetQty - sysQty) > CP_EPS
    needBin = StrComp(targetBin, sysBin, vbTextCompare) <> 0
    needBatch = StrComp(targetBatch, sysBatch, vbTextCompare) <> 0
    If Not needQty And Not needBin And Not needBatch Then
        decision = "CF_ONLY": pipeline = "CF_USER_HANDOFF": reasonText = "Target equals Sys Qty/Bin/Batch; stop for user CF."
    ElseIf needQty And Not needBin And Not needBatch And targetQty < sysQty - CP_EPS Then
        decision = "CHANGE_QTY": pipeline = "CHANGE_QTY -> CF_USER_HANDOFF": reasonText = "Qty decreased; stop before CF for user."
    Else
        decision = "POSTING_REQUIRED"
        If targetQty > sysQty + CP_EPS Then pipeline = "CHANGE_QTY"
        If needBin Or needBatch Then
            If pipeline <> "" Then pipeline = pipeline & " -> "
            pipeline = pipeline & "POST_CHANGE_BIN_BATCH"
        End If
        If pipeline = "" Then pipeline = "POST_CHANGE_BIN_BATCH"
        pipeline = "RELEASE_OTHER_WT (if needed) -> " & pipeline & " -> CREATE_OR_REACQUIRE_WT -> CF_USER_HANDOFF"
        reasonText = "Posting required; stop before CF for user confirmation."
    End If
End Sub

Private Sub CP_ResolveExceptions(ByVal wsWH As Worksheet, ByVal wsPlan As Worksheet)
    Dim postIndex As Object, blockedIndex As Object, wsPost As Worksheet
    Dim lastPlan As Long, r As Long, planRow As Long
    Dim material As String, targetBin As String, targetBatch As String, sourceKey As String
    Dim legs As Collection, leg As Variant, requiredQty As Double, sourceRequiredQty As Double
    Dim available As Double, totalStock As Double, selfHeldQty As Double, decision As String, gateText As String, reasonText As String
    Dim pipeline As String, planID As String, lineNo As Long, qtyMode As String

    Set wsPost = CP_RequireSheet(CP_POST_CACHE, True)
    Set postIndex = CP_BuildPostingIndex(wsPost)
    Set blockedIndex = CP_BuildBlockedPostingIndex(wsPost)
    CP_ResetLegs
    lastPlan = CP_LastRow(wsPlan, 1, 2)
    lineNo = 1
    For r = 2 To lastPlan
        ' Quantity-only plans also need one current, exact stock source before
        ' an executor may change PRDO quantity and create a replacement WT.
        If UCase$(CP_Text(wsPlan.Cells(r, 20).Value2)) <> "POSTING_REQUIRED" And _
           UCase$(CP_Text(wsPlan.Cells(r, 20).Value2)) <> "CHANGE_QTY" Then GoTo NextPlan
        material = CP_Text(wsPlan.Cells(r, 6).Value2)
        targetBin = CP_Text(wsPlan.Cells(r, 12).Value2)
        targetBatch = CP_Text(wsPlan.Cells(r, 14).Value2)
        qtyMode = CP_Text(wsPlan.Cells(r, CP_COL_QTY_MODE).Value2)
        sourceRequiredQty = CP_SourceRequiredQty(wsPlan, r)
        If sourceRequiredQty <= CP_EPS Then
            decision = "STOP_SOURCE_PLAN_INCOMPLETE"
            gateText = "SOURCE_QTY"
            reasonText = "Source quantity is not positive for the locked plan."
            CP_UpdatePlanDecision wsPlan, r, decision, gateText, reasonText, pipeline
            GoTo NextPlan
        End If
        planID = CP_Text(wsPlan.Cells(r, 1).Value2)
        Set legs = CP_BuildLegs(postIndex, material, targetBin, targetBatch, sourceRequiredQty)
        If legs Is Nothing Then
            decision = "STOP_SOURCE_PLAN_INCOMPLETE"
            gateText = "SOURCE_LEGS"
            reasonText = "Target Bin Posting stock does not resolve to one exact batch; enter Actual Batch when multiple batches exist."
            CP_UpdatePlanDecision wsPlan, r, decision, gateText, reasonText, pipeline
            GoTo NextPlan
        End If

        decision = "READY"
        gateText = "SOURCE_READY"
        reasonText = ""
        For Each leg In legs
            sourceKey = CStr(leg(0))
            requiredQty = CDbl(leg(1))
            If postIndex.Exists(sourceKey) Then
                available = CDbl(postIndex(sourceKey)(0))
                totalStock = CDbl(postIndex(sourceKey)(1))
            Else
                available = 0#: totalStock = 0#
            End If
            ' DELTA_ADD preserves the existing WT.  Its required source is
            ' only the incremental quantity and it must never be made ready by
            ' projecting a Cancel.  Replacement routes retain the established
            ' exact-own-WT projection and later reconcile before a write.
            selfHeldQty = 0#
            If qtyMode <> "DELTA_ADD" Then
                selfHeldQty = CP_SelfHeldQtyForPlanSource(wsWH, wsPlan, r, sourceKey)
                available = CP_ProjectAvailableAfterOwnCancel(available, totalStock, selfHeldQty)
            End If
            CP_WriteLeg planID, lineNo, sourceKey, requiredQty, available, totalStock
            lineNo = lineNo + 1
            If Not postIndex.Exists(sourceKey) Then
                If blockedIndex.Exists(sourceKey) Then
                    decision = "STOP_SOURCE_STOCK_TYPE_NOT_USABLE": gateText = "STOCK_TYPE"
                    reasonText = "Exact source exists only with non-FT stock type; QT/QI is not usable."
                Else
                    decision = "STOP_SOURCE_NOT_FOUND": gateText = "POSTING_SOURCE"
                    reasonText = "Exact Material|Bin|Batch source was not found in POSTING_CACHE."
                End If
            ElseIf totalStock + CP_EPS < requiredQty Then
                decision = "STOP_SOURCE_INSUFFICIENT": gateText = "TOTAL_STOCK"
                reasonText = "Exact source is short " & Format$(requiredQty - totalStock, "0.###") & _
                             " EA; required=" & Format$(requiredQty, "0.###") & _
                             "; proven capacity=" & Format$(totalStock, "0.###") & _
                             "; source=" & sourceKey & ". No alternate-bin allocation is automatic."
            ElseIf available + CP_EPS < requiredQty And decision = "READY" Then
                decision = "NEED_RELEASE": gateText = "AVAILABLE"
                reasonText = "Available Qty is below required Qty; release plan required."
            End If
        Next leg
        If targetBatch = "" And legs.count = 1 Then
            targetBatch = CP_KeyPart(CStr(legs(1)(0)), 3)
            wsPlan.Cells(r, 14).Value2 = targetBatch
            wsPlan.Cells(r, 23).Value2 = IIf(StrComp(targetBatch, CP_Text(wsPlan.Cells(r, 11).Value2), vbTextCompare) <> 0, "TRUE", "FALSE")
        End If
        If decision = "READY" Then
            If legs.count > 1 Then
                decision = "MULTI_SOURCE"
                reasonText = "Target Bin resolves to " & CStr(legs.count) & " active WH batch source legs."
            Else
                decision = CP_FinalDecision(wsPlan, r)
                If qtyMode = "DELTA_ADD" Then
                    reasonText = "Exact source is ready for additional Qty=" & _
                                 Format$(sourceRequiredQty, "0.###") & " EA; existing WT is retained."
                ElseIf selfHeldQty > CP_EPS Then
                    reasonText = "Exact source is ready after Cancel releases own WT reservation=" & _
                                 Format$(selfHeldQty, "0.###") & " EA."
                Else
                    reasonText = "Exact posting source is ready."
                End If
            End If
            gateText = "PASS"
        End If
        wsPlan.Cells(r, 15).Value2 = CP_PlanFirstAvailable(planID)
        wsPlan.Cells(r, 16).Value2 = CP_PlanFirstTotal(planID)
        wsPlan.Cells(r, 19).Value2 = legs.count
        CP_UpdatePlanDecision wsPlan, r, decision, gateText, reasonText, CP_Text(wsPlan.Cells(r, 25).Value2)
NextPlan:
    Next r
End Sub

Private Function CP_FinalDecision(ByVal wsPlan As Worksheet, ByVal planRow As Long) As String
    Dim needQty As Boolean, needBin As Boolean, needBatch As Boolean
    needQty = (UCase$(CP_Text(wsPlan.Cells(planRow, 21).Value2)) = "TRUE")
    needBin = (UCase$(CP_Text(wsPlan.Cells(planRow, 22).Value2)) = "TRUE")
    needBatch = (UCase$(CP_Text(wsPlan.Cells(planRow, 23).Value2)) = "TRUE")
    If needBin And needBatch Then
        CP_FinalDecision = "CHANGE_BIN_BATCH"
    ElseIf needBin Then
        CP_FinalDecision = "CHANGE_BIN"
    ElseIf needBatch Then
        CP_FinalDecision = "CHANGE_BATCH"
    ElseIf needQty Then
        CP_FinalDecision = "CHANGE_QTY"
    Else
        CP_FinalDecision = "CF_ONLY"
    End If
End Function

Private Sub CP_BuildReleasePlan(ByVal wsWH As Worksheet, ByVal wsPlan As Worksheet)
    Dim lastPlan As Long, r As Long, planID As String, decision As String
    Dim material As String, binText As String, batchText As String
    Dim available As Double, requiredQty As Double, needRelease As Double
    Dim legRow As Long, lastLeg As Long
    Dim totalNeedRelease As Double, firstAvailable As Double, needLegs As Long

    CP_ResetRelease
    lastPlan = CP_LastRow(wsPlan, 1, 2)
    For r = 2 To lastPlan
        decision = UCase$(CP_Text(wsPlan.Cells(r, 20).Value2))
        If decision <> "NEED_RELEASE" Then GoTo NextRelease
        planID = CP_Text(wsPlan.Cells(r, 1).Value2)
        lastLeg = CP_LastRow(ThisWorkbook.Worksheets(CP_PLAN_LEGS), 1, 2)
        totalNeedRelease = 0#
        firstAvailable = 0#
        needLegs = 0
        For legRow = 2 To lastLeg
            If CP_Text(ThisWorkbook.Worksheets(CP_PLAN_LEGS).Cells(legRow, 1).Value2) = planID And _
               UCase$(CP_Text(ThisWorkbook.Worksheets(CP_PLAN_LEGS).Cells(legRow, 7).Value2)) = "NEED_RELEASE" Then
                material = CP_KeyPart(CP_Text(ThisWorkbook.Worksheets(CP_PLAN_LEGS).Cells(legRow, 3).Value2), 1)
                binText = CP_KeyPart(CP_Text(ThisWorkbook.Worksheets(CP_PLAN_LEGS).Cells(legRow, 3).Value2), 2)
                batchText = CP_KeyPart(CP_Text(ThisWorkbook.Worksheets(CP_PLAN_LEGS).Cells(legRow, 3).Value2), 3)
                requiredQty = CP_Number(ThisWorkbook.Worksheets(CP_PLAN_LEGS).Cells(legRow, 4).Value2)
                available = CP_Number(ThisWorkbook.Worksheets(CP_PLAN_LEGS).Cells(legRow, 5).Value2)
                needRelease = requiredQty - available
                totalNeedRelease = totalNeedRelease + needRelease
                needLegs = needLegs + 1
                If needLegs = 1 Then firstAvailable = available
            End If
        Next legRow
        wsPlan.Cells(r, 17).Value2 = totalNeedRelease
        wsPlan.Cells(r, 18).Value2 = firstAvailable
        ' WH_ACTUAL is only a preview/debug source.  It must never be used
        ' to final-select a holder before the exact MON reader has produced
        ' PLAN_RELEASE evidence.
        CP_UpdatePlanDecision wsPlan, r, "NEED_RELEASE", "NEED_RELEASE_REQUEST", _
                              "Available is below Required; exact active holders must be read from MON before release can be finalized.", _
                              CP_Text(wsPlan.Cells(r, 25).Value2)
NextRelease:
    Next r
End Sub

Private Function CP_BuildPostingIndex(ByVal ws As Worksheet) As Object
    Dim d As Object, lastRow As Long, r As Long, keyText As String
    Dim available As Double, totalStock As Double, current As Variant
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    lastRow = CP_LastRow(ws, 1, 2)
    For r = 2 To lastRow
        keyText = CP_SourceKey(CP_Text(ws.Cells(r, 3).Value2), CP_Text(ws.Cells(r, 6).Value2), CP_Text(ws.Cells(r, 5).Value2))
        If keyText <> "||" And UCase$(CP_Text(ws.Cells(r, 7).Value2)) = "FT" Then
            available = CP_Number(ws.Cells(r, 4).Value2)
            totalStock = CP_Number(ws.Cells(r, 8).Value2)
            If d.Exists(keyText) Then
                current = d(keyText)
                d(keyText) = Array(CDbl(current(0)) + available, CDbl(current(1)) + totalStock)
            Else
                d.Add keyText, Array(available, totalStock)
            End If
        End If
    Next r
    Set CP_BuildPostingIndex = d
End Function

Private Function CP_BuildBlockedPostingIndex(ByVal ws As Worksheet) As Object
    Dim d As Object, lastRow As Long, r As Long, keyText As String, stockType As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    lastRow = CP_LastRow(ws, 1, 2)
    For r = 2 To lastRow
        keyText = CP_SourceKey(CP_Text(ws.Cells(r, 3).Value2), CP_Text(ws.Cells(r, 6).Value2), CP_Text(ws.Cells(r, 5).Value2))
        stockType = UCase$(CP_Text(ws.Cells(r, 7).Value2))
        If keyText <> "||" And stockType <> "" And stockType <> "FT" Then d(keyText) = True
    Next r
    Set CP_BuildBlockedPostingIndex = d
End Function

Private Function CP_BuildLegs(ByVal postIndex As Object, ByVal material As String, _
                              ByVal targetBin As String, ByVal targetBatch As String, _
                              ByVal targetQty As Double) As Collection
    Dim result As New Collection, keyText As Variant, matchKey As String, matches As Long
    If targetQty <= CP_EPS Then Set CP_BuildLegs = result: Exit Function
    If targetBatch <> "" Then
        result.Add Array(CP_SourceKey(material, targetBin, targetBatch), targetQty)
        Set CP_BuildLegs = result
        Exit Function
    End If
    If postIndex Is Nothing Then Exit Function
    For Each keyText In postIndex.keys
        If StrComp(CP_KeyPart(CStr(keyText), 1), material, vbTextCompare) = 0 And _
           StrComp(CP_KeyPart(CStr(keyText), 2), targetBin, vbTextCompare) = 0 And _
           CP_KeyPart(CStr(keyText), 3) <> "" Then
            matches = matches + 1
            matchKey = CStr(keyText)
        End If
    Next keyText
    ' A blank Actual Batch may be inferred only from exactly one usable FT
    ' Posting batch at the user-selected bin. Multiple batches remain a
    ' user decision; no FIFO or quantity-based batch guessing is allowed.
    If matches <> 1 Then Exit Function
    result.Add Array(matchKey, targetQty)
    Set CP_BuildLegs = result
End Function

' Adds only the reservation of the exact WT being reworked.  POSTING_CACHE
' already exposes physical available and total stock; its available quantity
' can be lower because this active WT holds stock.  The result is capped by
' the same source's total stock, so stale or duplicate data cannot fabricate
' capacity.
Private Function CP_ProjectAvailableAfterOwnCancel(ByVal available As Double, ByVal totalStock As Double, _
                                                    ByVal selfHeldQty As Double) As Double
    Dim projected As Double
    projected = available + selfHeldQty
    If totalStock > 0# And projected > totalStock Then projected = totalStock
    CP_ProjectAvailableAfterOwnCancel = projected
End Function

Private Function CP_SelfHeldQtyForPlanSource(ByVal wsWH As Worksheet, ByVal wsPlan As Worksheet, _
                                              ByVal planRow As Long, ByVal sourceKey As String) As Double
    Dim whRow As Long, heldKey As String
    Dim obState As String, wtState As String, rawStatus As String
    On Error GoTo ExitHere
    whRow = CLng(CP_Number(wsPlan.Cells(planRow, 27).Value2))
    If whRow < CP_FIRST_WH_ROW Then Exit Function
    obState = UCase$(CP_Text(wsWH.Cells(whRow, 25).Value2))
    wtState = UCase$(CP_Text(wsWH.Cells(whRow, 26).Value2))
    rawStatus = UCase$(CP_Text(wsWH.Cells(whRow, 10).Value2))
    heldKey = CP_SourceKey(CP_Text(wsWH.Cells(whRow, 6).Value2), _
                           CP_Text(wsWH.Cells(whRow, 8).Value2), _
                           CP_Text(wsWH.Cells(whRow, 9).Value2))
    If obState <> "OB_OPEN_OR_PARTIAL" Or wtState <> "WT_ACTIVE" Or rawStatus <> "" Then Exit Function
    If StrComp(heldKey, sourceKey, vbTextCompare) <> 0 Then Exit Function
    CP_SelfHeldQtyForPlanSource = CP_Number(wsWH.Cells(whRow, 7).Value2)
ExitHere:
End Function

Private Function CP_SelectReleaseCandidates(ByVal ws As Worksheet, ByVal material As String, _
                                            ByVal binText As String, ByVal batchText As String, _
                                            ByVal currentOB As String, ByVal currentWT As String, _
                                            ByVal needRelease As Double) As Collection
    Dim allCandidates As New Collection, bestOne As Variant, x As Variant
    Dim lastRow As Long, r As Long, qty As Double, bestQty As Double
    Dim stateText As String, obState As String, statusText As String
    Dim selected As New Collection, sumQty As Double
    lastRow = CP_LastRow(ws, 1, CP_FIRST_WH_ROW)
    For r = CP_FIRST_WH_ROW To lastRow
        obState = UCase$(CP_Text(ws.Cells(r, 25).Value2))
        stateText = UCase$(CP_Text(ws.Cells(r, 26).Value2))
        statusText = UCase$(CP_Text(ws.Cells(r, 10).Value2))
        qty = CP_Number(ws.Cells(r, 7).Value2)
        If obState = "OB_OPEN_OR_PARTIAL" And stateText = "WT_ACTIVE" And statusText = "" _
           And CP_Text(ws.Cells(r, 1).Value2) <> currentOB And CP_Text(ws.Cells(r, 5).Value2) <> currentWT _
           And StrComp(CP_Text(ws.Cells(r, 6).Value2), material, vbTextCompare) = 0 _
           And StrComp(CP_Text(ws.Cells(r, 8).Value2), binText, vbTextCompare) = 0 _
           And StrComp(CP_Text(ws.Cells(r, 9).Value2), batchText, vbTextCompare) = 0 _
           And qty > CP_EPS Then
            allCandidates.Add Array(CP_Text(ws.Cells(r, 1).Value2), CP_Text(ws.Cells(r, 5).Value2), qty)
            If qty + CP_EPS >= needRelease Then
                If bestQty = 0# Or qty < bestQty Then bestQty = qty: bestOne = Array(CP_Text(ws.Cells(r, 1).Value2), CP_Text(ws.Cells(r, 5).Value2), qty)
            End If
        End If
    Next r
    If bestQty > 0# Then
        selected.Add bestOne
        Set CP_SelectReleaseCandidates = selected
        Exit Function
    End If
    CP_SortCandidates allCandidates
    For Each x In allCandidates
        selected.Add x
        sumQty = sumQty + CDbl(x(2))
        If sumQty + CP_EPS >= needRelease Then Exit For
    Next x
    Set CP_SelectReleaseCandidates = selected
End Function

Private Sub CP_SortCandidates(ByRef items As Collection)
    Dim i As Long, j As Long, a As Variant, b As Variant, sorted As New Collection
    If items.count <= 1 Then Exit Sub
    Do While items.count > 0
        j = 1
        For i = 2 To items.count
            If CDbl(items(i)(2)) < CDbl(items(j)(2)) Then j = i
        Next i
        sorted.Add items(j)
        items.Remove j
    Loop
    For Each a In sorted: items.Add a: Next a
End Sub

Private Function CP_WriteReleaseCandidates(ByVal planID As String, ByVal candidates As Collection, _
                                           ByVal material As String, ByVal binText As String, _
                                           ByVal batchText As String, ByVal needRelease As Double, _
                                           ByVal available As Double) As Double
    Dim ws As Worksheet, r As Long, x As Variant, runID As String, selectOrder As Long
    Set ws = CP_RequireSheet(CP_PLAN_RELEASE, True)
    r = CP_LastRow(ws, 1, 2) + 1
    runID = CP_Text(ThisWorkbook.Worksheets(CP_WH).Range("W2").Value2)
    For Each x In candidates
        selectOrder = selectOrder + 1
        ws.Cells(r, 1).Value2 = planID
        ws.Cells(r, 2).Value2 = 1
        ws.Cells(r, 3).Value2 = CP_SourceKey(material, binText, batchText)
        ws.Cells(r, 4).Value2 = x(0)
        ws.Cells(r, 6).Value2 = x(1)
        ws.Cells(r, 7).Value2 = material
        ws.Cells(r, 8).Value2 = binText
        ws.Cells(r, 9).Value2 = batchText
        ws.Cells(r, 10).Value2 = x(2)
        ws.Cells(r, 11).Value2 = "WT_ACTIVE"
        ws.Cells(r, 12).Value2 = "TRUE"
        ws.Cells(r, 13).Value2 = "TRUE"
        ws.Cells(r, 14).Value2 = selectOrder
        ws.Cells(r, 15).Value2 = needRelease
        ws.Cells(r, 16).Value2 = available + CDbl(x(2))
        ws.Cells(r, 17).Value2 = "LOCAL_SELECTED"
        ws.Cells(r, 18).Value2 = "Local WH_ACTUAL candidate; MON holder discovery not loaded."
        ws.Cells(r, 19).Value2 = runID
        ws.Cells(r, 20).Value2 = "WH_ACTUAL"
        ws.Cells(r, 21).Value2 = available
        CP_WriteReleaseCandidates = CP_WriteReleaseCandidates + CDbl(x(2))
        r = r + 1
    Next x
End Function

Private Function CP_ReleaseText(ByVal candidates As Collection) As String
    Dim x As Variant
    For Each x In candidates
        If CP_ReleaseText <> "" Then CP_ReleaseText = CP_ReleaseText & ", "
        CP_ReleaseText = CP_ReleaseText & "OB=" & x(0) & "/WT=" & x(1) & "/Qty=" & x(2)
    Next x
End Function

Private Function CP_PlanFirstAvailable(ByVal planID As String) As Double
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = CP_RequireSheet(CP_PLAN_LEGS, True)
    lastRow = CP_LastRow(ws, 1, 2)
    For r = 2 To lastRow
        If CP_Text(ws.Cells(r, 1).Value2) = planID Then
            CP_PlanFirstAvailable = CP_Number(ws.Cells(r, 5).Value2)
            Exit Function
        End If
    Next r
End Function

Private Function CP_PlanFirstTotal(ByVal planID As String) As Double
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = CP_RequireSheet(CP_PLAN_LEGS, True)
    lastRow = CP_LastRow(ws, 1, 2)
    For r = 2 To lastRow
        If CP_Text(ws.Cells(r, 1).Value2) = planID Then
            CP_PlanFirstTotal = CP_Number(ws.Cells(r, 6).Value2)
            Exit Function
        End If
    Next r
End Function

Private Function CP_KeyPart(ByVal keyText As String, ByVal partNo As Long) As String
    Dim parts As Variant
    parts = Split(keyText, "|")
    If partNo >= 1 And partNo <= UBound(parts) + 1 Then CP_KeyPart = CStr(parts(partNo - 1))
End Function

' Refresh runtime group keys after exception resolution.  A blank Actual Batch
' may be filled from exactly one Posting source, so keys and fingerprints must
' be rebuilt before the plan is locked.
Private Sub CP_RefreshPlanGroupKeys(ByVal wsPlan As Worksheet)
    Dim r As Long, lastRow As Long, material As String, targetBin As String, targetBatch As String
    On Error GoTo ExitHere
    lastRow = CP_LastRow(wsPlan, 1, 2)
    For r = 2 To lastRow
        material = CP_Text(wsPlan.Cells(r, 6).Value2)
        If material = "" Then GoTo NextRow
        targetBin = CP_Text(wsPlan.Cells(r, 12).Value2)
        targetBatch = CP_Text(wsPlan.Cells(r, 14).Value2)
        wsPlan.Cells(r, CP_COL_DOC_GROUP).Value2 = "DOC|" & CP_Text(wsPlan.Cells(r, 3).Value2)
        wsPlan.Cells(r, CP_COL_LOCATION_GROUP).Value2 = "POSTSRC|" & material & "|" & targetBin & "|" & targetBatch
        wsPlan.Cells(r, CP_COL_CREATE_GROUP).Value2 = "CREATE|" & CP_Text(wsPlan.Cells(r, 3).Value2) & _
            "|" & CP_Text(wsPlan.Cells(r, 4).Value2) & _
            "|MAT=" & material & "|SRC_BIN=" & CP_Text(wsPlan.Cells(r, 12).Value2) & _
            "|SRC_BATCH=" & CP_Text(wsPlan.Cells(r, 13).Value2) & _
            "|DST_BIN=" & CP_Text(wsPlan.Cells(r, 10).Value2) & _
            "|DST_BATCH=" & CP_Text(wsPlan.Cells(r, 14).Value2)
        wsPlan.Cells(r, CP_COL_PLAN_FINGERPRINT).Value2 = CP_PlanFingerprint(wsPlan, r)
NextRow:
    Next r
ExitHere:
End Sub

' Write one aggregate quantity contract for each OB/Document/Item/Material
' group.  Untouched members remain in both totals; CANCEL_TO_ZERO contributes a
' target of zero.  Button #4 consumes these values and never recomputes them.
Private Sub CP_BuildQtyGroupContract(ByVal wsPlan As Worksheet)
    Dim groups As Object, lastRow As Long, r As Long, keyText As String
    Dim material As String, g As Object, decisionText As String
    Dim currentQty As Double, targetQty As Double
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    lastRow = CP_LastRow(wsPlan, 1, 2)
    For r = 2 To lastRow
        material = CP_Text(wsPlan.Cells(r, 6).Value2)
        If material = "" Then GoTo NextRow
        keyText = "QTY|" & CP_Text(wsPlan.Cells(r, 2).Value2) & "|" & _
            CP_Text(wsPlan.Cells(r, 3).Value2) & "|" & CP_Text(wsPlan.Cells(r, 4).Value2) & "|" & material
        If Not groups.Exists(keyText) Then
            Set g = CreateObject("Scripting.Dictionary")
            g.CompareMode = vbTextCompare
            g.Add "Current", 0#
            g.Add "Target", 0#
            g.Add "Members", 0&
            groups.Add keyText, g
        End If
        Set g = groups.item(keyText)
        currentQty = CP_Number(wsPlan.Cells(r, 7).Value2)
        targetQty = CP_Number(wsPlan.Cells(r, 8).Value2)
        decisionText = UCase$(CP_Text(wsPlan.Cells(r, 20).Value2))
        If decisionText = "CANCEL_TO_ZERO" Then targetQty = 0#
        g("Current") = CDbl(g("Current")) + currentQty
        g("Target") = CDbl(g("Target")) + targetQty
        g("Members") = CLng(g("Members")) + 1
NextRow:
    Next r

    For r = 2 To lastRow
        material = CP_Text(wsPlan.Cells(r, 6).Value2)
        If material = "" Then GoTo NextWrite
        keyText = "QTY|" & CP_Text(wsPlan.Cells(r, 2).Value2) & "|" & _
            CP_Text(wsPlan.Cells(r, 3).Value2) & "|" & CP_Text(wsPlan.Cells(r, 4).Value2) & "|" & material
        Set g = groups.item(keyText)
        wsPlan.Cells(r, CP_COL_QTY_GROUP_ID).Value2 = keyText
        wsPlan.Cells(r, CP_COL_GROUP_CURRENT).Value2 = CDbl(g("Current"))
        wsPlan.Cells(r, CP_COL_GROUP_TARGET).Value2 = CDbl(g("Target"))
        wsPlan.Cells(r, CP_COL_GROUP_DELTA).Value2 = CDbl(g("Target")) - CDbl(g("Current"))
        wsPlan.Cells(r, CP_COL_GROUP_MEMBERS).Value2 = CLng(g("Members"))
        wsPlan.Cells(r, CP_COL_GROUP_NEED_QTY).Value2 = IIf(Abs(CDbl(g("Target")) - CDbl(g("Current"))) > CP_EPS, "TRUE", "FALSE")
        wsPlan.Cells(r, CP_COL_PLAN_FINGERPRINT).Value2 = CP_PlanFingerprint(wsPlan, r)
NextWrite:
    Next r
End Sub

Private Sub CP_WritePlanRow(ByVal wsPlan As Worksheet, ByVal outRow As Long, ByVal wsWH As Worksheet, _
                            ByVal whRow As Long, ByVal targetQty As Double, ByVal targetBin As String, _
                            ByVal targetBatch As String, ByVal decision As String, ByVal gateText As String, _
                            ByVal reasonText As String, ByVal pipeline As String)
    Dim sysQty As Double, sysBin As String, sysBatch As String
    sysQty = CP_Number(wsWH.Cells(whRow, 7).Value2)
    sysBin = CP_Text(wsWH.Cells(whRow, 8).Value2)
    sysBatch = CP_Text(wsWH.Cells(whRow, 9).Value2)
    wsPlan.Cells(outRow, 1).Value2 = "P" & Format$(outRow - 1, "0000")
    wsPlan.Cells(outRow, 2).Value2 = CP_Text(wsWH.Cells(whRow, 1).Value2)
    wsPlan.Cells(outRow, 3).Value2 = CP_Text(wsWH.Cells(whRow, 2).Value2)
    wsPlan.Cells(outRow, 4).Value2 = CP_Text(wsWH.Cells(whRow, 4).Value2)
    wsPlan.Cells(outRow, 5).Value2 = CP_Text(wsWH.Cells(whRow, 5).Value2)
    wsPlan.Cells(outRow, 6).Value2 = CP_Text(wsWH.Cells(whRow, 6).Value2)
    wsPlan.Cells(outRow, 7).Value2 = sysQty
    wsPlan.Cells(outRow, 8).Value2 = targetQty
    wsPlan.Cells(outRow, 9).Value2 = sysBin
    wsPlan.Cells(outRow, 10).Value2 = targetBin
    wsPlan.Cells(outRow, 11).Value2 = sysBatch
    wsPlan.Cells(outRow, 12).Value2 = targetBin
    wsPlan.Cells(outRow, 13).Value2 = sysBatch
    wsPlan.Cells(outRow, 14).Value2 = targetBatch
    wsPlan.Cells(outRow, 20).Value2 = decision
    wsPlan.Cells(outRow, 24).Value2 = gateText
    wsPlan.Cells(outRow, 25).Value2 = pipeline
    wsPlan.Cells(outRow, 26).Value2 = reasonText
    wsPlan.Cells(outRow, 21).Value2 = IIf(Abs(targetQty - sysQty) > CP_EPS, "TRUE", "FALSE")
    wsPlan.Cells(outRow, 22).Value2 = IIf(StrComp(targetBin, sysBin, vbTextCompare) <> 0, "TRUE", "FALSE")
    wsPlan.Cells(outRow, 23).Value2 = IIf(StrComp(targetBatch, sysBatch, vbTextCompare) <> 0, "TRUE", "FALSE")
    wsPlan.Cells(outRow, 27).Value2 = whRow
    wsPlan.Cells(outRow, CP_COL_QTY_MODE).Value2 = CP_DeriveQtyMode(targetQty, sysQty, targetBin, sysBin, targetBatch, sysBatch)
    wsPlan.Cells(outRow, CP_COL_DOC_GROUP).Value2 = "DOC|" & CP_Text(wsWH.Cells(whRow, 2).Value2)
    wsPlan.Cells(outRow, CP_COL_LOCATION_GROUP).Value2 = "POSTSRC|" & CP_Text(wsWH.Cells(whRow, 6).Value2) & _
        "|" & targetBin & "|" & targetBatch
    wsPlan.Cells(outRow, CP_COL_CREATE_GROUP).Value2 = "CREATE|" & CP_Text(wsWH.Cells(whRow, 2).Value2) & _
        "|" & CP_Text(wsWH.Cells(whRow, 3).Value2) & _
        "|MAT=" & CP_Text(wsWH.Cells(whRow, 6).Value2) & "|SRC_BIN=" & sysBin & _
        "|SRC_BATCH=" & sysBatch & "|DST_BIN=" & targetBin & "|DST_BATCH=" & targetBatch
    wsPlan.Cells(outRow, CP_COL_PLAN_FINGERPRINT).Value2 = CP_PlanFingerprint(wsPlan, outRow)
    wsPlan.Cells(outRow, CP_COL_QTY_DELTA).Value2 = targetQty - sysQty
End Sub

Private Sub CP_UpdatePlanDecision(ByVal wsPlan As Worksheet, ByVal r As Long, ByVal decision As String, _
                                  ByVal gateText As String, ByVal reasonText As String, ByVal pipeline As String)
    Dim whRow As Long
    wsPlan.Cells(r, 20).Value2 = decision
    wsPlan.Cells(r, 24).Value2 = gateText
    wsPlan.Cells(r, 25).Value2 = pipeline
    wsPlan.Cells(r, 26).Value2 = reasonText
    CP_WriteExecutionContract wsPlan, r, decision
    whRow = CLng(wsPlan.Cells(r, 27).Value2)
    CP_WriteWHDecision ThisWorkbook.Worksheets(CP_WH), whRow, decision, reasonText, pipeline
    If whRow >= CP_FIRST_WH_ROW Then ThisWorkbook.Worksheets(CP_WH).Cells(whRow, 16).Value2 = gateText
End Sub

Private Sub CP_WriteWHDecision(ByVal ws As Worksheet, ByVal r As Long, ByVal decision As String, _
                               ByVal reasonText As String, ByVal pipeline As String)
    If r < CP_FIRST_WH_ROW Then Exit Sub
    ws.Cells(r, 15).Value2 = decision
    ws.Cells(r, 16).Value2 = "PRECHECK"
    ws.Cells(r, 17).Value2 = "CHECKED"
    ws.Cells(r, 18).Value2 = pipeline
    ws.Cells(r, 19).Value2 = reasonText
    ws.Cells(r, 29).Value2 = decision
    ws.Cells(r, 30).Value2 = reasonText
End Sub

Private Sub CP_MarkPostingUnavailable(ByVal wsPlan As Worksheet, ByVal reasonText As String)
    Dim lastRow As Long, r As Long, decisionText As String, gateText As String
    decisionText = "STOP_POST_SESSION_REQUIRED"
    gateText = "POST_SESSION"
    If InStr(1, reasonText, "STOP_STALE_POSTING_CACHE", vbTextCompare) > 0 Then
        decisionText = "STOP_STALE_POSTING_CACHE"
        gateText = "POSTING_CACHE"
    ElseIf InStr(1, reasonText, "STOP_POSTING_NO_STOCK", vbTextCompare) > 0 Then
        decisionText = "STOP_POSTING_NO_STOCK"
        gateText = "POSTING_CACHE"
    ElseIf InStr(1, reasonText, "STOP_POST_CAPTURE", vbTextCompare) > 0 Or _
           InStr(1, reasonText, "STOP_POST_EXPECTED", vbTextCompare) > 0 Then
        decisionText = "STOP_POST_CAPTURE_INCOMPLETE"
        gateText = "POST_CAPTURE"
    End If
    lastRow = CP_LastRow(wsPlan, 1, 2)
    For r = 2 To lastRow
        If UCase$(CP_Text(wsPlan.Cells(r, 20).Value2)) = "POSTING_REQUIRED" Or _
           UCase$(CP_Text(wsPlan.Cells(r, 20).Value2)) = "CHANGE_QTY" Then
            CP_UpdatePlanDecision wsPlan, r, decisionText, gateText, reasonText, _
                                  CP_Text(wsPlan.Cells(r, 25).Value2)
        End If
    Next r
End Sub

Private Sub CP_WriteLeg(ByVal planID As String, ByVal lineNo As Long, ByVal sourceKey As String, _
                        ByVal requiredQty As Double, ByVal available As Double, ByVal totalStock As Double)
    Dim ws As Worksheet, r As Long
    Set ws = CP_RequireSheet(CP_PLAN_LEGS, True)
    r = CP_LastRow(ws, 1, 2) + 1
    ws.Cells(r, 1).Value2 = planID
    ws.Cells(r, 2).Value2 = lineNo
    ws.Cells(r, 3).Value2 = sourceKey
    ws.Cells(r, 4).Value2 = requiredQty
    ws.Cells(r, 5).Value2 = available
    ws.Cells(r, 6).Value2 = totalStock
    ws.Cells(r, 7).Value2 = IIf(totalStock + CP_EPS < requiredQty, "STOP_SOURCE_INSUFFICIENT", _
                            IIf(available + CP_EPS < requiredQty, "NEED_RELEASE", "SOURCE_READY"))
End Sub

' CHECK owns this contract.  Button #4 may only consume READY_TO_EXECUTE;
' it must never reinterpret a business decision while SAP writes are in flight.
Private Sub CP_WriteExecutionContract(ByVal wsPlan As Worksheet, ByVal r As Long, ByVal decision As String)
    Dim qtyMode As String, execMode As String, execState As String
    qtyMode = UCase$(CP_Text(wsPlan.Cells(r, CP_COL_QTY_MODE).Value2))
    decision = UCase$(Trim$(decision))

    Select Case decision
        Case "CF_ONLY"
            execMode = "NONE"
            execState = "READY_FOR_CF"
        Case "DONE"
            execMode = "NONE"
            execState = "DONE_BY_USER_CF"
        Case "USER_SKIP"
            execMode = "NONE"
            execState = "USER_SKIP"
        Case "CANCEL_TO_ZERO"
            execMode = "CANCEL_PRDO_ZERO"
            execState = "READY_TO_EXECUTE"
        Case "NEED_RELEASE"
            execMode = "NONE"
            execState = "BLOCKED_RELEASE_REQUIRED"
        Case "MULTI_SOURCE"
            execMode = "NONE"
            execState = "BLOCKED_MULTI_SOURCE"
        Case "CHANGE_QTY"
            If qtyMode = "DELTA_ADD" Then
                ' Delta-add is a separate production route: retain the exact
                ' existing WT, raise the PRDO total, then create only the
                ' incremental open quantity.  It must never fall through to
                ' the legacy replacement route, which cancels the old WT.
                execMode = "KEEP_EXISTING_CREATE_DELTA"
                execState = "READY_TO_EXECUTE"
            Else
                execMode = "REPLACE_CANCEL_CREATE"
                execState = "READY_TO_EXECUTE"
            End If
        Case "CHANGE_BIN", "CHANGE_BATCH", "CHANGE_BIN_BATCH"
            execMode = "REPLACE_CANCEL_CREATE"
            execState = "READY_TO_EXECUTE"
        Case Else
            execMode = "NONE"
            If Left$(decision, 5) = "STOP_" Then
                execState = "BLOCKED_" & Mid$(decision, 6)
            ElseIf decision = "POSTING_REQUIRED" Then
                execState = "BLOCKED_SOURCE_UNRESOLVED"
            Else
                execState = "BLOCKED_PLAN_DECISION"
            End If
    End Select

    wsPlan.Cells(r, CP_COL_EXEC_MODE).Value2 = execMode
    wsPlan.Cells(r, CP_COL_EXEC_STATE).Value2 = execState
    If execState = "READY_TO_EXECUTE" Then
        wsPlan.Cells(r, 24).Value2 = "READY_TO_EXECUTE"
    ElseIf Left$(execState, 8) = "BLOCKED_" Then
        wsPlan.Cells(r, 24).Value2 = "BLOCKED"
    Else
        wsPlan.Cells(r, 24).Value2 = execState
    End If
End Sub

Private Function CP_DeriveQtyMode(ByVal targetQty As Double, ByVal sysQty As Double, _
                                  ByVal targetBin As String, ByVal sysBin As String, _
                                  ByVal targetBatch As String, ByVal sysBatch As String) As String
    If targetQty > sysQty + CP_EPS And _
       StrComp(targetBin, sysBin, vbTextCompare) = 0 And _
       StrComp(targetBatch, sysBatch, vbTextCompare) = 0 Then
        CP_DeriveQtyMode = "DELTA_ADD"
    ElseIf Abs(targetQty - sysQty) > CP_EPS Then
        CP_DeriveQtyMode = "FINAL_TOTAL"
    Else
        CP_DeriveQtyMode = "NO_QTY_CHANGE"
    End If
End Function

Private Function CP_SourceRequiredQty(ByVal wsPlan As Worksheet, ByVal r As Long) As Double
    Dim targetQty As Double, sysQty As Double
    targetQty = CP_Number(wsPlan.Cells(r, 8).Value2)
    sysQty = CP_Number(wsPlan.Cells(r, 7).Value2)
    If UCase$(CP_Text(wsPlan.Cells(r, CP_COL_QTY_MODE).Value2)) = "DELTA_ADD" Then
        CP_SourceRequiredQty = targetQty - sysQty
    Else
        CP_SourceRequiredQty = targetQty
    End If
End Function

Private Function CP_PlanFingerprint(ByVal wsPlan As Worksheet, ByVal r As Long) As String
    CP_PlanFingerprint = "OB=" & CP_Text(wsPlan.Cells(r, 2).Value2) & _
        "|DOC=" & CP_Text(wsPlan.Cells(r, 3).Value2) & _
        "|ITEM=" & CP_Text(wsPlan.Cells(r, 4).Value2) & _
        "|WT=" & CP_Text(wsPlan.Cells(r, 5).Value2) & _
        "|SYSQ=" & CP_Text(wsPlan.Cells(r, 7).Value2) & _
        "|TARGETQ=" & CP_Text(wsPlan.Cells(r, 8).Value2) & _
        "|SYSBIN=" & CP_Text(wsPlan.Cells(r, 9).Value2) & _
        "|TARGETBIN=" & CP_Text(wsPlan.Cells(r, 12).Value2) & _
        "|SYSBATCH=" & CP_Text(wsPlan.Cells(r, 11).Value2) & _
        "|TARGETBATCH=" & CP_Text(wsPlan.Cells(r, 14).Value2) & _
        "|QTY_GROUP=" & CP_Text(wsPlan.Cells(r, CP_COL_QTY_GROUP_ID).Value2) & _
        "|EXEC_MODE=" & CP_Text(wsPlan.Cells(r, CP_COL_EXEC_MODE).Value2)
End Function

Private Sub CP_ResetOutputs(ByVal wsPlan As Worksheet)
    wsPlan.Cells.ClearContents
    wsPlan.Range("A1:AZ1").Value = Array("PlanID", "OB", "Document", "WH", "WT", "Material", _
        "Sys Qty", "Target Qty", "Sys Bin", "Target Bin", "Sys Batch", "Source Bin", "Source Batch", "Target Batch", _
        "Available", "Total", "Need Release", "Projected Available", "Source Legs", "Decision", _
        "NeedQtyChange", "NeedBinChange", "NeedBatchChange", "Gate", "Action Pipeline", "Reason", "WH Actual Row", _
        "QtyMode", "ExecMode", "ExecState", "DocumentGroupId", "SourceGroupId", "CreateGroupId", "PlanFingerprint", "QtyDelta", _
        "", "", "", "", "", "", "", "", "", "", "", "QTY_GROUP_ID", "GROUP_CURRENT_TOTAL", _
        "GROUP_TARGET_TOTAL", "GROUP_DELTA_TOTAL", "GROUP_MEMBER_COUNT", "GROUP_NEED_QTY")
    wsPlan.rows(1).Font.Bold = True
    wsPlan.columns("A:AZ").NumberFormat = "@"
    wsPlan.columns("G:H").NumberFormat = "#,##0.###"
    wsPlan.columns("O:S").NumberFormat = "#,##0.###"
    CP_ResetLegs
    CP_ResetRelease
End Sub

Private Sub CP_ResetLegs()
    Dim ws As Worksheet
    Set ws = CP_RequireSheet(CP_PLAN_LEGS, True)
    ws.Cells.ClearContents
    ws.Range("A1:G1").Value = Array("PlanID", "Leg", "SourceKey", "Required Qty", "Available", "Total", "Gate")
    ws.rows(1).Font.Bold = True
End Sub

Private Sub CP_ResetRelease()
    Dim ws As Worksheet
    Set ws = CP_RequireSheet(CP_PLAN_RELEASE, True)
    ws.Cells.ClearContents
    ws.Range("A1:U1").Value = Array("PlanID", "Leg", "SourceKey", "Holder OB", "Holder WH", "Holder WT", _
        "Material", "Bin", "Batch", "Held Qty", "WT Status", "Eligible", "Selected", "Select Order", _
        "Need Release", "Projected Available", "Gate", "Reason", "Query Run ID", "Query Source", "Posting Available")
    ws.rows(1).Font.Bold = True
End Sub

Private Sub CP_WriteSummary(ByVal wsPlan As Worksheet, ByVal fastRows As Long, ByVal checkedRows As Long, _
                             ByVal exceptionMaterials As Long, ByVal started As Double, _
                             ByVal checkResult As String, ByVal readyRows As Long, ByVal stopRows As Long)
    wsPlan.Range("AJ1").Value2 = "PLAN LOCK"
    wsPlan.Range("AJ2").Value2 = "LOCKED"
    wsPlan.Range("AK1").Value2 = "PLAN RUN ID"
    wsPlan.Range("AK2").Value2 = CP_Text(ThisWorkbook.Worksheets(CP_WH).Range("W2").Value2)
    wsPlan.Range("AL1").Value2 = "CHECK STATUS"
    wsPlan.Range("AL2").Value2 = checkResult
    wsPlan.Range("AL3").Value2 = "Rows=" & CStr(checkedRows) & "; Fast=" & CStr(fastRows) & _
                                  "; Posting materials=" & CStr(exceptionMaterials) & _
                                  "; Ready=" & CStr(readyRows) & "; Stop=" & CStr(stopRows)
    wsPlan.Range("AL4").Value2 = "SAP_WRITE=0; CANCEL=0; SAVE=0; CREATE=0; CF=0"
    wsPlan.Range("AL5").Value2 = "Elapsed=" & Format$(RT_Elapsed(started), "0.000") & " s"
    wsPlan.Range("AN1").Value2 = "PLAN HASH"
    wsPlan.Range("AN2").Value2 = CP_PlanContractHash(wsPlan)
    wsPlan.Range("AO1").Value2 = "WH ACTUAL HASH"
    wsPlan.Range("AO2").Value2 = CP_WHActualContractHash(wsPlan)
    wsPlan.Range("AP1").Value2 = "WH ACTUAL TIMESTAMP"
    wsPlan.Range("AP2").Value2 = Now
    wsPlan.Range("AQ1").Value2 = "POSTING RUN ID"
    wsPlan.Range("AQ2").Value2 = CP_PostingRunID()
    wsPlan.Range("AR1").Value2 = "MON RUN ID"
    wsPlan.Range("AR2").Value2 = CP_Text(ThisWorkbook.Worksheets(CP_WH).Range("W2").Value2)
    wsPlan.Range("AS1").Value2 = "PLAN CREATED AT"
    wsPlan.Range("AS2").Value2 = Now
    wsPlan.Range("AT1").Value2 = "PLAN CONTRACT VERSION"
    wsPlan.Range("AT2").Value2 = "CHECK_LOCKED_V2"
    wsPlan.columns("A:AZ").AutoFit
End Sub

' The lock stores deterministic hashes of the immutable PLAN contract and the
' relevant WH_ACTUAL snapshot.  Button #4 recomputes both values before any
' SAP write, so a user edit after CHECK is a stale-plan stop rather than a
' reason for the executor to re-plan.
Private Function CP_PlanContractHash(ByVal wsPlan As Worksheet) As String
    Dim r As Long, c As Long, lastRow As Long, payload As String
    lastRow = CP_LastRow(wsPlan, 1, 2)
    For r = 2 To lastRow
        payload = payload & "|R" & CStr(r)
        For c = 1 To CP_COL_QTY_DELTA
            payload = payload & "|" & CP_Text(wsPlan.Cells(r, c).Value2)
        Next c
        For c = CP_COL_QTY_GROUP_ID To CP_COL_GROUP_NEED_QTY
            payload = payload & "|" & CP_Text(wsPlan.Cells(r, c).Value2)
        Next c
    Next r
    CP_PlanContractHash = CP_StableHash(payload)
End Function

Private Function CP_WHActualContractHash(ByVal wsPlan As Worksheet) As String
    Dim wsWH As Worksheet, r As Long, c As Long, whRow As Long, lastRow As Long, payload As String
    Set wsWH = ThisWorkbook.Worksheets(CP_WH)
    lastRow = CP_LastRow(wsPlan, 1, 2)
    For r = 2 To lastRow
        whRow = CLng(CP_Number(wsPlan.Cells(r, 27).Value2))
        payload = payload & "|P=" & CP_Text(wsPlan.Cells(r, 1).Value2) & "|W=" & CStr(whRow)
        ' Lock only the MON source snapshot and user-owned picking inputs.
        ' O:S and Y:AD are derived status/result fields written by CHECK and
        ' the executor.  Including them made a completed prior OB invalidate
        ' the same CHECK lock before the next OB could run.
        For c = 1 To 14
            payload = payload & "|" & CP_Text(wsWH.Cells(whRow, c).Value2)
        Next c
        For c = 20 To 22
            payload = payload & "|" & CP_Text(wsWH.Cells(whRow, c).Value2)
        Next c
    Next r
    CP_WHActualContractHash = CP_StableHash(payload)
End Function

Private Function CP_PostingRunID() As String
    On Error Resume Next
    CP_PostingRunID = CP_Text(ThisWorkbook.Worksheets("POSTING_SUMMARY").Cells(2, 6).Value2)
End Function

Private Function CP_StableHash(ByVal textIn As String) As String
    Const HASH_MOD As Double = 2147483629#
    Dim i As Long, codePoint As Long, valueOut As Double
    For i = 1 To Len(textIn)
        codePoint = AscW(Mid$(textIn, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536
        valueOut = valueOut * 131# + CDbl(codePoint)
        valueOut = valueOut - Fix(valueOut / HASH_MOD) * HASH_MOD
    Next i
    CP_StableHash = "H" & CStr(Fix(valueOut))
End Function

Private Function CP_RequireSheet(ByVal sheetName As String, ByVal veryHidden As Boolean) As Worksheet
    On Error Resume Next
    Set CP_RequireSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If CP_RequireSheet Is Nothing Then Err.Raise vbObjectError + 9812, , _
        "STOP_PLAN_SHEET_MISSING: " & sheetName & " must be pre-created in the template."
    If veryHidden Then CP_RequireSheet.Visible = xlSheetVeryHidden Else CP_RequireSheet.Visible = xlSheetVisible
End Function

Private Function CP_ValidatePostingCache(ByVal wsWH As Worksheet, ByRef messageText As String) As Boolean
    Dim wsCache As Worksheet, wsSummary As Worksheet, runID As String
    Dim lastRow As Long, summaryLast As Long, r As Long
    Dim cacheRunID As String, summaryRunID As String, readyText As String, snapshotResult As String
    Dim validRows As Long, summaryRows As Long
    Dim summaryExpected As Double, summaryCaptured As Double, summaryValidStockRows As Double
    Dim expectedAll As Double, capturedAll As Double, validAll As Double, haveCounts As Boolean
    On Error GoTo Failed
    runID = CP_Text(wsWH.Range("W2").Value2)
    Set wsCache = CP_RequireSheet(CP_POST_CACHE, True)
    Set wsSummary = CP_RequireSheet("POSTING_SUMMARY", True)
    summaryLast = CP_LastRow(wsSummary, 1, 2)
    For r = 2 To summaryLast
        If CP_Text(wsSummary.Cells(r, 1).Value2) <> "" Then
            summaryRows = summaryRows + 1
            summaryRunID = CP_Text(wsSummary.Cells(r, 6).Value2)
            readyText = UCase$(CP_Text(wsSummary.Cells(r, 9).Value2))
            If StrComp(summaryRunID, runID, vbTextCompare) <> 0 Or readyText <> "TRUE" Then
                messageText = "STOP_STALE_POSTING_CACHE: POSTING_SUMMARY is not DATA_READY for current WH RunID."
                Exit Function
            End If
            If CP_Text(wsSummary.Cells(r, 13).Value2) = "" Or _
               CP_Text(wsSummary.Cells(r, 14).Value2) = "" Or _
               CP_Text(wsSummary.Cells(r, 15).Value2) = "" Then
                messageText = "STOP_POST_CAPTURE_INCOMPLETE: POSTING_SUMMARY is missing independent capture counts."
                Exit Function
            End If
            summaryExpected = CP_Number(wsSummary.Cells(r, 13).Value2)
            summaryCaptured = CP_Number(wsSummary.Cells(r, 14).Value2)
            summaryValidStockRows = CP_Number(wsSummary.Cells(r, 15).Value2)
            If summaryExpected < 0# Or summaryCaptured < 0# Or summaryValidStockRows < 0# Or _
               Abs(summaryExpected - summaryCaptured) > CP_EPS Then
                messageText = "STOP_POST_CAPTURE_INCOMPLETE: expected/captured count is not exact for current POSTING cache."
                Exit Function
            End If
            If Not haveCounts Then
                expectedAll = summaryExpected
                capturedAll = summaryCaptured
                validAll = summaryValidStockRows
                haveCounts = True
            ElseIf Abs(expectedAll - summaryExpected) > CP_EPS Or _
                   Abs(capturedAll - summaryCaptured) > CP_EPS Or _
                   Abs(validAll - summaryValidStockRows) > CP_EPS Then
                messageText = "STOP_POST_CAPTURE_INCOMPLETE: POSTING_SUMMARY contains inconsistent global capture counts."
                Exit Function
            End If
        End If
    Next r
    If summaryRows = 0 Or Not haveCounts Then
        messageText = "STOP_STALE_POSTING_CACHE: POSTING_SUMMARY has no current material result."
        Exit Function
    End If
    lastRow = CP_LastRow(wsCache, 1, 2)
    For r = 2 To lastRow
        cacheRunID = CP_Text(wsCache.Cells(r, 1).Value2)
        snapshotResult = UCase$(CP_Text(wsCache.Cells(r, 19).Value2))
        If StrComp(cacheRunID, runID, vbTextCompare) <> 0 Then
            messageText = "STOP_STALE_POSTING_CACHE: cache LoadRunID does not equal WH RunID."
            Exit Function
        End If
        If snapshotResult <> "VALID_STOCK" Then
            messageText = "STOP_POST_CAPTURE_INCOMPLETE: POSTING_CACHE contains a non-stock raw row."
            Exit Function
        End If
        If UCase$(CP_Text(wsCache.Cells(r, 7).Value2)) <> "FT" Then
            messageText = "STOP_SOURCE_STOCK_TYPE_NOT_USABLE: POSTING_CACHE contains a non-FT stock row."
            Exit Function
        End If
        If CP_Text(wsCache.Cells(r, 3).Value2) = "" Then
            messageText = "STOP_POST_CAPTURE_INCOMPLETE: POSTING_CACHE has a stock row without Material."
            Exit Function
        End If
        validRows = validRows + 1
    Next r
    If validAll = 0# Then
        If validRows <> 0 Then
            messageText = "STOP_POST_CAPTURE_INCOMPLETE: POSTING_CACHE rows do not match zero valid-stock result."
            Exit Function
        End If
        messageText = "STOP_POSTING_NO_STOCK: full SAP capture returned no valid stock rows."
        Exit Function
    End If
    If validRows = 0 Then
        messageText = "STOP_STALE_POSTING_CACHE: POSTING_CACHE is empty after a nonzero valid-stock result."
        Exit Function
    End If
    If Abs(CDbl(validRows) - validAll) > CP_EPS Then
        messageText = "STOP_POST_CAPTURE_INCOMPLETE: POSTING_CACHE valid-stock row count does not match POSTING_SUMMARY."
        Exit Function
    End If
    CP_ValidatePostingCache = True
    Exit Function
Failed:
    messageText = "STOP_STALE_POSTING_CACHE: validation failed: " & Err.Description
End Function

Private Function CP_IsPostingRefreshRequired(ByVal messageText As String) As Boolean
    CP_IsPostingRefreshRequired = _
        InStr(1, messageText, "STOP_STALE_POSTING_CACHE", vbTextCompare) > 0 Or _
        InStr(1, messageText, "STOP_POST_SESSION_REQUIRED", vbTextCompare) > 0 Or _
        InStr(1, messageText, "STOP_POST_CAPTURE_INCOMPLETE", vbTextCompare) > 0
End Function

Private Sub CP_PlanOutcome(ByVal wsPlan As Worksheet, ByRef readyRows As Long, _
                           ByRef stopRows As Long, ByRef checkResult As String)
    Dim lastRow As Long, r As Long, decision As String, execState As String
    lastRow = CP_LastRow(wsPlan, 1, 2)
    checkResult = "PLAN_NOT_EXECUTABLE"
    For r = 2 To lastRow
        decision = UCase$(CP_Text(wsPlan.Cells(r, 20).Value2))
        execState = UCase$(CP_Text(wsPlan.Cells(r, CP_COL_EXEC_STATE).Value2))
        If execState = "READY_TO_EXECUTE" Then
            readyRows = readyRows + 1
        ElseIf Left$(execState, 8) = "BLOCKED_" Or Left$(decision, 5) = "STOP_" Then
            stopRows = stopRows + 1
        End If
    Next r
    If readyRows > 0 And stopRows = 0 Then
        checkResult = "PLAN_EXECUTABLE"
    ElseIf readyRows > 0 And stopRows > 0 Then
        checkResult = "PLAN_WITH_STOPS"
    ElseIf readyRows = 0 And stopRows = 0 Then
        checkResult = "PLAN_NO_SAP_WORK"
    End If
End Sub

' PLAN is rebuilt by every CHECK.  Copy its decision evidence to the
' append-only ledger before the next CHECK replaces the working snapshot.
Private Sub CP_RecordCheckHistory(ByVal wsPlan As Worksheet, ByVal checkRunID As String)
    Dim wsWH As Worksheet, r As Long, lastRow As Long, whRow As Long
    Dim itemText As String, execState As String, detailText As String
    On Error GoTo ExitHere
    Set wsWH = ThisWorkbook.Worksheets(CP_WH)
    lastRow = CP_LastRow(wsPlan, 1, 2)
    For r = 2 To lastRow
        whRow = CLng(CP_Number(wsPlan.Cells(r, 27).Value2))
        itemText = ""
        If whRow >= CP_FIRST_WH_ROW Then itemText = CP_Text(wsWH.Cells(whRow, 3).Value2)
        execState = CP_Text(wsPlan.Cells(r, CP_COL_EXEC_STATE).Value2)
        detailText = "DECISION=" & CP_Text(wsPlan.Cells(r, 20).Value2) & _
            "|GATE=" & CP_Text(wsPlan.Cells(r, 24).Value2) & _
            "|REASON=" & CP_Text(wsPlan.Cells(r, 26).Value2) & _
            "|SOURCE_LEGS=" & CP_Text(wsPlan.Cells(r, 19).Value2)
        CF_HistoryAppend checkRunID, "CHECK", "PLAN_CLASSIFIED", _
            CP_Text(wsPlan.Cells(r, 1).Value2), CP_Text(wsPlan.Cells(r, 2).Value2), _
            CP_Text(wsPlan.Cells(r, 3).Value2), itemText, CP_Text(wsPlan.Cells(r, 6).Value2), _
            CP_Text(wsPlan.Cells(r, 5).Value2), "CHECK_PLAN", CP_Text(wsPlan.Cells(r, 20).Value2), _
            CP_Text(wsPlan.Cells(r, 7).Value2), CP_Text(wsPlan.Cells(r, 8).Value2), _
            CP_Text(wsPlan.Cells(r, 15).Value2), execState, detailText
    Next r
ExitHere:
End Sub

Private Function CP_SourceKey(ByVal material As String, ByVal binText As String, ByVal batchText As String) As String
    CP_SourceKey = Trim$(material) & "|" & Trim$(binText) & "|" & Trim$(batchText)
End Function

Private Function CP_BinKey(ByVal material As String, ByVal binText As String) As String
    CP_BinKey = Trim$(material) & "|" & Trim$(binText)
End Function

Private Function CP_LastRow(ByVal ws As Worksheet, ByVal colNo As Long, ByVal firstRow As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.rows.count, colNo).End(xlUp).row
    If r < firstRow Then CP_LastRow = firstRow - 1 Else CP_LastRow = r
End Function

Private Function CP_Text(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    CP_Text = Trim$(CStr(valueIn))
End Function

Private Function CP_Number(ByVal valueIn As Variant) As Double
    If IsNumeric(valueIn) Then CP_Number = CDbl(valueIn) Else CP_Number = RT_Number(valueIn)
End Function

Private Sub CP_Assert(ByVal condition As Boolean, ByVal caseName As String)
    If Not condition Then Err.Raise vbObjectError + 9820, , "CHECK_PLANNER_PROBE_FAIL: " & caseName
End Sub
