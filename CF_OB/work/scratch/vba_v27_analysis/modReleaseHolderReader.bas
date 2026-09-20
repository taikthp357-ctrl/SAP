Attribute VB_Name = "modReleaseHolderReader"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub RH_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub RH_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

' V1.2.4 - read-only release holder acquisition and deterministic selector.
' This module never presses Cancel/Write/Save/Create/CF.  It only reads MON
' and writes candidate/selection evidence to the existing PLAN_RELEASE sheet.

Private Const RH_PLAN_RELEASE As String = "PLAN_RELEASE"
Private Const RH_WH As String = "WH_ACTUAL"
Private Const RH_PLAN As String = "PLAN"
Private Const RH_PLAN_LEGS As String = "PLAN_SOURCE_LEGS"
Private Const RH_FIRST_ROW As Long = 2
Private Const RH_EPS As Double = 0.000001

Public Sub CF_LoadReleaseHoldersReadOnly()
    Dim messageText As String
    If Not RH_LoadReleaseHolders(messageText) Then
        If Not gRT_Silent Then MsgBox messageText, vbExclamation, "RELEASE HOLDER READER"
    ElseIf Not gRT_Silent Then
        MsgBox messageText, vbInformation, "RELEASE HOLDER READER"
    End If
End Sub

Public Sub CF_SelectReleasePlanReadOnly()
    Dim messageText As String
    If Not RH_SelectReleasePlan(messageText) Then
        If Not gRT_Silent Then MsgBox messageText, vbExclamation, "RELEASE SELECTOR"
    ElseIf Not gRT_Silent Then
        MsgBox messageText, vbInformation, "RELEASE SELECTOR"
    End If
End Sub

Public Sub CF_ReleaseHolderReader_CompileProbe()
    Dim keyText As String
    keyText = RH_SourceKey("M", "B", "C")
    If keyText <> "M|B|C" Then Err.Raise vbObjectError + 9920, , "RELEASE_HOLDER_READER_COMPILE_PROBE_FAIL"
End Sub

Public Function RH_LoadReleaseHolders(ByRef messageText As String) As Boolean
    Dim wsPlan As Worksheet, wsLegs As Worksheet, wsRelease As Worksheet
    Dim ses As Object, upperGrid As Object, holderGrid As Object
    Dim material As String, binText As String, batchText As String, currentOB As String, currentWT As String
    Dim requiredQty As Double, available As Double, totalStock As Double, needRelease As Double
    Dim planID As String, sourceKey As String, runID As String, sessionMessage As String
    Dim r As Long, foundNeed As Boolean, rawCount As Long
    Dim resultText As String
    Dim requestMap As Object, requestKey As Variant, requestValue As Variant, requestList As Collection

    On Error GoTo Failed
    messageText = ""
    If Not RH_ReadConfig(messageText) Then Exit Function
    If Not SAPRun_RequireRuntime("RELEASE_HOLDER_READ", _
        "RELEASE_MON_MATERIAL_ID,RELEASE_MON_BIN_ID,RELEASE_MON_BATCH_ID,RELEASE_MON_EXECUTE_ID,RELEASE_MON_RESULT_GRID_ID,RELEASE_MON_RESULT_COUNT_ID,RELEASE_MON_RESULT_COUNT_PATTERN,RELEASE_MON_OB_COL,RELEASE_MON_WH_COL,RELEASE_MON_WT_COL,RELEASE_MON_MATERIAL_COL,RELEASE_MON_BATCH_COL,RELEASE_MON_BIN_COL,RELEASE_MON_QTY_COL,RELEASE_MON_STATUS_COL", messageText) Then Exit Function
    If Not SAPRun_RequireReadOnly("READ", messageText) Then Exit Function
    If Not CF_CheckPlanner_ValidateSnapshot(messageText) Then Exit Function
    Set wsPlan = ThisWorkbook.Worksheets(RH_PLAN)
    Set wsLegs = ThisWorkbook.Worksheets(RH_PLAN_LEGS)
    Set wsRelease = ThisWorkbook.Worksheets(RH_PLAN_RELEASE)
    runID = RH_Text(ThisWorkbook.Worksheets(RH_WH).Range("W2").Value2)
    If runID = "" Then messageText = "STOP_STALE_WH_SNAPSHOT: WH RunID is blank.": Exit Function

    Set requestMap = RH_BuildReleaseRequests(wsPlan, wsLegs, messageText)
    If requestMap Is Nothing Then Exit Function
    foundNeed = (requestMap.count > 0)
    If Not foundNeed Then
        messageText = "RELEASE_HOLDER_READER_NOT_REQUIRED: no NEED_RELEASE source was produced by the local planner."
        RH_LoadReleaseHolders = True
        Exit Function
    End If

    Set ses = CF_MON_GetExistingReaderSession(sessionMessage)
    If ses Is Nothing Then messageText = sessionMessage: Exit Function
    If Not CF_MON_WaitExistingReaderIdle(ses, RH_ConfigTimeout()) Then
        messageText = "STOP_STAGE_TIMEOUT: WAIT_MON_IDLE before holder query exceeded configured timeout."
        Exit Function
    End If

    RH_ResetReleaseEvidence wsRelease
    For Each requestKey In requestMap.keys
        sourceKey = CStr(requestKey)
        material = RH_KeyPart(sourceKey, 1): binText = RH_KeyPart(sourceKey, 2): batchText = RH_KeyPart(sourceKey, 3)
        Set upperGrid = Nothing: Set holderGrid = Nothing
        If Not RH_QueryMONExact(ses, material, binText, batchText, upperGrid, holderGrid, resultText) Then
            messageText = resultText: Exit Function
        End If
        Set requestList = requestMap(requestKey)
        For Each requestValue In requestList
            planID = CStr(requestValue(0))
            needRelease = CDbl(requestValue(2)): available = CDbl(requestValue(3))
            rawCount = RH_ReadHolderRows(holderGrid, planID, CLng(requestValue(1)), sourceKey, material, binText, batchText, _
                                         RH_PlanCurrentOB(wsPlan, planID), RH_PlanCurrentWT(wsPlan, planID), _
                                         needRelease, available, runID, wsRelease)
            If rawCount = 0 Then
                RH_WriteReleaseStop wsRelease, planID, sourceKey, material, binText, batchText, _
                                    needRelease, available, runID, "STOP_RELEASE_SOURCE_NOT_PROVEN", _
                                    "MON returned no eligible active holder rows for exact Material|Bin|Batch."
                messageText = "STOP_RELEASE_SOURCE_NOT_PROVEN: MON returned no eligible active holder rows for " & sourceKey
                Exit Function
            End If
        Next requestValue
    Next requestKey
    If Not RH_DedupeReleaseEvidence(wsRelease, messageText) Then Exit Function
    If Not RH_SelectReleasePlan(messageText) Then Exit Function
    If Not CF_FinalizeReleasePlanLocal(messageText) Then Exit Function

    RT_Log "RELEASE_HOLDER_READER", "", "", "PASS", _
           "NEED_RELEASE holder query completed; candidates written to PLAN_RELEASE; " & _
           "SAP_WRITE=0; CANCEL=0; SAVE=0; CREATE=0; CF=0"
    messageText = "RELEASE_HOLDER_READER_PASS: exact MON holder rows were captured and selected in PLAN_RELEASE; no SAP action executed."
    RH_LoadReleaseHolders = True
    Exit Function
Failed:
    messageText = "RELEASE_HOLDER_READER_FAIL: Stage=READ_HOLDER; Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function RH_SelectReleasePlan(ByRef messageText As String) As Boolean
    Dim ws As Worksheet, lastRow As Long, planID As String, sourceKey As String
    Dim d As Object, r As Long, needRelease As Double, available As Double
    Dim key As Variant, parts As Variant, reservedWT As Object
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(RH_PLAN_RELEASE)
    lastRow = RH_LastRow(ws, 1, RH_FIRST_ROW)
    If lastRow < RH_FIRST_ROW Then messageText = "RELEASE_SELECTOR_STOP: PLAN_RELEASE has no holder candidates.": Exit Function
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set reservedWT = CreateObject("Scripting.Dictionary"): reservedWT.CompareMode = vbTextCompare
    For r = RH_FIRST_ROW To lastRow
        planID = RH_Text(ws.Cells(r, 1).Value2)
        sourceKey = RH_Text(ws.Cells(r, 3).Value2)
        If planID <> "" And sourceKey <> "" Then
            If Not d.Exists(planID & "|" & sourceKey) Then d.Add planID & "|" & sourceKey, True
        End If
        If UCase$(RH_Text(ws.Cells(r, 13).Value2)) = "TRUE" Then reservedWT(RH_ReservationKey(ws, r)) = True
    Next r
    For Each key In d.keys
        parts = Split(CStr(key), "|")
        planID = CStr(parts(0))
        sourceKey = Mid$(CStr(key), Len(planID) + 2)
        needRelease = RH_FirstNeed(ws, planID, sourceKey)
        available = RH_FirstAvailable(ws, planID, sourceKey)
        If needRelease <= RH_EPS Then
            messageText = "RELEASE_SELECTOR_STOP: Need Release is not positive for " & planID
            Exit Function
        End If
        If Not RH_SelectForPlan(ws, planID, sourceKey, needRelease, available, reservedWT, messageText) Then Exit Function
    Next key
    messageText = "RELEASE_SELECTOR_PASS: deterministic whole-WT selection completed in PLAN_RELEASE; no Cancel executed."
    RH_SelectReleasePlan = True
    Exit Function
Failed:
    messageText = "RELEASE_SELECTOR_FAIL: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_FinalizeReleasePlanLocal(ByRef messageText As String) As Boolean
    Dim wsPlan As Worksheet, wsLegs As Worksheet, wsRelease As Worksheet
    Dim lastPlan As Long, lastLeg As Long, lastRelease As Long, r As Long, legRow As Long, relRow As Long
    Dim planID As String, decision As String, sourceKey As String, gateText As String
    Dim requiredQty As Double, available As Double, projected As Double, needRelease As Double
    Dim allReady As Boolean, hasNeed As Boolean, selectedQty As Double, legNo As Long
    On Error GoTo Failed
    Set wsPlan = ThisWorkbook.Worksheets(RH_PLAN)
    Set wsLegs = ThisWorkbook.Worksheets(RH_PLAN_LEGS)
    Set wsRelease = ThisWorkbook.Worksheets(RH_PLAN_RELEASE)
    lastPlan = RH_LastRow(wsPlan, 1, RH_FIRST_ROW)
    lastLeg = RH_LastRow(wsLegs, 1, RH_FIRST_ROW)
    lastRelease = RH_LastRow(wsRelease, 1, RH_FIRST_ROW)
    For r = RH_FIRST_ROW To lastPlan
        planID = RH_Text(wsPlan.Cells(r, 1).Value2)
        If planID <> "" Then
            hasNeed = False: allReady = True
            For legRow = RH_FIRST_ROW To lastLeg
                If RH_Text(wsLegs.Cells(legRow, 1).Value2) = planID And UCase$(RH_Text(wsLegs.Cells(legRow, 7).Value2)) = "NEED_RELEASE" Then
                    hasNeed = True: legNo = CLng(RH_Number(wsLegs.Cells(legRow, 2).Value2))
                    sourceKey = RH_Text(wsLegs.Cells(legRow, 3).Value2)
                    requiredQty = RH_Number(wsLegs.Cells(legRow, 4).Value2)
                    available = RH_Number(wsLegs.Cells(legRow, 5).Value2)
                    selectedQty = 0#
                    For relRow = RH_FIRST_ROW To lastRelease
                        If RH_Text(wsRelease.Cells(relRow, 1).Value2) = planID And RH_Text(wsRelease.Cells(relRow, 3).Value2) = sourceKey And _
                           UCase$(RH_Text(wsRelease.Cells(relRow, 13).Value2)) = "TRUE" Then selectedQty = selectedQty + RH_Number(wsRelease.Cells(relRow, 10).Value2)
                    Next relRow
                    projected = available + selectedQty
                    If projected + RH_EPS < requiredQty Then allReady = False
                End If
            Next legRow
            If hasNeed Then
                If allReady Then
                    decision = "RELEASE_READY": gateText = "RELEASE_READY"
                    wsPlan.Cells(r, 18).Value2 = projected
                    wsPlan.Cells(r, 17).Value2 = 0#
                    wsPlan.Cells(r, 26).Value2 = "Final local release evidence is complete; no SAP write or Cancel executed."
                Else
                    decision = "NEED_RELEASE": gateText = "STOP_RELEASE_INSUFFICIENT"
                    wsPlan.Cells(r, 26).Value2 = "PLAN_RELEASE evidence does not cover the required quantity; no release action executed."
                End If
                wsPlan.Cells(r, 20).Value2 = decision
                wsPlan.Cells(r, 24).Value2 = gateText
            End If
        End If
    Next r
    messageText = "FINAL_RELEASE_PLAN_PASS: PLAN updated from existing PLAN_RELEASE evidence; evidence preserved; no SAP action executed."
    CF_FinalizeReleasePlanLocal = True
    Exit Function
Failed:
    messageText = "FINAL_RELEASE_PLAN_FAIL: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function RH_QueryMONExact(ByVal ses As Object, ByVal material As String, ByVal binText As String, _
                                  ByVal batchText As String, ByRef upperGrid As Object, _
                                  ByRef holderGrid As Object, ByRef messageText As String) As Boolean
    Dim gridId As String, colOB As String, colWT As String, colMat As String, colBatch As String
    Dim colBin As String, colQty As String, colStatus As String, colWH As String
    Dim resultText As String
    On Error GoTo Failed
    gridId = RH_Config("RELEASE_MON_RESULT_GRID_ID")
    colOB = RH_Config("RELEASE_MON_OB_COL")
    colWT = RH_Config("RELEASE_MON_WT_COL")
    colMat = RH_Config("RELEASE_MON_MATERIAL_COL")
    colBatch = RH_Config("RELEASE_MON_BATCH_COL")
    colBin = RH_Config("RELEASE_MON_BIN_COL")
    colQty = RH_Config("RELEASE_MON_QTY_COL")
    colStatus = RH_Config("RELEASE_MON_STATUS_COL")
    colWH = RH_Config("RELEASE_MON_WH_COL")
    If gridId = "" Or colOB = "" Or colWH = "" Or colWT = "" Or colMat = "" Or colBatch = "" Or colBin = "" Or colQty = "" Or colStatus = "" Then
        messageText = "STOP_RELEASE_LAYOUT_UNPROVEN: configure exact RELEASE_MON_*_COL keys before holder query."
        Exit Function
    End If

    If Not RH_ConfigureMONFilter(ses, material, binText, batchText, resultText) Then
        messageText = resultText
        Exit Function
    End If
    If Not CF_MON_WaitExistingReaderIdle(ses, RH_ConfigTimeout()) Then
        messageText = "STOP_STAGE_TIMEOUT: MON exact holder query did not become idle in time."
        Exit Function
    End If
    Set upperGrid = RH_FindById(ses, gridId)
    If upperGrid Is Nothing Then
        messageText = "STOP_CONTROL_NOT_FOUND: RELEASE_MON_RESULT_GRID_ID=" & gridId
        Exit Function
    End If
    If Not RH_HasColumn(upperGrid, colOB) Or Not RH_HasColumn(upperGrid, colWH) Or Not RH_HasColumn(upperGrid, colWT) Or _
       Not RH_HasColumn(upperGrid, colMat) Or Not RH_HasColumn(upperGrid, colBatch) Or _
       Not RH_HasColumn(upperGrid, colBin) Or Not RH_HasColumn(upperGrid, colQty) Or _
       Not RH_HasColumn(upperGrid, colStatus) Then
        messageText = "STOP_CONTROL_NOT_FOUND: one or more RELEASE_MON_*_COL IDs are missing from the exact result grid."
        Exit Function
    End If
    If Not RH_ExposeHolderRows(ses, upperGrid, resultText) Then
        messageText = resultText
        Exit Function
    End If
    Set holderGrid = upperGrid
    RH_QueryMONExact = True
    Exit Function
Failed:
    messageText = "STOP_RELEASE_READER_ERROR: MON exact query failed. Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function RH_ConfigureMONFilter(ByVal ses As Object, ByVal material As String, ByVal binText As String, _
                                       ByVal batchText As String, ByRef messageText As String) As Boolean
    Dim matID As String, binID As String, batchID As String, btnID As String
    Dim ctl As Object, btn As Object
    On Error GoTo Failed
    matID = RH_Config("RELEASE_MON_MATERIAL_ID")
    binID = RH_Config("RELEASE_MON_BIN_ID")
    batchID = RH_Config("RELEASE_MON_BATCH_ID")
    btnID = RH_Config("RELEASE_MON_EXECUTE_ID")
    If matID = "" Or binID = "" Or batchID = "" Or btnID = "" Then
        messageText = "STOP_RELEASE_LAYOUT_UNPROVEN: configure RELEASE_MON_MATERIAL_ID, RELEASE_MON_BIN_ID, RELEASE_MON_BATCH_ID, RELEASE_MON_EXECUTE_ID."
        Exit Function
    End If
    Set ctl = RH_FindById(ses, matID)
    If ctl Is Nothing Then messageText = "STOP_CONTROL_NOT_FOUND: RELEASE_MON_MATERIAL_ID=" & matID: Exit Function
    RH_SetValue ctl, ""
    If Not RH_SetAndVerify ctl, material Then messageText = "STOP_RELEASE_FILTER_NOT_PROVEN: MON Material read-back mismatch.": Exit Function
    Set ctl = RH_FindById(ses, binID)
    If ctl Is Nothing Then messageText = "STOP_CONTROL_NOT_FOUND: RELEASE_MON_BIN_ID=" & binID: Exit Function
    RH_SetValue ctl, ""
    If Not RH_SetAndVerify ctl, binText Then messageText = "STOP_RELEASE_FILTER_NOT_PROVEN: MON Bin read-back mismatch.": Exit Function
    Set ctl = RH_FindById(ses, batchID)
    If ctl Is Nothing Then messageText = "STOP_CONTROL_NOT_FOUND: RELEASE_MON_BATCH_ID=" & batchID: Exit Function
    RH_SetValue ctl, ""
    If Not RH_SetAndVerify ctl, batchText Then messageText = "STOP_RELEASE_FILTER_NOT_PROVEN: MON Batch read-back mismatch.": Exit Function
    Set btn = RH_FindById(ses, btnID)
    If btn Is Nothing Then messageText = "STOP_CONTROL_NOT_FOUND: RELEASE_MON_EXECUTE_ID=" & btnID: Exit Function
    btn.Press
    RH_ConfigureMONFilter = True
    Exit Function
Failed:
    messageText = "STOP_RELEASE_READER_ERROR: MON filter failed. Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function RH_ReadHolderRows(ByVal holderGrid As Object, ByVal planID As String, ByVal legNo As Long, _
                                   ByVal sourceKey As String, ByVal material As String, ByVal binText As String, ByVal batchText As String, _
                                   ByVal currentOB As String, ByVal currentWT As String, ByVal needRelease As Double, _
                                   ByVal available As Double, _
                                   ByVal runID As String, ByVal ws As Worksheet) As Long
    Dim lastRow As Long, r As Long, ob As String, whNo As String, wt As String, rowMat As String, rowBin As String, rowBatch As String
    Dim statusText As String, qty As Double, eligible As Boolean, outRow As Long
    lastRow = CLng(holderGrid.rowCount)
    For r = 0 To lastRow - 1
        ob = RH_GridText(holderGrid, r, RH_Config("RELEASE_MON_OB_COL"))
        whNo = RH_GridText(holderGrid, r, RH_Config("RELEASE_MON_WH_COL"))
        wt = RH_GridText(holderGrid, r, RH_Config("RELEASE_MON_WT_COL"))
        rowMat = RH_GridText(holderGrid, r, RH_Config("RELEASE_MON_MATERIAL_COL"))
        rowBin = RH_GridText(holderGrid, r, RH_Config("RELEASE_MON_BIN_COL"))
        rowBatch = RH_GridText(holderGrid, r, RH_Config("RELEASE_MON_BATCH_COL"))
        statusText = UCase$(RH_GridText(holderGrid, r, RH_Config("RELEASE_MON_STATUS_COL")))
        qty = RH_Number(RH_GridText(holderGrid, r, RH_Config("RELEASE_MON_QTY_COL")))
        eligible = (StrComp(rowMat, material, vbTextCompare) = 0 And _
                    StrComp(rowBin, binText, vbTextCompare) = 0 And _
                    StrComp(rowBatch, batchText, vbTextCompare) = 0 And qty > RH_EPS _
                    And statusText = "" _
                    And StrComp(ob, currentOB, vbTextCompare) <> 0 And StrComp(wt, currentWT, vbTextCompare) <> 0)
        If eligible Then
            outRow = RH_LastRow(ws, 1, RH_FIRST_ROW) + 1
            ws.Cells(outRow, 1).Value2 = planID
            ws.Cells(outRow, 2).Value2 = legNo
            ws.Cells(outRow, 3).Value2 = sourceKey
            ws.Cells(outRow, 4).Value2 = ob
            ws.Cells(outRow, 5).Value2 = whNo
            ws.Cells(outRow, 6).Value2 = wt
            ws.Cells(outRow, 7).Value2 = rowMat
            ws.Cells(outRow, 8).Value2 = rowBin
            ws.Cells(outRow, 9).Value2 = rowBatch
            ws.Cells(outRow, 10).Value2 = qty
            ws.Cells(outRow, 11).Value2 = IIf(statusText = "", "WT_ACTIVE", statusText)
            ws.Cells(outRow, 12).Value2 = "TRUE"
            ws.Cells(outRow, 13).Value2 = "FALSE"
            ws.Cells(outRow, 15).Value2 = needRelease
            ws.Cells(outRow, 16).Value2 = available
            ws.Cells(outRow, 21).Value2 = available
            ws.Cells(outRow, 17).Value2 = "CANDIDATE_ONLY"
            ws.Cells(outRow, 18).Value2 = "Exact MON Material|Bin|Batch holder row; read-only."
            ws.Cells(outRow, 19).Value2 = runID
            ws.Cells(outRow, 20).Value2 = "MON"
            RH_ReadHolderRows = RH_ReadHolderRows + 1
        End If
    Next r
End Function

Private Function RH_SelectForPlan(ByVal ws As Worksheet, ByVal planID As String, ByVal sourceKey As String, _
                                  ByVal needRelease As Double, ByVal available As Double, _
                                  ByVal reservedWT As Object, ByRef messageText As String) As Boolean
    Dim rows As Collection, r As Long, lastRow As Long, qty As Double, total As Double, bestRow As Long, bestQty As Double
    Dim i As Long, outRow As Long, projected As Double
    Dim chosenCount As Long, candidateRow As Long, candidateQty As Double, candidateWT As String
    Dim reservationKey As String
    On Error GoTo Failed
    Set rows = New Collection
    lastRow = RH_LastRow(ws, 1, RH_FIRST_ROW)
    For r = RH_FIRST_ROW To lastRow
        If RH_Text(ws.Cells(r, 1).Value2) = planID And RH_Text(ws.Cells(r, 3).Value2) = sourceKey _
           And UCase$(RH_Text(ws.Cells(r, 12).Value2)) = "TRUE" And _
           UCase$(RH_Text(ws.Cells(r, 11).Value2)) = "WT_ACTIVE" Then rows.Add r
    Next r
    If rows.count = 0 Then messageText = "STOP_RELEASE_SOURCE_NOT_PROVEN: no eligible holders for " & sourceKey: Exit Function
    For i = 1 To rows.count
        r = CLng(rows(i)): qty = RH_Number(ws.Cells(r, 10).Value2)
        reservationKey = RH_ReservationKey(ws, r)
        If qty + RH_EPS >= needRelease And Not reservedWT.Exists(reservationKey) Then
            If bestRow = 0 Or qty < bestQty Or (Abs(qty - bestQty) <= RH_EPS And _
               StrComp(RH_Text(ws.Cells(r, 6).Value2), RH_Text(ws.Cells(bestRow, 6).Value2), vbTextCompare) < 0) Then
                bestRow = r: bestQty = qty
            End If
        End If
    Next i
    For r = RH_FIRST_ROW To lastRow
        If RH_Text(ws.Cells(r, 1).Value2) = planID And RH_Text(ws.Cells(r, 3).Value2) = sourceKey Then ws.Cells(r, 13).Value2 = "FALSE"
    Next r
    If bestRow > 0 Then
        ws.Cells(bestRow, 13).Value2 = "TRUE": ws.Cells(bestRow, 14).Value2 = 1
        total = bestQty
        reservedWT(RH_ReservationKey(ws, bestRow)) = True
    Else
        Do While total + RH_EPS < needRelease
            candidateRow = 0: candidateQty = 0#: candidateWT = ""
             For i = 1 To rows.count
                 r = CLng(rows(i))
                 If UCase$(RH_Text(ws.Cells(r, 13).Value2)) <> "TRUE" Then
                     qty = RH_Number(ws.Cells(r, 10).Value2)
                     reservationKey = RH_ReservationKey(ws, r)
                     If Not reservedWT.Exists(reservationKey) And _
                        (candidateRow = 0 Or qty > candidateQty Or (Abs(qty - candidateQty) <= RH_EPS And _
                        StrComp(RH_Text(ws.Cells(r, 6).Value2), candidateWT, vbTextCompare) < 0)) Then
                        candidateRow = r: candidateQty = qty: candidateWT = RH_Text(ws.Cells(r, 6).Value2)
                    End If
                End If
            Next i
            If candidateRow = 0 Then Exit Do
            chosenCount = chosenCount + 1
            ws.Cells(candidateRow, 13).Value2 = "TRUE": ws.Cells(candidateRow, 14).Value2 = chosenCount
            reservedWT(RH_ReservationKey(ws, candidateRow)) = True
            total = total + candidateQty
        Loop
    End If
    projected = available + total
    For r = RH_FIRST_ROW To lastRow
        If RH_Text(ws.Cells(r, 1).Value2) = planID And RH_Text(ws.Cells(r, 3).Value2) = sourceKey Then
            ws.Cells(r, 16).Value2 = projected
            If projected + RH_EPS >= available + needRelease Then
                ws.Cells(r, 17).Value2 = "RELEASE_READY"
                ws.Cells(r, 18).Value2 = "Projected Available=" & CStr(projected) & "; whole WT selection; no Cancel executed."
            Else
                ws.Cells(r, 17).Value2 = "STOP_RELEASE_INSUFFICIENT"
                ws.Cells(r, 18).Value2 = "Projected Available=" & CStr(projected) & "; active holder qty is insufficient."
            End If
        End If
    Next r
    If projected + RH_EPS < available + needRelease Then
        messageText = "STOP_RELEASE_INSUFFICIENT: eligible active holders do not cover Need Release for " & sourceKey
        Exit Function
    End If
    RH_SelectForPlan = True
    Exit Function
Failed:
    messageText = "RELEASE_SELECTOR_FAIL: Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function RH_ReadConfig(ByRef messageText As String) As Boolean
    If RH_Config("RELEASE_MON_RESULT_GRID_ID") = "" Or RH_Config("RELEASE_MON_OB_COL") = "" Or RH_Config("RELEASE_MON_WH_COL") = "" Or RH_Config("RELEASE_MON_WT_COL") = "" Or _
       RH_Config("RELEASE_MON_MATERIAL_COL") = "" Or RH_Config("RELEASE_MON_BATCH_COL") = "" Or RH_Config("RELEASE_MON_BIN_COL") = "" Or _
       RH_Config("RELEASE_MON_QTY_COL") = "" Or RH_Config("RELEASE_MON_STATUS_COL") = "" Or RH_Config("RELEASE_MON_RESULT_COUNT_ID") = "" Or _
       RH_Config("RELEASE_MON_RESULT_COUNT_PATTERN") = "" Or RH_Config("RELEASE_MON_MATERIAL_ID") = "" Or _
       RH_Config("RELEASE_MON_BIN_ID") = "" Or RH_Config("RELEASE_MON_BATCH_ID") = "" Or RH_Config("RELEASE_MON_EXECUTE_ID") = "" Then
        messageText = "STOP_RELEASE_LAYOUT_UNPROVEN: holder reader requires exact RELEASE_MON_* CONFIG mapping."
        Exit Function
    End If
    RH_ReadConfig = True
End Function

Private Sub RH_ResetReleaseEvidence(ByVal ws As Worksheet)
    ws.Range("A2:U2000").ClearContents
End Sub

Private Sub RH_WriteReleaseStop(ByVal ws As Worksheet, ByVal planID As String, ByVal sourceKey As String, _
                                ByVal material As String, ByVal binText As String, ByVal batchText As String, _
                                ByVal needRelease As Double, ByVal available As Double, ByVal runID As String, ByVal gateText As String, _
                                ByVal reasonText As String)
    Dim r As Long
    r = RH_LastRow(ws, 1, RH_FIRST_ROW) + 1
    ws.Cells(r, 1).Value2 = planID: ws.Cells(r, 3).Value2 = sourceKey
    ws.Cells(r, 7).Value2 = material: ws.Cells(r, 8).Value2 = binText: ws.Cells(r, 9).Value2 = batchText
    ws.Cells(r, 15).Value2 = needRelease: ws.Cells(r, 17).Value2 = gateText
    ws.Cells(r, 18).Value2 = reasonText: ws.Cells(r, 19).Value2 = runID: ws.Cells(r, 20).Value2 = "MON"
    ws.Cells(r, 21).Value2 = RH_Number(available)
End Sub

Private Function RH_BuildReleaseRequests(ByVal wsPlan As Worksheet, ByVal wsLegs As Worksheet, _
                                         ByRef messageText As String) As Object
    Dim d As Object, requests As Collection, lastPlan As Long, lastLeg As Long
    Dim r As Long, legNo As Long, planID As String, sourceKey As String, gateText As String
    Dim requiredQty As Double, available As Double, needRelease As Double, item As Variant
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    lastPlan = RH_LastRow(wsPlan, 1, RH_FIRST_ROW)
    lastLeg = RH_LastRow(wsLegs, 1, RH_FIRST_ROW)
    For r = RH_FIRST_ROW To lastLeg
        planID = RH_Text(wsLegs.Cells(r, 1).Value2)
        gateText = UCase$(RH_Text(wsLegs.Cells(r, 7).Value2))
        If planID <> "" And gateText = "NEED_RELEASE" Then
            legNo = CLng(RH_Number(wsLegs.Cells(r, 2).Value2))
            sourceKey = RH_Text(wsLegs.Cells(r, 3).Value2)
            requiredQty = RH_Number(wsLegs.Cells(r, 4).Value2)
            available = RH_Number(wsLegs.Cells(r, 5).Value2)
            needRelease = requiredQty - available
            If sourceKey = "" Or needRelease <= RH_EPS Then
                messageText = "STOP_RELEASE_SOURCE_NOT_PROVEN: invalid NEED_RELEASE leg for " & planID
                Exit Function
            End If
            If Not d.Exists(sourceKey) Then
                Set requests = New Collection
                d.Add sourceKey, requests
            Else
                Set requests = d(sourceKey)
            End If
            requests.Add Array(planID, legNo, needRelease, available)
        End If
    Next r
    Set RH_BuildReleaseRequests = d
End Function

Private Function RH_PlanCurrentOB(ByVal wsPlan As Worksheet, ByVal planID As String) As String
    Dim r As Long, lastRow As Long
    lastRow = RH_LastRow(wsPlan, 1, RH_FIRST_ROW)
    For r = RH_FIRST_ROW To lastRow
        If RH_Text(wsPlan.Cells(r, 1).Value2) = planID Then RH_PlanCurrentOB = RH_Text(wsPlan.Cells(r, 2).Value2): Exit Function
    Next r
End Function

Private Function RH_PlanCurrentWT(ByVal wsPlan As Worksheet, ByVal planID As String) As String
    Dim r As Long, lastRow As Long
    lastRow = RH_LastRow(wsPlan, 1, RH_FIRST_ROW)
    For r = RH_FIRST_ROW To lastRow
        If RH_Text(wsPlan.Cells(r, 1).Value2) = planID Then RH_PlanCurrentWT = RH_Text(wsPlan.Cells(r, 5).Value2): Exit Function
    Next r
End Function

Private Function RH_ReservationKey(ByVal ws As Worksheet, ByVal rowNo As Long) As String
    RH_ReservationKey = RH_Text(ws.Cells(rowNo, 4).Value2) & "|" & RH_Text(ws.Cells(rowNo, 5).Value2) & "|" & RH_Text(ws.Cells(rowNo, 6).Value2)
End Function

Private Function RH_DedupeReleaseEvidence(ByVal ws As Worksheet, ByRef messageText As String) As Boolean
    Dim d As Object, lastRow As Long, r As Long, keyText As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    lastRow = RH_LastRow(ws, 1, RH_FIRST_ROW)
    For r = RH_FIRST_ROW To lastRow
        If RH_Text(ws.Cells(r, 1).Value2) <> "" And RH_Text(ws.Cells(r, 4).Value2) <> "" Then
            keyText = RH_Text(ws.Cells(r, 4).Value2) & "|" & RH_Text(ws.Cells(r, 5).Value2) & "|" & RH_Text(ws.Cells(r, 6).Value2) & "|" & RH_Text(ws.Cells(r, 7).Value2) & "|" & RH_Text(ws.Cells(r, 8).Value2) & "|" & RH_Text(ws.Cells(r, 9).Value2)
            If d.Exists(keyText) Then
                If Abs(RH_Number(ws.Cells(r, 10).Value2) - CDbl(d(keyText))) > RH_EPS Then
                    messageText = "STOP_AMBIGUOUS_HOLDER: duplicate holder key has conflicting quantities: " & keyText
                    Exit Function
                End If
                ws.rows(r).ClearContents
            Else
                d.Add keyText, RH_Number(ws.Cells(r, 10).Value2)
            End If
        End If
    Next r
    RH_DedupeReleaseEvidence = True
End Function

Private Function RH_SetAndVerify(ByVal ctl As Object, ByVal valueText As String) As Boolean
    RH_SetValue ctl, valueText
    RH_SetAndVerify = (StrComp(RH_ControlText(ctl), valueText, vbTextCompare) = 0)
End Function

Private Function RH_ControlText(ByVal ctl As Object) As String
    On Error Resume Next
    RH_ControlText = Trim$(CStr(ctl.text & ""))
    If RH_ControlText = "" Then RH_ControlText = Trim$(CStr(ctl.key & ""))
    On Error GoTo 0
End Function

Private Function RH_ExposeHolderRows(ByVal ses As Object, ByVal grid As Object, ByRef messageText As String) As Boolean
    Dim expectedText As String, expectedRows As Long, countID As String, patternText As String
    Dim lastCount As Long, stable As Long, started As Double, rowCount As Long
    Dim maxRows As Long
    countID = RH_Config("RELEASE_MON_RESULT_COUNT_ID")
    patternText = RH_Config("RELEASE_MON_RESULT_COUNT_PATTERN")
    If countID = "" Or patternText = "" Then
        messageText = "STOP_RELEASE_CAPTURE_UNPROVEN: configure RELEASE_MON_RESULT_COUNT_ID/PATTERN before holder query."
        Exit Function
    End If
    expectedText = RH_ControlText(RH_FindById(ses, countID))
    If Not RH_ParseCount(expectedText, patternText, expectedRows) Then
        messageText = "STOP_RELEASE_CAPTURE_UNPROVEN: MON independent holder result count could not be parsed."
        Exit Function
    End If
    If expectedRows = 0 Then
        If CLng(grid.rowCount) = 0 Then RH_ExposeHolderRows = True Else messageText = "STOP_RELEASE_CAPTURE_UNPROVEN: independent MON count is zero but grid is not empty."
        Exit Function
    End If
    maxRows = ModuleHelper.ToolConfigLong("RELEASE_MAX_EXPECTED_ROWS", 2000, 1, 20000)
    If expectedRows > maxRows Then
        messageText = "STOP_RELEASE_CAPTURE_UNPROVEN: independent MON count=" & CStr(expectedRows) & _
                      " exceeds RELEASE_MAX_EXPECTED_ROWS=" & CStr(maxRows)
        Exit Function
    End If
    started = Timer: lastCount = -1
    Do
        If RH_Elapsed(started) >= RH_ConfigTimeout() Then Exit Do
        rowCount = CLng(grid.rowCount)
        If rowCount = expectedRows Then RH_ExposeHolderRows = True: Exit Function
        If rowCount > expectedRows Then messageText = "STOP_RELEASE_CAPTURE_UNPROVEN: MON RowCount exceeds independent holder count.": Exit Function
        If rowCount = lastCount Then stable = stable + 1 Else stable = 0: lastCount = rowCount
        If stable >= 2 Then
            On Error Resume Next
            grid.FirstVisibleRow = rowCount + 1
            On Error GoTo 0
        End If
        RH_Sleep 150: DoEvents
    Loop
    messageText = "STOP_RELEASE_CAPTURE_UNPROVEN: MON holder result did not expose all expected rows."
End Function

Private Function RH_ParseCount(ByVal rawText As String, ByVal patternText As String, ByRef resultCount As Long) As Boolean
    Dim p As Long, tailText As String, digits As String, i As Long, ch As String
    p = InStr(1, rawText, patternText, vbTextCompare)
    If p = 0 Then Exit Function
    tailText = Mid$(rawText, p + Len(patternText))
    For i = 1 To Len(tailText)
        ch = Mid$(tailText, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch ElseIf digits <> "" Then Exit For
    Next i
    If digits = "" Then Exit Function
    resultCount = CLng(digits): RH_ParseCount = True
End Function

Private Function RH_PlanHasNeedRelease(ByVal ws As Worksheet, ByVal planID As String) As Boolean
    Dim lastRow As Long, r As Long
    lastRow = RH_LastRow(ws, 1, RH_FIRST_ROW)
    For r = RH_FIRST_ROW To lastRow
        If RH_Text(ws.Cells(r, 1).Value2) = planID And _
           UCase$(RH_Text(ws.Cells(r, 7).Value2)) = "NEED_RELEASE" Then
            RH_PlanHasNeedRelease = True
            Exit Function
        End If
    Next r
End Function

Private Function RH_PlanNeedRelease(ByVal ws As Worksheet, ByVal planID As String) As Double
    Dim lastRow As Long, r As Long
    lastRow = RH_LastRow(ws, 1, RH_FIRST_ROW)
    For r = RH_FIRST_ROW To lastRow
        If RH_Text(ws.Cells(r, 1).Value2) = planID And UCase$(RH_Text(ws.Cells(r, 7).Value2)) = "NEED_RELEASE" Then RH_PlanNeedRelease = RH_Number(ws.Cells(r, 4).Value2) - RH_Number(ws.Cells(r, 5).Value2): Exit Function
    Next r
End Function

Private Function RH_FirstNeed(ByVal ws As Worksheet, ByVal planID As String, ByVal sourceKey As String) As Double
    Dim r As Long, lastRow As Long
    lastRow = RH_LastRow(ws, 1, RH_FIRST_ROW)
    For r = RH_FIRST_ROW To lastRow
        If RH_Text(ws.Cells(r, 1).Value2) = planID And RH_Text(ws.Cells(r, 3).Value2) = sourceKey Then RH_FirstNeed = RH_Number(ws.Cells(r, 15).Value2): Exit Function
    Next r
End Function

Private Function RH_FirstAvailable(ByVal ws As Worksheet, ByVal planID As String, ByVal sourceKey As String) As Double
    Dim r As Long, lastRow As Long
    lastRow = RH_LastRow(ws, 1, RH_FIRST_ROW)
    For r = RH_FIRST_ROW To lastRow
        If RH_Text(ws.Cells(r, 1).Value2) = planID And RH_Text(ws.Cells(r, 3).Value2) = sourceKey Then RH_FirstAvailable = RH_Number(ws.Cells(r, 21).Value2): Exit Function
    Next r
End Function

Private Function RH_HasColumn(ByVal grid As Object, ByVal colId As String) As Boolean
    Dim order As Object, i As Long, keyText As String
    On Error Resume Next: Set order = grid.columnOrder: On Error GoTo 0
    If order Is Nothing Then Exit Function
    For i = 0 To CLng(order.count) - 1
        On Error Resume Next
        keyText = CStr(order.ElementAt(i) & "")
        If Err.Number <> 0 Then Err.Clear: keyText = CStr(order.item(i) & "")
        On Error GoTo 0
        If StrComp(keyText, colId, vbTextCompare) = 0 Then RH_HasColumn = True: Exit Function
    Next i
End Function

Private Function RH_GridText(ByVal grid As Object, ByVal rowNo As Long, ByVal colId As String) As String
    On Error Resume Next: RH_GridText = Trim$(CStr(grid.GetCellValue(rowNo, colId) & "")): On Error GoTo 0
End Function

Private Function RH_FindById(ByVal ses As Object, ByVal idText As String) As Object
    If ses Is Nothing Or Trim$(idText) = "" Then Exit Function
    On Error Resume Next
    Set RH_FindById = ses.FindById(idText, False)
    On Error GoTo 0
End Function

Private Sub RH_SetValue(ByVal ctl As Object, ByVal valueText As String)
    If ctl Is Nothing Then Exit Sub
    On Error Resume Next
    ctl.text = valueText
    If Err.Number <> 0 Then Err.Clear: ctl.key = valueText
    Err.Clear
    On Error GoTo 0
End Sub

Private Function RH_Config(ByVal keyName As String, Optional ByVal defaultValue As String = "") As String
    RH_Config = Trim$(ModuleHelper.ToolConfig(keyName, defaultValue))
End Function

Private Function RH_ConfigTimeout() As Double
    RH_ConfigTimeout = ModuleHelper.ToolConfigDouble("RELEASE_QUERY_TIMEOUT_SEC", 30#, 5#, 120#)
End Function

Private Function RH_Elapsed(ByVal started As Double) As Double
    If Timer >= started Then
        RH_Elapsed = Timer - started
    Else
        RH_Elapsed = 86400# - started + Timer
    End If
End Function

Private Function RH_Text(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    RH_Text = Trim$(CStr(valueIn))
End Function

Private Function RH_Number(ByVal valueIn As Variant) As Double
    If IsNumeric(valueIn) Then RH_Number = CDbl(valueIn) Else RH_Number = RT_Number(valueIn)
End Function

Private Function RH_KeyPart(ByVal keyText As String, ByVal partNo As Long) As String
    Dim p As Variant: p = Split(keyText, "|")
    If partNo >= 1 And partNo <= UBound(p) + 1 Then RH_KeyPart = CStr(p(partNo - 1))
End Function

Private Function RH_SourceKey(ByVal material As String, ByVal binText As String, ByVal batchText As String) As String
    RH_SourceKey = Trim$(material) & "|" & Trim$(binText) & "|" & Trim$(batchText)
End Function

Private Function RH_LastRow(ByVal ws As Worksheet, ByVal colNo As Long, ByVal firstRow As Long) As Long
    Dim r As Long
    r = ws.Cells(ws.rows.count, colNo).End(xlUp).row
    If r < firstRow Then RH_LastRow = firstRow - 1 Else RH_LastRow = r
End Function
