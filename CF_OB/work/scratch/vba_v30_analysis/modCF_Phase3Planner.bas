Attribute VB_Name = "modCF_Phase3Planner"
Option Explicit

Private Const P3_PLAN_SHEET As String = "PHASE3_PLAN"
Private Const P3_TEST_SHEET As String = "PHASE3_TESTS"
Private Const P3_RAW_SHEET As String = "POSTING_RAW"
Private Const P3_POST_SHEET As String = "POSTING"
Private Const P3_WH_SHEET As String = "DATA_CORE"
Private P3_TargetOB As String

Public Sub CF_Phase3_ReadOnly()
    Dim messageText As String
    On Error GoTo Failed
    If Not P3_PreflightPilot(messageText) Then
        P3_WritePilotStop messageText
        MsgBox messageText, vbExclamation, "PHASE 3 STOP"
        Exit Sub
    End If
    If Not P3_RebuildPostingStrict(messageText) Then
        MsgBox messageText, vbExclamation, "PHASE 3 STOP"
        Exit Sub
    End If
    If Not P3_ReadState(messageText) Then
        MsgBox messageText, vbExclamation, "PHASE 3 STOP"
        Exit Sub
    End If
    If Not P3_BuildDecision(messageText) Then
        MsgBox messageText, vbExclamation, "PHASE 3 STOP"
        Exit Sub
    End If
    If Not P3_BuildReleasePlan(messageText) Then
        MsgBox messageText, vbExclamation, "PHASE 3 STOP"
        Exit Sub
    End If
    If Not P3_ValidatePlan(messageText) Then
        MsgBox messageText, vbExclamation, "PHASE 3 STOP"
        Exit Sub
    End If
    MsgBox messageText, vbInformation, "PHASE 3 READ-ONLY"
    Exit Sub
Failed:
    MsgBox "PHASE3_READONLY_FAIL: " & CStr(Err.Number) & ": " & Err.Description, _
           vbExclamation, "PHASE 3 STOP"
End Sub

Public Sub CF_Phase3_ReadOnly_NoDialog()
    Dim messageText As String
    Dim ws As Worksheet
    If Not P3_PreflightPilot(messageText) Then
        P3_WritePilotStop messageText
        Exit Sub
    End If
    If Not P3_RebuildPostingStrict(messageText) Then
        Set ws = P3_EnsureSheet(P3_PLAN_SHEET, False)
        ws.Range("V1").Value2 = messageText
        Exit Sub
    End If
    If Not P3_ReadState(messageText) Then
        Set ws = P3_EnsureSheet(P3_PLAN_SHEET, False)
        ws.Range("V1").Value2 = messageText
        Exit Sub
    End If
    If Not P3_BuildDecision(messageText) Then
        Set ws = P3_EnsureSheet(P3_PLAN_SHEET, False)
        ws.Range("V1").Value2 = messageText
        Exit Sub
    End If
    If Not P3_BuildReleasePlan(messageText) Then
        Set ws = P3_EnsureSheet(P3_PLAN_SHEET, False)
        ws.Range("V1").Value2 = messageText
        Exit Sub
    End If
    If Not P3_ValidatePlan(messageText) Then
        Set ws = P3_EnsureSheet(P3_PLAN_SHEET, False)
        ws.Range("V1").Value2 = messageText
        Exit Sub
    End If
    Set ws = P3_EnsureSheet(P3_PLAN_SHEET, False)
    ws.Range("V1").Value2 = messageText
End Sub

Private Function P3_PreflightPilot(ByRef messageText As String) As Boolean
    Dim ws As Worksheet, ob As String
    Dim lastRow As Long, r As Long, foundCount As Long
    On Error GoTo Failed
    ob = P3_Text(CF_GetInputSheet().Range("A7").Value2)
    If ob = "" Then
        messageText = "PILOT_OB_MISSING: INPUT!A7 is blank. STOP."
        Exit Function
    End If
    Set ws = ThisWorkbook.Worksheets(P3_WH_SHEET)
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 5 To lastRow
        If StrComp(P3_Text(ws.Cells(r, 1).Value2), ob, vbTextCompare) = 0 Then
            foundCount = foundCount + 1
        End If
    Next r
    If foundCount = 0 Then
        messageText = "PILOT_OB_NOT_FOUND_IN_WH_ACTUAL: OB=" & ob & _
                      "; cached WH/WT/Material/Qty/Bin/Batch chain is missing. STOP."
        Exit Function
    End If
    P3_TargetOB = ob
    messageText = "PILOT_OB_FOUND: OB=" & ob & "; rows=" & CStr(foundCount)
    P3_PreflightPilot = True
    Exit Function
Failed:
    messageText = "PILOT_PREFLIGHT_FAIL: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub P3_WritePilotStop(ByVal messageText As String)
    Dim ws As Worksheet, ob As String
    On Error Resume Next
    Set ws = P3_EnsureSheet(P3_PLAN_SHEET, False)
    ob = P3_Text(CF_GetInputSheet().Range("A7").Value2)
    ws.Cells.ClearContents
    ws.Range("A1:T1").value = Array("Source Row", "OB", "Document", "WH", "WT", "Material", "Sys Batch", "Sys Bin", "Sys Qty", "Actual Qty", "Actual Bin", "Actual Batch", "Posting Available", "Total Stock", "Projected Available", "Decision", "Release Status", "Release WT List", "Action", "Note")
    ws.Cells(2, 2).Value2 = ob
    ws.Cells(2, 16).Value2 = "REVIEW_STOP"
    ws.Cells(2, 19).Value2 = "STOP"
    ws.Cells(2, 20).Value2 = messageText & "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
    ws.Range("V1").Value2 = "PILOT STOP: " & messageText
    RT_Log "PHASE3_PILOT", ob, "", "FAIL", messageText & _
        "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
End Sub

Public Sub CF_RebuildPostingStrict()
    Dim messageText As String
    If Not P3_RebuildPostingStrict(messageText) Then
        MsgBox messageText, vbExclamation, "POSTING MAP STOP"
    Else
        MsgBox messageText, vbInformation, "POSTING MAP"
    End If
End Sub

Private Function P3_RebuildPostingStrict(ByRef messageText As String) As Boolean
    Dim wsRaw As Worksheet, wsPost As Worksheet
    Dim lastRaw As Long, r As Long, c As Long, outRow As Long
    Dim rawValid As Long, mapped As Long
    On Error GoTo Failed
    Set wsRaw = ThisWorkbook.Worksheets(P3_RAW_SHEET)
    Set wsPost = ThisWorkbook.Worksheets(P3_POST_SHEET)
    lastRaw = wsRaw.Cells(wsRaw.rows.count, 1).End(xlUp).row
    For r = 2 To lastRaw
        If P3_ValidRawRow(wsRaw, r) Then rawValid = rawValid + 1
    Next r
    wsPost.Range("A2:W2000").ClearContents
    outRow = 2
    For r = 2 To lastRaw
        If P3_ValidRawRow(wsRaw, r) Then
            For c = 1 To 6
                wsPost.Cells(outRow, c).Value2 = wsRaw.Cells(r, c).Value2
            Next c
            outRow = outRow + 1
            If outRow > 2000 Then Exit For
        End If
    Next r
    mapped = outRow - 2
    If rawValid <> mapped Then
        wsPost.Range("A2:W2000").ClearContents
        messageText = "POSTING_MAP_COUNT_MISMATCH: RAW valid=" & CStr(rawValid) & _
                      "; POSTING mapped=" & CStr(mapped) & ". STOP."
        P3_RebuildPostingStrict = False
        Exit Function
    End If
    messageText = "POSTING MAP PASS: RAW valid=" & CStr(rawValid) & _
                  "; POSTING mapped=" & CStr(mapped) & "; Available=0 retained."
    P3_RebuildPostingStrict = True
    Exit Function
Failed:
    messageText = "POSTING_MAP_FAIL: " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Sub CF_ReadState()
    Dim messageText As String
    If Not P3_ReadState(messageText) Then MsgBox messageText, vbExclamation, "READ STATE"
End Sub

Public Sub CF_BuildDecision()
    Dim messageText As String
    If Not P3_ReadState(messageText) Then
        MsgBox messageText, vbExclamation, "BUILD DECISION"
    ElseIf Not P3_BuildDecision(messageText) Then
        MsgBox messageText, vbExclamation, "BUILD DECISION"
    Else
        MsgBox messageText, vbInformation, "BUILD DECISION"
    End If
End Sub

Public Sub CF_BuildReleasePlan()
    Dim messageText As String
    If Not P3_BuildReleasePlan(messageText) Then MsgBox messageText, vbExclamation, "RELEASE PLAN"
End Sub

Public Sub CF_ValidatePlan()
    Dim messageText As String
    If Not P3_ValidatePlan(messageText) Then
        MsgBox messageText, vbExclamation, "VALIDATE PLAN"
    Else
        MsgBox messageText, vbInformation, "VALIDATE PLAN"
    End If
End Sub

Public Sub CF_Phase3_RunUnitTests()
    Dim ws As Worksheet, r As Long, passed As Long, total As Long
    Dim got As String, expected As String, detailText As String
    Set ws = P3_EnsureSheet(P3_TEST_SHEET, False)
    ws.Cells.ClearContents
    ws.Range("A1:F1").value = Array("Case", "Expected", "Actual", "Status", "Detail", "SAP Write")
    r = 2

    got = P3_Decide("BIW", "B1", "BIN-A", 100, "B1", "BIN-A", 100, 100, 150, 1, True, False, got, detailText)
    P3_TestRow ws, r, "DIRECT_CONFIRM", "CF_ONLY", got, detailText, passed, total
    got = P3_Decide("BIW", "B1", "BIN-A", 100, "B1", "BIN-A", 80, 100, 150, 1, True, False, got, detailText)
    P3_TestRow ws, r, "CHANGE_QTY", "CHANGE_QTY", got, detailText, passed, total
    got = P3_Decide("BIW", "B1", "BIN-A", 100, "B1", "BIN-B", 100, 100, 150, 1, True, False, got, detailText)
    P3_TestRow ws, r, "CHANGE_BIN", "CHANGE_BIN", got, detailText, passed, total
    got = P3_Decide("BIW", "B1", "BIN-A", 100, "B2", "BIN-B", 100, 100, 150, 1, True, False, got, detailText)
    P3_TestRow ws, r, "POST_CHANGE_BATCH_BIN", "POST_CHANGE_BATCH_BIN", got, detailText, passed, total
    got = P3_Decide("BIW", "B1", "BIN-A", 100, "B1", "BIN-A", 100, 0, 150, 1, True, False, got, detailText)
    P3_TestRow ws, r, "AVAILABLE_ZERO_TOTAL_ENOUGH", "NEED_MON", got, detailText, passed, total
    got = P3_Decide("BIW", "B1", "BIN-A", 100, "B1", "BIN-A", 100, 40, 150, 1, True, False, got, detailText)
    P3_TestRow ws, r, "ONE_OTHER_WT", "NEED_MON", got, detailText, passed, total
    got = P3_Decide("BIW", "B1", "BIN-A", 100, "B1", "BIN-A", 100, 40, 150, 1, True, False, got, detailText)
    P3_TestRow ws, r, "MULTI_OTHER_WT", "NEED_MON", got, detailText, passed, total
    got = P3_Decide("BIW", "B1", "BIN-A", 100, "B1", "BIN-A", 100, 20, 80, 1, True, False, got, detailText)
    P3_TestRow ws, r, "NOT_ENOUGH_STOCK", "NOT_ENOUGH_STOCK", got, detailText, passed, total
    got = P3_Decide("BIW", "B1", "BIN-A", 100, "B1", "BIN-A", 100, 100, 150, 2, True, False, got, detailText)
    P3_TestRow ws, r, "DUPLICATE_SOURCE", "REVIEW_STOP", got, detailText, passed, total
    got = P3_Decide("BIW", "B1", "", 100, "B1", "", 100, 100, 150, 2, True, False, got, detailText)
    P3_TestRow ws, r, "MULTI_BIN_AMBIGUOUS", "REVIEW_STOP", got, detailText, passed, total
    got = P3_Decide("BIW", "B1", "BIN-A", 100, "B1", "BIN-A", 100, 100, 150, 1, False, True, got, detailText)
    P3_TestRow ws, r, "MISSING_MAP", "REVIEW_STOP", got, detailText, passed, total

    ws.columns("A:F").AutoFit
    ws.Range("F2:F" & CStr(r - 1)).value = "0"
    ws.Range("H1").Value2 = "UNIT TESTS: " & CStr(passed) & "/" & CStr(total) & _
                             " PASS. SAP Write=0."
End Sub

Private Sub P3_TestRow(ByVal ws As Worksheet, ByRef r As Long, ByVal caseName As String, _
                       ByVal expected As String, ByVal actual As String, ByVal detailText As String, _
                       ByRef passed As Long, ByRef total As Long)
    total = total + 1
    ws.Cells(r, 1).Value2 = caseName
    ws.Cells(r, 2).Value2 = expected
    ws.Cells(r, 3).Value2 = actual
    ws.Cells(r, 4).Value2 = IIf(expected = actual, "PASS", "FAIL")
    ws.Cells(r, 5).Value2 = detailText
    ws.Cells(r, 6).Value2 = "0"
    If expected = actual Then passed = passed + 1
    r = r + 1
End Sub

Private Function P3_Decide(ByVal material As String, ByVal sysBatch As String, _
                           ByVal sysBin As String, ByVal sysQty As Double, _
                           ByVal actualBatch As String, ByVal actualBin As String, _
                           ByVal actualQty As Double, ByVal available As Double, _
                           ByVal totalStock As Double, ByVal candidateCount As Long, _
                           ByVal mapOK As Boolean, ByVal missingMap As Boolean, _
                           ByRef got As String, ByRef detailText As String) As String
    Dim qtyDiff As Boolean, binDiff As Boolean, batchDiff As Boolean
    If Not mapOK Or missingMap Then
        got = "REVIEW_STOP": detailText = "MAP_FAIL_CLOSED": P3_Decide = got: Exit Function
    End If
    If candidateCount <> 1 Then
        got = "REVIEW_STOP": detailText = "SOURCE_CANDIDATES=" & CStr(candidateCount): P3_Decide = got: Exit Function
    End If
    If actualQty <= 0# Then
        got = "REVIEW_STOP": detailText = "ACTUAL_QTY_INVALID": P3_Decide = got: Exit Function
    End If
    If totalStock < actualQty Then
        got = "NOT_ENOUGH_STOCK": detailText = "TOTAL_STOCK=" & CStr(totalStock): P3_Decide = got: Exit Function
    End If
    qtyDiff = (Abs(actualQty - sysQty) > 0.000001)
    binDiff = (UCase$(Trim$(actualBin)) <> UCase$(Trim$(sysBin)))
    batchDiff = (UCase$(Trim$(actualBatch)) <> UCase$(Trim$(sysBatch)))
    If available < actualQty Then
        got = "NEED_MON": detailText = "AVAILABLE=" & CStr(available) & "; TOTAL=" & CStr(totalStock)
    ElseIf batchDiff Or binDiff Then
        If batchDiff And binDiff Then
            got = "POST_CHANGE_BATCH_BIN"
        ElseIf binDiff Then
            got = "CHANGE_BIN"
        Else
            got = "POST_CHANGE_BATCH_BIN"
        End If
        detailText = "POST_STEP_REQUIRED"
    ElseIf qtyDiff Then
        got = "CHANGE_QTY": detailText = "QTY_DIFF"
    Else
        got = "CF_ONLY": detailText = "DIRECT_MATCH"
    End If
    P3_Decide = got
End Function

Private Function P3_ReadState(ByRef messageText As String) As Boolean
    Dim wsRaw As Worksheet, wsPost As Worksheet
    Dim rawValid As Long, mapped As Long
    On Error GoTo Failed
    Set wsRaw = ThisWorkbook.Worksheets(P3_RAW_SHEET)
    Set wsPost = ThisWorkbook.Worksheets(P3_POST_SHEET)
    rawValid = P3_CountValidRaw(wsRaw)
    mapped = P3_CountMappedPosting(wsPost)
    If rawValid <> mapped Then
        messageText = "POSTING_MAP_COUNT_MISMATCH: RAW valid=" & CStr(rawValid) & _
                      "; POSTING mapped=" & CStr(mapped) & ". STOP."
        P3_ReadState = False
        Exit Function
    End If
    If rawValid = 0 Then
        messageText = "POSTING_MAP_EMPTY: no valid Product/Batch/Bin/Total Stock rows. STOP."
        P3_ReadState = False
        Exit Function
    End If
    messageText = "STATE PASS: RAW valid=" & CStr(rawValid) & "; POSTING mapped=" & CStr(mapped)
    P3_ReadState = True
    Exit Function
Failed:
    messageText = "READ_STATE_FAIL: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function P3_CountValidRaw(ByVal ws As Worksheet) As Long
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If P3_ValidRawRow(ws, r) Then P3_CountValidRaw = P3_CountValidRaw + 1
    Next r
End Function

Private Function P3_CountMappedPosting(ByVal ws As Worksheet) As Long
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If P3_Text(ws.Cells(r, 1).Value2) <> "" Then P3_CountMappedPosting = P3_CountMappedPosting + 1
    Next r
End Function

Private Function P3_ValidRawRow(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim material As String, batch As String, binText As String
    material = P3_Text(ws.Cells(r, 1).Value2)
    batch = P3_Text(ws.Cells(r, 3).Value2)
    binText = P3_Text(ws.Cells(r, 4).Value2)
    P3_ValidRawRow = (material <> "" And batch <> "" And binText <> "" And _
                      P3_NumberState(ws.Cells(r, 2).Value2) = "VALUE" And _
                      P3_NumberState(ws.Cells(r, 6).Value2) = "VALUE" And _
                      P3_Number(ws.Cells(r, 6).Value2) >= 0#)
End Function

Private Function P3_BuildDecision(ByRef messageText As String) As Boolean
    Dim wsWH As Worksheet, wsPost As Worksheet, wsPlan As Worksheet
    Dim lastWH As Long, lastPost As Long, whRow As Long, postRow As Long, planRow As Long
    Dim material As String, sysBatch As String, sysBin As String, actualBatch As String, actualBin As String
    Dim areaText As String, ob As String, doc As String, wt As String, whNo As String, statusText As String
    Dim qtyState As String, sysQtyState As String, availableState As String, totalState As String
    Dim actualQty As Double, sysQty As Double, available As Double, totalStock As Double
    Dim candidateCount As Long, firstCandidate As Long, areaCount As Long
    Dim resultText As String, actionText As String, noteText As String, targetBin As String
    Dim qtyDiff As Boolean, binDiff As Boolean, batchDiff As Boolean
    Dim rawValid As Long, mapped As Long
    On Error GoTo Failed
    Set wsWH = ThisWorkbook.Worksheets(P3_WH_SHEET)
    Set wsPost = ThisWorkbook.Worksheets(P3_POST_SHEET)
    If Not AP_NormalizeAndValidateAll(wsWH, noteText) Then Err.Raise vbObjectError + 9850, , noteText
    Set wsPlan = P3_EnsureSheet(P3_PLAN_SHEET, False)
    wsPlan.Cells.ClearContents
    wsPlan.Range("A1:T1").value = Array("Source Row", "OB", "Document", "WH", "WT", "Material", "Sys Batch", "Sys Bin", "Sys Qty", "Actual Qty", "Actual Bin", "Actual Batch", "Posting Available", "Total Stock", "Projected Available", "Decision", "Release Status", "Release WT List", "Action", "Note")
    wsPost.Range("J2:W2000").ClearContents
    lastWH = wsWH.Cells(wsWH.rows.count, 6).End(xlUp).row
    lastPost = wsPost.Cells(wsPost.rows.count, 1).End(xlUp).row
    planRow = 2
    For whRow = 5 To lastWH
        ob = P3_Text(wsWH.Cells(whRow, 1).Value2)
        If P3_TargetOB <> "" And StrComp(ob, P3_TargetOB, vbTextCompare) <> 0 Then GoTo P3_NextDecisionRow
        material = P3_Text(wsWH.Cells(whRow, 6).Value2)
        qtyState = P3_NumberState(wsWH.Cells(whRow, 12).Value2)
        actualBin = AP_ResolvedBin(wsWH, whRow, P3_Text(wsWH.Cells(whRow, 8).Value2))
        areaText = P3_Text(wsWH.Cells(whRow, 21).Value2)
        actualBatch = P3_Text(wsWH.Cells(whRow, 22).Value2)
        If material <> "" And (qtyState <> "BLANK" Or actualBin <> "" Or actualBatch <> "" Or areaText <> "") Then
            ob = P3_Text(wsWH.Cells(whRow, 1).Value2)
            doc = P3_Text(wsWH.Cells(whRow, 2).Value2)
            whNo = P3_Text(wsWH.Cells(whRow, 4).Value2)
            wt = P3_Text(wsWH.Cells(whRow, 5).Value2)
            sysBatch = P3_Text(wsWH.Cells(whRow, 9).Value2)
            sysBin = P3_Text(wsWH.Cells(whRow, 8).Value2)
            sysQtyState = P3_NumberState(wsWH.Cells(whRow, 7).Value2)
            If actualBatch = "" Then actualBatch = sysBatch
            resultText = "REVIEW_STOP": actionText = "REVIEW_STOP": noteText = ""
            targetBin = actualBin
            candidateCount = 0: firstCandidate = 0: areaCount = 0
            If qtyState <> "VALUE" Or P3_Number(wsWH.Cells(whRow, 12).Value2) <= 0# Then
                noteText = "ACTUAL_QTY_BLANK_OR_ERROR"
            ElseIf actualBatch = "" Then
                noteText = "BATCH_BLANK"
            Else
                actualQty = P3_Number(wsWH.Cells(whRow, 12).Value2)
                For postRow = 2 To lastPost
                    If StrComp(P3_Text(wsPost.Cells(postRow, 1).Value2), material, vbTextCompare) = 0 And _
                       StrComp(P3_Text(wsPost.Cells(postRow, 3).Value2), actualBatch, vbTextCompare) = 0 Then
                        If targetBin <> "" Then
                            If StrComp(P3_Text(wsPost.Cells(postRow, 4).Value2), targetBin, vbTextCompare) = 0 Then
                                candidateCount = candidateCount + 1
                                If firstCandidate = 0 Then firstCandidate = postRow
                            End If
                        ElseIf P3_BinInArea(P3_Text(wsPost.Cells(postRow, 4).Value2), areaText) Then
                            areaCount = areaCount + 1
                            candidateCount = candidateCount + 1
                            If firstCandidate = 0 Then firstCandidate = postRow
                        End If
                    End If
                Next postRow
                If targetBin = "" And candidateCount = 1 Then targetBin = P3_Text(wsPost.Cells(firstCandidate, 4).Value2)
                If candidateCount <> 1 Then
                    noteText = "TARGET_SOURCE_CANDIDATES=" & CStr(candidateCount)
                ElseIf sysQtyState <> "VALUE" Then
                    noteText = "SYS_QTY_ERROR"
                ElseIf P3_NumberState(wsPost.Cells(firstCandidate, 2).Value2) <> "VALUE" Or _
                       P3_NumberState(wsPost.Cells(firstCandidate, 6).Value2) <> "VALUE" Then
                    noteText = "POSTING_AVAILABLE_OR_TOTAL_ERROR"
                Else
                    sysQty = P3_Number(wsWH.Cells(whRow, 7).Value2)
                    available = P3_Number(wsPost.Cells(firstCandidate, 2).Value2)
                    totalStock = P3_Number(wsPost.Cells(firstCandidate, 6).Value2)
                    qtyDiff = (Abs(actualQty - sysQty) > 0.000001)
                    binDiff = (UCase$(targetBin) <> UCase$(sysBin))
                    batchDiff = (UCase$(actualBatch) <> UCase$(sysBatch))
                    resultText = P3_Decide(material, sysBatch, sysBin, sysQty, actualBatch, targetBin, actualQty, available, totalStock, 1, True, False, resultText, noteText)
                    If resultText = "CF_ONLY" Then actionText = "CF_ONLY"
                    If resultText = "CHANGE_QTY" Then actionText = "CHANGE_QTY"
                    If resultText = "CHANGE_BIN" Then actionText = "CHANGE_BIN"
                    If resultText = "POST_CHANGE_BATCH_BIN" Then actionText = "POST_CHANGE_BATCH_BIN"
                    If resultText = "NEED_MON" Then actionText = "NEED_MON"
                    If resultText = "NOT_ENOUGH_STOCK" Then actionText = "REVIEW_STOP"
                    noteText = noteText & "; Match=Product+Batch+SourceBin; Available=" & CStr(available) & "; TotalStock=" & CStr(totalStock)
                    If resultText = "NEED_MON" Then noteText = noteText & "; Cancel not executed"
                    wsPost.Cells(firstCandidate, 10).Value2 = ob
                    wsPost.Cells(firstCandidate, 11).Value2 = whNo
                    wsPost.Cells(firstCandidate, 12).Value2 = sysQty
                    wsPost.Cells(firstCandidate, 13).Value2 = actualQty
                    wsPost.Cells(firstCandidate, 14).Value2 = sysBin
                    wsPost.Cells(firstCandidate, 15).Value2 = targetBin
                    wsPost.Cells(firstCandidate, 16).Value2 = sysBatch
                    wsPost.Cells(firstCandidate, 17).Value2 = actualBatch
                    wsPost.Cells(firstCandidate, 18).Value2 = actualQty
                    wsPost.Cells(firstCandidate, 19).Value2 = available
                    wsPost.Cells(firstCandidate, 20).Value2 = IIf(resultText = "NEED_MON", "YES", "NO")
                    wsPost.Cells(firstCandidate, 21).Value2 = actionText
                    wsPost.Cells(firstCandidate, 22).Value2 = resultText
                    wsPost.Cells(firstCandidate, 23).Value2 = noteText
                End If
            End If
            wsPlan.Cells(planRow, 1).Value2 = whRow
            wsPlan.Cells(planRow, 2).Value2 = ob
            wsPlan.Cells(planRow, 3).Value2 = doc
            wsPlan.Cells(planRow, 4).Value2 = whNo
            wsPlan.Cells(planRow, 5).Value2 = wt
            wsPlan.Cells(planRow, 6).Value2 = material
            wsPlan.Cells(planRow, 7).Value2 = sysBatch
            wsPlan.Cells(planRow, 8).Value2 = sysBin
            wsPlan.Cells(planRow, 9).Value2 = IIf(sysQtyState = "VALUE", P3_Number(wsWH.Cells(whRow, 7).Value2), "")
            wsPlan.Cells(planRow, 10).Value2 = IIf(qtyState = "VALUE", P3_Number(wsWH.Cells(whRow, 12).Value2), "")
            wsPlan.Cells(planRow, 11).Value2 = targetBin
            wsPlan.Cells(planRow, 12).Value2 = actualBatch
            If firstCandidate > 0 Then
                wsPlan.Cells(planRow, 13).Value2 = P3_Number(wsPost.Cells(firstCandidate, 2).Value2)
                wsPlan.Cells(planRow, 14).Value2 = P3_Number(wsPost.Cells(firstCandidate, 6).Value2)
            End If
            wsPlan.Cells(planRow, 16).Value2 = resultText
            wsPlan.Cells(planRow, 19).Value2 = actionText
            wsPlan.Cells(planRow, 20).Value2 = noteText
            wsWH.Cells(whRow, 15).Value2 = resultText
            wsWH.Cells(whRow, 17).Value2 = "CHECKED"
            wsWH.Cells(whRow, 18).Value2 = actionText
            wsWH.Cells(whRow, 19).Value2 = noteText
            planRow = planRow + 1
        End If
P3_NextDecisionRow:
    Next whRow
    wsPlan.columns("A:T").AutoFit
    messageText = "DECISION PASS: rows=" & CStr(planRow - 2)
    P3_BuildDecision = True
    Exit Function
Failed:
    messageText = "BUILD_DECISION_FAIL: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function P3_BuildReleasePlan(ByRef messageText As String) As Boolean
    Dim wsPlan As Worksheet, wsWH As Worksheet
    Dim lastPlan As Long, lastWH As Long, r As Long, q As Long
    Dim decisionText As String, material As String, batchText As String, binText As String
    Dim currentOB As String, currentDoc As String, currentWT As String, currentWH As String
    Dim wt As String, ob As String, doc As String, statusText As String
    Dim qty As Double, releaseQty As Double, projected As Double
    Dim wtList As String, keyCount As Long
    On Error GoTo Failed
    Set wsPlan = ThisWorkbook.Worksheets(P3_PLAN_SHEET)
    Set wsWH = ThisWorkbook.Worksheets(P3_WH_SHEET)
    lastPlan = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    lastWH = wsWH.Cells(wsWH.rows.count, 6).End(xlUp).row
    For r = 2 To lastPlan
        decisionText = P3_Text(wsPlan.Cells(r, 16).Value2)
        If decisionText = "NEED_MON" Then
            material = P3_Text(wsPlan.Cells(r, 6).Value2)
            batchText = P3_Text(wsPlan.Cells(r, 12).Value2)
            binText = P3_Text(wsPlan.Cells(r, 11).Value2)
            currentOB = P3_Text(wsPlan.Cells(r, 2).Value2)
            currentDoc = P3_Text(wsPlan.Cells(r, 3).Value2)
            currentWH = P3_Text(wsPlan.Cells(r, 4).Value2)
            currentWT = P3_Text(wsPlan.Cells(r, 5).Value2)
            releaseQty = 0#: wtList = "": keyCount = 0
            For q = 5 To lastWH
                ob = P3_Text(wsWH.Cells(q, 1).Value2)
                doc = P3_Text(wsWH.Cells(q, 2).Value2)
                wt = P3_Text(wsWH.Cells(q, 5).Value2)
                statusText = P3_Text(wsWH.Cells(q, 10).Value2)
                If P3_Text(wsWH.Cells(q, 6).Value2) = material And _
                   P3_Text(wsWH.Cells(q, 9).Value2) = batchText And _
                   P3_Text(wsWH.Cells(q, 8).Value2) = binText And _
                   P3_Text(wsWH.Cells(q, 4).Value2) <> currentWH And _
                   ob <> currentOB And doc <> currentDoc And wt <> currentWT And _
                   P3_CancelableStatus(statusText) And P3_NumberState(wsWH.Cells(q, 7).Value2) = "VALUE" Then
                    qty = P3_Number(wsWH.Cells(q, 7).Value2)
                    If qty > 0# Then
                        releaseQty = releaseQty + qty
                        keyCount = keyCount + 1
                        If wtList <> "" Then wtList = wtList & ","
                        wtList = wtList & wt
                    End If
                End If
            Next q
            projected = P3_Number(wsPlan.Cells(r, 13).Value2) + releaseQty
            wsPlan.Cells(r, 15).Value2 = projected
            wsPlan.Cells(r, 17).Value2 = IIf(keyCount > 0, IIf(projected >= P3_Number(wsPlan.Cells(r, 10).Value2), "PROJECTED_ENOUGH_AFTER_CANCEL", "PROJECTED_STILL_SHORT"), "NO_EXACT_OTHER_WT")
            wsPlan.Cells(r, 18).Value2 = wtList
            wsPlan.Cells(r, 20).Value2 = P3_Text(wsPlan.Cells(r, 20).Value2) & "; OtherWTQty=" & CStr(releaseQty) & "; Cancel not executed"
        End If
    Next r
    messageText = "RELEASE PLAN PASS: no Cancel/Write executed"
    P3_BuildReleasePlan = True
    Exit Function
Failed:
    messageText = "BUILD_RELEASE_PLAN_FAIL: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function P3_ValidatePlan(ByRef messageText As String) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long, decisionText As String
    Dim bad As Long, writeCount As Long
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets(P3_PLAN_SHEET)
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        decisionText = P3_Text(ws.Cells(r, 16).Value2)
        If decisionText <> "CF_ONLY" And decisionText <> "CHANGE_QTY" And _
           decisionText <> "CHANGE_BIN" And decisionText <> "POST_CHANGE_BATCH_BIN" And _
           decisionText <> "NEED_MON" And decisionText <> "NOT_ENOUGH_STOCK" And _
           decisionText <> "REVIEW_STOP" Then bad = bad + 1
        If InStr(1, P3_Text(ws.Cells(r, 20).Value2), "Cancel executed", vbTextCompare) > 0 Then writeCount = writeCount + 1
    Next r
    If bad > 0 Or writeCount > 0 Then
        messageText = "PLAN_VALIDATE_FAIL: bad decisions=" & CStr(bad) & "; SAP actions=" & CStr(writeCount)
        P3_ValidatePlan = False
    Else
        messageText = "PLAN PASS: rows=" & CStr(lastRow - 1) & "; SAP Write/Cancel/Save/CF=0"
        P3_ValidatePlan = True
    End If
    Exit Function
Failed:
    messageText = "VALIDATE_PLAN_FAIL: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function P3_CancelableStatus(ByVal statusText As String) As Boolean
    Dim s As String
    s = UCase$(Trim$(statusText))
    If s = "" Or s = "CANCELLED" Or s = "CANCELED" Or s = "DELETED" Or _
       s = "COMPLETED" Or s = "CLOSED" Or s = "FINISHED" Or s = "POSTED" Then
        P3_CancelableStatus = False
    Else
        P3_CancelableStatus = True
    End If
End Function

Private Function P3_EnsureSheet(ByVal sheetName As String, ByVal veryHidden As Boolean) As Worksheet
    On Error Resume Next
    Set P3_EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If P3_EnsureSheet Is Nothing Then
        Set P3_EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        P3_EnsureSheet.Name = sheetName
    End If
    If veryHidden Then P3_EnsureSheet.Visible = xlSheetVeryHidden
End Function

Private Function P3_Text(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then
        P3_Text = ""
    Else
        P3_Text = Trim$(CStr(valueIn))
    End If
End Function

Private Function P3_NumberState(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Or P3_Text(valueIn) = "" Then
        P3_NumberState = "BLANK"
    ElseIf IsNumeric(valueIn) Then
        P3_NumberState = "VALUE"
    Else
        P3_NumberState = "ERROR"
    End If
End Function

Private Function P3_Number(ByVal valueIn As Variant) As Double
    If P3_NumberState(valueIn) = "VALUE" Then P3_Number = CDbl(valueIn) Else P3_Number = 0#
End Function

Private Function P3_BinInArea(ByVal binText As String, ByVal areaText As String) As Boolean
    Dim a As String, b As String
    a = UCase$(Trim$(areaText))
    b = UCase$(Trim$(binText))
    If a = "P" Then a = "APFL-P"
    If a = "ZF" Then a = "APFL-ZF"
    If a = "HR" Then a = "APFL-HR"
    P3_BinInArea = (a <> "" And b <> "" And Left$(b, Len(a)) = a)
End Function

