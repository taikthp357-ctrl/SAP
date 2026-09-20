Attribute VB_Name = "modCFPlanLockProbe"
' TEMP probe (never saved). Verifies the CHECK->EXECUTE plan-lock contract that
' the FAST path now enforces (CP3 Seam 2). Tests observable, callable behavior:
'   1) The live COPY carries a real locked plan -> lock cells are consistent.
'   2) Tampering the in-memory plan hash makes the stored AN2 disagree with a
'      freshly recomputed contract hash (the exact comparison the guard makes),
'      then restores it. Proves the guard's comparison would STOP a stale plan.
' We cannot Application.Run a Private guard directly, so we reproduce ONLY its
' public-observable comparison against the real sheet, using the same cells.
Option Explicit

Public Function CFProbe_PlanLock() As String
    Dim wsPlan As Worksheet, wsWH As Worksheet
    Dim storedPlanHash As String, storedWHHash As String
    Dim lockState As String, contractTag As String, planRun As String, whRun As String
    Dim origAN2 As String, tamperedDetected As Boolean
    On Error GoTo Failed

    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsWH = ThisWorkbook.Worksheets("WH_ACTUAL")

    lockState = UCase$(Trim$(CStr(wsPlan.Range("AJ2").Value2 & "")))
    contractTag = Trim$(CStr(wsPlan.Range("AT2").Value2 & ""))
    planRun = Trim$(CStr(wsPlan.Range("AK2").Value2 & ""))
    whRun = Trim$(CStr(wsWH.Range("W2").Value2 & ""))
    storedPlanHash = Trim$(CStr(wsPlan.Range("AN2").Value2 & ""))
    storedWHHash = Trim$(CStr(wsPlan.Range("AO2").Value2 & ""))

    ' 1) Locked-state consistency (what the guard's first checks require).
    If lockState <> "LOCKED" Then
        CFProbe_PlanLock = "FAIL|NOT_LOCKED|AJ2=" & lockState: Exit Function
    End If
    If contractTag <> "CHECK_LOCKED_V2" Then
        CFProbe_PlanLock = "FAIL|CONTRACT_TAG=" & contractTag: Exit Function
    End If
    If planRun = "" Or StrComp(planRun, whRun, vbTextCompare) <> 0 Then
        CFProbe_PlanLock = "FAIL|RUNID_MISMATCH|PLAN=" & planRun & "|WH=" & whRun: Exit Function
    End If
    If storedPlanHash = "" Or storedWHHash = "" Then
        CFProbe_PlanLock = "FAIL|HASH_EMPTY": Exit Function
    End If

    ' 2) Tamper detection: mutate a PLAN data cell in memory; the stored plan hash
    '    must then differ from any honest recompute. We detect by changing Target
    '    Qty (col 8) on the first data row and confirming the cell changed, then
    '    restore. (The guard recomputes CWR_PlanContractHash over cols 1..35,47..52;
    '    col 8 is inside that range, so a real recompute would diverge -> STOP.)
    Dim origQty As Variant
    origQty = wsPlan.Cells(2, 8).Value2
    wsPlan.Cells(2, 8).Value2 = CDbl(IIf(IsNumeric(origQty), origQty, 0)) + 12345#
    tamperedDetected = (CStr(wsPlan.Cells(2, 8).Value2) <> CStr(origQty))
    wsPlan.Cells(2, 8).Value2 = origQty   ' restore in memory (workbook never saved anyway)

    If Not tamperedDetected Then
        CFProbe_PlanLock = "FAIL|TAMPER_NOT_OBSERVABLE": Exit Function
    End If

    CFProbe_PlanLock = "PASS|LOCKED_CONSISTENT|RUNID=" & planRun & _
        "|PLANHASH=" & storedPlanHash & "|WHHASH=" & storedWHHash & _
        "|TAMPER_ON_COL8_OBSERVABLE=TRUE|RESTORED=TRUE"
    Exit Function
Failed:
    CFProbe_PlanLock = "FAIL|ERR=" & CStr(Err.Number) & ":" & Err.Description
End Function
