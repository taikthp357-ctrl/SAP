Attribute VB_Name = "modCFX_Gates"
Option Explicit

' ============================================================================
' modCFX_Gates  -  the MINIMAL safety-gate set (spec section 14).
'
'   Three layers only:
'     1) INPUT GATE     - exact input present & unambiguous before any write.
'     2) ACTION GATE    - immediately-before-write precondition, read from RAM
'                         truth (not a re-read of a possibly-stale live grid).
'     3) POST-VERIFY     - after a write, confirm the intended effect once.
'
'   Historical regression metadata is an OFFLINE BUILD GATE, never a runtime
'   blocker (spec + ACCEPTANCE failure policy).  Removed from the live path.
'
'   The active-WT gate here REPLACES the defect in
'   CWR_BatchAssertNoActiveOutboundWT, whose STOP on a cancelled WT that read
'   back blank on a stale MON grid caused the WT 5042889690 false block.
' ============================================================================

' ---- INPUT GATE -------------------------------------------------------------
' Every member must carry exact, unambiguous targets before the plan locks.
Public Function CFX_InputGate(ByVal groupID As String, ByRef reason As String) As Boolean
    Dim g As Object, members As Collection, m As Object, i As Long
    Set g = CFX_Group(groupID)
    Set members = g("Members")
    For i = 1 To members.count
        Set m = members(i)
        If CStr(m("Material")) = "" Or CStr(m("Item")) = "" Then
            reason = "STOP_INPUT_INCOMPLETE|GROUP=" & groupID & "|PLAN=" & CStr(m("PlanID"))
            CFX_InputGate = False: Exit Function
        End If
        ' A target bin entered => must be an exact, resolvable bin (any area
        ' family that SAP/stock proves valid; no P/ZF/HR hard rejection - S10).
        If CStr(m("TargetBin")) <> "" And Not CFX_IsExactBin(CStr(m("TargetBin"))) Then
            reason = "STOP_INPUT_AMBIGUOUS_BIN|GROUP=" & groupID & "|BIN=" & CStr(m("TargetBin"))
            CFX_InputGate = False: Exit Function
        End If
    Next i
    CFX_InputGate = True
End Function

' Exact-bin test: a fully qualified bin string (has the AREA-NN-NN-NN shape
' under the configured root) is accepted regardless of area family, so exact
' bins like APFL-CC01-01-01 are NOT rejected merely for not being P/ZF/HR.
Public Function CFX_IsExactBin(ByVal bin As String) As Boolean
    Dim s As String
    s = UCase$(Trim$(bin))
    If s = "" Then Exit Function
    ' Must contain the root separator and at least one structured segment.
    If InStr(s, "-") = 0 Then Exit Function
    CFX_IsExactBin = (Len(s) >= 5)
End Function

' ---- ACTION GATE: no *genuinely* active conflicting outbound WT -------------
' RAM-truth first.  A WT that this run drove terminal (A/C) is never "active",
' regardless of what a stale grid row shows.  Only a FRESH contradiction (RAM
' says terminal but a controlled single refresh still shows it open) is allowed
' to stop - and even then it is a fresh-state contradiction, not stale cache.
Public Function CFX_ActionGateNoActiveConflict(ByVal groupID As String, _
        ByVal docNo As String, ByRef reason As String) As Boolean
    Dim g As Object, members As Collection, m As Object, i As Long
    Dim grid As Collection, attempt As Long, cleanPass As Boolean
    Dim oldWT As String, localReason As String
    Const MAX_ATTEMPTS As Long = 2      ' initial read + at most one refresh (S21)
    Set g = CFX_Group(groupID)
    Set members = g("Members")

    For attempt = 1 To MAX_ATTEMPTS
        If attempt = 1 Then
            Set grid = CFX_ReadWTGrid(docNo)
        Else
            Set grid = CFX_RefreshWTGrid(docNo)   ' one controlled refresh
        End If
        cleanPass = True
        For i = 1 To members.count
            Set m = members(i)
            If Not CFX_MemberConflictClear(m, grid, docNo, localReason) Then
                oldWT = CStr(m("WTNumber"))
                If oldWT <> "" And CFX_WTIsTerminal(oldWT) Then
                    ' RAM truth: we drove this WT terminal.  A grid row that
                    ' shows it open is stale.  Refresh once, then re-judge.
                    If attempt < MAX_ATTEMPTS Then
                        CFX_Log "WT_MAP_STALE_REFRESH", "WARN", _
                            "WT=" & oldWT & "|RAM=" & CFX_WTState(oldWT) & _
                            "|GRID=ACTIVE|GROUP=" & groupID
                        CFX_MetaInc "Warns"
                        cleanPass = False
                        Exit For          ' break member scan, go refresh
                    Else
                        ' Fresh grid STILL contradicts RAM terminal => a real
                        ' fresh-state contradiction (an allowed hard stop).
                        reason = "STOP_FRESH_STATE_CONTRADICTION|WT=" & oldWT & _
                            "|RAM=" & CFX_WTState(oldWT) & _
                            "|GRID=ACTIVE_AFTER_REFRESH|GROUP=" & groupID
                        CFX_ActionGateNoActiveConflict = False
                        Exit Function
                    End If
                Else
                    ' RAM does not consider it terminal => genuine unexpected
                    ' active peer WT.  Hard stop (fail-closed).
                    reason = "STOP_ACTIVE_OUTBOUND_WT_PRESENT|" & localReason & _
                        "|GROUP=" & groupID
                    CFX_ActionGateNoActiveConflict = False
                    Exit Function
                End If
            End If
        Next i
        If cleanPass Then
            CFX_ActionGateNoActiveConflict = True
            Exit Function
        End If
    Next attempt
    ' Exhausted attempts without a clean pass and without an explicit stop:
    ' treat as fresh-state contradiction (fail-closed).
    reason = "STOP_FRESH_STATE_CONTRADICTION|GROUP=" & groupID & "|ATTEMPTS=" & CStr(MAX_ATTEMPTS)
    CFX_ActionGateNoActiveConflict = False
End Function

' Returns True if the grid holds NO conflicting-active row for this member.
' A conflicting row = same item+material, status blank (open), whose WT equals
' the member's OldWT - EXCEPT when the plan intends to RETAIN that WT.  A WT is
' retained by any non-cancelling ExecMode: NONE (CF-only, S14) leaves it active,
' and KEEP_EXISTING_CREATE_DELTA (delta-add, S02) leaves it active by design.
' Only REPLACE_CANCEL_CREATE / CANCEL_PRDO_ZERO expect it to have gone terminal,
' so for those an open old-WT row is a genuine (possibly stale) conflict.
Private Function CFX_MemberConflictClear(ByVal m As Object, ByVal grid As Collection, _
        ByVal docNo As String, ByRef reason As String) As Boolean
    Dim row As Object, i As Long, rowWT As String, rowStat As String, retainKeep As Boolean, em As String
    CFX_MemberConflictClear = True
    em = UCase$(Trim$(CStr(m("ExecMode"))))
    retainKeep = (em <> "REPLACE_CANCEL_CREATE" And em <> "CANCEL_PRDO_ZERO")
    For i = 1 To grid.count
        Set row = grid(i)
        rowStat = CStr(row(CFX_K_TOSTAT))
        rowWT = CStr(row(CFX_K_TANUM))
        If rowWT <> "" And Trim$(rowStat) = "" _
           And CFX_Eq(CStr(row(CFX_K_ITEMNO)), CStr(m("Item"))) _
           And CFX_Eq(CStr(row(CFX_K_MATNR)), CStr(m("Material"))) Then
            If CFX_Eq(rowWT, CStr(m("WTNumber"))) Then
                If Not retainKeep Then
                    reason = "CONFLICT_OLD_WT_OPEN|WT=" & rowWT & "|ITEM=" & CStr(m("Item"))
                    CFX_MemberConflictClear = False
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' ---- POST-VERIFY: confirm one intended effect after a write -----------------
Public Function CFX_PostVerifyPrdoQty(ByVal docNo As String, ByVal item As String, _
        ByVal expectedQty As Double, ByRef reason As String) As Boolean
    Dim actual As Double
    actual = CFX_ReadPrdoQty(docNo, item)
    If Abs(actual - expectedQty) > 0.000001 Then
        reason = "STOP_POSTVERIFY_PRDO_QTY|DOC=" & docNo & "|ITEM=" & item & _
            "|EXPECTED=" & CStr(expectedQty) & "|ACTUAL=" & CStr(actual)
        CFX_PostVerifyPrdoQty = False
        Exit Function
    End If
    CFX_PostVerifyPrdoQty = True
End Function

' Verify Create delta via WHTask OpenQty (spec: not by reopening MON/OB).
' OpenQty = committed PRDO target total - already-active WT qty for the document
' (i.e. the quantity still awaiting a Create).  Before the Create phase this is
' the delta to allocate (S02: 600-580=20); after it, zero.
Public Function CFX_PostVerifyOpenQty(ByVal groupID As String, ByVal docNo As String, _
        ByVal expectedOpenQty As Double, ByRef reason As String) As Boolean
    Dim grid As Collection, row As Object, i As Long, activeQty As Double, openQty As Double
    If Not CFX_ReturnToWarehouseTask(reason) Then
        CFX_PostVerifyOpenQty = False: Exit Function
    End If
    Set grid = CFX_ReadWTGrid(docNo)
    For i = 1 To grid.count
        Set row = grid(i)
        If Trim$(CStr(row(CFX_K_TOSTAT))) = "" Then _
            activeQty = activeQty + CFX_Num(row(CFX_K_VSOLM))   ' blank status = active
    Next i
    openQty = CFX_ReadCommittedPrdoTotal(docNo) - activeQty
    If openQty < 0 Then openQty = 0#
    ' expectedOpenQty < 0 means "do not assert a specific value" (e.g. CF-only).
    If expectedOpenQty >= 0 And Abs(openQty - expectedOpenQty) > 0.000001 Then
        reason = "STOP_POSTVERIFY_OPEN_QTY|GROUP=" & groupID & _
            "|EXPECTED=" & CStr(expectedOpenQty) & "|ACTUAL=" & CStr(openQty)
        CFX_PostVerifyOpenQty = False
        Exit Function
    End If
    CFX_PostVerifyOpenQty = True
End Function

' ---- shared tiny helpers ----------------------------------------------------
Private Function CFX_Eq(ByVal a As String, ByVal b As String) As Boolean
    CFX_Eq = (StrComp(Trim$(a), Trim$(b), vbTextCompare) = 0)
End Function
Private Function CFX_Num(ByVal v As Variant) As Double
    On Error Resume Next
    CFX_Num = CDbl(v)
    On Error GoTo 0
End Function
