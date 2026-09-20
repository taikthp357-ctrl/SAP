Option Explicit

' ============================================================================
' modCFX_Executor  -  the deterministic Group State Machine (Stage 3).
'
'   Runs the LOCKED plan in ONE fixed phase order, from RAM truth, through the
'   SAP adapter (LIVE or REPLAY).  It never re-plans and never re-decides: it
'   consumes the CHECK ExecMode verbatim (spec sections 3-8):
'
'     NONE                       -> CF-only / no SAP write
'     CANCEL_PRDO_ZERO           -> cancel old WT, save PRDO 0, no create
'     REPLACE_CANCEL_CREATE      -> cancel old WT, save PRDO target, create full
'     KEEP_EXISTING_CREATE_DELTA -> retain old WT, save PRDO total, create delta
'
'   FIXED PHASE ORDER (never reordered):
'     0. Session roles asserted (MON != PRDO)          [S22]
'     1. Document bound once (DOCUMENT_OPEN_COUNT=1)
'     2. CANCEL phase        (all cancels first)       -> drives WT terminal
'     3. WT-state settle     (one monotonic update)
'     4. ACTION GATE         (RAM-truth no-active-conflict, replaces bug gate)
'     5. PRDO phase          (stage targets, ONE save per document)
'     6. POST-VERIFY qty     (WHTask OpenQty, not MON reopen)
'     7. CREATE phase        (exact bin/batch/qty allocations)
'     8. WHTask verify + CF user handoff (default)
'
'   The whole flow is exercisable offline: with Mode=REPLAY every adapter call
'   mutates the in-memory SAP world, so the 22 acceptance scenarios assert here.
' ============================================================================

' ---- top-level: execute every locked group for the bound OB -----------------
Public Function CFX_ExecuteRun(ByVal obText As String, ByRef reason As String) As Boolean
    Dim gKey As Variant, ok As Boolean
    If Not CFX_HasRun() Then
        reason = "STOP_NO_RUN|CALL_CFX_NewRun_FIRST"
        Exit Function
    End If
    If Not CBool(CFX_Meta()("PlanLocked")) Then
        reason = "STOP_PLAN_NOT_LOCKED|OB=" & obText
        Exit Function
    End If

    ' Phase 0: session roles must be distinct BEFORE any write (S22).
    If Not CFX_AssertSessionRoles(reason) Then
        CFX_MetaInc "HardStops"
        CFX_Log "SESSION_ROLE_ASSERT", "STOP", reason
        Exit Function
    End If

    CFX_MetaSet "Phase", "EXECUTING"
    CFX_Log "OB_EXECUTING", "PASS", _
        "PHASE_ORDER=CANCEL|WT_SETTLE|ACTION_GATE|PRDO_ONE_SAVE|POSTVERIFY|CREATE|WHTASK|CF_HANDOFF|OB=" & obText

    Dim groups As Object
    Set groups = CFX_Groups()
    For Each gKey In groups.keys
        ok = CFX_ExecuteGroup(CStr(gKey), reason)
        If Not ok Then
            CFX_MetaInc "HardStops"
            CFX_MetaSet "Phase", "STOPPED"
            CFX_Log "OB_STOPPED", "STOP", reason & "|GROUP=" & CStr(gKey)
            Exit Function
        End If
    Next gKey

    CFX_MetaSet "Phase", "READY_FOR_CF"
    CFX_Log "OB_READY_FOR_CF", "PASS", "CF=USER_HANDOFF|CF_SENT=FALSE|OB=" & obText
    CFX_ExecuteRun = True
End Function

' ============================================================================
' PER-GROUP fixed-phase machine.  One group == one document's qty target.
' ============================================================================
Private Function CFX_ExecuteGroup(ByVal groupID As String, ByRef reason As String) As Boolean
    Dim g As Object, members As Collection, docNo As String
    Set g = CFX_Group(groupID)
    Set members = g("Members")
    docNo = CStr(g("DocNo"))

    ' Phase 1: bind the document exactly once for this group.
    If Not CFX_OpenDocument(docNo, reason) Then
        reason = "STOP_DOCUMENT_OPEN|" & reason & "|GROUP=" & groupID
        Exit Function
    End If

    ' Phase 2: CANCEL phase - every member that the plan says to cancel, first.
    If Not CFX_PhaseCancel(g, members, docNo, reason) Then Exit Function

    ' Phase 3: WT-state settle - one monotonic update per cancelled WT.
    If Not CFX_PhaseSettleWT(members, reason) Then Exit Function

    ' Phase 4: ACTION GATE - RAM-truth no-active-conflict (replaces bug gate).
    '          Only reached when a create/qty write is pending for this group.
    If CFX_GroupHasWrite(g, members) Then
        If Not CFX_ActionGateNoActiveConflict(groupID, docNo, reason) Then
            Exit Function
        End If
    End If

    ' Phase 5-6: PRDO one-save + post-verify (only if any qty target changes).
    If Not CFX_PhasePrdo(g, members, docNo, reason) Then Exit Function

    ' Phase 7: CREATE phase - exact bin/batch/qty allocations.
    If Not CFX_PhaseCreate(g, members, docNo, reason) Then Exit Function

    ' Phase 8: WHTask verify (open-qty), then leave CF to the user.
    If Not CFX_PhaseWhTaskVerify(g, members, docNo, reason) Then Exit Function

    g("Phase") = "READY_FOR_CF"
    CFX_ExecuteGroup = True
End Function

' ---- Phase 2: CANCEL --------------------------------------------------------
' Cancel every member whose ExecMode requires releasing the old WT.  Idempotent:
' a WT already terminal in RAM (rerun after cancel, S17) is NOT cancelled again.
Private Function CFX_PhaseCancel(ByVal g As Object, ByVal members As Collection, _
        ByVal docNo As String, ByRef reason As String) As Boolean
    Dim i As Long, m As Object, oldWT As String, em As String
    For i = 1 To members.count
        Set m = members(i)
        em = CStr(m("ExecMode"))
        oldWT = CStr(m("WTNumber"))
        If CFX_ExecModeCancels(em) And oldWT <> "" Then
            If CFX_WTIsTerminal(oldWT) Then
                ' Already released in a prior run - reconcile, do not re-cancel.
                CFX_Log "CANCEL_SKIP_IDEMPOTENT", "PASS", _
                    "WT=" & oldWT & "|RAM=" & CFX_WTState(oldWT) & "|PLAN=" & CStr(m("PlanID"))
            Else
                If Not CFX_CancelWT(oldWT, docNo, reason) Then
                    reason = "STOP_CANCEL_FAILED|WT=" & oldWT & "|" & reason
                    Exit Function
                End If
                m("Decision") = "CANCELLED"
            End If
        End If
    Next i
    CFX_PhaseCancel = True
End Function

' ---- Phase 3: WT-state settle ----------------------------------------------
' After the cancel phase, drive each cancelled WT to its terminal state ONCE via
' the monotonic guard.  A grid that later shows it "open" is stale, not truth.
Private Function CFX_PhaseSettleWT(ByVal members As Collection, ByRef reason As String) As Boolean
    Dim i As Long, m As Object, oldWT As String, grid As Collection
    Dim remQty As Double, remKnown As Boolean, newState As String, r As Long, row As Object
    For i = 1 To members.count
        Set m = members(i)
        oldWT = CStr(m("WTNumber"))
        If oldWT = "" Then GoTo NextMember
        If StrComp(CStr(m("Decision")), "CANCELLED", vbTextCompare) <> 0 Then GoTo NextMember

        ' Remaining OPEN qty = VSOLM of rows for this WT that are still open
        ' (blank status).  After our cancel the WT row is status A, so this is
        ' naturally 0 (WT_CANCELLED_ZERO) unless a genuine open remainder exists.
        Set grid = CFX_ReadWTGrid(CStr(m("DocNo")))
        remQty = 0#: remKnown = True
        For r = 1 To grid.count
            Set row = grid(r)
            If StrComp(CStr(row(CFX_K_TANUM)), oldWT, vbTextCompare) = 0 _
               And Trim$(CStr(row(CFX_K_TOSTAT))) = "" Then
                remQty = remQty + CFX_Num(row(CFX_K_VSOLM))
            End If
        Next r
        newState = CFX_StateFromRawStatus("A", remKnown, remQty)   ' cancelled => A
        If Not CFX_WTSetState(oldWT, newState, reason) Then
            reason = "STOP_WT_SETTLE|" & reason
            Exit Function
        End If
        CFX_WT(oldWT)("RetainedQty") = remQty
        CFX_Log "WT_SETTLED", "PASS", "WT=" & oldWT & "|STATE=" & newState & "|RETAINED=" & CStr(remQty)
NextMember:
    Next i
    CFX_PhaseSettleWT = True
End Function

' ---- Phase 5-6: PRDO one-save + post-verify --------------------------------
' Stage every changed member's target for this document, then commit exactly one
' Save.  ExpectedOpenQty is asserted via WHTask OpenQty, never a MON reopen.
Private Function CFX_PhasePrdo(ByVal g As Object, ByVal members As Collection, _
        ByVal docNo As String, ByRef reason As String) As Boolean
    Dim i As Long, m As Object, anyChange As Boolean, needSave As Boolean, expectOpen As Double
    anyChange = False: needSave = False
    Dim itemTarget As Double
    For i = 1 To members.count
        Set m = members(i)
        If CFX_MemberChangesQty(m) Then
            anyChange = True
            If CBool(m("HasGroupStamp")) Then
                itemTarget = CDbl(m("GroupTarget"))
            Else
                itemTarget = CDbl(m("TargetQty"))
            End If
            CFX_StagePrdoQty docNo, CStr(m("Item")), itemTarget
            ' Idempotency (S18): only a target that isn't already committed needs
            ' the Save.  A rerun where PRDO is already at target does not re-save.
            If Abs(CFX_ReadPrdoQty(docNo, CStr(m("Item"))) - itemTarget) > 0.000001 Then
                needSave = True
            End If
        End If
    Next i
    If Not anyChange Then
        CFX_PhasePrdo = True         ' CF-only / cancel-only groups: no save
        Exit Function
    End If

    If needSave Then
        ' Exactly one Save for the document.
        If Not CFX_CommitPrdoSave(docNo, reason) Then
            reason = "STOP_PRDO_SAVE|" & reason & "|DOC=" & docNo
            Exit Function
        End If
    Else
        CFX_Log "PRDO_SAVE_SKIP_IDEMPOTENT", "PASS", "DOC=" & docNo & "|ALREADY_AT_TARGET"
    End If

    ' Post-verify each changed item's committed qty.
    For i = 1 To members.count
        Set m = members(i)
        If CFX_MemberChangesQty(m) Then
            If CBool(m("HasGroupStamp")) Then
                itemTarget = CDbl(m("GroupTarget"))
            Else
                itemTarget = CDbl(m("TargetQty"))
            End If
            If Not CFX_PostVerifyPrdoQty(docNo, CStr(m("Item")), itemTarget, reason) Then
                Exit Function
            End If
        End If
    Next i

    ' Open-qty expectation = exactly what the CREATE phase is STILL about to
    ' allocate, i.e. planned-create minus anything a prior run already created:
    '   KEEP_EXISTING_CREATE_DELTA -> target - sys   (old WT retained active)
    '   REPLACE_CANCEL_CREATE      -> full target    (old WT cancelled)
    '   CANCEL_PRDO_ZERO           -> 0              (nothing created)
    ' Subtracting already-created qty makes a rerun-after-create (S19) expect 0
    ' open, matching truth, while a fresh run (S02) still expects the full delta.
    ' The verifier independently derives open = committed target - active WT qty.
    expectOpen = CFX_GroupPlannedCreateQty(members) - CFX_GroupAlreadyCreatedQty(docNo, members)
    If expectOpen < 0 Then expectOpen = 0#
    If Not CFX_PostVerifyOpenQty(CStr(g("GroupID")), docNo, expectOpen, reason) Then
        Exit Function
    End If
    CFX_PhasePrdo = True
End Function

' ---- Phase 7: CREATE --------------------------------------------------------
' Create the exact allocations the plan calls for.  Idempotent (S19): if RAM
' already holds a matching fresh active WT for this member, do not create again.
' Delta-add creates only (target-current); replace creates the full target.
Private Function CFX_PhaseCreate(ByVal g As Object, ByVal members As Collection, _
        ByVal docNo As String, ByRef reason As String) As Boolean
    Dim i As Long, m As Object, em As String, qty As Double, alloc As Object
    Dim newWT As String, fp As String
    For i = 1 To members.count
        Set m = members(i)
        em = CStr(m("ExecMode"))
        If Not CFX_ExecModeCreates(em) Then GoTo NextMember

        If StrComp(em, "KEEP_EXISTING_CREATE_DELTA", vbTextCompare) = 0 Then
            qty = CDbl(m("TargetQty")) - CDbl(m("SysQty"))     ' delta only
        Else
            qty = CDbl(m("TargetQty"))                         ' full target
        End If
        If qty <= 0.000001 Then
            CFX_Log "CREATE_SKIP_ZERO", "PASS", "PLAN=" & CStr(m("PlanID")) & "|QTY=" & CStr(qty)
            GoTo NextMember
        End If

        fp = CFX_AllocFingerprint(m, qty)
        If CFX_AllocAlreadySatisfied(docNo, m, qty) Then
            CFX_Log "CREATE_SKIP_IDEMPOTENT", "PASS", "PLAN=" & CStr(m("PlanID")) & "|FP=" & fp
            GoTo NextMember
        End If

        Set alloc = CFX_NewAllocation(CStr(m("Material")), CStr(m("Item")), _
            CStr(m("TargetBin")), CStr(m("TargetBatch")), qty, fp)
        If Not CFX_CreateWT(alloc, docNo, newWT, reason) Then
            reason = "STOP_CREATE_FAILED|PLAN=" & CStr(m("PlanID")) & "|" & reason
            Exit Function
        End If
        alloc("Created") = True
        m("NewWT") = newWT
        ' The freshly created WT is active RAM truth.
        CFX_WTUpsert newWT, "", CFX_WT_ACTIVE, qty, "CREATE"
        Dim allocs As Collection
        Set allocs = g("Allocations")
        allocs.Add alloc
        CFX_Log "CREATE_WT", "PASS", "WT=" & newWT & "|BIN=" & CStr(m("TargetBin")) & _
            "|BATCH=" & CStr(m("TargetBatch")) & "|QTY=" & CStr(qty) & "|PLAN=" & CStr(m("PlanID"))
NextMember:
    Next i
    CFX_PhaseCreate = True
End Function

' ---- Phase 8: WHTask verify (open-qty) then CF handoff ----------------------
Private Function CFX_PhaseWhTaskVerify(ByVal g As Object, ByVal members As Collection, _
        ByVal docNo As String, ByRef reason As String) As Boolean
    ' Return to the WHTask surface (never reopen OB as MON) and confirm nothing
    ' is left open beyond plan.  expectedOpenQty<0 => do not assert a value.
    Dim allocs As Collection
    Set allocs = g("Allocations")
    If allocs.count = 0 And Not CFX_GroupChangesAnyQty(members) Then
        CFX_PhaseWhTaskVerify = True     ' CF-only group: nothing to verify
        Exit Function
    End If
    If Not CFX_PostVerifyOpenQty(CStr(g("GroupID")), docNo, 0#, reason) Then
        Exit Function
    End If
    CFX_PhaseWhTaskVerify = True
End Function

' ============================================================================
' EXEC MODE predicates - the frozen CHECK decision vocabulary.
' ============================================================================
Private Function CFX_ExecModeCancels(ByVal em As String) As Boolean
    Select Case UCase$(Trim$(em))
        Case "CANCEL_PRDO_ZERO", "REPLACE_CANCEL_CREATE"
            CFX_ExecModeCancels = True
    End Select
End Function

Private Function CFX_ExecModeCreates(ByVal em As String) As Boolean
    Select Case UCase$(Trim$(em))
        Case "REPLACE_CANCEL_CREATE", "KEEP_EXISTING_CREATE_DELTA"
            CFX_ExecModeCreates = True
    End Select
End Function

Private Function CFX_MemberChangesQty(ByVal m As Object) As Boolean
    Dim em As String
    em = UCase$(Trim$(CStr(m("ExecMode"))))
    Select Case em
        Case "CANCEL_PRDO_ZERO", "REPLACE_CANCEL_CREATE", "KEEP_EXISTING_CREATE_DELTA"
            CFX_MemberChangesQty = True
    End Select
End Function

Private Function CFX_GroupHasWrite(ByVal g As Object, ByVal members As Collection) As Boolean
    Dim i As Long, m As Object
    For i = 1 To members.count
        Set m = members(i)
        If CFX_ExecModeCreates(CStr(m("ExecMode"))) Or CFX_MemberChangesQty(m) Then
            CFX_GroupHasWrite = True
            Exit Function
        End If
    Next i
End Function

Private Function CFX_GroupChangesAnyQty(ByVal members As Collection) As Boolean
    Dim i As Long, m As Object
    For i = 1 To members.count
        Set m = members(i)
        If CFX_MemberChangesQty(m) Then
            CFX_GroupChangesAnyQty = True
            Exit Function
        End If
    Next i
End Function

' Sum of what the CREATE phase will allocate for this group, by ExecMode.  This
' is the open-qty the PRDO save should leave awaiting creation (pre-Create).
Private Function CFX_GroupPlannedCreateQty(ByVal members As Collection) As Double
    Dim i As Long, m As Object, em As String, q As Double, total As Double
    For i = 1 To members.count
        Set m = members(i)
        em = UCase$(Trim$(CStr(m("ExecMode"))))
        Select Case em
            Case "KEEP_EXISTING_CREATE_DELTA"
                q = CDbl(m("TargetQty")) - CDbl(m("SysQty"))
            Case "REPLACE_CANCEL_CREATE"
                q = CDbl(m("TargetQty"))
            Case Else
                q = 0#                       ' CANCEL_PRDO_ZERO / NONE create nothing
        End Select
        If q > 0 Then total = total + q
    Next i
    CFX_GroupPlannedCreateQty = total
End Function

' Of that planned create, the qty a PRIOR run already allocated (a fresh active
' WT already matching the member's exact bin/batch/qty fingerprint).  Lets the
' pre-Create OpenQty assertion net out to 0 on a rerun-after-create (S19).
Private Function CFX_GroupAlreadyCreatedQty(ByVal docNo As String, ByVal members As Collection) As Double
    Dim i As Long, m As Object, em As String, q As Double, total As Double
    For i = 1 To members.count
        Set m = members(i)
        em = UCase$(Trim$(CStr(m("ExecMode"))))
        Select Case em
            Case "KEEP_EXISTING_CREATE_DELTA"
                q = CDbl(m("TargetQty")) - CDbl(m("SysQty"))
            Case "REPLACE_CANCEL_CREATE"
                q = CDbl(m("TargetQty"))
            Case Else
                q = 0#
        End Select
        If q > 0.000001 Then
            If CFX_AllocAlreadySatisfied(docNo, m, q) Then total = total + q
        End If
    Next i
    CFX_GroupAlreadyCreatedQty = total
End Function

' ============================================================================
' ALLOCATION idempotency (S19 - rerun after create: do not duplicate a WT).
' ============================================================================
Private Function CFX_AllocFingerprint(ByVal m As Object, ByVal qty As Double) As String
    CFX_AllocFingerprint = CStr(m("Material")) & "|" & CStr(m("Item")) & "|" & _
        CStr(m("TargetBin")) & "|" & CStr(m("TargetBatch")) & "|QTY=" & CStr(qty)
End Function

Private Function CFX_AllocAlreadySatisfied(ByVal docNo As String, ByVal m As Object, _
        ByVal qty As Double) As Boolean
    Dim grid As Collection, row As Object, i As Long, fp As String
    Dim rowMat As String, rowItem As String, rowBin As String, rowBatch As String, rowQty As Double
    fp = CFX_AllocFingerprint(m, qty)
    Set grid = CFX_ReadWTGrid(docNo)
    For i = 1 To grid.count
        Set row = grid(i)
        If Trim$(CStr(row(CFX_K_TOSTAT))) = "" Then    ' active only
            rowMat = CStr(row(CFX_K_MATNR))
            rowItem = CStr(row(CFX_K_ITEMNO))
            rowBin = CStr(row(CFX_K_VLPLA))
            rowBatch = CStr(row(CFX_K_CHARG))
            rowQty = CFX_Num(row(CFX_K_VSOLM))
            If StrComp(rowMat, CStr(m("Material")), vbTextCompare) = 0 And _
               CFX_ItemEq(rowItem, CStr(m("Item"))) And _
               StrComp(rowBin, CStr(m("TargetBin")), vbTextCompare) = 0 And _
               StrComp(rowBatch, CStr(m("TargetBatch")), vbTextCompare) = 0 And _
               Abs(rowQty - qty) < 0.000001 Then
                CFX_AllocAlreadySatisfied = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function CFX_ItemEq(ByVal a As String, ByVal b As String) As Boolean
    ' Item numbers may have leading zeros stripped; normalize both.
    CFX_ItemEq = (CLng(val(a)) = CLng(val(b)))
End Function

Private Function CFX_Num(ByVal v As Variant) As Double
    On Error Resume Next
    CFX_Num = CDbl(v)
    On Error GoTo 0
End Function

' Static compile probe - forces a full body compile of this module without
' touching run state, sheets or SAP.  Mirrors modCFX_Entry.CFX_Entry_CompileProbe.
' If this THROWS via Application.Run, modCFX_Executor has a body compile error;
' if it returns "CFX_EXEC_OK", the module compiles and any earlier throw was runtime.
Public Function CFX_Exec_CompileProbe() As String
    CFX_Exec_CompileProbe = "CFX_EXEC_OK|SAP_WRITE=0"
End Function

' Runtime probe: exercises CFX_ExecuteRun exactly as the harness reaches it, but
' traps the VBA error so Application.Run returns the REAL Err.Number|Description
' instead of an opaque HRESULT.  Fresh run, NO locked plan -> a correct build
' returns "RT_OK|EXEC=False|..." (graceful STOP).  Touches no SAP/sheet/save.
Public Function CFX_Exec_RuntimeProbe() As String
    Dim reason As String, ok As Boolean
    On Error GoTo Trap
    CFX_ClearRun
    CFX_NewRun "EXECPROBE", CFX_MODE_REPLAY
    ok = CFX_ExecuteRun("EXECPROBE-OB", reason)
    CFX_Exec_RuntimeProbe = "RT_OK|EXEC=" & CStr(ok) & "|REASON=" & reason
    Exit Function
Trap:
    CFX_Exec_RuntimeProbe = "RT_ERR|" & CStr(Err.Number) & "|" & Err.Description
End Function
