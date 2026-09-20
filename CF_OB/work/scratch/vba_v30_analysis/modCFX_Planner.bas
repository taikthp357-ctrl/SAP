Attribute VB_Name = "modCFX_Planner"
Option Explicit

' ============================================================================
' modCFX_Planner  -  Group Planner (Stage 1).
'
'   Consumes the LOCKED CHECK plan and builds the RAM group/member model, seeds
'   WT state from the plan's known truth, computes each group's target ONCE, and
'   locks the plan fingerprint.  No SAP writes and NO live-grid re-read: the plan
'   is frozen at CHECK, so the executor never re-derives it (spec sections 3-5).
'
'   Two entry points share one core (CFX_PlanCore) so grouping / target /
'   fingerprint / lock logic is identical and fully exercisable offline:
'     CFX_PlanFromRawMembers - REPLAY + tests feed member dicts directly.
'     CFX_PlanFromPlanSheet   - LIVE reads the locked PLAN + WH_ACTUAL sheets.
' ============================================================================

' ---- PLAN sheet column map (mirrors modCodexWTRework CWR_PLAN_COL_*) ---------
Private Const CFX_PC_PLANID As Long = 1
Private Const CFX_PC_OB As Long = 2
Private Const CFX_PC_DOC As Long = 3
Private Const CFX_PC_OLDWT As Long = 5
Private Const CFX_PC_MATERIAL As Long = 6
Private Const CFX_PC_SYSQTY As Long = 7
Private Const CFX_PC_TARGETQTY As Long = 8
Private Const CFX_PC_SYSBIN As Long = 9
Private Const CFX_PC_TARGETBIN As Long = 10
Private Const CFX_PC_SYSBATCH As Long = 11
Private Const CFX_PC_TARGETBATCH As Long = 14
Private Const CFX_PC_NEEDQTY As Long = 21
Private Const CFX_PC_NEEDBIN As Long = 22
Private Const CFX_PC_NEEDBATCH As Long = 23
Private Const CFX_PC_GATE As Long = 24
Private Const CFX_PC_WHROW As Long = 27
Private Const CFX_PC_QTY_MODE As Long = 28
Private Const CFX_PC_EXEC_MODE As Long = 29
Private Const CFX_PC_DOC_GROUP As Long = 31
Private Const CFX_PC_LOCATION_GROUP As Long = 32
Private Const CFX_PC_CREATE_GROUP As Long = 33
Private Const CFX_PC_FINGERPRINT As Long = 34
Private Const CFX_PC_QTY_DELTA As Long = 35
Private Const CFX_PC_QTY_GROUP_ID As Long = 47
Private Const CFX_PC_GROUP_CURRENT As Long = 48
Private Const CFX_PC_GROUP_TARGET As Long = 49
Private Const CFX_PC_GROUP_DELTA As Long = 50
Private Const CFX_PC_GROUP_MEMBERS As Long = 51
Private Const CFX_PC_GROUP_NEED_QTY As Long = 52
' WH_ACTUAL sheet columns used by the LIVE reader
Private Const CFX_WH_COL_ITEM As Long = 3
Private Const CFX_WH_COL_WTSTATUS As Long = 10

' ============================================================================
' RAW MEMBER contract - the single shape both entry points produce.  Tests and
' the REPLAY harness build these directly; the LIVE reader fills them from cells.
' ============================================================================
Public Function CFX_NewRawMember() As Object
    Dim r As Object
    Set r = NewCiDict()
    r("PlanID") = "": r("OB") = "": r("DocNo") = "": r("Item") = ""
    r("WH") = "": r("Material") = ""
    r("OldWT") = "": r("OldWTStatus") = ""          ' raw TOSTAT: ""=open,A,C
    r("SysQty") = 0#: r("TargetQty") = 0#: r("QtyDelta") = 0#
    r("SysBin") = "": r("TargetBin") = ""
    r("SysBatch") = "": r("TargetBatch") = ""
    r("QtyMode") = "": r("ExecMode") = ""
    r("NeedQty") = False: r("NeedBin") = False: r("NeedBatch") = False
    r("QtyGroupId") = "": r("DocGroup") = "": r("LocationGroup") = "": r("CreateGroup") = ""
    r("GroupCurrent") = 0#: r("GroupTarget") = 0#: r("GroupDelta") = 0#
    r("GroupMembers") = 0&: r("GroupNeedQty") = False
    r("HasGroupStamp") = False       ' True => trust GroupCurrent/GroupTarget as frozen aggregates
    r("PlanFingerprint") = ""
    Set CFX_NewRawMember = r
End Function

' ============================================================================
' CORE - build RAM groups/members from a Collection of raw-member dicts, seed WT
' state, compute each group target once, lock the plan fingerprint.
' ============================================================================
Public Function CFX_PlanCore(ByVal rawMembers As Collection, ByRef reason As String) As Boolean
    Dim raw As Object, g As Object, m As Object, i As Long
    Dim gid As String, docNo As String, wh As String, fp As String

    If rawMembers Is Nothing Or rawMembers.count = 0 Then
        reason = "STOP_PLAN_EMPTY|NO_MEMBERS_TO_PLAN"
        Exit Function
    End If
    CFX_MetaSet "Phase", "PLANNING"

    For i = 1 To rawMembers.count
        Set raw = rawMembers(i)
        gid = CFX_GroupKey(raw)
        docNo = CStr(raw("DocNo"))
        wh = CStr(raw("WH"))
        Set g = CFX_GroupUpsert(gid, docNo, wh)
        ' A quantity group is one document's outbound item target.  It must never
        ' straddle two documents - that would break the one-save-per-document law.
        If StrComp(CStr(g("DocNo")), docNo, vbTextCompare) <> 0 Then
            reason = "STOP_PLAN_GROUP_SPANS_DOCS|GROUP=" & gid & _
                "|DOC_A=" & CStr(g("DocNo")) & "|DOC_B=" & docNo
            CFX_PlanRollback
            Exit Function
        End If
        ' Capture the frozen group aggregate once (first stamped row wins; any
        ' later row that disagrees is a plan-integrity stop).
        If CBool(raw("HasGroupStamp")) Then
            If Not g.Exists("StampTarget") Then
                g("StampTarget") = CDbl(raw("GroupTarget"))
                g("StampCurrent") = CDbl(raw("GroupCurrent"))
            ElseIf Abs(CDbl(g("StampTarget")) - CDbl(raw("GroupTarget"))) > 0.000001 _
                 Or Abs(CDbl(g("StampCurrent")) - CDbl(raw("GroupCurrent"))) > 0.000001 Then
                reason = "STOP_PLAN_GROUP_TARGET_INCONSISTENT|GROUP=" & gid & _
                    "|T_A=" & CStr(g("StampTarget")) & "|T_B=" & CStr(raw("GroupTarget"))
                CFX_PlanRollback
                Exit Function
            End If
        End If
        Set m = CFX_PlanBuildMember(raw)
        CFX_MemberAdd gid, m
        CFX_PlanSeedWT m
    Next i

    ' Freeze each group's current/target/delta ONCE from the plan.
    If Not CFX_PlanComputeTargets(reason) Then
        CFX_PlanRollback
        Exit Function
    End If

    ' Lock (or, on rerun, prove the fingerprint has not drifted - idempotency).
    fp = CFX_PlanFingerprint()
    If CBool(CFX_Meta()("PlanLocked")) Then
        If StrComp(CStr(CFX_Meta()("PlanFingerprint")), fp, vbTextCompare) <> 0 Then
            reason = "STOP_PLAN_FINGERPRINT_DRIFT|LOCKED=" & CStr(CFX_Meta()("PlanFingerprint")) & _
                "|NOW=" & fp
            Exit Function
        End If
    Else
        CFX_MetaSet "PlanFingerprint", fp
        CFX_MetaSet "PlanLocked", True
    End If
    CFX_MetaSet "Phase", "PLAN_LOCKED"
    Dim groups As Object
    Set groups = CFX_Groups()
    CFX_Log "OB_PLAN_LOCKED", "PASS", "GROUPS=" & CStr(groups.count) & _
        "|MEMBERS=" & CStr(CFX_CountMembers()) & "|FINGERPRINT=" & fp & "|SAP_WRITE=0"
    CFX_PlanCore = True
End Function

' Build a MemberState from a raw-member dict (WTNumber = the OldWT to reconcile;
' the ACTION GATE compares grid rows against this).
Private Function CFX_PlanBuildMember(ByVal raw As Object) As Object
    Dim m As Object
    Set m = CFX_NewMember()
    m("PlanID") = CStr(raw("PlanID"))
    m("OB") = CStr(raw("OB"))
    m("DocNo") = CStr(raw("DocNo"))
    m("Item") = CStr(raw("Item"))
    m("WH") = CStr(raw("WH"))
    m("Material") = CStr(raw("Material"))
    m("SysQty") = CDbl(raw("SysQty"))
    m("TargetQty") = CDbl(raw("TargetQty"))
    m("SysBin") = CStr(raw("SysBin"))
    m("TargetBin") = CStr(raw("TargetBin"))
    m("SysBatch") = CStr(raw("SysBatch"))
    m("TargetBatch") = CStr(raw("TargetBatch"))
    m("WTNumber") = CStr(raw("OldWT"))
    ' Carry the CHECK decision vocabulary verbatim so the executor acts on the
    ' frozen plan, never re-deciding at write time.
    m("QtyMode") = UCase$(CStr(raw("QtyMode")))
    m("ExecMode") = UCase$(CStr(raw("ExecMode")))
    m("NeedQty") = CBool(raw("NeedQty"))
    m("NeedBin") = CBool(raw("NeedBin"))
    m("NeedBatch") = CBool(raw("NeedBatch"))
    m("QtyDelta") = CDbl(raw("QtyDelta"))
    m("OldWTStatus") = UCase$(Trim$(CStr(raw("OldWTStatus"))))
    m("HasGroupStamp") = CBool(raw("HasGroupStamp"))
    m("GroupTarget") = CDbl(raw("GroupTarget"))
    m("GroupCurrent") = CDbl(raw("GroupCurrent"))
    m("GroupDelta") = CDbl(raw("GroupDelta"))
    m("Decision") = "PLANNED"
    Set CFX_PlanBuildMember = m
End Function

' Seed RAM WT state from the plan's known status for this member's OldWT.  This
' is the RAM truth the ACTION GATE trusts over a possibly-stale live grid.
Private Sub CFX_PlanSeedWT(ByVal m As Object)
    Dim wt As String, raw As String, st As String, remKnown As Boolean, remQty As Double
    wt = CStr(m("WTNumber"))
    If wt = "" Then Exit Sub
    raw = CStr(m("OldWTStatus"))
    ' At plan time remaining qty is only "known" for an A row we can measure;
    ' leave PENDING otherwise so the executor's cancel proves the zero/remaining.
    remKnown = False: remQty = 0#
    st = CFX_StateFromRawStatus(raw, remKnown, remQty)
    CFX_WTUpsert wt, raw, st, 0#, "PLAN"
End Sub

' Group key = the frozen QTY_GROUP_ID when present, else a stable per-member key.
Private Function CFX_GroupKey(ByVal raw As Object) As String
    Dim gid As String
    gid = Trim$(CStr(raw("QtyGroupId")))
    If gid = "" Then gid = "QG|" & CStr(raw("DocNo")) & "|" & CStr(raw("Item")) & "|" & CStr(raw("PlanID"))
    CFX_GroupKey = gid
End Function

' Compute CurrentQty / TargetQty / Delta ONCE per group.  When the plan stamped
' a GroupTarget on the rows, honour it and assert every member agrees; otherwise
' derive Current = sum(SysQty), Target = sum(TargetQty).
Private Function CFX_PlanComputeTargets(ByRef reason As String) As Boolean
    Dim gKey As Variant, g As Object, members As Collection, m As Object, i As Long
    Dim sumSys As Double, sumTgt As Double, stampT As Double, stampC As Double, haveStamp As Boolean
    Dim groups As Object
    Set groups = CFX_Groups()
    For Each gKey In groups.keys
        Set g = CFX_Group(CStr(gKey))
        Set members = g("Members")
        sumSys = 0#: sumTgt = 0#: haveStamp = False: stampT = 0#: stampC = 0#
        For i = 1 To members.count
            Set m = members(i)
            sumSys = sumSys + CDbl(m("SysQty"))
            sumTgt = sumTgt + CDbl(m("TargetQty"))
        Next i
        ' The stamped group target (col 49) is identical on every row of a group;
        ' read it from the first member's raw carry if present (>0 or explicit).
        If members.count > 0 Then
            If g.Exists("StampTarget") Then
                haveStamp = True: stampT = CDbl(g("StampTarget")): stampC = CDbl(g("StampCurrent"))
            End If
        End If
        If haveStamp Then
            g("CurrentQty") = stampC
            g("TargetQty") = stampT
        Else
            g("CurrentQty") = sumSys
            g("TargetQty") = sumTgt
        End If
        g("Delta") = CDbl(g("TargetQty")) - CDbl(g("CurrentQty"))
        g("Locked") = True
        g("Phase") = "PLAN_LOCKED"
    Next gKey
    CFX_PlanComputeTargets = True
End Function

Private Sub CFX_PlanRollback()
    ' Clear any partial groups so a failed plan never leaves half-built RAM.
    Dim groups As Object
    Set groups = CFX_Groups()
    groups.RemoveAll
End Sub

' ============================================================================
' FINGERPRINT - deterministic over sorted groups + sorted members.  Same plan
' => same fingerprint => a rerun is proven identical (idempotency, spec s.17).
' ============================================================================
Public Function CFX_PlanFingerprint() As String
    Dim gKey As Variant, keys() As String, n As Long, i As Long, j As Long, tmp As String
    Dim g As Object, out As String, groups As Object
    Set groups = CFX_Groups()
    n = groups.count
    If n = 0 Then Exit Function
    ReDim keys(0 To n - 1)
    i = 0
    For Each gKey In groups.keys
        keys(i) = CStr(gKey): i = i + 1
    Next gKey
    For i = 0 To n - 2                         ' stable sort of group keys
        For j = i + 1 To n - 1
            If StrComp(keys(i), keys(j), vbTextCompare) > 0 Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i
    For i = 0 To n - 1
        Set g = CFX_Group(keys(i))
        out = out & "G[" & keys(i) & "|DOC=" & CStr(g("DocNo")) & _
            "|CUR=" & CStr(g("CurrentQty")) & "|TGT=" & CStr(g("TargetQty")) & _
            "|DEL=" & CStr(g("Delta")) & "]" & CFX_MembersFingerprint(g("Members")) & ";"
    Next i
    CFX_PlanFingerprint = out
End Function

Private Function CFX_MembersFingerprint(ByVal members As Collection) As String
    Dim n As Long, i As Long, j As Long, keys() As String, tmp As String, out As String, m As Object
    n = members.count
    If n = 0 Then Exit Function
    ReDim keys(0 To n - 1)
    For i = 1 To n
        Set m = members(i)
        keys(i - 1) = CStr(m("PlanID")) & "~" & CStr(m("Item")) & "~" & CStr(m("WTNumber")) & _
            "~M=" & CStr(m("Material")) & "~SQ=" & CStr(m("SysQty")) & "~TQ=" & CStr(m("TargetQty")) & _
            "~SB=" & CStr(m("SysBin")) & "~TB=" & CStr(m("TargetBin")) & _
            "~SBt=" & CStr(m("SysBatch")) & "~TBt=" & CStr(m("TargetBatch")) & _
            "~QM=" & CStr(m("QtyMode")) & "~EM=" & CStr(m("ExecMode"))
    Next i
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If StrComp(keys(i), keys(j), vbTextCompare) > 0 Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i
    For i = 0 To n - 1
        out = out & "{" & keys(i) & "}"
    Next i
    CFX_MembersFingerprint = out
End Function

' ============================================================================
' ENTRY - REPLAY / tests: plan directly from raw-member dicts.
' ============================================================================
Public Function CFX_PlanFromRawMembers(ByVal rawMembers As Collection, ByRef reason As String) As Boolean
    CFX_PlanFromRawMembers = CFX_PlanCore(rawMembers, reason)
End Function

' ============================================================================
' ENTRY - LIVE: read the LOCKED PLAN + WH_ACTUAL sheets for one OB, project each
' READY_TO_EXECUTE row into a raw member, then run the shared core.  (Only the
' cell projection is LIVE-specific; grouping/target/lock stay identical.)
' ============================================================================
Public Function CFX_PlanFromPlanSheet(ByVal wsPlan As Object, ByVal wsWH As Object, _
        ByVal obText As String, ByRef reason As String) As Boolean
    Dim rawMembers As Collection, raw As Object, r As Long, lastRow As Long
    Dim gate As String, whRow As Long
    On Error GoTo Failed
    Set rawMembers = New Collection
    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(-4162).row   ' -4162 = xlUp
    For r = 2 To lastRow
        If StrComp(CStr(wsPlan.Cells(r, CFX_PC_OB).Value2), obText, vbTextCompare) = 0 Then
            gate = UCase$(Trim$(CStr(wsPlan.Cells(r, CFX_PC_GATE).Value2)))
            If gate = "READY_TO_EXECUTE" Then
                whRow = CLng(val(wsPlan.Cells(r, CFX_PC_WHROW).Value2))
                Set raw = CFX_PlanReadRow(wsPlan, wsWH, r, whRow)
                rawMembers.Add raw
            End If
        End If
    Next r
    If rawMembers.count = 0 Then
        reason = "STOP_PLAN_NO_READY_ROWS|OB=" & obText
        Exit Function
    End If
    CFX_PlanFromPlanSheet = CFX_PlanCore(rawMembers, reason)
    Exit Function
Failed:
    reason = "STOP_PLAN_SHEET_READ_ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function CFX_PlanReadRow(ByVal wsPlan As Object, ByVal wsWH As Object, _
        ByVal r As Long, ByVal whRow As Long) As Object
    Dim raw As Object
    Set raw = CFX_NewRawMember()
    raw("PlanID") = CStr(wsPlan.Cells(r, CFX_PC_PLANID).Value2)
    raw("OB") = CStr(wsPlan.Cells(r, CFX_PC_OB).Value2)
    raw("DocNo") = CStr(wsPlan.Cells(r, CFX_PC_DOC).Value2)
    raw("Material") = CStr(wsPlan.Cells(r, CFX_PC_MATERIAL).Value2)
    raw("OldWT") = CStr(wsPlan.Cells(r, CFX_PC_OLDWT).Value2)
    raw("SysQty") = val(wsPlan.Cells(r, CFX_PC_SYSQTY).Value2)
    raw("TargetQty") = val(wsPlan.Cells(r, CFX_PC_TARGETQTY).Value2)
    raw("SysBin") = CStr(wsPlan.Cells(r, CFX_PC_SYSBIN).Value2)
    raw("TargetBin") = CStr(wsPlan.Cells(r, CFX_PC_TARGETBIN).Value2)
    raw("SysBatch") = CStr(wsPlan.Cells(r, CFX_PC_SYSBATCH).Value2)
    raw("TargetBatch") = CStr(wsPlan.Cells(r, CFX_PC_TARGETBATCH).Value2)
    raw("QtyMode") = CStr(wsPlan.Cells(r, CFX_PC_QTY_MODE).Value2)
    raw("ExecMode") = CStr(wsPlan.Cells(r, CFX_PC_EXEC_MODE).Value2)
    raw("NeedQty") = (UCase$(Trim$(CStr(wsPlan.Cells(r, CFX_PC_NEEDQTY).Value2))) = "TRUE")
    raw("NeedBin") = (UCase$(Trim$(CStr(wsPlan.Cells(r, CFX_PC_NEEDBIN).Value2))) = "TRUE")
    raw("NeedBatch") = (UCase$(Trim$(CStr(wsPlan.Cells(r, CFX_PC_NEEDBATCH).Value2))) = "TRUE")
    raw("QtyDelta") = val(wsPlan.Cells(r, CFX_PC_QTY_DELTA).Value2)
    raw("QtyGroupId") = CStr(wsPlan.Cells(r, CFX_PC_QTY_GROUP_ID).Value2)
    raw("DocGroup") = CStr(wsPlan.Cells(r, CFX_PC_DOC_GROUP).Value2)
    raw("LocationGroup") = CStr(wsPlan.Cells(r, CFX_PC_LOCATION_GROUP).Value2)
    raw("PlanFingerprint") = CStr(wsPlan.Cells(r, CFX_PC_FINGERPRINT).Value2)
    raw("GroupCurrent") = val(wsPlan.Cells(r, CFX_PC_GROUP_CURRENT).Value2)
    raw("GroupTarget") = val(wsPlan.Cells(r, CFX_PC_GROUP_TARGET).Value2)
    raw("GroupDelta") = val(wsPlan.Cells(r, CFX_PC_GROUP_DELTA).Value2)
    raw("HasGroupStamp") = True
    ' Item + raw WT status live on the WH_ACTUAL row the plan points at.
    If whRow >= 2 Then
        raw("Item") = CStr(wsWH.Cells(whRow, CFX_WH_COL_ITEM).Value2)
        raw("OldWTStatus") = CStr(wsWH.Cells(whRow, CFX_WH_COL_WTSTATUS).Value2)
    End If
    Set CFX_PlanReadRow = raw
End Function
