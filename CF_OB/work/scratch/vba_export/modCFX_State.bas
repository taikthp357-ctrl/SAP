Option Explicit

' ============================================================================
' modCFX_State  -  RAM State Store for the CF OB deterministic group executor.
'
' PURPOSE
'   Single in-memory source of truth for one Button #4 run.  Replaces the
'   "sheets-as-database + re-read live grid" pattern that let a stale MON grid
'   re-flag a cancelled WT as active (root cause of the WT 5042889690 stop).
'
'   Everything here is pure VBA (no SAP, no ThisWorkbook writes) so it is fully
'   exercisable in REPLAY mode offline.  State objects are Scripting.Dictionary
'   instances (matching the existing code style) so they serialize to a string
'   for evidence capture and replay assertions.
'
' STATE SHAPE
'   RunState (module singleton g_CFXRun)
'     .Meta   : Dictionary  RunID, Mode, Phase, PlanLocked, PlanFingerprint,
'                           MonSessionId, PrdoSessionId, counters...
'     .WTs    : Dictionary  key = WT number  -> WTState
'     .Groups : Dictionary  key = GroupID     -> GroupState
'   WTState      : WTNumber, RawStatus, State, RetainedQty, Terminal, Origin
'   GroupState   : GroupID, DocNo, WH, Members(Collection of MemberState),
'                  CurrentQty, TargetQty, Delta, Locked, Phase, Allocations
'   MemberState  : PlanID, OB, DocNo, Item, WH, Material, SysQty, TargetQty,
'                  SysBin, TargetBin, SysBatch, TargetBatch, Decision, WTNumber
'   AllocationState: Material, Item, Bin, Batch, Qty, Created, Fingerprint
'
' MONOTONIC INVARIANT (spec section 5)
'   A terminal WT (A -> WT_CANCELLED_*, C -> WT_STATE_DONE) can NEVER transition
'   back to WT_ACTIVE within the same run.  CFX_WTSetState is the ONLY writer of
'   WT state and it enforces this guard, returning False + reason on violation.
' ============================================================================

' ---- WT state vocabulary (mirrors modWTStateEngine, centralised here) -------
Public Const CFX_WT_ACTIVE As String = "WT_ACTIVE"
Public Const CFX_WT_CANCELLED_PENDING As String = "WT_CANCELLED_PENDING_REMAINING"
Public Const CFX_WT_CANCELLED_REMAINING As String = "WT_CANCELLED_REMAINING"
Public Const CFX_WT_CANCELLED_ZERO As String = "WT_CANCELLED_ZERO"
Public Const CFX_WT_DONE As String = "WT_STATE_DONE"
Public Const CFX_WT_UNKNOWN As String = "WT_UNKNOWN"

' ---- Run mode ---------------------------------------------------------------
Public Const CFX_MODE_LIVE As String = "LIVE"
Public Const CFX_MODE_REPLAY As String = "REPLAY"

' ---- Module singleton -------------------------------------------------------
Private g_CFXRun As Object   ' RunState dictionary; Nothing until CFX_NewRun

' ============================================================================
' RUN LIFECYCLE
' ============================================================================
Public Function CFX_NewRun(ByVal runID As String, ByVal mode As String) As Object
    Dim run As Object, meta As Object
    Set run = CreateObject("Scripting.Dictionary")
    run.CompareMode = vbTextCompare
    Set meta = CreateObject("Scripting.Dictionary")
    meta.CompareMode = vbTextCompare
    meta("RunID") = runID
    meta("Mode") = IIf(mode = CFX_MODE_REPLAY, CFX_MODE_REPLAY, CFX_MODE_LIVE)
    meta("Phase") = "INIT"
    meta("PlanLocked") = False
    meta("PlanFingerprint") = ""
    meta("MonSessionId") = ""
    meta("PrdoSessionId") = ""
    ' Bounded-navigation + action counters (evidence, never a SAP trigger).
    meta("DocumentOpenCount") = 0&
    meta("MonOpenCount") = 0&
    meta("WhTaskOpenCount") = 0&
    meta("WtMapRefreshes") = 0&
    meta("PrdoSaves") = 0&
    meta("Creates") = 0&
    meta("Cancels") = 0&
    meta("HardStops") = 0&
    meta("Warns") = 0&
    Set run("Meta") = meta
    Set run("WTs") = NewCiDict()
    Set run("Groups") = NewCiDict()
    Set run("Log") = New Collection
    Set g_CFXRun = run
    Set CFX_NewRun = run
End Function

Public Function CFX_Run() As Object
    ' Returns the active run (raises if none) - callers must CFX_NewRun first.
    If g_CFXRun Is Nothing Then Err.Raise vbObjectError + 7000, "modCFX_State", _
        "CFX_Run: no active run; call CFX_NewRun first."
    Set CFX_Run = g_CFXRun
End Function

Public Function CFX_HasRun() As Boolean
    CFX_HasRun = Not (g_CFXRun Is Nothing)
End Function

Public Sub CFX_ClearRun()
    Set g_CFXRun = Nothing
End Sub

Public Function CFX_Meta() As Object
    Set CFX_Meta = CFX_Run()("Meta")
End Function

' ============================================================================
' SMALL HELPERS
' ============================================================================
Public Function NewCiDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set NewCiDict = d
End Function

Public Function CFX_IsTerminalState(ByVal state As String) As Boolean
    Select Case state
        Case CFX_WT_CANCELLED_PENDING, CFX_WT_CANCELLED_REMAINING, _
             CFX_WT_CANCELLED_ZERO, CFX_WT_DONE
            CFX_IsTerminalState = True
        Case Else
            CFX_IsTerminalState = False
    End Select
End Function

' Deterministic normaliser (mirrors modWTStateEngine, single source).
'   blank/open  -> WT_ACTIVE
'   "C"         -> WT_STATE_DONE (terminal, satisfied)
'   "A"         -> cancelled; remaining-qty unknown at load => PENDING,
'                  else ZERO (rem=0) or REMAINING (rem>0)
Public Function CFX_StateFromRawStatus(ByVal rawStatus As String, _
        ByVal remainingKnown As Boolean, ByVal remainingQty As Double) As String
    Dim s As String
    s = UCase$(Trim$(rawStatus))
    Select Case s
        Case ""
            CFX_StateFromRawStatus = CFX_WT_ACTIVE
        Case "C"
            CFX_StateFromRawStatus = CFX_WT_DONE
        Case "A"
            If Not remainingKnown Then
                CFX_StateFromRawStatus = CFX_WT_CANCELLED_PENDING
            ElseIf remainingQty > 0.000001 Then
                CFX_StateFromRawStatus = CFX_WT_CANCELLED_REMAINING
            Else
                CFX_StateFromRawStatus = CFX_WT_CANCELLED_ZERO
            End If
        Case Else
            CFX_StateFromRawStatus = CFX_WT_UNKNOWN
    End Select
End Function

Public Sub CFX_MetaInc(ByVal key As String, Optional ByVal by As Long = 1)
    Dim m As Object
    Set m = CFX_Meta()
    If Not m.Exists(key) Then m(key) = 0&
    m(key) = CLng(m(key)) + by
End Sub

Public Sub CFX_MetaSet(ByVal key As String, ByVal val As Variant)
    Dim m As Object
    Set m = CFX_Meta()
    m(key) = val
End Sub

Public Sub CFX_Log(ByVal evName As String, ByVal result As String, ByVal detail As String)
    Dim lg As Collection
    Set lg = CFX_Run()("Log")
    lg.Add evName & "|" & result & "|" & detail
End Sub

' ============================================================================
' WT STATE  (the monotonic terminal guard lives here)
' ============================================================================
Public Function CFX_WTUpsert(ByVal wtNumber As String, ByVal rawStatus As String, _
        ByVal state As String, Optional ByVal retainedQty As Double = 0#, _
        Optional ByVal origin As String = "") As Object
    Dim wts As Object, wt As Object
    Set wts = CFX_Run()("WTs")
    If wtNumber = "" Then Err.Raise vbObjectError + 7010, "modCFX_State", _
        "CFX_WTUpsert: blank WT number."
    If wts.Exists(wtNumber) Then
        Set wt = wts(wtNumber)
    Else
        Set wt = NewCiDict()
        wt("WTNumber") = wtNumber
        wt("State") = CFX_WT_UNKNOWN
        Set wts(wtNumber) = wt
    End If
    wt("RawStatus") = rawStatus
    wt("RetainedQty") = retainedQty
    If origin <> "" Then wt("Origin") = origin
    wt("State") = state
    wt("Terminal") = CFX_IsTerminalState(state)
    Set CFX_WTUpsert = wt
End Function

Public Function CFX_WTExists(ByVal wtNumber As String) As Boolean
    Dim wts As Object
    Set wts = CFX_Run()("WTs")
    CFX_WTExists = wts.Exists(wtNumber)
End Function

Public Function CFX_WT(ByVal wtNumber As String) As Object
    Dim wts As Object
    Set wts = CFX_Run()("WTs")
    If Not wts.Exists(wtNumber) Then Err.Raise vbObjectError + 7011, "modCFX_State", _
        "CFX_WT: unknown WT " & wtNumber
    Set CFX_WT = wts(wtNumber)
End Function

Public Function CFX_WTState(ByVal wtNumber As String) As String
    If CFX_WTExists(wtNumber) Then
        CFX_WTState = CStr(CFX_WT(wtNumber)("State"))
    Else
        CFX_WTState = CFX_WT_UNKNOWN
    End If
End Function

Public Function CFX_WTIsTerminal(ByVal wtNumber As String) As Boolean
    If CFX_WTExists(wtNumber) Then CFX_WTIsTerminal = CBool(CFX_WT(wtNumber)("Terminal"))
End Function

' The ONLY sanctioned WT state transition writer.  Enforces the spec section-5
' monotonic invariant: once a WT is terminal (A cancelled / C done) it can never
' be driven back to WT_ACTIVE inside the same run.  Returns False + reason.
Public Function CFX_WTSetState(ByVal wtNumber As String, ByVal newState As String, _
        ByRef reason As String) As Boolean
    Dim wt As Object, cur As String
    If Not CFX_WTExists(wtNumber) Then
        ' First observation of this WT - accept as-is, no prior to violate.
        CFX_WTUpsert wtNumber, "", newState
        CFX_WTSetState = True
        Exit Function
    End If
    Set wt = CFX_WT(wtNumber)
    cur = CStr(wt("State"))
    If cur = newState Then
        CFX_WTSetState = True
        Exit Function
    End If
    If CFX_IsTerminalState(cur) And newState = CFX_WT_ACTIVE Then
        reason = "MONOTONIC_TERMINAL_VIOLATION|WT=" & wtNumber & _
                 "|FROM=" & cur & "|TO=" & newState
        CFX_WTSetState = False
        Exit Function
    End If
    wt("State") = newState
    wt("Terminal") = CFX_IsTerminalState(newState)
    CFX_WTSetState = True
End Function

' ============================================================================
' GROUP + MEMBER STATE
' ============================================================================
Public Function CFX_GroupUpsert(ByVal groupID As String, ByVal docNo As String, _
        ByVal wh As String) As Object
    Dim groups As Object, g As Object
    Set groups = CFX_Run()("Groups")
    If groups.Exists(groupID) Then
        Set g = groups(groupID)
    Else
        Set g = NewCiDict()
        g("GroupID") = groupID
        g("DocNo") = docNo
        g("WH") = wh
        Set g("Members") = New Collection
        Set g("Allocations") = New Collection
        g("CurrentQty") = 0#
        g("TargetQty") = 0#
        g("Delta") = 0#
        g("Locked") = False
        g("Phase") = "PLANNED"
        Set groups(groupID) = g
    End If
    Set CFX_GroupUpsert = g
End Function

Public Function CFX_Group(ByVal groupID As String) As Object
    Dim groups As Object
    Set groups = CFX_Run()("Groups")
    If Not groups.Exists(groupID) Then Err.Raise vbObjectError + 7012, _
        "modCFX_State", "CFX_Group: unknown group " & groupID
    Set CFX_Group = groups(groupID)
End Function

Public Function CFX_Groups() As Object
    Set CFX_Groups = CFX_Run()("Groups")
End Function

Public Function CFX_MemberAdd(ByVal groupID As String, ByVal member As Object) As Object
    Dim g As Object, mems As Collection
    Set g = CFX_Group(groupID)
    Set mems = g("Members")
    mems.Add member
    Set CFX_MemberAdd = member
End Function

Public Function CFX_NewMember() As Object
    Dim m As Object
    Set m = NewCiDict()
    m("PlanID") = "": m("OB") = "": m("DocNo") = "": m("Item") = ""
    m("WH") = "": m("Material") = ""
    m("SysQty") = 0#: m("TargetQty") = 0#
    m("SysBin") = "": m("TargetBin") = ""
    m("SysBatch") = "": m("TargetBatch") = ""
    m("Decision") = "": m("WTNumber") = ""
    Set CFX_NewMember = m
End Function

Public Function CFX_NewAllocation(ByVal material As String, ByVal item As String, _
        ByVal bin As String, ByVal batch As String, ByVal qty As Double, _
        ByVal fingerprint As String) As Object
    Dim a As Object
    Set a = NewCiDict()
    a("Material") = material: a("Item") = item
    a("Bin") = bin: a("Batch") = batch
    a("Qty") = qty: a("Created") = False
    a("Fingerprint") = fingerprint
    Set CFX_NewAllocation = a
End Function

' ============================================================================
' EVIDENCE  (matches ACCEPTANCE_TESTS.md capture columns)
' ============================================================================
Public Function CFX_EvidenceLine(ByVal scenario As String, ByVal result As String, _
        ByVal notes As String) As String
    Dim m As Object, finalStates As String, wtKey As Variant, wts As Object, groups As Object, wtObj As Object
    Set m = CFX_Meta()
    Set wts = CFX_Run()("WTs")
    Set groups = CFX_Run()("Groups")
    For Each wtKey In wts.keys
        Set wtObj = CFX_WT(CStr(wtKey))
        finalStates = finalStates & IIf(finalStates = "", "", ";") & _
            CStr(wtKey) & "=" & CStr(wtObj("State"))
    Next wtKey
    CFX_EvidenceLine = scenario & "," & result & "," & _
        CStr(groups.count) & "," & CStr(CFX_CountMembers()) & "," & _
        CStr(m("DocumentOpenCount")) & "," & CStr(m("MonOpenCount")) & "," & _
        CStr(m("WhTaskOpenCount")) & "," & CStr(m("WtMapRefreshes")) & "," & _
        CStr(m("PrdoSaves")) & "," & CStr(m("Creates")) & "," & _
        CStr(m("HardStops")) & "," & CStr(m("Warns")) & "," & _
        Chr$(34) & finalStates & Chr$(34) & "," & Chr$(34) & notes & Chr$(34)
End Function

Public Function CFX_CountMembers() As Long
    Dim gKey As Variant, total As Long, groups As Object, g As Object, mems As Collection
    Set groups = CFX_Run()("Groups")
    For Each gKey In groups.keys
        Set g = CFX_Group(CStr(gKey))
        Set mems = g("Members")
        total = total + mems.count
    Next gKey
    CFX_CountMembers = total
End Function
