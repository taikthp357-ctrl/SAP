Attribute VB_Name = "modCFX_SapAdapter"
Option Explicit

' ============================================================================
' modCFX_SapAdapter  -  the ONLY module that talks to SAP.
'
'   LIVE   : delegates to the proven CWR_* primitives in modCodexWTRework
'            (Cancel, PRDO Save, Create WT, MON/WHTask reads).  Those actions
'            are preserved verbatim; this layer only sequences and observes.
'   REPLAY : mutates an in-memory "SAP world" seeded from a fixture so the whole
'            executor + gates + state machine run offline and the 22 acceptance
'            scenarios can be asserted without a live system.
'
' SESSION ROLES (spec sections 11-12)
'   Two FIXED logical sessions: MON (read/monitor) and PRDO (current document).
'   They must resolve to DIFFERENT physical sessions.  CFX_AssertSessionRoles
'   fails closed on collision (scenario S22).  After a document is bound, the
'   warehouse return path is CFX_ReturnToWarehouseTask - never "reopen OB as MON"
'   (which is what produced MON_SESSION_REUSE ... SAME_TAB=TRUE).
' ============================================================================

Private g_World As Object       ' REPLAY SAP world (Nothing in LIVE)
Private g_MonSessionId As String
Private g_PrdoSessionId As String

' ---- WT grid column keys (mirror CWR_Cell keys) -----------------------------
Public Const CFX_K_TANUM As String = "TANUM"
Public Const CFX_K_TOSTAT As String = "TOSTAT"
Public Const CFX_K_ITEMNO As String = "ITEMNO"
Public Const CFX_K_MATNR As String = "MATNR"
Public Const CFX_K_CHARG As String = "CHARG"
Public Const CFX_K_VLPLA As String = "VLPLA"
Public Const CFX_K_NLPLA As String = "NLPLA"
Public Const CFX_K_PROCTY As String = "PROCTY"
Public Const CFX_K_VSOLM As String = "VSOLM"
Public Const CFX_K_DOCNO As String = "DOCNO"

' ============================================================================
' MODE
' ============================================================================
Public Function CFX_IsReplay() As Boolean
    CFX_IsReplay = (StrComp(CStr(CFX_Meta()("Mode")), CFX_MODE_REPLAY, vbTextCompare) = 0)
End Function

' ============================================================================
' REPLAY WORLD  (seeded by modCFX_Replay before a scenario run)
' ============================================================================
Public Function CFX_WorldNew() As Object
    Dim w As Object
    Set w = NewCiDict()
    Set w("Rows") = New Collection      ' each = dict of column->value (WT grid)
    Set w("PrdoQty") = NewCiDict()      ' key doc|item -> committed target qty
    Set w("PrdoStage") = NewCiDict()    ' key doc|item -> staged (uncommitted) qty
    Set w("StaleActive") = NewCiDict()  ' key WT -> remaining reads to lie "open" (S21)
    w("DocOpen") = False
    w("BoundDoc") = ""
    Set g_World = w
    Set CFX_WorldNew = w
End Function

' Seed a document's committed PRDO baseline for an item (the qty a prior SAVE
' already put on the document).  The WHTask OpenQty verify nets committed-target
' against active-WT qty, so every active WT must have a matching committed base
' or a mixed group (CF-only + changed members, S14) can't reconcile.
Public Sub CFX_WorldSeedPrdo(ByVal docNo As String, ByVal item As String, ByVal qty As Double)
    Dim d As Object
    Set d = CFX_World()("PrdoQty")
    d(docNo & "|" & item) = qty
End Sub

' Arrange a one-time (N-read) STALE cache for a WT: the next N grid reads report
' it status-blank ("open") even though its true stored status is terminal.  A
' controlled refresh past N returns the true row.  Models S21 exactly: RAM says
' A, the sheet cache still shows Active; one refresh reveals the fresh truth.
Public Sub CFX_WorldSetStaleActive(ByVal wt As String, ByVal reads As Long)
    Dim d As Object
    Set d = CFX_World()("StaleActive")
    d(wt) = reads
End Sub

Public Function CFX_World() As Object
    If g_World Is Nothing Then Err.Raise vbObjectError + 7100, "modCFX_SapAdapter", _
        "CFX_World: no REPLAY world; seed it first."
    Set CFX_World = g_World
End Function

Public Sub CFX_WorldAddWT(ByVal tanum As String, ByVal tostat As String, _
        ByVal item As String, ByVal material As String, ByVal batch As String, _
        ByVal bin As String, ByVal dest As String, ByVal procty As String, _
        ByVal qty As Double, ByVal docNo As String)
    Dim row As Object, rows As Collection
    Set row = NewCiDict()
    row(CFX_K_TANUM) = tanum: row(CFX_K_TOSTAT) = tostat
    row(CFX_K_ITEMNO) = item: row(CFX_K_MATNR) = material
    row(CFX_K_CHARG) = batch: row(CFX_K_VLPLA) = bin
    row(CFX_K_NLPLA) = dest: row(CFX_K_PROCTY) = procty
    row(CFX_K_VSOLM) = qty: row(CFX_K_DOCNO) = docNo
    Set rows = CFX_World()("Rows")
    rows.Add row
End Sub

' ============================================================================
' SESSION ROLES
' ============================================================================
Public Sub CFX_BindSessionRoles(ByVal monId As String, ByVal prdoId As String)
    g_MonSessionId = monId
    g_PrdoSessionId = prdoId
    CFX_MetaSet "MonSessionId", monId
    CFX_MetaSet "PrdoSessionId", prdoId
End Sub

' Fail-closed session-role collision check (S22).  MON must not equal PRDO.
Public Function CFX_AssertSessionRoles(ByRef reason As String) As Boolean
    Dim monId As String, prdoId As String
    monId = CStr(CFX_Meta()("MonSessionId"))
    prdoId = CStr(CFX_Meta()("PrdoSessionId"))
    If monId = "" Or prdoId = "" Then
        reason = "SESSION_ROLE_UNBOUND|MON=" & monId & "|PRDO=" & prdoId
        CFX_AssertSessionRoles = False
        Exit Function
    End If
    If StrComp(monId, prdoId, vbTextCompare) = 0 Then
        reason = "SESSION_ROLE_COLLISION|MON=PRDO=" & monId & "|NEVER_AUTO_REUSE"
        CFX_AssertSessionRoles = False
        Exit Function
    End If
    CFX_AssertSessionRoles = True
End Function

' The sanctioned warehouse-return path after a document is bound.  It NEVER
' reopens the OB as MON; it activates the WHTask surface on the MON session.
Public Function CFX_ReturnToWarehouseTask(ByRef reason As String) As Boolean
    CFX_MetaInc "WhTaskOpenCount"
    If CFX_IsReplay() Then
        CFX_ReturnToWarehouseTask = True
        Exit Function
    End If
    ' LIVE: delegate to the proven WHTask navigation primitive.
    ' Proven primitive to call here (modCF_WarehouseTask / CWR_* WHTask helper):
    '   CFWT_ActivateWarehouseTaskTab / CWR_ReturnToWarehouseTask
    On Error Resume Next
    Application.run "CWR_ReturnToWarehouseTaskLive"
    If Err.Number <> 0 Then
        reason = "WHTASK_RETURN_LIVE_ERROR|" & Err.Description
        On Error GoTo 0
        CFX_ReturnToWarehouseTask = False
        Exit Function
    End If
    On Error GoTo 0
    CFX_ReturnToWarehouseTask = True
End Function

' ============================================================================
' READ:  filtered WT grid for the current document (the read the buggy gate
'        used to do against a possibly-stale MON grid).  Returns a Collection
'        of column dicts.  In REPLAY this reflects prior cancels/creates, so it
'        is never stale relative to the actions this run performed.
' ============================================================================
Public Function CFX_ReadWTGrid(ByVal docNo As String) As Collection
    Dim out As New Collection, row As Object, src As Collection, i As Long
    Dim stale As Object, wt As String, proj As Object, k As Variant, staleNow As Boolean
    CFX_MetaInc "MonOpenCount"
    If CFX_IsReplay() Then
        Set src = CFX_World()("Rows")
        Set stale = CFX_World()("StaleActive")
        For i = 1 To src.count
            Set row = src(i)
            If CStr(row(CFX_K_DOCNO)) = docNo Or docNo = "" Then
                wt = CStr(row(CFX_K_TANUM))
                staleNow = False                         ' guard: no non-short-circuit read
                If stale.Exists(wt) Then
                    If CLng(stale(wt)) > 0 Then staleNow = True
                End If
                If staleNow Then
                    ' Emit a stale projection (blank status) WITHOUT mutating the
                    ' stored truth, then burn one stale read (S21 cache lie).
                    Set proj = NewCiDict()
                    For Each k In row.keys
                        proj(CStr(k)) = row(CStr(k))
                    Next k
                    proj(CFX_K_TOSTAT) = ""              ' lie: looks open
                    stale(wt) = CLng(stale(wt)) - 1
                    out.Add proj
                Else
                    out.Add row
                End If
            End If
        Next i
        Set CFX_ReadWTGrid = out
        Exit Function
    End If
    ' LIVE: delegate to proven CWR_GetMONSession + CWR_LoadFilteredWTGrid.
    ' (Mapped by CFX_LiveReadWTGrid, which wraps those primitives and converts
    '  the ALV shell into the same column-dict Collection shape.)
    Set CFX_ReadWTGrid = CFX_LiveReadWTGrid(docNo)
End Function

' One controlled refresh of the WT grid (scenario S21 - stale cache).  Emits a
' refresh count so evidence proves we refreshed at most once before trusting.
Public Function CFX_RefreshWTGrid(ByVal docNo As String) As Collection
    CFX_MetaInc "WtMapRefreshes"
    Set CFX_RefreshWTGrid = CFX_ReadWTGrid(docNo)
End Function

' ============================================================================
' WRITE:  Cancel WT  (proven action preserved in LIVE; simulated in REPLAY)
' ============================================================================
Public Function CFX_CancelWT(ByVal wtNumber As String, ByVal docNo As String, _
        ByRef reason As String) As Boolean
    Dim src As Collection, row As Object, i As Long, found As Boolean
    If CFX_IsReplay() Then
        Set src = CFX_World()("Rows")
        For i = 1 To src.count
            Set row = src(i)
            If StrComp(CStr(row(CFX_K_TANUM)), wtNumber, vbTextCompare) = 0 Then
                row(CFX_K_TOSTAT) = "A"      ' cancelled -> terminal A
                found = True
            End If
        Next i
        If Not found Then
            reason = "REPLAY_CANCEL_WT_NOT_FOUND|WT=" & wtNumber
            CFX_CancelWT = False
            Exit Function
        End If
        CFX_MetaInc "Cancels"
        CFX_CancelWT = True
        Exit Function
    End If
    ' LIVE: call proven cancel primitive (Codex_* / CWR_* cancel path).
    CFX_CancelWT = CFX_LiveCancelWT(wtNumber, docNo, reason)
    If CFX_CancelWT Then CFX_MetaInc "Cancels"
End Function

' ============================================================================
' WRITE:  PRDO Save - split into STAGE (enter each item's target qty) and COMMIT
'         (the single Save click per document).  This is what enforces the
'         "one PRDO Save per document" law: many stages, exactly one commit.
' ============================================================================
Public Sub CFX_StagePrdoQty(ByVal docNo As String, ByVal item As String, _
        ByVal targetQty As Double)
    ' Enter a target into the (REPLAY) PRDO grid; NOT yet committed to SAP.
    Dim key As String, d As Object
    key = docNo & "|" & item
    If CFX_IsReplay() Then
        Set d = CFX_World()("PrdoStage")
        d(key) = targetQty
        Exit Sub
    End If
    CFX_LiveStagePrdoQty docNo, item, targetQty
End Sub

' The single Save click for a document: promotes all staged item targets for
' that document to committed, counting exactly one PRDO save.
Public Function CFX_CommitPrdoSave(ByVal docNo As String, ByRef reason As String) As Boolean
    Dim stage As Object, saved As Object, k As Variant, kd As String
    If CFX_IsReplay() Then
        Set stage = CFX_World()("PrdoStage")
        Set saved = CFX_World()("PrdoQty")
        For Each k In stage.keys
            kd = CStr(k)
            If Left$(kd, Len(docNo) + 1) = docNo & "|" Then
                saved(kd) = stage(kd)
            End If
        Next k
        CFX_MetaInc "PrdoSaves"
        CFX_CommitPrdoSave = True
        Exit Function
    End If
    CFX_CommitPrdoSave = CFX_LiveCommitPrdoSave(docNo, reason)
    If CFX_CommitPrdoSave Then CFX_MetaInc "PrdoSaves"
End Function

Public Function CFX_ReadPrdoQty(ByVal docNo As String, ByVal item As String) As Double
    Dim key As String, d As Object
    key = docNo & "|" & item
    If CFX_IsReplay() Then
        Set d = CFX_World()("PrdoQty")
        If d.Exists(key) Then _
            CFX_ReadPrdoQty = CDbl(d(key))
        Exit Function
    End If
    CFX_ReadPrdoQty = CFX_LiveReadPrdoQty(docNo, item)
End Function

' Sum of all COMMITTED PRDO targets for a document (across its items).  Used by
' the WHTask OpenQty verify: OpenQty = committed target - already-active WT qty.
Public Function CFX_ReadCommittedPrdoTotal(ByVal docNo As String) As Double
    Dim saved As Object, k As Variant, kd As String, total As Double
    If CFX_IsReplay() Then
        Set saved = CFX_World()("PrdoQty")
        For Each k In saved.keys
            kd = CStr(k)
            If Left$(kd, Len(docNo) + 1) = docNo & "|" Then total = total + CDbl(saved(kd))
        Next k
        CFX_ReadCommittedPrdoTotal = total
        Exit Function
    End If
    CFX_ReadCommittedPrdoTotal = CFX_LiveReadCommittedPrdoTotal(docNo)
End Function

' ============================================================================
' WRITE:  Create WT allocation (exact bin/batch/qty preserved)
' ============================================================================
Public Function CFX_CreateWT(ByVal alloc As Object, ByVal docNo As String, _
        ByRef newWT As String, ByRef reason As String) As Boolean
    Dim rows As Collection
    If CFX_IsReplay() Then
        Set rows = CFX_World()("Rows")
        newWT = "9" & Format$(rows.count + 1, "000000000")
        CFX_WorldAddWT newWT, "", CStr(alloc("Item")), CStr(alloc("Material")), _
            CStr(alloc("Batch")), CStr(alloc("Bin")), "", "", CDbl(alloc("Qty")), docNo
        CFX_MetaInc "Creates"
        CFX_CreateWT = True
        Exit Function
    End If
    CFX_CreateWT = CFX_LiveCreateWT(alloc, docNo, newWT, reason)
    If CFX_CreateWT Then CFX_MetaInc "Creates"
End Function

' ============================================================================
' DOCUMENT bind (normal DOCUMENT_OPEN_COUNT=1 enforced by executor via counter)
' ============================================================================
Public Function CFX_OpenDocument(ByVal docNo As String, ByRef reason As String) As Boolean
    CFX_MetaInc "DocumentOpenCount"
    If CFX_IsReplay() Then
        CFX_World()("DocOpen") = True
        CFX_World()("BoundDoc") = docNo
        CFX_OpenDocument = True
        Exit Function
    End If
    CFX_OpenDocument = CFX_LiveOpenDocument(docNo, reason)
End Function

' ============================================================================
' LIVE delegation stubs - each wraps an existing proven primitive.  They are
' only reached when Mode=LIVE; REPLAY never calls them.  Kept thin on purpose:
' the real SAP GUI scripting stays in modCodexWTRework (unchanged).
' ============================================================================
Private Function CFX_LiveReadWTGrid(ByVal docNo As String) As Collection
    ' TODO(LIVE-WIRING): wrap CWR_GetMONSession + CWR_LoadFilteredWTGrid and
    ' project the ALV shell rows into column-dicts. Left unwired until live tab.
    Err.Raise vbObjectError + 7150, "modCFX_SapAdapter", _
        "LIVE WT grid read not wired yet - run in REPLAY, or wire to CWR primitives."
End Function
Private Function CFX_LiveCancelWT(ByVal wtNumber As String, ByVal docNo As String, ByRef reason As String) As Boolean
    Err.Raise vbObjectError + 7151, "modCFX_SapAdapter", "LIVE cancel not wired yet."
End Function
Private Sub CFX_LiveStagePrdoQty(ByVal docNo As String, ByVal item As String, ByVal targetQty As Double)
    Err.Raise vbObjectError + 7152, "modCFX_SapAdapter", "LIVE PRDO stage not wired yet."
End Sub
Private Function CFX_LiveCommitPrdoSave(ByVal docNo As String, ByRef reason As String) As Boolean
    Err.Raise vbObjectError + 7156, "modCFX_SapAdapter", "LIVE PRDO commit not wired yet."
End Function
Private Function CFX_LiveReadPrdoQty(ByVal docNo As String, ByVal item As String) As Double
    Err.Raise vbObjectError + 7153, "modCFX_SapAdapter", "LIVE PRDO read not wired yet."
End Function
Private Function CFX_LiveReadCommittedPrdoTotal(ByVal docNo As String) As Double
    Err.Raise vbObjectError + 7157, "modCFX_SapAdapter", "LIVE PRDO total read not wired yet."
End Function
Private Function CFX_LiveCreateWT(ByVal alloc As Object, ByVal docNo As String, ByRef newWT As String, ByRef reason As String) As Boolean
    Err.Raise vbObjectError + 7154, "modCFX_SapAdapter", "LIVE create not wired yet."
End Function
Private Function CFX_LiveOpenDocument(ByVal docNo As String, ByRef reason As String) As Boolean
    Err.Raise vbObjectError + 7155, "modCFX_SapAdapter", "LIVE open document not wired yet."
End Function

