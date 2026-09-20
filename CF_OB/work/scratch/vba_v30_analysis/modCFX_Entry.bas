Attribute VB_Name = "modCFX_Entry"
Option Explicit

' ============================================================================
' modCFX_Entry  -  the ONLY bridge from the workbook UI (Button #4) into the
' new deterministic CFX engine (State -> Planner -> Gates -> Executor -> Adapter).
'
'   ADDITIVE BY DESIGN.  The legacy proven path in
'   modCodexWTRework.CF_ExecutePlanToCFHandoff is preserved verbatim; Button #4
'   still binds to that same macro.  A one-line guard at the top of that macro
'   consults CFX_EntryEnabled():
'       FALSE (default) -> legacy CWR_* path runs, unchanged.
'       TRUE            -> this module's CFX_ExecutePlanToCFHandoff_Live() runs.
'
'   FAIL-CLOSED FOR OFFLINE WORK.  Until the staged company SAP test window, the
'   LIVE session bind (CFX_BindLiveSessionRoles) returns a STOP, so even with the
'   flag ON the new engine can NEVER perform a live SAP write.  This honours the
'   standing constraints: no live writes during offline work; no live-pass claim
'   until the company tests are performed.  Offline verification runs entirely
'   through Mode=REPLAY (modCFX_Replay.CFX_RunAllReplay), which never calls here.
' ============================================================================

' Defined-name the operator flips to TRUE only when the engine is cleared live.
Private Const CFX_ENGINE_FLAG_NAME As String = "CFX_UseNewEngine"

' ============================================================================
' FEATURE FLAG - default FALSE keeps Button #4 on the legacy proven path.
' ============================================================================
Public Function CFX_EntryEnabled() As Boolean
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.names(CFX_ENGINE_FLAG_NAME).RefersToRange.Value2
    On Error GoTo 0
    CFX_EntryEnabled = (UCase$(Trim$(CStr(v & ""))) = "TRUE")
End Function

' ============================================================================
' LIVE execute entry (new engine).  Returns a result string in the same shape
' the legacy core returns, so the Button #4 handler can display it uniformly.
' No SAP write occurs unless BOTH the flag is TRUE and the LIVE session bind +
' LIVE adapter are wired - neither of which is true during offline work.
' ============================================================================
Public Function CFX_ExecutePlanToCFHandoff_Live() As String
    Dim reason As String, obText As String, wsPlan As Object, wsWH As Object
    On Error GoTo Failed

    ' Fresh RAM run in LIVE mode (REPLAY world is never created here).
    CFX_ClearRun
    CFX_NewRun "LIVE-" & Format$(Now, "yyyymmdd_hhnnss"), CFX_MODE_LIVE

    ' Exact single input OB (INPUT!A7), same source the legacy path uses.
    If Not RT_GetSingleInputOB(obText, reason) Then
        CFX_ExecutePlanToCFHandoff_Live = "STOP|" & reason & "|SAP_WRITE=0|CF=USER_HANDOFF"
        Exit Function
    End If

    ' Fixed session roles MUST resolve to two DISTINCT physical sessions before
    ' any write (spec s.22).  Deferred until the company test window -> STOP.
    If Not CFX_BindLiveSessionRoles(obText, reason) Then
        CFX_ExecutePlanToCFHandoff_Live = "STOP|" & reason & "|SAP_WRITE=0|CF=USER_HANDOFF"
        Exit Function
    End If

    ' Plan from the LOCKED PLAN + WH_ACTUAL sheets (no SAP writes in planning).
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
    If Not CFX_PlanFromPlanSheet(wsPlan, wsWH, obText, reason) Then
        CFX_ExecutePlanToCFHandoff_Live = "STOP|" & reason & "|SAP_WRITE=0|CF=USER_HANDOFF"
        Exit Function
    End If

    ' Execute the locked plan through the LIVE adapter.  CF stays a user handoff.
    If Not CFX_ExecuteRun(obText, reason) Then
        CFX_ExecutePlanToCFHandoff_Live = "STOP|" & reason & "|CF=USER_HANDOFF"
        Exit Function
    End If

    CFX_ExecutePlanToCFHandoff_Live = "READY_FOR_CF_MANUAL_HANDOFF|ENGINE=CFX|OB=" & obText & _
        "|CF_SENT=FALSE|CF=USER_HANDOFF"
    Exit Function
Failed:
    CFX_ExecutePlanToCFHandoff_Live = "FAIL|CFX_ENTRY|" & CStr(Err.Number) & "|" & Err.Description & _
        "|SAP_WRITE=0|CF=USER_HANDOFF"
End Function

' ============================================================================
' LIVE session-role bind.  The proven MON resolver is CWR_GetMONSession and the
' PRDO session is the current-document session; both are Private inside
' modCodexWTRework and open/attach a live SAP tab.  Wiring them is the LAST step,
' performed ONLY in the staged company SAP window.  Until then this fails closed
' so the LIVE engine can never reach a write during offline work.
'
' When wired (company test window), the body will:
'   1. resolve the MON physical session id  (proven CWR MON resolver),
'   2. resolve the PRDO physical session id (proven current-document session),
'   3. CFX_BindSessionRoles monId, prdoId,
'   4. rely on the executor's Phase-0 CFX_AssertSessionRoles to fail-closed on
'      MON=PRDO collision (never auto-reuse the PRDO tab as MON).
' ============================================================================
Private Function CFX_BindLiveSessionRoles(ByVal obText As String, ByRef reason As String) As Boolean
    reason = "STOP_LIVE_SESSION_BIND_NOT_WIRED|OB=" & obText & _
        "|Enable only in the staged company SAP test window. Offline work performs no live write." & _
        "|Verify offline via REPLAY (CFX_RunAllReplay)."
    CFX_BindLiveSessionRoles = False
End Function

' Static compile/readiness probe - touches no SAP, changes no sheet.  Confirms
' the whole CFX stack links and the flag defaults OFF.
Public Function CFX_Entry_CompileProbe() As String
    Dim flagState As String
    flagState = IIf(CFX_EntryEnabled(), "ON", "OFF")
    CFX_Entry_CompileProbe = "CFX_ENTRY_OK|FLAG=" & flagState & "|DEFAULT=OFF|SAP_WRITE=0"
End Function
