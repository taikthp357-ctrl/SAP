Attribute VB_Name = "modCFExecContext"
' ============================================================================
' CF Execution Context  -  CHECK vs EXECUTE separation (V2.7 CP3)
' ============================================================================
' Central authority for the two-button contract:
'   CHECK  (nut "CHECK")  = read-only planning. Produces a locked PlanID whose
'                           fingerprint binds INPUT + PLAN. MUST NOT write SAP.
'   EXECUTE(nut "XU LY")  = may write SAP, but ONLY for the exact PlanID whose
'                           fingerprint still matches the current INPUT + PLAN.
'
' Design notes:
'   - Mode is process-global (a single run is either CHECK or EXECUTE).
'   - AssertCheckIsReadOnly() is the guard every SAP-write wrapper calls FIRST.
'     In CHECK mode it raises a trappable error so no write can proceed.
'   - PlanID fingerprint reuses the proven MAPFP-style FNV hash from
'     modSAPRunGuard (SAPRun_CalculateFingerprint) so the algorithm is battle
'     tested; here we canonicalize PLAN + INPUT content instead of CONFIG keys.
'   - Nothing here opens Excel dialogs or touches SAP. Pure state + hashing.
' ============================================================================
Option Explicit

Public Enum CFExecMode
    CFMODE_NONE = 0
    CFMODE_CHECK = 1
    CFMODE_EXECUTE = 2
End Enum

Private g_mode As CFExecMode
Private g_lockedPlanID As String
Private g_lockedPlanFingerprint As String
Private g_lockedInputFingerprint As String
Private g_checkID As String

Public Const CFEXEC_ERR_CHECK_READONLY As Long = vbObjectError + 1701

' ============================================================================
' Mode accessors
' ============================================================================
Public Sub CFExec_SetMode(ByVal m As CFExecMode)
    g_mode = m
End Sub

Public Function CFExec_Mode() As CFExecMode
    CFExec_Mode = g_mode
End Function

Public Function CFExec_IsCheck() As Boolean
    CFExec_IsCheck = (g_mode = CFMODE_CHECK)
End Function

Public Function CFExec_IsExecute() As Boolean
    CFExec_IsExecute = (g_mode = CFMODE_EXECUTE)
End Function

Public Function CFExec_ModeText() As String
    Select Case g_mode
        Case CFMODE_CHECK: CFExec_ModeText = "CHECK"
        Case CFMODE_EXECUTE: CFExec_ModeText = "EXECUTE"
        Case Else: CFExec_ModeText = "NONE"
    End Select
End Function

' ============================================================================
' The read-only guard  -  every SAP-write wrapper calls this FIRST
' ============================================================================
' In CHECK mode this raises a trappable error (fail-closed). Callers should let
' it propagate to their own error handler, which logs SAP_WRITE=0 and STOPs.
' actionName is echoed into the error text for the audit trail.
' ============================================================================
Public Sub CFExec_AssertCheckIsReadOnly(ByVal actionName As String)
    If g_mode = CFMODE_CHECK Then
        Err.Raise CFEXEC_ERR_CHECK_READONLY, "modCFExecContext", _
            "STOP_CHECK_IS_READ_ONLY|BLOCKED_WRITE=" & actionName & _
            "|CHECK must never write SAP. SAP_WRITE=0."
    End If
    ' NONE mode is also not allowed to write (fail-closed): a write may only
    ' happen inside an explicit EXECUTE run.
    If g_mode <> CFMODE_EXECUTE Then
        Err.Raise CFEXEC_ERR_CHECK_READONLY, "modCFExecContext", _
            "STOP_EXECUTE_MODE_REQUIRED|BLOCKED_WRITE=" & actionName & _
            "|No active EXECUTE context. SAP_WRITE=0."
    End If
End Sub

' Non-raising probe (for callers that prefer a boolean gate).
Public Function CFExec_WriteAllowed(ByRef messageText As String) As Boolean
    If g_mode = CFMODE_EXECUTE Then
        CFExec_WriteAllowed = True
    ElseIf g_mode = CFMODE_CHECK Then
        messageText = "STOP_CHECK_IS_READ_ONLY|SAP_WRITE=0"
    Else
        messageText = "STOP_EXECUTE_MODE_REQUIRED|SAP_WRITE=0"
    End If
End Function

' ============================================================================
' PlanID fingerprint  -  binds CHECK output to EXECUTE input
' ============================================================================
' Canonicalizes the plan payload (built by CHECK) and hashes it with the same
' FNV algorithm modSAPRunGuard uses, tagged CFPLANFP1 so the two namespaces
' never collide. planPayload is a caller-built canonical string (one field per
' line) covering everything EXECUTE must not silently diverge from: OB list,
' per-member OldWT/SysQty/TargetQty/ExecMode/Bin/Batch, group totals.
' ============================================================================
Public Function CFExec_CalcPlanFingerprint(ByVal planPayload As String) As String
    Dim i As Long, h As Double, codePoint As Long
    h = 104729#
    For i = 1 To Len(planPayload)
        codePoint = AscW(Mid$(planPayload, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536
        h = h * 65599# + CDbl(codePoint)
        h = h - Int(h / 2147483647#) * 2147483647#
    Next i
    CFExec_CalcPlanFingerprint = "CFPLANFP1-" & Hex$(CLng(h)) & "-" & CStr(Len(planPayload))
End Function

' Lock the plan after a successful CHECK. Generates CheckID + PlanID and stores
' both fingerprints. EXECUTE later re-derives the fingerprints from the live
' INPUT+PLAN and refuses to run if either differs.
Public Sub CFExec_LockPlan(ByVal inputPayload As String, ByVal planPayload As String)
    g_checkID = "CHK-" & Format$(Now, "yyyymmddhhnnss") & "-" & Right$("000" & CStr(Int(Rnd() * 1000)), 3)
    g_lockedInputFingerprint = CFExec_CalcPlanFingerprint(inputPayload)
    g_lockedPlanFingerprint = CFExec_CalcPlanFingerprint(planPayload)
    g_lockedPlanID = "PLAN-" & Right$(g_lockedPlanFingerprint, 12) & "-" & Right$(g_lockedInputFingerprint, 8)
End Sub

Public Function CFExec_LockedPlanID() As String
    CFExec_LockedPlanID = g_lockedPlanID
End Function

Public Function CFExec_CheckID() As String
    CFExec_CheckID = g_checkID
End Function

' EXECUTE gate: verify the current INPUT+PLAN still match what CHECK locked.
' Returns False (with reason) on any mismatch or missing lock -> caller STOPs.
Public Function CFExec_VerifyPlanUnchanged(ByVal inputPayload As String, _
        ByVal planPayload As String, ByRef messageText As String) As Boolean
    Dim curInputFp As String, curPlanFp As String
    messageText = ""
    If g_lockedPlanID = "" Then
        messageText = "STOP_NO_LOCKED_PLAN|Run CHECK first; EXECUTE has no locked PlanID."
        Exit Function
    End If
    curInputFp = CFExec_CalcPlanFingerprint(inputPayload)
    curPlanFp = CFExec_CalcPlanFingerprint(planPayload)
    If StrComp(curInputFp, g_lockedInputFingerprint, vbBinaryCompare) <> 0 Then
        messageText = "STOP_PLAN_STALE_RECHECK_REQUIRED|INPUT changed after CHECK" & _
            "|locked=" & g_lockedInputFingerprint & "|now=" & curInputFp
        Exit Function
    End If
    If StrComp(curPlanFp, g_lockedPlanFingerprint, vbBinaryCompare) <> 0 Then
        messageText = "STOP_PLAN_STALE_RECHECK_REQUIRED|PLAN changed after CHECK" & _
            "|locked=" & g_lockedPlanFingerprint & "|now=" & curPlanFp
        Exit Function
    End If
    CFExec_VerifyPlanUnchanged = True
End Function

' Reset all context (called by LAM MOI / new run).
Public Sub CFExec_Reset()
    g_mode = CFMODE_NONE
    g_lockedPlanID = ""
    g_lockedPlanFingerprint = ""
    g_lockedInputFingerprint = ""
    g_checkID = ""
End Sub

' ============================================================================
' Self-test  -  offline proof of the CHECK/EXECUTE contract
' ============================================================================
Public Function CFExecCtx_SelfTest() As String
    Dim msg As String
    Randomize
    CFExec_Reset

    ' 1) CHECK mode must block a write.
    CFExec_SetMode CFMODE_CHECK
    On Error Resume Next
    Err.Clear
    CFExec_AssertCheckIsReadOnly "CANCEL_WT"
    If Err.Number <> CFEXEC_ERR_CHECK_READONLY Then
        On Error GoTo 0
        CFExecCtx_SelfTest = "FAIL|CHECK_DID_NOT_BLOCK_WRITE|errNo=" & CStr(Err.Number)
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    ' 2) NONE mode must also block a write (fail-closed).
    CFExec_SetMode CFMODE_NONE
    On Error Resume Next
    Err.Clear
    CFExec_AssertCheckIsReadOnly "PRDO_SAVE"
    If Err.Number <> CFEXEC_ERR_CHECK_READONLY Then
        On Error GoTo 0
        CFExecCtx_SelfTest = "FAIL|NONE_DID_NOT_BLOCK_WRITE"
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    ' 3) Lock a plan in CHECK, then EXECUTE with identical payload must pass.
    CFExec_LockPlan "OB=1011846094", "M1|3500->0|CANCEL;M2|4500->8000|KEEP;GROUP=8000"
    If CFExec_LockedPlanID() = "" Then
        CFExecCtx_SelfTest = "FAIL|PLANID_NOT_SET"
        Exit Function
    End If
    CFExec_SetMode CFMODE_EXECUTE
    If Not CFExec_VerifyPlanUnchanged("OB=1011846094", "M1|3500->0|CANCEL;M2|4500->8000|KEEP;GROUP=8000", msg) Then
        CFExecCtx_SelfTest = "FAIL|IDENTICAL_PLAN_REJECTED|" & msg
        Exit Function
    End If

    ' 4) EXECUTE must be allowed to write now.
    On Error Resume Next
    Err.Clear
    CFExec_AssertCheckIsReadOnly "WT_CREATE"
    If Err.Number <> 0 Then
        On Error GoTo 0
        CFExecCtx_SelfTest = "FAIL|EXECUTE_BLOCKED_WRITE|errNo=" & CStr(Err.Number)
        Exit Function
    End If
    On Error GoTo 0

    ' 5) A changed PLAN must be rejected as stale.
    If CFExec_VerifyPlanUnchanged("OB=1011846094", "M1|3500->0|CANCEL;M2|4500->9000|KEEP;GROUP=9000", msg) Then
        CFExecCtx_SelfTest = "FAIL|STALE_PLAN_ACCEPTED"
        Exit Function
    End If

    ' 6) A changed INPUT must be rejected as stale.
    If CFExec_VerifyPlanUnchanged("OB=9999999999", "M1|3500->0|CANCEL;M2|4500->8000|KEEP;GROUP=8000", msg) Then
        CFExecCtx_SelfTest = "FAIL|STALE_INPUT_ACCEPTED"
        Exit Function
    End If

    CFExec_Reset
    CFExecCtx_SelfTest = "PASS"
End Function
