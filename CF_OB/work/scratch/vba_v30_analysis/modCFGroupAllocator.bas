Attribute VB_Name = "modCFGroupAllocator"
' ============================================================================
' CF Group Allocator — CONTRACT_A ≠ CONTRACT_B separation (V2.7 CP2)
' ============================================================================
' CONTRACT_A (PRDO aggregate): sum(MemberQty) = GroupTarget → controls PRDO Save
' CONTRACT_B (member allocation): each member FinalVerifiedQty = TargetQty → controls Create WT
' Group complete ONLY when both A and B pass.
'
' This module replaces the V2.6 buggy logic that gated on GROUP_DELTA=0.
' ============================================================================

Option Explicit

Private Const CFALLOC_EPS As Double = 0.000001

' ============================================================================
' Compute required Create quantity for ONE member
' ============================================================================
' Formula: RequiredCreateQty = max(TargetQty - EffectiveQtyAfterPlannedCancelReplace, 0)
'
' EffectiveQtyAfterPlannedCancelReplace depends on ExecMode:
'   - CANCEL_PRDO_ZERO: EffectiveQty = 0 (member will be cancelled)
'   - REPLACE_CANCEL_CREATE: EffectiveQty = 0 (old WT cancelled, full TargetQty created)
'   - KEEP_EXISTING_CREATE_DELTA: EffectiveQty = SysQty (existing WT retained)
'
' Example case: 3500+4500 → 0+8000
'   Member_A: TargetQty=0, ExecMode=CANCEL → Effective=0 → RequiredCreate=max(0-0,0)=0 ✓
'   Member_B: TargetQty=8000, ExecMode=KEEP, SysQty=4500 → Effective=4500 → RequiredCreate=max(8000-4500,0)=3500 ✓
' ============================================================================
Public Function CFAlloc_ComputeRequiredCreate(ByVal memberRec As Object) As Double
    Dim targetQty As Double, effectiveQty As Double, execMode As String
    Dim sysQty As Double, requiredCreate As Double

    On Error GoTo ErrorHandler

    targetQty = CDbl(memberRec("TargetQty"))
    sysQty = CDbl(memberRec("SysQty"))
    execMode = UCase$(CStr(memberRec("ExecMode")))

    ' Compute EffectiveQtyAfterPlannedCancelReplace
    Select Case execMode
        Case "CANCEL_PRDO_ZERO"
            effectiveQty = 0#
        Case "REPLACE_CANCEL_CREATE"
            effectiveQty = 0#
        Case "KEEP_EXISTING_CREATE_DELTA"
            effectiveQty = sysQty
        Case Else
            ' Unknown mode: assume KEEP (conservative fallback)
            effectiveQty = sysQty
    End Select

    ' RequiredCreateQty = max(TargetQty - EffectiveQty, 0)
    requiredCreate = targetQty - effectiveQty
    If requiredCreate < 0# Then requiredCreate = 0#

    CFAlloc_ComputeRequiredCreate = requiredCreate
    Exit Function

ErrorHandler:
    CFAlloc_ComputeRequiredCreate = 0#
End Function

' ============================================================================
' CONTRACT_A check: PRDO aggregate pass
' ============================================================================
' Returns TRUE when sum(MemberQty) = GroupTarget (within epsilon).
' This controls whether PRDO Save is required.
'
' For a single-member group: member SysQty = GroupTarget → A passes.
' For multi-member group: sum of all member SysQty = GroupTarget → A passes.
'
' Note: This uses GroupLiveQty (SAP PRDO readback total), NOT local SysQty sum.
' ============================================================================
Public Function CFAlloc_IsContractAPass(ByVal memberRec As Object) As Boolean
    Dim groupLiveQty As Double, groupTarget As Double

    On Error GoTo ErrorHandler

    groupLiveQty = CDbl(memberRec("GroupLiveQty"))
    groupTarget = CDbl(memberRec("GroupTarget"))

    CFAlloc_IsContractAPass = (Abs(groupLiveQty - groupTarget) <= CFALLOC_EPS)
    Exit Function

ErrorHandler:
    CFAlloc_IsContractAPass = False
End Function

' ============================================================================
' CONTRACT_B check: member allocation complete
' ============================================================================
' Returns TRUE when this member's FinalVerifiedQty = TargetQty (within epsilon).
' This controls whether Create WT is required for THIS member.
'
' A member passes CONTRACT_B when:
'   - Its RequiredCreateQty = 0 (no more Create needed)
'   - AND its TargetQty is correct (either already achieved or will be after pending Create)
'
' Simplified: RequiredCreateQty = 0 means member allocation complete.
' ============================================================================
Public Function CFAlloc_IsContractBPass(ByVal memberRec As Object) As Boolean
    Dim requiredCreate As Double

    On Error GoTo ErrorHandler

    requiredCreate = CFAlloc_ComputeRequiredCreate(memberRec)
    CFAlloc_IsContractBPass = (Abs(requiredCreate) <= CFALLOC_EPS)
    Exit Function

ErrorHandler:
    CFAlloc_IsContractBPass = False
End Function

' ============================================================================
' Group completion check: both CONTRACT_A and CONTRACT_B pass
' ============================================================================
' A group is complete when:
'   1. CONTRACT_A passes: PRDO aggregate = GroupTarget (no PRDO write needed)
'   2. CONTRACT_B passes for ALL members: every member allocation complete (no Create needed)
'
' This replaces the V2.6 buggy logic that only checked GROUP_DELTA=0.
' ============================================================================
Public Function CFAlloc_IsGroupComplete(ByVal members As Collection) As Boolean
    Dim contractA_Pass As Boolean, contractB_AllPass As Boolean
    Dim rec As Object

    On Error GoTo ErrorHandler

    If members.count = 0 Then
        CFAlloc_IsGroupComplete = False
        Exit Function
    End If

    ' Check CONTRACT_A (only need to check one member, they all share GroupLiveQty/GroupTarget)
    Set rec = members.item(1)
    contractA_Pass = CFAlloc_IsContractAPass(rec)

    ' Check CONTRACT_B for ALL members
    contractB_AllPass = True
    For Each rec In members
        If Not CFAlloc_IsContractBPass(rec) Then
            contractB_AllPass = False
            Exit For
        End If
    Next rec

    ' Group complete = A AND B
    CFAlloc_IsGroupComplete = (contractA_Pass And contractB_AllPass)
    Exit Function

ErrorHandler:
    CFAlloc_IsGroupComplete = False
End Function

' ============================================================================
' PRDOAggregateWriteRequired — CONTRACT_A write flag (separate from allocation)
' ============================================================================
' TRUE when the PRDO aggregate total does NOT yet equal GroupTarget, i.e. a
' PRDO Save is still required to reconcile the document aggregate.
'   PRDOAggregateWriteRequired = (GroupCurrentTotal <> GroupTargetTotal)
'
' MUST NOT be conflated with AllocationWriteRequired. A group can have the PRDO
' aggregate already at target (A satisfied) while individual members still need
' Create WT (B not satisfied) — the exact bug this CP2 fixes.
' ============================================================================
Public Function CFAlloc_PRDOAggregateWriteRequired(ByVal memberRec As Object) As Boolean
    Dim groupLiveQty As Double, groupTarget As Double

    On Error GoTo ErrorHandler

    groupLiveQty = CDbl(memberRec("GroupLiveQty"))
    groupTarget = CDbl(memberRec("GroupTarget"))

    ' Write required when aggregate is not yet at target.
    CFAlloc_PRDOAggregateWriteRequired = (Abs(groupLiveQty - groupTarget) > CFALLOC_EPS)
    Exit Function

ErrorHandler:
    ' Fail safe: assume a write is required so we never silently skip a save.
    CFAlloc_PRDOAggregateWriteRequired = True
End Function

' ============================================================================
' AllocationWriteRequired — CONTRACT_B write flag (separate from PRDO aggregate)
' ============================================================================
' TRUE when ANY member of the group still needs a physical allocation action:
'   - RequiredCreateQty > 0 (a Create WT is owed), OR
'   - PlannedCancelQty > 0 (a Cancel is owed), OR
'   - BinChange (NeedBin), OR
'   - BatchChange (NeedBatch)
'
' MUST NOT be conflated with PRDOAggregateWriteRequired.
' ============================================================================
Public Function CFAlloc_AllocationWriteRequired(ByVal members As Collection) As Boolean
    Dim rec As Object, execMode As String

    On Error GoTo ErrorHandler

    For Each rec In members
        If CBool(rec("Terminal")) Then GoTo NextMember

        ' RequiredCreateQty > 0 → a Create WT is owed.
        If CFAlloc_ComputeRequiredCreate(rec) > CFALLOC_EPS Then
            CFAlloc_AllocationWriteRequired = True
            Exit Function
        End If

        ' PlannedCancelQty > 0 → a Cancel is owed (CANCEL_PRDO_ZERO / REPLACE release SysQty).
        execMode = UCase$(CStr(rec("ExecMode")))
        If execMode = "CANCEL_PRDO_ZERO" Or execMode = "REPLACE_CANCEL_CREATE" Then
            If CDbl(rec("SysQty")) > CFALLOC_EPS Then
                CFAlloc_AllocationWriteRequired = True
                Exit Function
            End If
        End If

        ' Bin or Batch change → a location move is owed.
        If CBool(rec("NeedBin")) Or CBool(rec("NeedBatch")) Then
            CFAlloc_AllocationWriteRequired = True
            Exit Function
        End If
NextMember:
    Next rec

    CFAlloc_AllocationWriteRequired = False
    Exit Function

ErrorHandler:
    ' Fail safe: assume allocation is required so we never silently skip a Create.
    CFAlloc_AllocationWriteRequired = True
End Function

' ============================================================================
' Diagnostic: Get allocation status text for a member
' ============================================================================
' Returns human-readable allocation status for logging/debugging.
' ============================================================================
Public Function CFAlloc_MemberStatusText(ByVal memberRec As Object) As String
    Dim reqCreate As Double, contractB As Boolean, txt As String

    On Error GoTo ErrorHandler

    reqCreate = CFAlloc_ComputeRequiredCreate(memberRec)
    contractB = CFAlloc_IsContractBPass(memberRec)

    txt = "TARGET=" & CStr(memberRec("TargetQty")) & _
          "|SYS=" & CStr(memberRec("SysQty")) & _
          "|MODE=" & CStr(memberRec("ExecMode")) & _
          "|REQ_CREATE=" & CStr(reqCreate) & _
          "|CONTRACT_B=" & IIf(contractB, "PASS", "FAIL")

    CFAlloc_MemberStatusText = txt
    Exit Function

ErrorHandler:
    CFAlloc_MemberStatusText = "ERROR"
End Function

' ============================================================================
' Diagnostic: Get group allocation status text
' ============================================================================
' Returns human-readable group status for logging/debugging.
' ============================================================================
Public Function CFAlloc_GroupStatusText(ByVal members As Collection) As String
    Dim contractA As Boolean, groupComplete As Boolean, txt As String
    Dim rec As Object, totalReqCreate As Double

    On Error GoTo ErrorHandler

    If members.count = 0 Then
        CFAlloc_GroupStatusText = "EMPTY_GROUP"
        Exit Function
    End If

    Set rec = members.item(1)
    contractA = CFAlloc_IsContractAPass(rec)
    groupComplete = CFAlloc_IsGroupComplete(members)

    totalReqCreate = 0#
    For Each rec In members
        totalReqCreate = totalReqCreate + CFAlloc_ComputeRequiredCreate(rec)
    Next rec

    ' Re-anchor to first member (For Each leaves rec = Nothing after completion)
    Set rec = members.item(1)
    txt = "GROUP_LIVE=" & CStr(rec("GroupLiveQty")) & _
          "|GROUP_TARGET=" & CStr(rec("GroupTarget")) & _
          "|MEMBERS=" & CStr(members.count) & _
          "|TOTAL_REQ_CREATE=" & CStr(totalReqCreate) & _
          "|CONTRACT_A=" & IIf(contractA, "PASS", "FAIL") & _
          "|GROUP_COMPLETE=" & IIf(groupComplete, "PASS", "FAIL")

    CFAlloc_GroupStatusText = txt
    Exit Function

ErrorHandler:
    CFAlloc_GroupStatusText = "ERROR"
End Function

' ============================================================================
' Self-test — proves the CONTRACT_A ≠ CONTRACT_B bug case (offline, no SAP)
' ============================================================================
' Case 3500+4500 → 0+8000 (shared PRDO item, one member cancelled, one grown):
'   Member_A: TargetQty=0, SysQty=3500, ExecMode=CANCEL_PRDO_ZERO
'             → Effective=0, RequiredCreate=max(0-0,0)=0
'   Member_B: TargetQty=8000, SysQty=4500, ExecMode=KEEP_EXISTING_CREATE_DELTA
'             → Effective=4500, RequiredCreate=max(8000-4500,0)=3500
'   GroupTarget = 8000, GroupLive (after PRDO save) = 8000 → CONTRACT_A PASS
'   Before Member_B create: CONTRACT_B FAIL → group NOT complete (the fix).
' Returns "PASS" or a "FAIL|..." description.
' ============================================================================
Public Function CFAlloc_SelfTest() As String
    Dim mA As Object, mB As Object, members As Collection
    Dim reqA As Double, reqB As Double, msg As String

    On Error GoTo ErrorHandler

    Set mA = CFAlloc_MakeTestMember(0#, 3500#, "CANCEL_PRDO_ZERO", 8000#, 8000#)
    Set mB = CFAlloc_MakeTestMember(8000#, 4500#, "KEEP_EXISTING_CREATE_DELTA", 8000#, 8000#)

    reqA = CFAlloc_ComputeRequiredCreate(mA)
    reqB = CFAlloc_ComputeRequiredCreate(mB)

    If Abs(reqA - 0#) > CFALLOC_EPS Then
        CFAlloc_SelfTest = "FAIL|MEMBER_A_REQ_CREATE=" & CStr(reqA) & "|EXPECTED=0"
        Exit Function
    End If
    If Abs(reqB - 3500#) > CFALLOC_EPS Then
        CFAlloc_SelfTest = "FAIL|MEMBER_B_REQ_CREATE=" & CStr(reqB) & "|EXPECTED=3500"
        Exit Function
    End If

    ' CONTRACT_A passes (aggregate 8000 = target 8000) for both members.
    If Not CFAlloc_IsContractAPass(mA) Then
        CFAlloc_SelfTest = "FAIL|CONTRACT_A_SHOULD_PASS"
        Exit Function
    End If
    ' CONTRACT_B: A complete (req=0), B not complete (req=3500).
    If Not CFAlloc_IsContractBPass(mA) Then
        CFAlloc_SelfTest = "FAIL|CONTRACT_B_MEMBER_A_SHOULD_PASS"
        Exit Function
    End If
    If CFAlloc_IsContractBPass(mB) Then
        CFAlloc_SelfTest = "FAIL|CONTRACT_B_MEMBER_B_SHOULD_FAIL_BEFORE_CREATE"
        Exit Function
    End If

    ' Group must NOT be complete while Member_B still owes a Create (the core fix:
    ' the old GROUP_DELTA=0 logic would have marked this group complete).
    Set members = New Collection
    members.Add mA
    members.Add mB
    If CFAlloc_IsGroupComplete(members) Then
        CFAlloc_SelfTest = "FAIL|GROUP_SHOULD_NOT_BE_COMPLETE_WHILE_MEMBER_B_OWES_CREATE"
        Exit Function
    End If
    If Not CFAlloc_AllocationWriteRequired(members) Then
        CFAlloc_SelfTest = "FAIL|ALLOCATION_WRITE_SHOULD_BE_REQUIRED"
        Exit Function
    End If

    ' After Member_B's create completes, model it as retained at full target.
    mB("SysQty") = 8000#
    If Not CFAlloc_IsContractBPass(mB) Then
        CFAlloc_SelfTest = "FAIL|CONTRACT_B_MEMBER_B_SHOULD_PASS_AFTER_CREATE"
        Exit Function
    End If
    If Not CFAlloc_IsGroupComplete(members) Then
        CFAlloc_SelfTest = "FAIL|GROUP_SHOULD_BE_COMPLETE_AFTER_MEMBER_B_CREATE"
        Exit Function
    End If

    CFAlloc_SelfTest = "PASS"
    Exit Function

ErrorHandler:
    CFAlloc_SelfTest = "FAIL|ERR=" & CStr(Err.Number) & ":" & Err.Description
End Function

' Build an in-memory member record with the fields the allocator reads.
Private Function CFAlloc_MakeTestMember(ByVal targetQty As Double, ByVal sysQty As Double, _
                                        ByVal execMode As String, ByVal groupLiveQty As Double, _
                                        ByVal groupTarget As Double) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = vbTextCompare
    m.Add "TargetQty", targetQty
    m.Add "SysQty", sysQty
    m.Add "ExecMode", execMode
    m.Add "GroupLiveQty", groupLiveQty
    m.Add "GroupTarget", groupTarget
    m.Add "Terminal", False
    m.Add "NeedBin", False
    m.Add "NeedBatch", False
    Set CFAlloc_MakeTestMember = m
End Function
