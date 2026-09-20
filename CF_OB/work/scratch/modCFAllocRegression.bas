Attribute VB_Name = "modCFAllocRegression"
' ============================================================================
' CF Allocator Regression Harness (V2.7 CP2)  -  offline, no SAP.
'   Exercises modCFGroupAllocator against the Qty regression suite:
'     - Single-member: no-op / increase-delta / large-delta / decrease-replace
'       / cancel-to-zero / bin-change / idempotent-rerun
'     - Group: net-zero reallocation (the 3500+4500->0+8000 core bug),
'       net-zero after create, multi-member increase.
'   Asserts exact RequiredCreateQty and CONTRACT_A/B/complete flags.
'   NEVER touches SAP. Pure in-memory logic test.
' ============================================================================
Option Explicit

Private Const REG_EPS As Double = 0.000001
Private g_pass As Long
Private g_fail As Long
Private g_detail As String

Private Function REG_Make(ByVal targetQty As Double, ByVal sysQty As Double, _
        ByVal execMode As String, ByVal groupLive As Double, ByVal groupTarget As Double, _
        Optional ByVal needBin As Boolean = False, Optional ByVal needBatch As Boolean = False, _
        Optional ByVal terminal As Boolean = False) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = vbTextCompare
    m.Add "TargetQty", targetQty
    m.Add "SysQty", sysQty
    m.Add "ExecMode", execMode
    m.Add "GroupLiveQty", groupLive
    m.Add "GroupTarget", groupTarget
    m.Add "NeedBin", needBin
    m.Add "NeedBatch", needBatch
    m.Add "Terminal", terminal
    Set REG_Make = m
End Function

Private Sub REG_AssertCreate(ByVal id As String, ByVal m As Object, ByVal expectCreate As Double, _
        ByVal expectBpass As Boolean)
    Dim gotCreate As Double, gotBpass As Boolean, ok As Boolean
    gotCreate = CFAlloc_ComputeRequiredCreate(m)
    gotBpass = CFAlloc_IsContractBPass(m)
    ok = (Abs(gotCreate - expectCreate) <= REG_EPS) And (gotBpass = expectBpass)
    If ok Then
        g_pass = g_pass + 1
        g_detail = g_detail & id & "|PASS|CREATE=" & CStr(gotCreate) & "|Bpass=" & CStr(gotBpass) & vbCrLf
    Else
        g_fail = g_fail + 1
        g_detail = g_detail & id & "|FAIL|CREATE=" & CStr(gotCreate) & "(exp " & CStr(expectCreate) & _
            ")|Bpass=" & CStr(gotBpass) & "(exp " & CStr(expectBpass) & ")" & vbCrLf
    End If
End Sub

Private Sub REG_AssertGroup(ByVal id As String, ByVal members As Collection, _
        ByVal expectApass As Boolean, ByVal expectComplete As Boolean, _
        ByVal expectAllocWrite As Boolean, ByVal expectTotalCreate As Double)
    Dim gotA As Boolean, gotComplete As Boolean, gotAlloc As Boolean
    Dim tot As Double, rec As Object, ok As Boolean
    gotA = CFAlloc_IsContractAPass(members.item(1))
    gotComplete = CFAlloc_IsGroupComplete(members)
    gotAlloc = CFAlloc_AllocationWriteRequired(members)
    tot = 0#
    For Each rec In members
        tot = tot + CFAlloc_ComputeRequiredCreate(rec)
    Next rec
    ok = (gotA = expectApass) And (gotComplete = expectComplete) And _
         (gotAlloc = expectAllocWrite) And (Abs(tot - expectTotalCreate) <= REG_EPS)
    If ok Then
        g_pass = g_pass + 1
        g_detail = g_detail & id & "|PASS|A=" & CStr(gotA) & "|complete=" & CStr(gotComplete) & _
            "|allocWrite=" & CStr(gotAlloc) & "|totCreate=" & CStr(tot) & vbCrLf
    Else
        g_fail = g_fail + 1
        g_detail = g_detail & id & "|FAIL|A=" & CStr(gotA) & "(exp " & CStr(expectApass) & _
            ")|complete=" & CStr(gotComplete) & "(exp " & CStr(expectComplete) & _
            ")|allocWrite=" & CStr(gotAlloc) & "(exp " & CStr(expectAllocWrite) & _
            ")|totCreate=" & CStr(tot) & "(exp " & CStr(expectTotalCreate) & ")" & vbCrLf
    End If
End Sub

' ============================================================================
' PUBLIC ENTRY  -  run the whole suite, return summary + write detail file
' ============================================================================
Public Function CFReg_RunAll() As String
    Dim m As Object, members As Collection
    g_pass = 0: g_fail = 0: g_detail = ""

    ' --- Single-member Qty cases (mirror replay S01/S02/S03/S06/S07/S09/S18) ---
    ' R01 no-op: 600->600 KEEP, aggregate ok -> no create, B pass
    Set m = REG_Make(600#, 600#, "KEEP_EXISTING_CREATE_DELTA", 600#, 600#)
    REG_AssertCreate "R01_NOOP", m, 0#, True
    ' R02 increase-delta: 580->600 KEEP -> create 20, B fail until create
    Set m = REG_Make(600#, 580#, "KEEP_EXISTING_CREATE_DELTA", 600#, 600#)
    REG_AssertCreate "R02_DELTA_ADD_20", m, 20#, False
    ' R03 large delta: 200->600 KEEP -> create 400
    Set m = REG_Make(600#, 200#, "KEEP_EXISTING_CREATE_DELTA", 600#, 600#)
    REG_AssertCreate "R03_DELTA_ADD_400", m, 400#, False
    ' R04 decrease-replace: 580->400 REPLACE -> cancel old, create full 400
    Set m = REG_Make(400#, 580#, "REPLACE_CANCEL_CREATE", 400#, 400#)
    REG_AssertCreate "R04_DECREASE_REPLACE_400", m, 400#, False
    ' R05 cancel-to-zero: 580->0 CANCEL -> create 0, B pass (nothing to create)
    Set m = REG_Make(0#, 580#, "CANCEL_PRDO_ZERO", 0#, 0#)
    REG_AssertCreate "R05_CANCEL_ZERO", m, 0#, True
    ' R06 bin-change only, same qty: 580->580 REPLACE (cancel+recreate full)
    Set m = REG_Make(580#, 580#, "REPLACE_CANCEL_CREATE", 580#, 580#, True, False)
    REG_AssertCreate "R06_BIN_CHANGE_REPLACE_580", m, 580#, False
    ' R07 increase via replace (bug-class S08): 580->640 REPLACE -> create 640
    Set m = REG_Make(640#, 580#, "REPLACE_CANCEL_CREATE", 640#, 640#)
    REG_AssertCreate "R07_REPLACE_INCREASE_640", m, 640#, False

    ' --- GROUP cases ---
    ' G01 CORE BUG 3500+4500 -> 0+8000: net delta 0, but member B owes create 3500.
    Set members = New Collection
    members.Add REG_Make(0#, 3500#, "CANCEL_PRDO_ZERO", 8000#, 8000#)          ' WT_A cancel
    members.Add REG_Make(8000#, 4500#, "KEEP_EXISTING_CREATE_DELTA", 8000#, 8000#) ' WT_B keep+create 3500
    REG_AssertGroup "G01_NETZERO_3500_4500", members, True, False, True, 3500#

    ' G02 after member_B create completes (SysQty now 8000) -> group complete, no alloc.
    Set members = New Collection
    members.Add REG_Make(0#, 0#, "CANCEL_PRDO_ZERO", 8000#, 8000#)              ' WT_A already cancelled (Sys 0)
    members.Add REG_Make(8000#, 8000#, "KEEP_EXISTING_CREATE_DELTA", 8000#, 8000#) ' WT_B fully allocated
    REG_AssertGroup "G02_NETZERO_AFTER_CREATE", members, True, True, False, 0#

    ' G03 multi-member increase: 580+500 -> 600+600 (both KEEP, create 20+100)
    Set members = New Collection
    members.Add REG_Make(600#, 580#, "KEEP_EXISTING_CREATE_DELTA", 1080#, 1200#)
    members.Add REG_Make(600#, 500#, "KEEP_EXISTING_CREATE_DELTA", 1080#, 1200#)
    REG_AssertGroup "G03_MULTI_INCREASE", members, False, False, True, 120#

    ' G04 pure decrease group 300+700 -> 500+500 (one increase, one decrease-replace)
    Set members = New Collection
    members.Add REG_Make(500#, 300#, "KEEP_EXISTING_CREATE_DELTA", 1000#, 1000#) ' +200
    members.Add REG_Make(500#, 700#, "REPLACE_CANCEL_CREATE", 1000#, 1000#)      ' replace full 500
    REG_AssertGroup "G04_MIXED_REALLOC", members, True, False, True, 700#

    Dim summary As String
    summary = "REG_" & IIf(g_fail = 0, "ALL_PASS", "FAIL") & _
        "|TOTAL=" & CStr(g_pass + g_fail) & "|PASS=" & CStr(g_pass) & "|FAIL=" & CStr(g_fail)

    ' Write detail to file (offline evidence).
    On Error Resume Next
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile("g:\SAP\CF_OB\work\REPORTS\cp2_alloc_regression.log", True)
    ts.Write summary & vbCrLf & vbCrLf & g_detail
    ts.Close
    On Error GoTo 0

    CFReg_RunAll = summary
End Function
