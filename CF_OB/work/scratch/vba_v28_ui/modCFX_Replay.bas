Attribute VB_Name = "modCFX_Replay"
Option Explicit

' ============================================================================
' modCFX_Replay  -  offline acceptance harness (Stage 5).
'
'   Seeds an in-memory SAP world + a locked plan for each of the 28 REPLAY
'   scenarios, runs the REAL State -> Planner -> Gates -> Executor -> Adapter
'   stack in Mode=REPLAY, and asserts the spec's key invariant per scenario.
'   Emits the ACCEPTANCE_TESTS.md evidence line for every scenario.
' ============================================================================

Private g_Results As Collection
Private g_FailCount As Long
Private g_PassCount As Long

Private Const RB_OB As String = "1011832412"
Private Const RB_DOC As String = "1012087422"
Private Const RB_WH As String = "6026342483"
Private Const RB_MAT As String = "BIW32023716AA"

Public Sub RB_Begin(ByVal scenario As String)
    If g_Results Is Nothing Then Set g_Results = New Collection
    CFX_ClearRun
    CFX_NewRun "REPLAY-" & scenario, CFX_MODE_REPLAY
    CFX_WorldNew
    CFX_BindSessionRoles "MON_SES_1", "PRDO_SES_2"
End Sub

Public Function RB_Member(ByVal planID As String, ByVal item As String, _
        ByVal oldWT As String, ByVal oldWTStatus As String, _
        ByVal sysQty As Double, ByVal targetQty As Double, _
        ByVal sysBin As String, ByVal targetBin As String, _
        ByVal sysBatch As String, ByVal targetBatch As String, _
        ByVal qtyMode As String, ByVal execMode As String) As Object
    Dim r As Object
    Set r = CFX_NewRawMember()
    r("PlanID") = planID: r("OB") = RB_OB: r("DocNo") = RB_DOC
    r("Item") = item: r("WH") = RB_WH: r("Material") = RB_MAT
    r("OldWT") = oldWT: r("OldWTStatus") = oldWTStatus
    r("SysQty") = sysQty: r("TargetQty") = targetQty
    r("SysBin") = sysBin: r("TargetBin") = targetBin
    r("SysBatch") = sysBatch: r("TargetBatch") = targetBatch
    r("QtyMode") = qtyMode: r("ExecMode") = execMode
    r("QtyDelta") = targetQty - sysQty
    r("QtyGroupId") = "QG|" & RB_DOC & "|" & item
    r("NeedQty") = (execMode <> "NONE")
    Set RB_Member = r
End Function

Public Sub RB_SeedWT(ByVal wt As String, ByVal status As String, ByVal item As String, _
        ByVal bin As String, ByVal batch As String, ByVal qty As Double)
    CFX_WorldAddWT wt, status, item, RB_MAT, batch, bin, "", "", qty, RB_DOC
    CFX_WorldSeedPrdo RB_DOC, item, CFX_ReadPrdoQty(RB_DOC, item) + qty
End Sub

Public Function RB_Plan(ByVal members As Collection, ByRef reason As String) As Boolean
    RB_Plan = CFX_PlanFromRawMembers(members, reason)
End Function

Public Sub RB_FinishExpect(ByVal scenario As String, ByVal expectPass As Boolean, _
        ByVal execOk As Boolean, ByVal reason As String, ByVal assertNote As String)
    Dim passed As Boolean, result As String, notes As String
    passed = (execOk = expectPass)
    result = IIf(passed, "PASS", "STOP")
    notes = assertNote
    If Not passed Then notes = "ASSERT_FAIL|EXPECT_PASS=" & CStr(expectPass) & _
        "|GOT=" & CStr(execOk) & "|" & reason & "|" & assertNote
    If passed Then g_PassCount = g_PassCount + 1 Else g_FailCount = g_FailCount + 1
    If g_Results Is Nothing Then Set g_Results = New Collection
    g_Results.Add CFX_EvidenceLine(scenario, result, notes)
End Sub

Public Sub RB_RunSingle(ByVal scenario As String, ByVal m As Object, _
        ByVal expectPass As Boolean, ByVal assertNote As String)
    Dim members As Collection, reason As String, planOk As Boolean, execOk As Boolean
    Set members = New Collection
    members.Add m
    planOk = RB_Plan(members, reason)
    If Not planOk Then
        RB_FinishExpect scenario, expectPass, False, "PLAN_STOP|" & reason, assertNote
        Exit Sub
    End If
    execOk = CFX_ExecuteRun(RB_OB, reason)
    RB_FinishExpect scenario, expectPass, execOk, reason, assertNote
End Sub

' ============================================================================
' SCENARIO BODIES (S01-S28)
' ============================================================================

Public Sub RS_S01()
    RB_Begin "S01"
    RB_SeedWT "5040000001", "", "10", "APFL-P09-01-06", "260816-02", 600#
    Dim m As Object
    Set m = RB_Member("P01", "10", "5040000001", "", 600#, 600#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "NO_QTY_CHANGE", "NONE")
    RB_RunSingle "S01", m, True, "CF_ONLY|NO_REOPEN|DOC_OPEN=1"
End Sub

Public Sub RS_S02()
    RB_Begin "S02"
    RB_SeedWT "5040000002", "", "10", "APFL-P09-01-06", "260816-02", 580#
    Dim m As Object
    Set m = RB_Member("P02", "10", "5040000002", "", 580#, 600#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "DELTA_ADD", "KEEP_EXISTING_CREATE_DELTA")
    RB_RunSingle "S02", m, True, "INCREASE_QTY|DELTA_ADD|CREATE_20|PRDO_TOTAL=600"
End Sub

Public Sub RS_S03()
    RB_Begin "S03"
    RB_SeedWT "5040000003", "", "10", "APFL-P09-01-06", "260816-02", 200#
    Dim m As Object
    Set m = RB_Member("P03", "10", "5040000003", "", 200#, 600#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "DELTA_ADD", "KEEP_EXISTING_CREATE_DELTA")
    RB_RunSingle "S03", m, True, "LARGE_DELTA_ADD|RETAIN_200|CREATE_400|PRDO_600"
End Sub

Public Sub RS_S04()
    RB_Begin "S04"
    RB_SeedWT "5040000041", "", "10", "APFL-P09-01-06", "260816-02", 300#
    RB_SeedWT "5040000042", "", "10", "APFL-P09-01-06", "260816-02", 280#
    Dim m As Object
    Set m = RB_Member("P04", "10", "5040000042", "", 280#, 300#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "DELTA_ADD", "KEEP_EXISTING_CREATE_DELTA")
    m("GroupCurrent") = 580#
    m("GroupTarget") = 600#
    m("GroupDelta") = 20#
    m("HasGroupStamp") = True
    RB_RunSingle "S04", m, True, "DELTA_WITH_PRIOR_SPLIT|RETAIN_ALL|CREATE_20|PRDO_600"
End Sub

Public Sub RS_S05()
    RB_Begin "S05"
    RB_SeedWT "5040000051", "", "10", "APFL-P09-01-06", "260816-02", 200#
    RB_SeedWT "5040000052", "", "10", "APFL-P09-01-06", "260816-02", 200#
    RB_SeedWT "5040000053", "", "10", "APFL-P09-01-06", "260816-02", 200#
    Dim members As Collection, reason As String, execOk As Boolean, planOk As Boolean
    Set members = New Collection
    members.Add RB_Member("P05a", "10", "5040000051", "", 200#, 200#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "NO_QTY_CHANGE", "NONE")
    members.Add RB_Member("P05b", "10", "5040000052", "", 200#, 200#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "NO_QTY_CHANGE", "NONE")
    members.Add RB_Member("P05c", "10", "5040000053", "", 200#, 200#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "NO_QTY_CHANGE", "NONE")
    Dim m1 As Object, m2 As Object, m3 As Object
    Set m1 = members(1): m1("QtyGroupId") = "QG|" & RB_DOC & "|10"
    Set m2 = members(2): m2("QtyGroupId") = "QG|" & RB_DOC & "|10"
    Set m3 = members(3): m3("QtyGroupId") = "QG|" & RB_DOC & "|10"
    planOk = RB_Plan(members, reason)
    If Not planOk Then RB_FinishExpect "S05", True, False, "PLAN_STOP|" & reason, "": Exit Sub
    execOk = CFX_ExecuteRun(RB_OB, reason)
    RB_FinishExpect "S05", True, execOk, reason, "MULTI_WT_ONE_LINE|AGGREGATE_600|ONE_GROUP"
End Sub

Public Sub RS_S06()
    RB_Begin "S06"
    RB_SeedWT "5040000061", "", "10", "APFL-P14-01-08", "260716-02", 580#
    Dim m As Object
    Set m = RB_Member("P06", "10", "5040000061", "", 580#, 0#, _
        "APFL-P14-01-08", "APFL-P14-01-08", "260716-02", "260716-02", "FINAL_TOTAL", "CANCEL_PRDO_ZERO")
    RB_RunSingle "S06", m, True, "CANCEL_TO_ZERO|NO_CREATE|PRDO_ZERO"
End Sub

Public Sub RS_S07()
    RB_Begin "S07"
    RB_SeedWT "5040000071", "", "10", "APFL-P14-01-08", "260716-02", 580#
    Dim m As Object
    Set m = RB_Member("P07", "10", "5040000071", "", 580#, 400#, _
        "APFL-P14-01-08", "APFL-P09-01-06", "260716-02", "260816-02", "FINAL_TOTAL", "REPLACE_CANCEL_CREATE")
    RB_RunSingle "S07", m, True, "DECREASE_QTY|CANCEL_OLD|PRDO_400|CREATE_400"
End Sub

Public Sub RS_S08()
    RB_Begin "S08"
    RB_SeedWT "5042889690", "", "10", "APFL-P14-01-08", "260716-02", 580#
    Dim m As Object
    Set m = RB_Member("P08", "10", "5042889690", "", 580#, 640#, _
        "APFL-P14-01-08", "APFL-P09-01-06", "260716-02", "260816-02", "FINAL_TOTAL", "REPLACE_CANCEL_CREATE")
    RB_RunSingle "S08", m, True, "CURRENT_BUG_CLASS|CANCEL_580|PRDO_640|NO_STALE_STOP|CREATE_640"
End Sub

Public Sub RS_S09()
    RB_Begin "S09"
    RB_SeedWT "5040000091", "", "10", "APFL-P14-01-08", "260716-02", 580#
    Dim m As Object
    Set m = RB_Member("P09", "10", "5040000091", "", 580#, 580#, _
        "APFL-P14-01-08", "APFL-P09-01-06", "260716-02", "260716-02", "NO_QTY_CHANGE", "REPLACE_CANCEL_CREATE")
    RB_RunSingle "S09", m, True, "BIN_CHANGE_ONLY|CANCEL_580|NO_PRDO_SAVE|CREATE_580"
End Sub

Public Sub RS_S10()
    RB_Begin "S10"
    RB_SeedWT "5040000010", "", "10", "APFL-P14-01-08", "260816-02", 300#
    Dim m As Object
    Set m = RB_Member("P10", "10", "5040000010", "", 300#, 300#, _
        "APFL-P14-01-08", "APFL-CC01-01-01", "260816-02", "260816-02", "NO_QTY_CHANGE", "REPLACE_CANCEL_CREATE")
    RB_RunSingle "S10", m, True, "APFL-CC01_ACCEPTED|NO_P_ZF_HR_REJECTION"
End Sub

Public Sub RS_S11()
    RB_Begin "S11"
    RB_SeedWT "5040000011", "A", "10", "APFL-P14-01-08", "260816-02", 0#
    Dim m As Object
    Set m = RB_Member("P11", "10", "5040000011", "A", 200#, 0#, _
        "APFL-P14-01-08", "APFL-P14-01-08", "260816-02", "260816-02", "NO_QTY_CHANGE", "NONE")
    RB_RunSingle "S11", m, True, "A_UNUSED|NO_NEW_WT|NO_HARD_STOP"
End Sub

Public Sub RS_S12()
    RB_Begin "S12"
    RB_SeedWT "5040000012", "A", "10", "APFL-P14-01-08", "260816-02", 0#
    Dim m As Object
    Set m = RB_Member("P12", "10", "5040000012", "A", 0#, 80#, _
        "APFL-P14-01-08", "APFL-P09-01-06", "260816-02", "260816-02", "FINAL_TOTAL", "REPLACE_CANCEL_CREATE")
    RB_RunSingle "S12", m, True, "A_RECREATE|CREATE_80|OLD_WT_RETAINED=0|CANCEL_IDEMPOTENT_SKIP"
End Sub

Public Sub RS_S13()
    RB_Begin "S13"
    RB_SeedWT "5040000013", "A", "10", "APFL-P14-01-08", "260816-02", 0#
    Dim m As Object
    Set m = RB_Member("P13", "10", "5040000013", "A", 0#, 150#, _
        "APFL-P14-01-08", "APFL-CC01-02-02", "260816-02", "260816-02", "FINAL_TOTAL", "REPLACE_CANCEL_CREATE")
    RB_RunSingle "S13", m, True, "A_REPICK_NEW_BIN|CREATE_150|OLD_SOURCE_IGNORED"
End Sub

Public Sub RS_S14()
    RB_Begin "S14"
    RB_SeedWT "5040000141", "", "10", "APFL-P09-01-06", "260816-02", 600#
    RB_SeedWT "5040000142", "", "20", "APFL-P09-02-06", "260816-02", 580#
    RB_SeedWT "5040000143", "", "30", "APFL-P09-03-06", "260716-02", 400#
    RB_SeedWT "5040000144", "", "40", "APFL-P09-04-06", "260816-02", 300#
    Dim members As Collection, reason As String, execOk As Boolean, planOk As Boolean
    Dim g As String
    g = "QG|" & RB_DOC & "|10"
    Set members = New Collection
    members.Add RB_Member("P14a", "10", "5040000141", "", 600#, 600#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "NO_QTY_CHANGE", "NONE")
    members.Add RB_Member("P14b", "20", "5040000142", "", 580#, 600#, _
        "APFL-P09-02-06", "APFL-P09-02-06", "260816-02", "260816-02", "DELTA_ADD", "KEEP_EXISTING_CREATE_DELTA")
    members.Add RB_Member("P14c", "30", "5040000143", "", 400#, 400#, _
        "APFL-P09-03-06", "APFL-P09-03-06", "260716-02", "260816-02", "NO_QTY_CHANGE", "REPLACE_CANCEL_CREATE")
    members.Add RB_Member("P14d", "40", "5040000144", "", 300#, 0#, _
        "APFL-P09-04-06", "APFL-P09-04-06", "260816-02", "260816-02", "FINAL_TOTAL", "CANCEL_PRDO_ZERO")
    Dim i As Long, mObj As Object
    For i = 1 To members.count
        Set mObj = members(i)
        mObj("QtyGroupId") = g
    Next i
    planOk = RB_Plan(members, reason)
    If Not planOk Then RB_FinishExpect "S14", True, False, "PLAN_STOP|" & reason, "": Exit Sub
    execOk = CFX_ExecuteRun(RB_OB, reason)
    RB_FinishExpect "S14", True, execOk, reason, "FOUR_MEMBERS|ONE_GROUP|DOC_OPEN=1|NO_MEMBER_WRITE_IN_PLAN"
End Sub

Public Sub RS_S15()
    RB_Begin "S15"
    RB_SeedWT "5040000151", "", "10", "APFL-P09-01-06", "260816-02", 580#
    RB_SeedWT "5040000152", "", "20", "APFL-P09-02-06", "260816-02", 500#
    Dim members As Collection, reason As String, execOk As Boolean, planOk As Boolean, g As String
    g = "QG|" & RB_DOC & "|10"
    Set members = New Collection
    members.Add RB_Member("P15a", "10", "5040000151", "", 580#, 600#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "DELTA_ADD", "KEEP_EXISTING_CREATE_DELTA")
    members.Add RB_Member("P15b", "20", "5040000152", "", 500#, 400#, _
        "APFL-P09-02-06", "APFL-P09-02-06", "260816-02", "260816-02", "FINAL_TOTAL", "REPLACE_CANCEL_CREATE")
    Dim mObj1 As Object, mObj2 As Object
    Set mObj1 = members(1): mObj1("QtyGroupId") = g
    Set mObj2 = members(2): mObj2("QtyGroupId") = g
    planOk = RB_Plan(members, reason)
    If Not planOk Then RB_FinishExpect "S15", True, False, "PLAN_STOP|" & reason, "": Exit Sub
    execOk = CFX_ExecuteRun(RB_OB, reason)
    RB_FinishExpect "S15", True, execOk, reason, "INCREASE_AND_DECREASE|ONE_SAVE_PER_DOC"
End Sub

Public Sub RS_S16()
    RB_Begin "S16"
    RB_SeedWT "5040000161", "", "10", "APFL-P14-01-08", "260716-02", 300#
    RB_SeedWT "5040000162", "", "20", "APFL-P14-02-08", "260716-02", 200#
    Dim members As Collection, reason As String, execOk As Boolean, planOk As Boolean, g As String
    g = "QG|" & RB_DOC & "|10"
    Set members = New Collection
    members.Add RB_Member("P16a", "10", "5040000161", "", 300#, 300#, _
        "APFL-P14-01-08", "APFL-P09-01-06", "260716-02", "260816-02", "NO_QTY_CHANGE", "REPLACE_CANCEL_CREATE")
    members.Add RB_Member("P16b", "20", "5040000162", "", 200#, 200#, _
        "APFL-P14-02-08", "APFL-CC01-01-01", "260716-02", "260816-02", "NO_QTY_CHANGE", "REPLACE_CANCEL_CREATE")
    Dim mObj1 As Object, mObj2 As Object
    Set mObj1 = members(1): mObj1("QtyGroupId") = g
    Set mObj2 = members(2): mObj2("QtyGroupId") = g
    planOk = RB_Plan(members, reason)
    If Not planOk Then RB_FinishExpect "S16", True, False, "PLAN_STOP|" & reason, "": Exit Sub
    execOk = CFX_ExecuteRun(RB_OB, reason)
    RB_FinishExpect "S16", True, execOk, reason, "TWO_CREATES|DISTINCT_BINS_BATCHES|NO_QTY_MERGE"
End Sub

Public Sub RS_S17()
    RB_Begin "S17"
    RB_SeedWT "5042889690", "A", "10", "APFL-P14-01-08", "260716-02", 580#
    Dim m As Object
    Set m = RB_Member("P17", "10", "5042889690", "A", 580#, 640#, _
        "APFL-P14-01-08", "APFL-P09-01-06", "260716-02", "260816-02", "FINAL_TOTAL", "REPLACE_CANCEL_CREATE")
    RB_RunSingle "S17", m, True, "RERUN_AFTER_CANCEL|IDEMPOTENT_SKIP|CANCELS=0|CREATE_640"
End Sub

Public Sub RS_S18()
    RB_Begin "S18"
    RB_SeedWT "5040000181", "", "10", "APFL-P09-01-06", "260816-02", 580#
    CFX_WorldSeedPrdo RB_DOC, "10", 600#
    Dim m As Object
    Set m = RB_Member("P18", "10", "5040000181", "", 580#, 600#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "DELTA_ADD", "KEEP_EXISTING_CREATE_DELTA")
    RB_RunSingle "S18", m, True, "RERUN_AFTER_SAVE|PRDO_SKIP_IDEMPOTENT|SAVES=0|CREATE_DELTA_20"
End Sub

Public Sub RS_S19()
    RB_Begin "S19"
    RB_SeedWT "5040000191", "", "10", "APFL-P09-01-06", "260816-02", 580#
    RB_SeedWT "9000000199", "", "10", "APFL-P09-01-06", "260816-02", 20#
    Dim m As Object
    Set m = RB_Member("P19", "10", "5040000191", "", 580#, 600#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "DELTA_ADD", "KEEP_EXISTING_CREATE_DELTA")
    RB_RunSingle "S19", m, True, "RERUN_AFTER_CREATE|ALLOC_ALREADY_SATISFIED|CREATES=0|NO_DUP_WT"
End Sub

Public Sub RS_S20()
    RB_Begin "S20"
    RB_SeedWT "5040000020", "C", "10", "APFL-P09-01-06", "260816-02", 350#
    Dim m As Object
    Set m = RB_Member("P20", "10", "5040000020", "C", 350#, 350#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "NO_QTY_CHANGE", "NONE")
    RB_RunSingle "S20", m, True, "EXISTING_C_TERMINAL|SATISFIED|NO_RECREATE"
End Sub

Public Sub RS_S21()
    RB_Begin "S21"
    RB_SeedWT "5042889690", "A", "10", "APFL-P14-01-08", "260716-02", 580#
    CFX_WorldSetStaleActive "5042889690", 1
    Dim m As Object
    Set m = RB_Member("P21", "10", "5042889690", "A", 580#, 640#, _
        "APFL-P14-01-08", "APFL-P09-01-06", "260716-02", "260816-02", "FINAL_TOTAL", "REPLACE_CANCEL_CREATE")
    RB_RunSingle "S21", m, True, "STALE_CACHE_REFRESHED_ONCE|NO_STALE_BLOCK|WARN=1|CREATE_640"
End Sub

Public Sub RS_S22()
    RB_Begin "S22"
    CFX_BindSessionRoles "SAME_SESSION_9", "SAME_SESSION_9"
    RB_SeedWT "5040000022", "", "10", "APFL-P09-01-06", "260816-02", 580#
    Dim m As Object
    Set m = RB_Member("P22", "10", "5040000022", "", 580#, 640#, _
        "APFL-P09-01-06", "APFL-P09-01-06", "260816-02", "260816-02", "FINAL_TOTAL", "REPLACE_CANCEL_CREATE")
    RB_RunSingle "S22", m, False, "SESSION_ROLE_COLLISION|STOP_BEFORE_WRITE|NEVER_AUTO_REUSE"
End Sub

Public Sub RS_S23()
    Dim detailText As String, decisionText As String, ok As Boolean
    RB_Begin "S23"
    decisionText = CWR_RegressionGateDecision("CREATE_WT_DELTA", True, "NOT_PROVEN", _
                                              "LIVE_PROVEN", False, 0, detailText)
    ok = (decisionText = "LIVE_PROVEN_INHERITED") And _
         InStr(1, detailText, "CONTRACT_CHANGED=False", vbTextCompare) > 0
    RB_FinishExpect "S23", True, ok, detailText, _
                    "HISTORICAL_PROOF_INHERITED|NO_STOP|SAP_WRITE_ALLOWED_BY_NORMAL_EXECUTOR"
End Sub

Public Sub RS_S24()
    Dim detailText As String, decisionText As String, ok As Boolean
    RB_Begin "S24"
    decisionText = CWR_RegressionGateDecision("CREATE_WT_DELTA", True, "NOT_PROVEN", _
                                              "LIVE_PROVEN", True, 0, detailText)
    ok = (decisionText = "LIVE_REPROOF_REQUIRED") And _
         InStr(1, detailText, "MAX_WRITE_ATTEMPTS=1", vbTextCompare) > 0 And _
         InStr(1, detailText, "PROOF_RUN=BOUNDED", vbTextCompare) > 0
    RB_FinishExpect "S24", True, ok, detailText, _
                    "CONTRACT_CHANGED|BOUNDED_REPROOF|MAX_WRITE_ATTEMPTS=1|NO_UNBOUNDED_RETRY"
End Sub

Public Sub RS_S25()
    Dim detailText As String, decisionText As String, executionAllowed As Boolean
    RB_Begin "S25"
    decisionText = CWR_RegressionGateDecision("CREATE_WT_DELTA", True, "PASS", _
                                              "LIVE_PROVEN", False, 1, detailText)
    executionAllowed = (decisionText <> "REGRESSION_FAIL")
    RB_FinishExpect "S25", False, executionAllowed, detailText, _
                    "REAL_REGRESSION_FAIL|STOP_BEFORE_WRITE|SAP_WRITE=0"
End Sub

Public Sub RS_S26()
    Dim detailText As String, decisionText As String, ok As Boolean
    RB_Begin "S26"
    decisionText = CWR_RegressionGateDecision("CHANGE_BIN", False, "NOT_PROVEN", _
                                              "LIVE_PROVEN", False, 0, detailText)
    ok = (decisionText = "NOT_RELEVANT")
    RB_FinishExpect "S26", True, ok, detailText, _
                    "IRRELEVANT_SCENARIO|NOT_RELEVANT|CURRENT_EXECUTION_NOT_BLOCKED"
End Sub

Public Sub RS_S27()
    Dim m As Object, reason As String, planOk As Boolean, execOk As Boolean
    Dim creates As Long
    RB_Begin "S27"
    RB_SeedWT "5040000271", "A", "10", "APFL-P14-01-08", "260716-02", 0#
    RB_SeedWT "9000000277", "", "10", "APFL-ZF01-01-01", "260816-02", 640#
    Set m = RB_Member("P27", "10", "5040000271", "A", 580#, 640#, _
                      "APFL-P14-01-08", "APFL-ZF01-01-01", "260716-02", "260816-02", _
                      "FINAL_TOTAL", "REPLACE_CANCEL_CREATE")
    Dim members As Collection
    Set members = New Collection: members.Add m
    planOk = RB_Plan(members, reason)
    If planOk Then execOk = CFX_ExecuteRun(RB_OB, reason)
    creates = CLng(CFX_Meta()("Creates"))
    RB_FinishExpect "S27", True, planOk And execOk And creates = 0, reason, _
                    "CHANGE_BIN_RERUN_AFTER_CREATE|CREATES=0|NO_DUPLICATE_WT"
End Sub

Public Sub RS_S28()
    Dim m As Object, members As Collection, reason As String
    Dim planOk As Boolean, execOk As Boolean, creates As Long, saves As Long
    RB_Begin "S28"
    RB_SeedWT "9000000288", "", "10", "APFL-HR01-01-01", "260816-02", 640#
    Set m = RB_Member("P28", "10", "9000000288", "", 640#, 640#, _
                      "APFL-HR01-01-01", "APFL-HR01-01-01", "260816-02", "260816-02", _
                      "NO_QTY_CHANGE", "NONE")
    Set members = New Collection: members.Add m
    planOk = RB_Plan(members, reason)
    If planOk Then execOk = CFX_ExecuteRun(RB_OB, reason)
    creates = CLng(CFX_Meta()("Creates")): saves = CLng(CFX_Meta()("PrdoSaves"))
    RB_FinishExpect "S28", True, planOk And execOk And creates = 0 And saves = 0 And _
                    CStr(CFX_Meta()("Phase")) = "READY_FOR_CF", reason, _
                    "CHANGE_BIN_FINAL_HANDOFF|NO_MUTATION|CF_USER_HANDOFF"
End Sub

' ============================================================================
' PUBLIC ENTRY
' ============================================================================

Public Function CFX_RunAllReplay() As String
    Set g_Results = New Collection
    g_FailCount = 0: g_PassCount = 0

    Call RS_S01
    Call RS_S02
    Call RS_S03
    Call RS_S04
    Call RS_S05
    Call RS_S06
    Call RS_S07
    Call RS_S08
    Call RS_S09
    Call RS_S10
    Call RS_S11
    Call RS_S12
    Call RS_S13
    Call RS_S14
    Call RS_S15
    Call RS_S16
    Call RS_S17
    Call RS_S18
    Call RS_S19
    Call RS_S20
    Call RS_S21
    Call RS_S22
    Call RS_S23
    Call RS_S24
    Call RS_S25
    Call RS_S26
    Call RS_S27
    Call RS_S28

    Dim summary As String
    summary = "REPLAY_" & IIf(g_FailCount = 0, "ALL_PASS", "FAIL") & _
        "|TOTAL=" & CStr(g_PassCount + g_FailCount) & _
        "|PASS=" & CStr(g_PassCount) & "|FAIL=" & CStr(g_FailCount)
    CFX_RunAllReplay = summary
End Function

Public Function CFX_ReplayEvidenceCsv() As String
    Dim out As String, i As Long
    out = "Scenario,Result,GroupCount,MemberCount,DocumentOpenCount,MonOpenCount," & _
          "WHTaskClicks,WTMapRefreshes,PRDOSaves,Creates,HardStops,Warns,FinalWTStates,Notes"
    If g_Results Is Nothing Then
        CFX_ReplayEvidenceCsv = out
        Exit Function
    End If
    For i = 1 To g_Results.count
        out = out & vbCrLf & CStr(g_Results(i))
    Next i
    CFX_ReplayEvidenceCsv = out
End Function
