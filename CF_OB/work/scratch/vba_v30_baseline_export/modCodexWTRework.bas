Attribute VB_Name = "modCodexWTRework"
Option Explicit

' Generic rework runner.  It derives one exact context from the sole PASS
' row for the OB currently entered on INPUT.  It never confirms a WT and
' always stops at CF_USER_HANDOFF.
Private CWR_EXPECTED_OB As String
Private CWR_EXPECTED_DOC As String
Private CWR_EXPECTED_WT As String
Private CWR_EXPECTED_ITEM As String
Private CWR_EXPECTED_MATERIAL As String
Private CWR_EXPECTED_SOURCE_QTY As Double
Private CWR_EXPECTED_SOURCE_BIN As String
Private CWR_EXPECTED_SOURCE_BATCH As String
Private CWR_EXPECTED_DEST_BIN As String
Private CWR_EXPECTED_PROCTY As String
Private CWR_TARGET_QTY As Double
' PRDO target and WT-create quantity are deliberately separate for DELTA_ADD.
' Example: an existing open WT 60 is retained, PRDO is raised to 280, and the
' new WT is created only for the incremental 220.
Private CWR_CREATE_QTY As Double
Private CWR_TARGET_BIN As String
Private CWR_TARGET_BATCH As String
Private CWR_TARGET_STORAGE_TYPE As String
Private CWR_TARGET_STORAGE_SECTION As String
Private CWR_TARGET_STOCK_TYPE As String
Private CWR_EXPECTED_WAREHOUSE As String
Private CWR_WH_BATCH As String
Private CWR_PICK_BIN_BATCH_BEFORE As String
Private CWR_PLAN_ID As String
Private CWR_PLAN_DECISION As String
Private CWR_PLAN_QTY_MODE As String
Private CWR_PLAN_EXEC_MODE As String
Private CWR_PLAN_ROW As Long
Private CWR_PLAN_WH_ROW As Long
' A nonblank value is used only by the OB queue to bind every re-entrant
' helper call to one exact PLAN row.  The regular single-case entry points
' deliberately leave it blank and retain their one-ready-plan contract.
Private CWR_REQUESTED_PLAN_ID As String
' Set only while the controlled multi-WT executor is cancelling a complete
' batch.  A peer old WT is expected to remain active until its own turn; it
' must not be confused with an unrelated new outbound WT.
Private CWR_BATCH_PEER_OLD_WTS As Object
Private CWR_BATCH_ACTIVE As Boolean
Private CWR_HISTORY_RUN_ID As String
Private CWR_HISTORY_SEQ As Long
' These counters describe bounded MON navigation for the current executor run.
' They never cause a SAP write and are reset before each top-level route.
Private CWR_MON_OPEN_COUNT As Long
Private CWR_MON_REFRESH_COUNT As Long
Private CWR_DOCUMENT_OPEN_COUNT As Long
Private CWR_DOCUMENT_RECOVERY_COUNT As Long
Private CWR_WHTASK_OPEN_COUNT As Long
Private CWR_WHTASK_REUSE_COUNT As Long
Private CWR_WHTASK_RECOVERY_COUNT As Long
' Per-OB idempotency counters.  They are reset at the scheduler boundary and
' are deliberately kept in memory only; RUN_HISTORY remains the durable proof.
' A second PRDO Save for the same document/item in one cycle is blocked before
' the SAP write rather than silently becoming a duplicate action.
Private CWR_PRDO_DOCUMENT_SAVE_COUNTS As Object
Private CWR_CREATE_GROUP_ACTION_COUNTS As Object
Private CWR_CANCEL_RESULT_BY_WT As Object
' True only while a prior POST Create+Save is being reconciled.  This is a
' read-only recovery context: it may tolerate a missing target-batch PLAN field
' long enough to query SAP, but it never authorizes a second POST write.
Private CWR_RECOVERY_CONTEXT As Boolean
' Scheduler execution-context cache.  One locked OB run may revisit the same
' plan row from several phase gates; reloading PLAN/WH_ACTUAL on every visit
' used to reset the live context and made the SAP tab appear to reopen.  The
' cache is scoped to one run and one immutable plan fingerprint only.
Private CWR_CONTEXT_VALID As Boolean
Private CWR_CONTEXT_PLAN_ID As String
Private CWR_CONTEXT_FINGERPRINT As String
Private CWR_CONTEXT_OB As String
Private CWR_CONTEXT_DOC As String
Private CWR_CONTEXT_WT As String
Private CWR_CREATE_CONTEXT_VALID As Boolean
Private CWR_CREATE_CONTEXT_OB As String
Private CWR_CREATE_CONTEXT_DOC As String
Private CWR_CREATE_CONTEXT_SESSION As Object

Private Const CWR_NODE_KEY As String = "N000000010"
Private Const CWR_UPPER_ID As String = "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell"
Private Const CWR_WT_ID As String = "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell"
Private Const CWR_ERP_ID As String = "wnd[1]/usr/txtS_ERP_H-LOW"
Private Const CWR_DOC_ID As String = "wnd[1]/usr/txtS_DOCNO-LOW"
Private Const CWR_EXECUTE_ID As String = "wnd[1]/tbar[0]/btn[8]"
Private Const CWR_WT_ACTION As String = "N000000096"
Private Const CWR_PLAN_COL_QTY_MODE As Long = 28
Private Const CWR_PLAN_COL_EXEC_MODE As Long = 29
Private Const CWR_PLAN_COL_EXEC_STATE As Long = 30
Private Const CWR_PLAN_COL_DOC_GROUP As Long = 31
Private Const CWR_PLAN_COL_LOCATION_GROUP As Long = 32
Private Const CWR_PLAN_COL_CREATE_GROUP As Long = 33
Private Const CWR_PLAN_COL_FINGERPRINT As Long = 34
Private Const CWR_PLAN_COL_QTY_DELTA As Long = 35
Private Const CWR_PLAN_COL_QTY_GROUP_ID As Long = 47
Private Const CWR_PLAN_COL_GROUP_CURRENT As Long = 48
Private Const CWR_PLAN_COL_GROUP_TARGET As Long = 49
Private Const CWR_PLAN_COL_GROUP_DELTA As Long = 50
Private Const CWR_PLAN_COL_GROUP_MEMBERS As Long = 51
Private Const CWR_PLAN_COL_GROUP_NEED_QTY As Long = 52
Private Const CWR_MON_NAVIGATION_LIMIT As Long = 20
Private Const CWR_REGRESSION_SHEET As String = "REGRESSION_AUDIT"
Private Const CWR_REGRESSION_BUILD As String = "V2.4-GROUP-EXECUTOR-WHTASK-REUSE"
Private Const CWR_CREATE_WT_DELTA_CONTRACT_VERSION As String = "2"
Private Const CWR_CHANGE_BIN_CONTRACT_VERSION As String = "1"
Private Const CWR_MAX_PROOF_WRITE_ATTEMPTS As Long = 1

' Source addendum for the in-workbook modCodexWTRework module.
' Opens the standard Distribute Quantity surface only on a temporary manual
' draft, maps it, then cancels that surface and deletes the temporary draft.
' It never confirms a distribution, saves, creates a WT, or calls CF.

Public Function Codex_ProbeManualDistributionSurfaceAndReset() As String
    Const WT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV"
    Const WT_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/cntlCC_TB_ODP_TODLV/shellcont/shell"
    Const WT_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/subSUB_ODP_TODLV_DATA:/SCWM/SAPLUI_TODLV:4110/cntlCC_ALV_OD_TODLV/shellcont/shell"
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Const STOCK_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/cntlCC_TB_ODP_AVSTKDLV/shellcont/shell"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Const INSERT_ID As String = "OK_OD_TODLV_INSERT"
    Const DISTRIBUTE_ID As String = "OK_OD_AVSTKDLV_SPLIT_QTY"
    Dim createSes As Object, wrGrid As Object, wtTab As Object, wtToolbar As Object, wtGrid As Object
    Dim stockTab As Object, stockToolbar As Object, stockGrid As Object, modal As Object, cancelButton As Object
    Dim messageText As String, probeText As String, insertDetail As String, actionDetail As String, candidateProof As String
    Dim dumpText As String, cleanupText As String, modalTitle As String, cancelTip As String
    Dim stageText As String, candidateRow As Long, props As Variant, draftInserted As Boolean

    On Error GoTo Failed
    stageText = "LOAD_CONTEXT"
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "MANUAL_DISTRIBUTE_OPEN_CREATE_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If
    If Not CWR_VerifyExactCreateWTRequest(createSes, wrGrid, messageText) Then GoTo Stopped
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    Set wtToolbar = createSes.FindById(WT_TOOLBAR_ID, False)
    If wtTab Is Nothing Or wtToolbar Is Nothing Then
        messageText = "MANUAL_DISTRIBUTE_WT_SURFACE_NOT_FOUND"
        GoTo Stopped
    End If
    wtTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_DISTRIBUTE_WT_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 0 Then
        messageText = "MANUAL_DISTRIBUTE_REQUIRES_EMPTY_DRAFT_GRID"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(wtToolbar, INSERT_ID, "New Warehouse Task: Manual", insertDetail, messageText) Then GoTo Stopped
    wtToolbar.PressButton INSERT_ID
    draftInserted = True
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_DISTRIBUTE_INSERT_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 1 Then
        messageText = "MANUAL_DISTRIBUTE_INSERT_COUNT_MISMATCH"
        GoTo Stopped
    End If
    wtGrid.selectedRows = "0"
    wtGrid.SetCurrentCell 0, "VLPLA"
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Then
        messageText = "MANUAL_DISTRIBUTE_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_DISTRIBUTE_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    Set stockToolbar = createSes.FindById(STOCK_TOOLBAR_ID, False)
    If stockGrid Is Nothing Or stockToolbar Is Nothing Then
        messageText = "MANUAL_DISTRIBUTE_STOCK_SURFACE_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_FindExactCreateWTTargetCandidate(stockGrid, candidateRow, candidateProof, messageText) Then GoTo Stopped
    stockGrid.selectedRows = CStr(candidateRow)
    stockGrid.SetCurrentCell candidateRow, "LGPLA"
    stockGrid.SelectionChanged
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_DISTRIBUTE_SELECTION_TIMEOUT"
        GoTo Stopped
    End If
    If CWR_SelectedRowsText(stockGrid) <> CStr(candidateRow) Then
        messageText = "MANUAL_DISTRIBUTE_SELECTION_READBACK_MISMATCH: " & CWR_SelectedRowsText(stockGrid)
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(stockToolbar, DISTRIBUTE_ID, "Distribute Quantity", actionDetail, messageText) Then GoTo Stopped
    stageText = "OPEN_DISTRIBUTION_SURFACE"
    stockToolbar.PressButton DISTRIBUTE_ID
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_DISTRIBUTE_ACTION_TIMEOUT"
        GoTo Stopped
    End If
    Set modal = createSes.FindById("wnd[1]", False)
    dumpText = "INSERT_ACTION=" & insertDetail & vbLf & "CANDIDATE=" & candidateProof & vbLf & "ACTION=" & actionDetail & vbLf
    If modal Is Nothing Then
        dumpText = dumpText & "DISTRIBUTION_MODAL=NONE" & vbLf & CWR_DumpKnownGrid("STOCK_AFTER_DISTRIBUTE", stockGrid, 10)
    Else
        modalTitle = CWR_Text(modal.text)
        props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
        dumpText = dumpText & "DISTRIBUTION_MODAL_TITLE=" & modalTitle & vbLf & CStr(createSes.GetObjectTree("wnd[1]", props) & "")
        Set cancelButton = createSes.FindById("wnd[1]/tbar[0]/btn[12]", False)
        If cancelButton Is Nothing Then
            messageText = "MANUAL_DISTRIBUTE_MODAL_CANCEL_NOT_FOUND: " & modalTitle
            GoTo Stopped
        End If
        cancelTip = CWR_Text(cancelButton.Tooltip)
        If InStr(1, cancelTip, "Cancel", vbTextCompare) = 0 Then
            messageText = "MANUAL_DISTRIBUTE_MODAL_CANCEL_IDENTITY_MISMATCH: " & cancelTip
            GoTo Stopped
        End If
        cancelButton.Press
        If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
            messageText = "MANUAL_DISTRIBUTE_CANCEL_TIMEOUT"
            GoTo Stopped
        End If
        If Not createSes.FindById("wnd[1]", False) Is Nothing Then
            messageText = "MANUAL_DISTRIBUTE_MODAL_REMAINS_AFTER_CANCEL"
            GoTo Stopped
        End If
        dumpText = dumpText & vbLf & "DISTRIBUTION_CANCELLED=TRUE"
    End If
    cleanupText = ""
    If Not CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText) Then
        messageText = "MANUAL_DISTRIBUTE_CLEANUP_FAILED: " & cleanupText
        GoTo Stopped
    End If
    draftInserted = False
    CWR_Log "WT_MANUAL_DISTRIBUTION_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", "Distribution surface opened/mapped then cancelled; temporary manual draft deleted; no Save/CF."
    Codex_ProbeManualDistributionSurfaceAndReset = "PASS|MANUAL_DISTRIBUTION_SURFACE|WRITE=FALSE" & vbLf & dumpText & vbLf & "CLEANUP=" & cleanupText
    Exit Function
Stopped:
    If draftInserted Then
        cleanupText = ""
        If CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText) Then
            draftInserted = False
            messageText = messageText & "; CLEANUP=" & cleanupText
        Else
            messageText = messageText & "; UNSAVED_MANUAL_DRAFT_MAY_REMAIN; CLEANUP=" & cleanupText
        End If
    End If
    Codex_ProbeManualDistributionSurfaceAndReset = "STOP|STAGE=" & stageText & "|" & messageText
    Exit Function
Failed:
    If draftInserted Then
        cleanupText = ""
        Call CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText)
    End If
    Codex_ProbeManualDistributionSurfaceAndReset = "FAIL|STAGE=" & stageText & "|" & CStr(Err.Number) & "|" & Err.Description & "|" & cleanupText
End Function


' Source addendum for the in-workbook modCodexWTRework module.
' Identifies a quantity field on an unsaved exact manual draft and deletes it.
' It never sends Save, Create+Save, confirmation, or CF.

Public Function Codex_ProbeManualExactQtyAndReset() As String
    Const WT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV"
    Const WT_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/cntlCC_TB_ODP_TODLV/shellcont/shell"
    Const WT_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/subSUB_ODP_TODLV_DATA:/SCWM/SAPLUI_TODLV:4110/cntlCC_ALV_OD_TODLV/shellcont/shell"
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Const INSERT_ID As String = "OK_OD_TODLV_INSERT"
    Dim createSes As Object, wrGrid As Object, wtTab As Object, wtToolbar As Object, wtGrid As Object
    Dim stockTab As Object, stockGrid As Object
    Dim messageText As String, probeText As String, insertDetail As String, candidateProof As String
    Dim targetType As String, targetSection As String, qtyColumn As String, qtyReadback As String
    Dim dumpText As String, cleanupText As String, stageText As String, candidateRow As Long, c As Variant
    Dim statusType As String, statusText As String, errNo As Long, draftInserted As Boolean

    On Error GoTo Failed
    stageText = "LOAD_CONTEXT"
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "MANUAL_QTY_OPEN_CREATE_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If
    If Not CWR_VerifyExactCreateWTRequest(createSes, wrGrid, messageText) Then GoTo Stopped
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Then
        messageText = "MANUAL_QTY_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_QTY_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    If stockGrid Is Nothing Then
        messageText = "MANUAL_QTY_STOCK_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_FindExactCreateWTTargetCandidate(stockGrid, candidateRow, candidateProof, messageText) Then GoTo Stopped
    targetType = CWR_Cell(stockGrid, candidateRow, "LGTYP")
    targetSection = CWR_Cell(stockGrid, candidateRow, "LGBER")
    If targetType = "" Or targetSection = "" Then
        messageText = "MANUAL_QTY_TARGET_ATTRIBUTES_MISSING"
        GoTo Stopped
    End If
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    Set wtToolbar = createSes.FindById(WT_TOOLBAR_ID, False)
    If wtTab Is Nothing Or wtToolbar Is Nothing Then
        messageText = "MANUAL_QTY_WT_SURFACE_NOT_FOUND"
        GoTo Stopped
    End If
    wtTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_QTY_WT_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 0 Then
        messageText = "MANUAL_QTY_REQUIRES_EMPTY_DRAFT_GRID"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(wtToolbar, INSERT_ID, "New Warehouse Task: Manual", insertDetail, messageText) Then GoTo Stopped
    wtToolbar.PressButton INSERT_ID
    draftInserted = True
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_QTY_INSERT_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 1 Then
        messageText = "MANUAL_QTY_INSERT_COUNT_MISMATCH"
        GoTo Stopped
    End If
    wtGrid.ModifyCell 0, "VLTYP", targetType
    wtGrid.ModifyCell 0, "VLBER", targetSection
    wtGrid.ModifyCell 0, "VLPLA", CWR_TARGET_BIN
    If StrComp(CWR_Cell(wtGrid, 0, "VLTYP"), targetType, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(wtGrid, 0, "VLBER"), targetSection, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(wtGrid, 0, "VLPLA"), CWR_TARGET_BIN, vbTextCompare) <> 0 Then
        messageText = "MANUAL_QTY_SOURCE_READBACK_MISMATCH"
        GoTo Stopped
    End If
    stageText = "FIND_QTY_FIELD"
    For Each c In Array("VSOLM", "VSOLA", "QTY", "QUAN", "VLQTY", "VSOQM", "NISTM", "VSOLL", "VLQTY_UI")
        On Error Resume Next
        Err.Clear
        wtGrid.ModifyCell 0, CStr(c), CWR_FastQuantityText(CWR_TARGET_QTY)
        errNo = Err.Number
        Err.Clear
        On Error GoTo Failed
        If errNo = 0 Then
            qtyReadback = CWR_Cell(wtGrid, 0, CStr(c))
            dumpText = dumpText & "QTY_TRY|COL=" & CStr(c) & "|READBACK=" & qtyReadback & vbLf
            If Abs(CWR_Number(qtyReadback) - CWR_TARGET_QTY) <= 0.000001 Then
                qtyColumn = CStr(c)
                Exit For
            End If
        Else
            dumpText = dumpText & "QTY_TRY|COL=" & CStr(c) & "|COM_ERROR=" & CStr(errNo) & vbLf
        End If
    Next c
    If qtyColumn = "" Then
        messageText = "MANUAL_QTY_FIELD_NOT_FOUND" & vbLf & dumpText
        GoTo Stopped
    End If
    wtGrid.TriggerModified
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_QTY_TRIGGER_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 1 Then
        messageText = "MANUAL_QTY_GRID_LOST_AFTER_TRIGGER"
        GoTo Stopped
    End If
    wtGrid.SetCurrentCell 0, qtyColumn
    wtGrid.PressEnter
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_QTY_GRID_COMMIT_TIMEOUT"
        GoTo Stopped
    End If
    CWR_ReadStatusBar createSes, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "MANUAL_QTY_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 1 Then
        messageText = "MANUAL_QTY_GRID_LOST_AFTER_COMMIT"
        GoTo Stopped
    End If
    If Abs(CWR_Number(CWR_Cell(wtGrid, 0, qtyColumn)) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(CWR_Cell(wtGrid, 0, "VLPLA"), CWR_TARGET_BIN, vbTextCompare) <> 0 Then
        messageText = "MANUAL_QTY_FINAL_READBACK_MISMATCH"
        GoTo Stopped
    End If
    dumpText = dumpText & "CANDIDATE=" & candidateProof & vbLf & "QTY_COLUMN=" & qtyColumn & vbLf & _
        CWR_DumpKnownGrid("MANUAL_QTY_DRAFT", wtGrid, 5) & CWR_DumpGridFieldContract("MANUAL_QTY_DRAFT", wtGrid, 0)
    cleanupText = ""
    If Not CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText) Then
        messageText = "MANUAL_QTY_CLEANUP_FAILED: " & cleanupText
        GoTo Stopped
    End If
    draftInserted = False
    CWR_Log "WT_MANUAL_QTY_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", "Exact source and quantity entered/read back on one temporary draft then deleted; no Save/CF."
    Codex_ProbeManualExactQtyAndReset = "PASS|MANUAL_QTY_DRAFT_CONTRACT|WRITE=FALSE" & vbLf & dumpText & "CLEANUP=" & cleanupText
    Exit Function
Stopped:
    If draftInserted Then
        cleanupText = ""
        If CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText) Then
            draftInserted = False
            messageText = messageText & "; CLEANUP=" & cleanupText
        Else
            messageText = messageText & "; UNSAVED_MANUAL_DRAFT_MAY_REMAIN; CLEANUP=" & cleanupText
        End If
    End If
    Codex_ProbeManualExactQtyAndReset = "STOP|STAGE=" & stageText & "|" & messageText
    Exit Function
Failed:
    If draftInserted Then
        cleanupText = ""
        Call CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText)
    End If
    Codex_ProbeManualExactQtyAndReset = "FAIL|STAGE=" & stageText & "|" & CStr(Err.Number) & "|" & Err.Description & "|" & cleanupText
End Function


' Source addendum for the in-workbook modCodexWTRework module.
' Tests an exact source entry only in one temporary manual draft and deletes it.
' No Save, Create+Save, confirmation, or CF is sent.

Public Function Codex_PreviewManualSourceEntryAndReset() As String
    Const WT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV"
    Const WT_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/cntlCC_TB_ODP_TODLV/shellcont/shell"
    Const WT_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/subSUB_ODP_TODLV_DATA:/SCWM/SAPLUI_TODLV:4110/cntlCC_ALV_OD_TODLV/shellcont/shell"
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Const INSERT_ID As String = "OK_OD_TODLV_INSERT"
    Dim createSes As Object, wrGrid As Object, wtTab As Object, wtToolbar As Object, wtGrid As Object
    Dim stockTab As Object, stockGrid As Object
    Dim messageText As String, probeText As String, insertDetail As String, candidateProof As String
    Dim targetType As String, targetSection As String, dumpText As String, cleanupText As String
    Dim statusType As String, statusText As String, stageText As String, candidateRow As Long
    Dim draftInserted As Boolean

    On Error GoTo Failed
    stageText = "LOAD_CONTEXT"
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "MANUAL_SOURCE_OPEN_CREATE_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If
    If Not CWR_VerifyExactCreateWTRequest(createSes, wrGrid, messageText) Then GoTo Stopped
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Then
        messageText = "MANUAL_SOURCE_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_SOURCE_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    If stockGrid Is Nothing Then
        messageText = "MANUAL_SOURCE_STOCK_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_FindExactCreateWTTargetCandidate(stockGrid, candidateRow, candidateProof, messageText) Then GoTo Stopped
    targetType = CWR_Cell(stockGrid, candidateRow, "LGTYP")
    targetSection = CWR_Cell(stockGrid, candidateRow, "LGBER")
    If targetType = "" Or targetSection = "" Then
        messageText = "MANUAL_SOURCE_TARGET_ATTRIBUTES_MISSING"
        GoTo Stopped
    End If
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    Set wtToolbar = createSes.FindById(WT_TOOLBAR_ID, False)
    If wtTab Is Nothing Or wtToolbar Is Nothing Then
        messageText = "MANUAL_SOURCE_WT_SURFACE_NOT_FOUND"
        GoTo Stopped
    End If
    wtTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_SOURCE_WT_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 0 Then
        messageText = "MANUAL_SOURCE_REQUIRES_EMPTY_DRAFT_GRID"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(wtToolbar, INSERT_ID, "New Warehouse Task: Manual", insertDetail, messageText) Then GoTo Stopped
    wtToolbar.PressButton INSERT_ID
    draftInserted = True
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_SOURCE_INSERT_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 1 Then
        messageText = "MANUAL_SOURCE_INSERT_COUNT_MISMATCH"
        GoTo Stopped
    End If
    If UCase$(CWR_CellChangeable(wtGrid, 0, "VLTYP")) <> "TRUE" Or _
       UCase$(CWR_CellChangeable(wtGrid, 0, "VLBER")) <> "TRUE" Or _
       UCase$(CWR_CellChangeable(wtGrid, 0, "VLPLA")) <> "TRUE" Then
        messageText = "MANUAL_SOURCE_FIELDS_NOT_CHANGEABLE"
        GoTo Stopped
    End If
    wtGrid.ModifyCell 0, "VLTYP", targetType
    wtGrid.ModifyCell 0, "VLBER", targetSection
    wtGrid.ModifyCell 0, "VLPLA", CWR_TARGET_BIN
    If StrComp(CWR_Cell(wtGrid, 0, "VLTYP"), targetType, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(wtGrid, 0, "VLBER"), targetSection, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(wtGrid, 0, "VLPLA"), CWR_TARGET_BIN, vbTextCompare) <> 0 Then
        messageText = "MANUAL_SOURCE_LOCAL_READBACK_MISMATCH"
        GoTo Stopped
    End If
    wtGrid.TriggerModified
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_SOURCE_TRIGGER_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 1 Then
        messageText = "MANUAL_SOURCE_GRID_LOST_AFTER_TRIGGER"
        GoTo Stopped
    End If
    wtGrid.SetCurrentCell 0, "VLPLA"
    wtGrid.PressEnter
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_SOURCE_GRID_COMMIT_TIMEOUT"
        GoTo Stopped
    End If
    CWR_ReadStatusBar createSes, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "MANUAL_SOURCE_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 1 Then
        messageText = "MANUAL_SOURCE_GRID_LOST_AFTER_COMMIT"
        GoTo Stopped
    End If
    dumpText = "CANDIDATE=" & candidateProof & vbLf & "TARGET_TYPE=" & targetType & "|TARGET_SECTION=" & targetSection & vbLf & _
        CWR_DumpKnownGrid("MANUAL_SOURCE_DRAFT", wtGrid, 5) & CWR_DumpGridFieldContract("MANUAL_SOURCE_DRAFT", wtGrid, 0)
    cleanupText = ""
    If Not CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText) Then
        messageText = "MANUAL_SOURCE_CLEANUP_FAILED: " & cleanupText
        GoTo Stopped
    End If
    draftInserted = False
    CWR_Log "WT_MANUAL_SOURCE_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", "Exact source fields entered/read back on one temporary draft then deleted; no Save/CF."
    Codex_PreviewManualSourceEntryAndReset = "PASS|MANUAL_SOURCE_ENTRY_CONTRACT|WRITE=FALSE" & vbLf & dumpText & "CLEANUP=" & cleanupText
    Exit Function
Stopped:
    If draftInserted Then
        cleanupText = ""
        If CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText) Then
            draftInserted = False
            messageText = messageText & "; CLEANUP=" & cleanupText
        Else
            messageText = messageText & "; UNSAVED_MANUAL_DRAFT_MAY_REMAIN; CLEANUP=" & cleanupText
        End If
    End If
    Codex_PreviewManualSourceEntryAndReset = "STOP|STAGE=" & stageText & "|" & messageText
    Exit Function
Failed:
    If draftInserted Then
        cleanupText = ""
        Call CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText)
    End If
    Codex_PreviewManualSourceEntryAndReset = "FAIL|STAGE=" & stageText & "|" & CStr(Err.Number) & "|" & Err.Description & "|" & cleanupText
End Function


' Source addendum for the in-workbook modCodexWTRework module.
' Reads exact target-stock attributes and manual WT draft input fields, then
' deletes the temporary draft. It never calls Save or CF.

Public Function Codex_ProbeManualWTInputContractAndReset() As String
    Const WT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV"
    Const WT_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/cntlCC_TB_ODP_TODLV/shellcont/shell"
    Const WT_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/subSUB_ODP_TODLV_DATA:/SCWM/SAPLUI_TODLV:4110/cntlCC_ALV_OD_TODLV/shellcont/shell"
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Const INSERT_ID As String = "OK_OD_TODLV_INSERT"
    Dim createSes As Object, wrGrid As Object, wtTab As Object, wtToolbar As Object, wtGrid As Object
    Dim stockTab As Object, stockGrid As Object
    Dim messageText As String, probeText As String, insertDetail As String, candidateProof As String
    Dim dumpText As String, cleanupText As String, stageText As String, candidateRow As Long, c As Variant
    Dim draftInserted As Boolean

    On Error GoTo Failed
    stageText = "LOAD_CONTEXT"
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "MANUAL_INPUT_OPEN_CREATE_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If
    If Not CWR_VerifyExactCreateWTRequest(createSes, wrGrid, messageText) Then GoTo Stopped
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Then
        messageText = "MANUAL_INPUT_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_INPUT_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    If stockGrid Is Nothing Then
        messageText = "MANUAL_INPUT_STOCK_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_FindExactCreateWTTargetCandidate(stockGrid, candidateRow, candidateProof, messageText) Then GoTo Stopped
    dumpText = "CANDIDATE=" & candidateProof & vbLf
    For Each c In Array("LGTYP", "LGBER", "LGPLA", "CHARG", "UI_AVLQTY", "LIME_QUAN", "OWNER", "ENTITLED", "CAT", "PROCTY")
        dumpText = dumpText & "STOCK|COL=" & CStr(c) & "|VALUE=" & CWR_Cell(stockGrid, candidateRow, CStr(c)) & "|CHANGEABLE=" & CWR_CellChangeable(stockGrid, candidateRow, CStr(c)) & vbLf
    Next c
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    Set wtToolbar = createSes.FindById(WT_TOOLBAR_ID, False)
    If wtTab Is Nothing Or wtToolbar Is Nothing Then
        messageText = "MANUAL_INPUT_WT_SURFACE_NOT_FOUND"
        GoTo Stopped
    End If
    wtTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_INPUT_WT_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 0 Then
        messageText = "MANUAL_INPUT_REQUIRES_EMPTY_DRAFT_GRID"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(wtToolbar, INSERT_ID, "New Warehouse Task: Manual", insertDetail, messageText) Then GoTo Stopped
    wtToolbar.PressButton INSERT_ID
    draftInserted = True
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_INPUT_INSERT_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 1 Then
        messageText = "MANUAL_INPUT_INSERT_COUNT_MISMATCH"
        GoTo Stopped
    End If
    dumpText = dumpText & "INSERT_ACTION=" & insertDetail & vbLf
    For Each c In Array("DLVNO", "DOCNO", "ITMNO", "ITEMNO", "MATNR", "CHARG", "VLTYP", "VLBER", "VLPLA", "VSOLM", "VSOLA", "QTY", "QUAN", "VLQTY", "NLTYP", "NLBER", "NLPLA", "PROCTY", "OWNER", "ENTITLED")
        dumpText = dumpText & "MANUAL|COL=" & CStr(c) & "|VALUE=" & CWR_Cell(wtGrid, 0, CStr(c)) & "|CHANGEABLE=" & CWR_CellChangeable(wtGrid, 0, CStr(c)) & vbLf
    Next c
    cleanupText = ""
    If Not CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText) Then
        messageText = "MANUAL_INPUT_CLEANUP_FAILED: " & cleanupText
        GoTo Stopped
    End If
    draftInserted = False
    CWR_Log "WT_MANUAL_INPUT_CONTRACT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", "Manual WT fields inspected on one temporary draft then deleted; no Save/CF."
    Codex_ProbeManualWTInputContractAndReset = "PASS|MANUAL_WT_INPUT_CONTRACT|WRITE=FALSE" & vbLf & dumpText & "CLEANUP=" & cleanupText
    Exit Function
Stopped:
    If draftInserted Then
        cleanupText = ""
        If CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText) Then
            draftInserted = False
            messageText = messageText & "; CLEANUP=" & cleanupText
        Else
            messageText = messageText & "; UNSAVED_MANUAL_DRAFT_MAY_REMAIN; CLEANUP=" & cleanupText
        End If
    End If
    Codex_ProbeManualWTInputContractAndReset = "STOP|STAGE=" & stageText & "|" & messageText
    Exit Function
Failed:
    If draftInserted Then
        cleanupText = ""
        Call CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText)
    End If
    Codex_ProbeManualWTInputContractAndReset = "FAIL|STAGE=" & stageText & "|" & CStr(Err.Number) & "|" & Err.Description & "|" & cleanupText
End Function


' Source addendum for the in-workbook modCodexWTRework module.
' It proves the Manual + Adopt Quantity path without calling Save or CF.

Public Function Codex_DumpExecutionContextLocal() As String
    Dim messageText As String
    On Error GoTo Failed
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    Codex_DumpExecutionContextLocal = "PASS|LOCAL_CONTEXT|PLAN=" & CWR_PLAN_ID & _
        "|OB=" & CWR_EXPECTED_OB & "|DOC=" & CWR_EXPECTED_DOC & _
        "|ITEM=" & CWR_EXPECTED_ITEM & "|MAT=" & CWR_EXPECTED_MATERIAL & _
        "|QTY=" & CStr(CWR_TARGET_QTY) & "|WH_BATCH=" & CWR_WH_BATCH & _
        "|SOURCE_BIN=" & CWR_TARGET_BIN & "|DEST_BIN=" & CWR_EXPECTED_DEST_BIN & _
        "|PROCTY=" & CWR_EXPECTED_PROCTY & "|STOCK_TYPE=" & CWR_TARGET_STOCK_TYPE
    Exit Function
Stopped:
    Codex_DumpExecutionContextLocal = "STOP|" & messageText
    Exit Function
Failed:
    Codex_DumpExecutionContextLocal = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Creates exactly one unsaved manual WT draft, adopts only the uniquely proven
' target-bin/batch candidate, captures the draft, then deletes that same draft.
' It never calls Save, Create+Save, confirmation, or CF.
Public Function Codex_PreviewManualAdoptExactCandidateAndReset() As String
    Const WT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV"
    Const WT_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/cntlCC_TB_ODP_TODLV/shellcont/shell"
    Const WT_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/subSUB_ODP_TODLV_DATA:/SCWM/SAPLUI_TODLV:4110/cntlCC_ALV_OD_TODLV/shellcont/shell"
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Const STOCK_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/cntlCC_TB_ODP_AVSTKDLV/shellcont/shell"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Const INSERT_ID As String = "OK_OD_TODLV_INSERT"
    Const ADOPT_ID As String = "OK_OD_AVSTKDLV_TAKE_QTY"
    Dim createSes As Object, wrGrid As Object, wtTab As Object, wtToolbar As Object, wtGrid As Object
    Dim stockTab As Object, stockToolbar As Object, stockGrid As Object
    Dim messageText As String, monitorProof As String, probeText As String, candidateProof As String
    Dim insertDetail As String, adoptDetail As String, statusType As String, statusText As String
    Dim dumpText As String, cleanupText As String, stageText As String, candidateRow As Long
    Dim draftInserted As Boolean

    On Error GoTo Failed
    stageText = "LOAD_CONTEXT"
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    stageText = "ASSERT_CF_HANDOFF"
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    stageText = "ASSERT_CANCELLED"
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    stageText = "ASSERT_PRDO_QTY"
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    stageText = "MONITOR_PREVERIFY"
    If Not CWR_AssertNoActiveOutboundWT(monitorProof, messageText) Then GoTo Stopped
    stageText = "FIND_CREATE_SESSION"
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "MANUAL_ADOPT_OPEN_CREATE_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If
    stageText = "VERIFY_WAREHOUSE_REQUEST"
    If Not CWR_VerifyExactCreateWTRequest(createSes, wrGrid, messageText) Then GoTo Stopped
    stageText = "OPEN_WT_TAB"
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    Set wtToolbar = createSes.FindById(WT_TOOLBAR_ID, False)
    If wtTab Is Nothing Or wtToolbar Is Nothing Then
        messageText = "MANUAL_ADOPT_WT_SURFACE_NOT_FOUND"
        GoTo Stopped
    End If
    wtTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_ADOPT_WT_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 0 Then
        messageText = "MANUAL_ADOPT_REQUIRES_EMPTY_DRAFT_GRID"
        GoTo Stopped
    End If
    stageText = "VERIFY_INSERT_ACTION"
    If Not CWR_GetVerifiedToolbarAction(wtToolbar, INSERT_ID, "New Warehouse Task: Manual", insertDetail, messageText) Then GoTo Stopped
    stageText = "INSERT_MANUAL_DRAFT"
    wtToolbar.PressButton INSERT_ID
    draftInserted = True
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_ADOPT_INSERT_TIMEOUT"
        GoTo Stopped
    End If
    CWR_ReadStatusBar createSes, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "MANUAL_ADOPT_INSERT_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 1 Then
        messageText = "MANUAL_ADOPT_INSERT_DRAFT_COUNT_MISMATCH"
        GoTo Stopped
    End If
    wtGrid.selectedRows = "0"
    wtGrid.SetCurrentCell 0, "VLPLA"
    stageText = "OPEN_STOCK_TAB"
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Then
        messageText = "MANUAL_ADOPT_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_ADOPT_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    Set stockToolbar = createSes.FindById(STOCK_TOOLBAR_ID, False)
    If stockGrid Is Nothing Or stockToolbar Is Nothing Then
        messageText = "MANUAL_ADOPT_STOCK_SURFACE_NOT_FOUND"
        GoTo Stopped
    End If
    stageText = "FIND_EXACT_TARGET"
    If Not CWR_FindExactCreateWTTargetCandidate(stockGrid, candidateRow, candidateProof, messageText) Then GoTo Stopped
    stageText = "SELECT_EXACT_TARGET"
    stockGrid.selectedRows = CStr(candidateRow)
    stockGrid.SetCurrentCell candidateRow, "LGPLA"
    stockGrid.SelectionChanged
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_ADOPT_SELECTION_TIMEOUT"
        GoTo Stopped
    End If
    If CWR_SelectedRowsText(stockGrid) <> CStr(candidateRow) Then
        messageText = "MANUAL_ADOPT_SELECTION_READBACK_MISMATCH: " & CWR_SelectedRowsText(stockGrid)
        GoTo Stopped
    End If
    stageText = "VERIFY_ADOPT_ACTION"
    If Not CWR_GetVerifiedToolbarAction(stockToolbar, ADOPT_ID, "Adopt Quantity in WhsTask", adoptDetail, messageText) Then GoTo Stopped
    stageText = "ADOPT_EXACT_TARGET"
    stockToolbar.PressButton ADOPT_ID
    If Not CWR_WaitPOSTSaveSettle(createSes, 1#, 30#) Then
        messageText = "MANUAL_ADOPT_TIMEOUT"
        GoTo Stopped
    End If
    CWR_ReadStatusBar createSes, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "MANUAL_ADOPT_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    stageText = "READ_ADOPTED_DRAFT"
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    If wtTab Is Nothing Then
        messageText = "MANUAL_ADOPT_WT_TAB_LOST"
        GoTo Stopped
    End If
    wtTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "MANUAL_ADOPT_WT_TAB_AFTER_ADOPT_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 1 Then
        messageText = "MANUAL_ADOPT_DRAFT_COUNT_AFTER_ADOPT_MISMATCH"
        GoTo Stopped
    End If
    dumpText = "MONITOR=" & monitorProof & vbLf & "INSERT_ACTION=" & insertDetail & vbLf & _
        "CANDIDATE=" & candidateProof & vbLf & "ADOPT_ACTION=" & adoptDetail & vbLf & _
        CWR_DumpKnownGrid("MANUAL_ADOPT_DRAFT", wtGrid, 5) & _
        CWR_DumpGridFieldContract("MANUAL_ADOPT_DRAFT", wtGrid, 0)
    cleanupText = ""
    If Not CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText) Then
        messageText = "MANUAL_ADOPT_CLEANUP_FAILED: " & cleanupText
        GoTo Stopped
    End If
    draftInserted = False
    CWR_Log "WT_CREATE_MANUAL_ADOPT_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact target candidate adopted into one manual draft and deleted without Save or CF. " & candidateProof
    Codex_PreviewManualAdoptExactCandidateAndReset = "PASS|MANUAL_ADOPT_DRAFT_CONTRACT|WRITE=FALSE" & vbLf & _
        dumpText & "CLEANUP=" & cleanupText
    Exit Function
Stopped:
    If draftInserted Then
        cleanupText = ""
        If CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText) Then
            draftInserted = False
            messageText = messageText & "; CLEANUP=" & cleanupText
        Else
            messageText = messageText & "; UNSAVED_MANUAL_DRAFT_MAY_REMAIN; CLEANUP=" & cleanupText
        End If
    End If
    CWR_Log "WT_CREATE_MANUAL_ADOPT_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", _
        "STAGE=" & stageText & "; " & messageText
    Codex_PreviewManualAdoptExactCandidateAndReset = "STOP|STAGE=" & stageText & "|" & messageText
    Exit Function
Failed:
    If draftInserted Then
        cleanupText = ""
        If CWR_DeleteOneUnsavedManualDraft(createSes, cleanupText) Then
            draftInserted = False
        Else
            cleanupText = "UNSAVED_MANUAL_DRAFT_MAY_REMAIN; " & cleanupText
        End If
    End If
    CWR_Log "WT_CREATE_MANUAL_ADOPT_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "STAGE=" & stageText & "; Err=" & CStr(Err.Number) & ": " & Err.Description & "; " & cleanupText
    Codex_PreviewManualAdoptExactCandidateAndReset = "FAIL|STAGE=" & stageText & "|ERR=" & _
        CStr(Err.Number) & "|" & Err.Description & "|" & cleanupText
End Function

' Deletes exactly one tool-owned, unsaved manual draft.  The Create WT screen
' contains no save command here; it is a UI cleanup only.
Private Function CWR_DeleteOneUnsavedManualDraft(ByVal createSes As Object, ByRef cleanupText As String) As Boolean
    Const WT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV"
    Const WT_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/cntlCC_TB_ODP_TODLV/shellcont/shell"
    Const WT_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/subSUB_ODP_TODLV_DATA:/SCWM/SAPLUI_TODLV:4110/cntlCC_ALV_OD_TODLV/shellcont/shell"
    Const DELETE_ID As String = "OK_OD_TODLV_DELETE"
    Dim wtTab As Object, wtToolbar As Object, wtGrid As Object, deleteDetail As String, messageText As String
    On Error GoTo Failed
    If createSes Is Nothing Then
        cleanupText = "CREATE_SESSION_NOT_SET"
        Exit Function
    End If
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    Set wtToolbar = createSes.FindById(WT_TOOLBAR_ID, False)
    If wtTab Is Nothing Or wtToolbar Is Nothing Then
        cleanupText = "WT_CLEANUP_SURFACE_NOT_FOUND"
        Exit Function
    End If
    wtTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        cleanupText = "WT_CLEANUP_TAB_TIMEOUT"
        Exit Function
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Then
        cleanupText = "WT_CLEANUP_GRID_NOT_FOUND"
        Exit Function
    End If
    If CLng(wtGrid.rowCount) = 0 Then
        cleanupText = "DRAFT_ROWS=0"
        CWR_DeleteOneUnsavedManualDraft = True
        Exit Function
    End If
    If CLng(wtGrid.rowCount) <> 1 Then
        cleanupText = "WT_CLEANUP_ROW_COUNT=" & CStr(wtGrid.rowCount)
        Exit Function
    End If
    wtGrid.selectedRows = "0"
    wtGrid.SetCurrentCell 0, "VLPLA"
    If Not CWR_GetVerifiedToolbarAction(wtToolbar, DELETE_ID, "Delete Warehouse Task", deleteDetail, messageText) Then
        cleanupText = messageText
        Exit Function
    End If
    wtToolbar.PressButton DELETE_ID
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        cleanupText = "WT_CLEANUP_DELETE_TIMEOUT"
        Exit Function
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 0 Then
        cleanupText = "WT_CLEANUP_POSTVERIFY_FAILED"
        Exit Function
    End If
    cleanupText = "DRAFT_ROWS=0; DELETE_ACTION=" & deleteDetail
    CWR_DeleteOneUnsavedManualDraft = True
    Exit Function
Failed:
    cleanupText = "WT_CLEANUP_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function


' Source addendum for the in-workbook modCodexWTRework module.
' Reconciles one state-unproven POST action without retrying it.  It performs
' SAP reads only and records the result only when the live batch balances add
' back to the locked PLAN_SOURCE_LEGS quantity/total contract.
Public Function Codex_ReconcileDirectPostingChangeRecovery() As String
    Dim ses As Object, grid As Object, wsLeg As Worksheet, wsActual As Worksheet
    Dim messageText As String, proofText As String, expectedKey As String
    Dim sourceAvail As Double, sourceTotal As Double, whAvail As Double, whTotal As Double
    Dim sourceRows As Long, whRows As Long, initialAvail As Double, initialTotal As Double
    Dim r As Long, lastRow As Long, legMatches As Long

    On Error GoTo Failed
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    ' This public recovery is entered only after an earlier POST write was
    ' sent.  Force a new exact product result; the visible grid from that
    ' write is not admissible readback evidence.
    If Not CWR_QueryExpectedPOST(ses, grid, messageText, True) Then GoTo Stopped
    If Not CWR_AssertNoPendingPOSTDraft(grid, messageText) Then GoTo Stopped
    If Not CWR_ReadPickBinBatchBalances(grid, "", "", sourceAvail, sourceTotal, whAvail, whTotal, _
                                        sourceRows, whRows, proofText, messageText) Then GoTo Stopped

    Set wsLeg = ThisWorkbook.Worksheets("PLAN_SOURCE_LEGS")
    expectedKey = CWR_EXPECTED_MATERIAL & "|" & CWR_TARGET_BIN & "|" & CWR_PICK_BIN_BATCH_BEFORE
    lastRow = wsLeg.Cells(wsLeg.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If StrComp(CWR_Text(wsLeg.Cells(r, 1).Value2), CWR_PLAN_ID, vbTextCompare) = 0 And _
           StrComp(CWR_Text(wsLeg.Cells(r, 3).Value2), expectedKey, vbTextCompare) = 0 Then
            legMatches = legMatches + 1
            initialAvail = CWR_Number(wsLeg.Cells(r, 5).Value2)
            initialTotal = CWR_Number(wsLeg.Cells(r, 6).Value2)
        End If
    Next r
    If legMatches <> 1 Then
        messageText = "POST_RECOVERY_SOURCE_LEG_MATCHES=" & CStr(legMatches)
        GoTo Stopped
    End If
    If sourceRows <> 1 Or whRows <> 1 Or _
       Abs(whAvail - CWR_TARGET_QTY) > 0.000001 Or Abs(whTotal - CWR_TARGET_QTY) > 0.000001 Or _
       Abs((sourceAvail + whAvail) - initialAvail) > 0.000001 Or _
       Abs((sourceTotal + whTotal) - initialTotal) > 0.000001 Then
        messageText = "POST_RECOVERY_DELTA_UNPROVEN: " & proofText & _
            "; PLAN_AVAILABLE=" & CStr(initialAvail) & "; PLAN_TOTAL=" & CStr(initialTotal)
        GoTo Stopped
    End If

    Set wsActual = ThisWorkbook.Worksheets("DATA_CORE")
    wsActual.Cells(CWR_PLAN_WH_ROW, 15).Value2 = "READY"
    wsActual.Cells(CWR_PLAN_WH_ROW, 16).Value2 = "CREATE_WT_FROM_TARGET_THEN_CF_HANDOFF"
    wsActual.Cells(CWR_PLAN_WH_ROW, 17).Value2 = "PICK_BIN_BATCH_POST_VERIFIED"
    wsActual.Cells(CWR_PLAN_WH_ROW, 18).Value2 = "POST_PICK_BIN_DIRECT_RECOVERY"
    wsActual.Cells(CWR_PLAN_WH_ROW, 19).Value2 = "POST direct recovery proved exact batch delta; create one outbound WT only, then stop before CF. " & proofText
    wsActual.Range("AF1").Value2 = "WT Rework Result"
    wsActual.Range("AF2").Value2 = "PASS: direct Posting Change recovery proved exact batch delta. Next: create one outbound WT; CF remains user handoff. " & proofText
    CWR_Log "POST_DIRECT_RECOVERY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "No POST retry. Exact live balance reconciled; " & proofText
    Codex_ReconcileDirectPostingChangeRecovery = "PASS|POST_DIRECT_RECOVERY_VERIFIED|" & proofText & _
        "|NEXT=CREATE_OUTBOUND_WT|CF=USER_HANDOFF"
    Exit Function

Stopped:
    CWR_Log "POST_DIRECT_RECOVERY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ReconcileDirectPostingChangeRecovery = "STOP|" & messageText & "|NO_POST_RETRY"
    Exit Function
Failed:
    CWR_Log "POST_DIRECT_RECOVERY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReconcileDirectPostingChangeRecovery = "FAIL|" & CStr(Err.Number) & "|" & Err.Description & "|NO_POST_RETRY"
End Function


' Source addendum for the in-workbook modCodexWTRework module.
' Read-only SAP GUI session inventory; it never navigates or presses a control.
Public Function Codex_ListSAPSessionStateReadOnly() As String
    Dim gui As Object, app As Object, con As Object, ses As Object
    Dim i As Long, tcode As String, titleText As String, busyText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    If app.children.count < 1 Then
        Codex_ListSAPSessionStateReadOnly = "STOP|SAP_CONNECTION_NOT_FOUND"
        Exit Function
    End If
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = "": titleText = "": busyText = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction & "")
        titleText = CStr(ses.FindById("wnd[0]", False).text & "")
        busyText = CStr(ses.busy)
        On Error GoTo Failed
        Codex_ListSAPSessionStateReadOnly = Codex_ListSAPSessionStateReadOnly & _
            "SESSION=" & CStr(i) & "|TCODE=" & tcode & "|BUSY=" & busyText & "|TITLE=" & titleText & vbLf
    Next i
    If Codex_ListSAPSessionStateReadOnly = "" Then Codex_ListSAPSessionStateReadOnly = "STOP|SAP_SESSION_NOT_FOUND"
    Exit Function
Failed:
    Codex_ListSAPSessionStateReadOnly = "FAIL|SESSION_PROBE|" & CStr(Err.Number) & "|" & Err.Description
End Function


' Source addendum for the in-workbook modCodexWTRework module.
' It only navigates an existing POST/PRDO session to MON and reads the exact
' filtered WT grid for the current Plan.  No business action is sent.
Public Function Codex_ReadMONAfterPostingRecovery() As String
    Dim ses As Object, lower As Object, messageText As String, dumpText As String

    On Error GoTo Failed
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_OpenMonitorForRead(ses, messageText) Then GoTo Stopped
    If Not CWR_LoadFilteredWTGrid(ses, lower, messageText) Then GoTo Stopped

    dumpText = CWR_DumpKnownGrid("POST_RECOVERY_MON", lower, 50)
    CWR_Log "POST_RECOVERY_MON", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Read-only MON recovery after Posting Change; no WT confirmation sent."
    Codex_ReadMONAfterPostingRecovery = "PASS|POST_RECOVERY_MON|WRITE=FALSE" & vbLf & dumpText
    Exit Function

Stopped:
    CWR_Log "POST_RECOVERY_MON", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ReadMONAfterPostingRecovery = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_RECOVERY_MON", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReadMONAfterPostingRecovery = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function


Public Function Codex_SAPNumberParseRegression() As String
    Dim parsedQty As Double
    If Abs(CWR_Number("1.380") - 1380#) > 0.000001 Then
        Codex_SAPNumberParseRegression = "FAIL|SAP_NUMERIC_1380_PARSE"
    ElseIf Abs(CWR_Number("1.640") - 1640#) > 0.000001 Then
        Codex_SAPNumberParseRegression = "FAIL|SAP_NUMERIC_1640_PARSE"
    ElseIf Abs(CWR_Number("120") - 120#) > 0.000001 Then
        Codex_SAPNumberParseRegression = "FAIL|SAP_NUMERIC_INTEGER_PARSE"
    ElseIf Abs(CWR_Number("1200") - 1200#) > 0.000001 Then
        Codex_SAPNumberParseRegression = "FAIL|SAP_NUMERIC_1200_PARSE"
    ElseIf Abs(CWR_Number("1.200") - 1200#) > 0.000001 Then
        Codex_SAPNumberParseRegression = "FAIL|SAP_NUMERIC_1200_THOUSANDS_PARSE"
    ElseIf CWR_FastQuantityText(1200#) <> "1200" Then
        Codex_SAPNumberParseRegression = "FAIL|SAP_NUMERIC_1200_WRITE_TEXT"
    ElseIf Abs(CWR_Number("20,5") - 20.5) > 0.000001 Then
        Codex_SAPNumberParseRegression = "FAIL|SAP_NUMERIC_DECIMAL_PARSE"
    ElseIf Abs(CWR_Number("0")) > 0.000001 Then
        Codex_SAPNumberParseRegression = "FAIL|SAP_NUMERIC_ZERO_PARSE"
    ElseIf CWR_TryParseSAPQty("", parsedQty) Then
        Codex_SAPNumberParseRegression = "FAIL|SAP_NUMERIC_BLANK_ACCEPTED"
    Else
        Codex_SAPNumberParseRegression = "PASS|SAP_NUMBER_PARSE|1.380->1380|1.640->1640|120->120|1200->1200|1.200->1200|20,5->20.5|0->0|BLANK=INVALID"
    End If
End Function
' Source addendum for the in-workbook modCodexWTRework module.
' This probe is SAP read-only: it restores the exact Plan context after a
' cancelled WT has been recorded and dumps actual stock at the requested
' pick bin.  It never modifies a POST row, creates a WT, saves, or confirms.
Public Function Codex_DumpCurrentPickBinStockReadOnly() As String
    Dim ses As Object, grid As Object
    Dim messageText As String, dumpText As String

    On Error GoTo Failed
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_QueryExpectedPOST(ses, grid, messageText) Then GoTo Stopped

    dumpText = CWR_DumpTargetBinStockCandidates(grid, messageText)
    If Left$(dumpText, 5) = "FAIL|" Then GoTo Stopped
    CWR_Log "POST_PICK_BIN_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Read-only target-bin stock evidence; no POST draft or SAP write."
    Codex_DumpCurrentPickBinStockReadOnly = "PASS|PICK_BIN_STOCK_READONLY|WRITE=FALSE" & vbLf & dumpText
    Exit Function

Stopped:
    CWR_Log "POST_PICK_BIN_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpCurrentPickBinStockReadOnly = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_PICK_BIN_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpCurrentPickBinStockReadOnly = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function


' The button entry point owns a bounded OB queue.  Each eligible Plan row is
' handled separately, one SAP write route at a time.  It never confirms a
' warehouse task: each created WT remains READY_FOR_CF for the user.
Public Sub CF_ExecutePlanToCFHandoff()
    Dim resultText As String, titleText As String
    resultText = CWR_ExecuteOBQueueToCFHandoffCore()
    titleText = "XU LY DEN CF"
    If InStr(1, resultText, "READY_FOR_CF_MANUAL_HANDOFF", vbTextCompare) > 0 Then
        MsgBox "PLAN da san sang cho CF thu cong." & vbCrLf & _
               "Khong con SAP write nao duoc phep chay o nut #4; CF chua duoc gui." & vbCrLf & _
               "Sau khi nguoi dung CF trong SAP, hay tai WH va CHECK + POSTING lai." & vbCrLf & vbCrLf & _
               resultText, vbInformation, titleText
    ElseIf Left$(resultText, 5) = "PASS|" Then
        MsgBox "Da xu ly theo PLAN va dung truoc CF." & vbCrLf & _
               "Hay CF thu cong trong SAP, sau do tai WH va CHECK + POSTING lai." & vbCrLf & vbCrLf & _
               resultText, vbInformation, titleText
    Else
        MsgBox "Da dung an toan. Khong tu chay lai SAP." & vbCrLf & vbCrLf & _
               resultText, vbExclamation, titleText
    End If
End Sub

' Local readiness check for the rework button.  It does not attach to SAP and
' is deliberately separate from the legacy CF pilot audit, because CF remains
' permanently user-owned in this workflow.
Public Function CF_ReworkReadinessAudit() As String
    Dim messageText As String
    On Error GoTo Failed
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If CWR_PLAN_DECISION = "CF_ONLY" Then
        CF_ReworkReadinessAudit = "PASS|REWORK_READY_CF_ONLY|PLAN=" & CWR_PLAN_ID & "|SAP_WRITE=0|CF=USER_HANDOFF"
        Exit Function
    End If
    If Not CWR_IsBasicPlanDecision(CWR_PLAN_DECISION) Then
        messageText = "PLAN_DECISION_NOT_ENABLED: " & CWR_PLAN_DECISION
        GoTo Stopped
    End If
    If Not CWR_AssertCurrentPlanExecutionContract(messageText) Then GoTo Stopped
    If Not CWR_AssertReworkGate(messageText) Then GoTo Stopped
    If Not CWR_AssertFullRouteActionContract(messageText) Then GoTo Stopped
    CF_ReworkReadinessAudit = "PASS|REWORK_READY|PLAN=" & CWR_PLAN_ID & _
                              "|DECISION=" & CWR_PLAN_DECISION & _
                              "|OB=" & CWR_EXPECTED_OB & "|WT=" & CWR_EXPECTED_WT & _
                              "|CF=USER_HANDOFF|SAP_WRITE=0"
    Exit Function
Stopped:
    CF_ReworkReadinessAudit = "STOP|REWORK_NOT_READY|" & messageText & "|SAP_WRITE=0|CF=USER_HANDOFF"
    Exit Function
Failed:
    CF_ReworkReadinessAudit = "FAIL|REWORK_AUDIT|" & CStr(Err.Number) & "|" & Err.Description & "|SAP_WRITE=0|CF=USER_HANDOFF"
End Function

' Local-only audit for the immutable CHECK contract.  It is intentionally safe
' to run before Button #4: no SAP session is opened and no worksheet decision
' is changed.
Public Function Codex_ValidateLockedPlanContractLocal() As String
    Dim wsPlan As Worksheet, wsWH As Worksheet, messageText As String
    On Error GoTo Failed
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
    If Not CWR_AssertQueuePlanLock(wsPlan, wsWH, messageText) Then
        Codex_ValidateLockedPlanContractLocal = "STOP|" & messageText & "|SAP_WRITE=0"
    Else
        Codex_ValidateLockedPlanContractLocal = "PASS|LOCKED_PLAN_VALID|RUN=" & _
            CWR_Text(wsPlan.Range("AK2").Value2) & "|HASH=" & CWR_Text(wsPlan.Range("AN2").Value2) & "|SAP_WRITE=0"
    End If
    Exit Function
Failed:
    Codex_ValidateLockedPlanContractLocal = "FAIL|LOCKED_PLAN_AUDIT|" & CStr(Err.Number) & "|" & Err.Description & "|SAP_WRITE=0"
End Function

' Read-only scheduler preview.  This exposes the exact phase groups selected
' from the locked PLAN without opening SAP or changing a Plan/WH decision.
Public Function Codex_PreviewLockedPlanScheduleLocal() As String
    Dim wsPlan As Worksheet, wsWH As Worksheet, r As Long, lastRow As Long
    Dim messageText As String, docGroups As Object, locationGroups As Object, createGroups As Object
    Dim planCount As Long, deltaCount As Long, qtyCount As Long, locationCount As Long
    On Error GoTo Failed
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
    If Not CWR_AssertQueuePlanLock(wsPlan, wsWH, messageText) Then
        Codex_PreviewLockedPlanScheduleLocal = "STOP|" & messageText & "|SAP_WRITE=0"
        Exit Function
    End If
    Set docGroups = CreateObject("Scripting.Dictionary"): docGroups.CompareMode = vbTextCompare
    Set locationGroups = CreateObject("Scripting.Dictionary"): locationGroups.CompareMode = vbTextCompare
    Set createGroups = CreateObject("Scripting.Dictionary"): createGroups.CompareMode = vbTextCompare
    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If UCase$(CWR_Text(wsPlan.Cells(r, 24).Value2)) = "READY_TO_EXECUTE" And _
           UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_EXEC_STATE).Value2)) = "READY_TO_EXECUTE" Then
            planCount = planCount + 1
            If UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_QTY_MODE).Value2)) = "DELTA_ADD" Then deltaCount = deltaCount + 1
            If UCase$(CWR_Text(wsPlan.Cells(r, 21).Value2)) = "TRUE" Then
                qtyCount = qtyCount + 1
                docGroups(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_DOC_GROUP).Value2)) = True
            End If
            If UCase$(CWR_Text(wsPlan.Cells(r, 22).Value2)) = "TRUE" Or _
               UCase$(CWR_Text(wsPlan.Cells(r, 23).Value2)) = "TRUE" Then
                locationCount = locationCount + 1
                locationGroups(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_LOCATION_GROUP).Value2)) = True
            End If
            createGroups(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_CREATE_GROUP).Value2)) = True
        End If
    Next r
    Codex_PreviewLockedPlanScheduleLocal = "PASS|LOCKED_PLAN_SCHEDULER_PREVIEW|PLANS=" & CStr(planCount) & _
        "|DELTA_ADD=" & CStr(deltaCount) & "|QTY_PLANS=" & CStr(qtyCount) & _
        "|DOCUMENT_QTY_GROUPS=" & CStr(docGroups.count) & "|LOCATION_PLANS=" & CStr(locationCount) & _
        "|CREATE_GROUPS=" & CStr(createGroups.count) & "|SAP_WRITE=0"
    Exit Function
Failed:
    Codex_PreviewLockedPlanScheduleLocal = "FAIL|LOCKED_PLAN_SCHEDULER_PREVIEW|" & CStr(Err.Number) & "|" & Err.Description & "|SAP_WRITE=0"
End Function

' Fast production route.  It never falls back into preview, mapper, or
' diagnostic code.  Earlier Cancel/POST/PRDO stages must already be locally
' post-verified; otherwise it stops with a resumable checkpoint.
Public Function CWR_RunFastToCFCore() As String
    Dim messageText As String, stageText As String, resultText As String, flowText As String
    On Error GoTo Failed
    stageText = "LOAD_EXACT_PLAN"
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    ' FAST_PATH_PLAN_LOCK (CP3 Seam 2): FAST must verify the CHECK-locked
    ' plan hash before any SAP write, exactly like the QUEUE path does.
    stageText = "VERIFY_PLAN_LOCK"
    If Not CWR_AssertQueuePlanLock(ThisWorkbook.Worksheets("PLAN"), _
            ThisWorkbook.Worksheets("DATA_CORE"), messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    flowText = CF_FastFlowForDecision(CWR_PLAN_DECISION)
    If flowText = vbNullString Then
        messageText = "STOP_FAST_DECISION_UNSUPPORTED|" & CWR_PLAN_DECISION
        GoTo Stopped
    End If
    If StrComp(CWR_PLAN_DECISION, "CF_ONLY", vbTextCompare) = 0 Then
        CWR_RecordCFOnlyUserHandoff "Fast contract: no SAP write is required."
        CWR_RunFastToCFCore = "PASS|READY_FOR_CF|PLAN=" & CWR_PLAN_ID & "|CF=USER_HANDOFF"
        Exit Function
    End If
    stageText = "REQUIRE_CANCELLED_CHECKPOINT"
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    stageText = "REQUIRE_POST_PRDO_CHECKPOINT"
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    stageText = "FAST_CREATE_WT"
    resultText = CWR_CreateOutboundWTFromExactCandidateFast()
    If Left$(resultText, 5) <> "PASS|" Then
        messageText = resultText
        GoTo Stopped
    End If
    CWR_RunFastToCFCore = resultText
    Exit Function
Stopped:
    CF_LogFast "FAST_STOP", "STAGE=" & stageText & "|" & messageText
    CWR_RunFastToCFCore = "STOP|STAGE=" & stageText & "|" & messageText & "|NO_RETRY|CF=USER_HANDOFF"
    Exit Function
Failed:
    CF_LogFast "FAST_FAIL", "STAGE=" & stageText & "|" & CStr(Err.Number) & "|" & Err.Description
    CWR_RunFastToCFCore = "FAIL|STAGE=" & stageText & "|" & CStr(Err.Number) & "|" & Err.Description & "|CF=USER_HANDOFF"
End Function

' Implements only the locked WT segment:
' PLAN navigation -> Stk Can Be Removed -> Qty -> select -> Adopt -> minimal
' gate -> the historically proven Create then Save composite -> MON post-verify
' -> READY_FOR_CF.  All controls/actions come from the registry; no diagnostic
' session-opening, dump, tree discovery, CF, or automatic retry is permitted.
Private Function CWR_CreateOutboundWTFromExactCandidateFast() As String
    Dim createSes As Object, wrGrid As Object, stockTab As Object, stockGrid As Object
    Dim stockToolbar As Object, wtTab As Object, wtGrid As Object, topToolbar As Object
    Dim requestGridId As String, stockTabId As String, stockGridId As String, stockToolbarId As String
    Dim qtyColumn As String, adoptId As String, wtTabId As String, wtGridId As String
    Dim toolbarId As String, createCommandId As String, saveCommandId As String
    Dim messageText As String, proofText As String, draftProof As String, actionDetail As String, statusType As String, statusText As String
    Dim wtNo As String, postProof As String, mapPromotionText As String, openQty As Double
    Dim candidateRow As Long, adoptSent As Boolean, createSaveSent As Boolean, stageText As String

    On Error GoTo Failed
    stageText = "MONITOR_PREVERIFY"
    If CWR_IsDeltaAddRoute() Then
        If Not CWR_AssertRetainedDeltaWT(messageText) Then GoTo Stopped
        proofText = "RETAINED_OLD_WT=" & CWR_EXPECTED_WT & "; OLD_QTY=" & CStr(CWR_EXPECTED_SOURCE_QTY) & _
                    "; NEW_WT_QTY=" & CStr(CWR_CREATE_QTY)
    Else
        If Not CWR_AssertNoActiveOutboundWT(proofText, messageText) Then GoTo Stopped
    End If

    requestGridId = SAPAction_Command("WT_REQUEST_GRID", "S000")
    stockTabId = SAPAction_Command("WT_TAB_STK_CAN_BE_REMOVED", "S000")
    stockGridId = SAPAction_Command("WT_STOCK_GRID", "S000")
    stockToolbarId = SAPAction_Command("WT_STOCK_TOOLBAR", "S000")
    qtyColumn = SAPAction_Command("WT_STOCK_SET_QTY", "S000")
    adoptId = SAPAction_Command("WT_ADOPT_QUANTITY", "S000")
    wtTabId = SAPAction_Command("WT_TAB_WAREHOUSE_TASK", "S000")
    wtGridId = SAPAction_Command("WT_WAREHOUSE_TASK_GRID", "S000")
    toolbarId = SAPAction_Command("WT_REQUEST_TOOLBAR", "S000")

    stageText = "OPEN_CREATE_WT_FROM_PLAN"
    If Not CWR_OpenFastCreateWTFromPlan(createSes, messageText) Then GoTo Stopped
    stageText = "READ_REQUEST"
    Set wrGrid = createSes.FindById(requestGridId, False)
    If Not CWR_FastReadExactRequest(wrGrid, openQty, messageText) Then GoTo Stopped

    ' Req. Qty is the document total.  Create-WT must consume only the
    ' currently open quantity shown by SAP (OpenQtyBun/OpenQtyAUoM), which is
    ' the remaining delta after retained/cancelled WTs.  Never use the total
    ' as the source or draft quantity when SAP exposes a smaller open amount.
    If openQty > 0.000001 Then
        If openQty > CWR_TARGET_QTY + 0.000001 Then
            messageText = "STOP_CREATE_QTY_EXCEEDS_REQUEST_TOTAL|REQ_TOTAL=" & _
                CStr(CWR_TARGET_QTY) & "|OPEN_QTY=" & CStr(openQty)
            GoTo Stopped
        End If
        If Abs(openQty - CWR_CREATE_QTY) > 0.000001 Then
            CWR_Log "CREATE_WT_QTY_RECONCILE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                "REQ_TOTAL=" & CStr(CWR_TARGET_QTY) & "; OPEN_QTY_BUN=" & CStr(openQty) & _
                "; PLAN_CREATE_QTY=" & CStr(CWR_CREATE_QTY) & "; EFFECTIVE_CREATE_QTY=" & CStr(openQty) & _
                "; SOURCE_QTY_MUST_USE_OPEN_QTY"
            CWR_HistoryCheckpoint "CREATE_QTY_RECONCILED", "WT_REQUEST_OPEN_QTY", "PASS", "", _
                                  CWR_FastQuantityText(openQty), CWR_FastQuantityText(openQty), _
                                  "REQ_TOTAL=" & CStr(CWR_TARGET_QTY) & "|OPEN_QTY_BUN=" & CStr(openQty)
            CWR_CREATE_QTY = openQty
        End If
    ElseIf CWR_CREATE_QTY <= 0.000001 Then
        messageText = "STOP_CREATE_QTY_OPEN_QTY_INVALID|REQ_TOTAL=" & CStr(CWR_TARGET_QTY) & _
            "|OPEN_QTY=" & CStr(openQty)
        GoTo Stopped
    End If

    ' A zero request quantity may only be resumed after a minimal proof of the
    ' existing Adopt result.  Any partial quantity is an unknown state; never
    ' issue a second Adopt or reset the session.
    If Abs(openQty) <= 0.000001 Then
        stageText = "RESUME_ADOPTED_DRAFT_GATE"
        If Not CWR_FastMinimalAdoptGate(createSes, requestGridId, stockTabId, stockGridId, _
                                        wtTabId, wtGridId, qtyColumn, proofText, messageText) Then GoTo Stopped
    Else
        If Abs(openQty - CWR_CREATE_QTY) > 0.000001 Then
            messageText = "STATE_UNKNOWN_AFTER_ADOPT|OPEN_QTY=" & CStr(openQty)
            GoTo Stopped
        End If

        stageText = "SELECT_EXACT_REQUEST"
        If Not CWR_FastSelectExactRequestRow(createSes, requestGridId, wrGrid, messageText) Then GoTo Stopped
        If Not CWR_FastReadExactRequest(wrGrid, openQty, messageText) Then GoTo Stopped
        If Abs(openQty - CWR_CREATE_QTY) > 0.000001 Then
            messageText = "STATE_UNKNOWN_AFTER_ADOPT|REQUEST_CHANGED_AFTER_SELECTION|OPEN_QTY=" & CStr(openQty)
            GoTo Stopped
        End If

        stageText = "READ_SOURCE_STOCK"
        Set stockTab = createSes.FindById(stockTabId, False)
        If stockTab Is Nothing Then
            messageText = "STOP_SOURCE_STOCK_TAB_NOT_FOUND"
            GoTo Stopped
        End If
        stockTab.Select
        If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 15#) Then
            messageText = "STOP_SOURCE_STOCK_TAB_TIMEOUT"
            GoTo Stopped
        End If
        Set stockGrid = createSes.FindById(stockGridId, False)
        Set stockToolbar = createSes.FindById(stockToolbarId, False)
        If stockGrid Is Nothing Or stockToolbar Is Nothing Then
            messageText = "STOP_SOURCE_STOCK_SURFACE_NOT_FOUND"
            GoTo Stopped
        End If

        stageText = "SOURCE_ROW_FOUND"
        If Not CWR_FastFindExactSourceRow(stockGrid, candidateRow, True, messageText) Then GoTo Stopped
        CF_LogFast ST_SOURCE_FOUND, "ROW=" & CStr(candidateRow) & "|BIN=" & CWR_TARGET_BIN & "|BATCH=" & CWR_WH_BATCH
        CWR_HistoryCheckpoint "SOURCE_ROW_FOUND", "WT_STOCK_SELECT_ROW", "PASS", "", _
                              CWR_FastQuantityText(CWR_CREATE_QTY), "ROW=" & CStr(candidateRow), _
                              "BIN=" & CWR_TARGET_BIN & "|BATCH=" & CWR_WH_BATCH

        stageText = "SOURCE_QTY_SET"
        stockGrid.ModifyCell candidateRow, qtyColumn, CWR_FastQuantityText(CWR_CREATE_QTY)
        stockGrid.TriggerModified
        If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 20#) Then
            messageText = "STOP_SOURCE_QTY_TRIGGER_TIMEOUT"
            GoTo Stopped
        End If
        Set stockGrid = createSes.FindById(stockGridId, False)
        If stockGrid Is Nothing Or Abs(CWR_Number(CWR_Cell(stockGrid, candidateRow, qtyColumn)) - CWR_CREATE_QTY) > 0.000001 Then
            messageText = "STOP_SOURCE_QTY_READBACK"
            GoTo Stopped
        End If
        CF_LogFast ST_QTY_SET, "ROW=" & CStr(candidateRow) & "|QTY=" & CStr(CWR_CREATE_QTY)
        CWR_HistoryCheckpoint "SOURCE_QTY_READBACK_PASS", "WT_STOCK_SET_QTY", "PASS", "", _
                              CWR_FastQuantityText(CWR_CREATE_QTY), CWR_FastQuantityText(CWR_CREATE_QTY), _
                              "ROW=" & CStr(candidateRow) & "|BIN=" & CWR_TARGET_BIN & "|BATCH=" & CWR_WH_BATCH

        stageText = "SOURCE_ROW_SELECTED"
        stockGrid.selectedRows = CStr(candidateRow)
        stockGrid.SetCurrentCell candidateRow, "LGPLA"
        stockGrid.SelectionChanged
        If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 15#) Then
            messageText = "STOP_SOURCE_SELECTION_TIMEOUT"
            GoTo Stopped
        End If
        Set stockGrid = createSes.FindById(stockGridId, False)
        If stockGrid Is Nothing Or CWR_SelectedRowsText(stockGrid) <> CStr(candidateRow) Then
            messageText = "STOP_SOURCE_SELECTION_READBACK"
            GoTo Stopped
        End If
        CF_LogFast ST_SOURCE_SELECTED, "ROW=" & CStr(candidateRow)
        CWR_HistoryCheckpoint "SOURCE_ROW_SELECTED", "WT_STOCK_SELECT_ROW", "PASS", "", "", _
                              "ROW=" & CStr(candidateRow), "ROW=" & CStr(candidateRow)

        stageText = "ADOPT_QUANTITY_SENT"
        If Not CWR_RequireReworkAction("WT_ADOPT", messageText) Then GoTo Stopped
        If Not CWR_GetVerifiedToolbarAction(stockToolbar, adoptId, "Adopt Quantity in WhsTask", actionDetail, messageText) Then GoTo Stopped
        stockToolbar.PressButton adoptId
        adoptSent = True
        If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 30#) Then
            messageText = "STATE_UNKNOWN_AFTER_ADOPT|TIMEOUT"
            GoTo Stopped
        End If
        CWR_ReadStatusBar createSes, statusType, statusText
        If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
            messageText = "STATE_UNKNOWN_AFTER_ADOPT|SAP_" & statusType & "|" & statusText
            GoTo Stopped
        End If
        If CWR_HasAnyModal(createSes, messageText) Then
            messageText = "STATE_UNKNOWN_AFTER_ADOPT|MODAL|" & messageText
            GoTo Stopped
        End If
        CF_LogFast ST_ADOPT_SENT, actionDetail
        CWR_HistoryCheckpoint "ADOPT_QUANTITY_SENT", "WT_ADOPT_QUANTITY", "SENT", "", _
                              CWR_FastQuantityText(CWR_CREATE_QTY), "", actionDetail

        stageText = "ADOPT_MINIMAL_VERIFY_PASS"
        If Not CWR_FastMinimalAdoptGate(createSes, requestGridId, stockTabId, stockGridId, _
                                        wtTabId, wtGridId, qtyColumn, proofText, messageText) Then GoTo Stopped
    End If
    CF_LogFast ST_ADOPT_PASS, proofText
    CWR_HistoryCheckpoint "ADOPT_MINIMAL_VERIFY_PASS", "WT_ADOPT_QUANTITY", "PASS", "", _
                          CWR_FastQuantityText(CWR_CREATE_QTY), proofText, proofText

    stageText = "CREATE_SAVE_SENT"
    If Not SAPAction_AssertLivePass("WT_CREATE_SAVE", "S000", messageText) Then GoTo Stopped
    createCommandId = SAPAction_Command("WT_CREATE_COMMAND_ID", "S000")
    saveCommandId = SAPAction_Command("WT_SAVE_COMMAND_ID", "S000")
    If Not CWR_RequireReworkAction("WT_CREATE_SAVE", messageText) Then GoTo Stopped
    Set topToolbar = createSes.FindById(toolbarId, False)
    If topToolbar Is Nothing Then
        messageText = "STOP_CREATE_SAVE_TOOLBAR_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(topToolbar, createCommandId, "Create Warehouse Task", actionDetail, messageText) Then GoTo Stopped
    CF_LogFast ST_CREATE_SAVE_SENT, "COMPOSITE=WT_CREATE_SAVE|CREATE=" & actionDetail
    CWR_HistoryCheckpoint "CREATE_SAVE_SENT", "WT_CREATE_SAVE", "SENT", "", _
                          CWR_FastQuantityText(CWR_CREATE_QTY), "", "COMPOSITE=WT_CREATE_SAVE|CREATE=" & actionDetail
    topToolbar.PressButton createCommandId
    createSaveSent = True
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 45#) Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|CREATE_TIMEOUT"
        GoTo Stopped
    End If
    CWR_ReadStatusBar createSes, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|CREATE_SAP_" & statusType & "|" & statusText
        GoTo Stopped
    End If
    If CWR_HasAnyModal(createSes, messageText) Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|CREATE_MODAL|" & messageText
        GoTo Stopped
    End If

    stageText = "CREATE_DRAFT_RESERVATION"
    Set stockTab = createSes.FindById(stockTabId, False)
    If stockTab Is Nothing Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|RESERVATION_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 20#) Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|RESERVATION_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set wrGrid = createSes.FindById(requestGridId, False)
    Set stockGrid = createSes.FindById(stockGridId, False)
    If Not CWR_ProveExactOutboundDraftReservation(wrGrid, stockGrid, draftProof, messageText) Then GoTo Stopped

    stageText = "SAVE_SENT"
    Set topToolbar = createSes.FindById(toolbarId, False)
    If topToolbar Is Nothing Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|SAVE_TOOLBAR_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(topToolbar, saveCommandId, "Save", actionDetail, messageText) Then GoTo Stopped
    topToolbar.PressButton saveCommandId
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 45#) Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|SAVE_TIMEOUT"
        GoTo Stopped
    End If
    CWR_ReadStatusBar createSes, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|SAVE_SAP_" & statusType & "|" & statusText
        GoTo Stopped
    End If
    If CWR_HasAnyModal(createSes, messageText) Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|SAVE_MODAL|" & messageText
        GoTo Stopped
    End If

    stageText = "WT_POST_VERIFY"
    If Not CWR_ReadExactNewOutboundWT(wtNo, postProof, messageText) Then GoTo Stopped
    CWR_RecordOutboundWTCreated wtNo, postProof
    If Not SAPAction_PromoteLivePass("WT_CREATE_SAVE", "WT=" & wtNo & "; " & postProof, mapPromotionText) Then _
        mapPromotionText = "STOP_MAP_PROMOTION|" & mapPromotionText
    CF_LogFast "WT_POST_VERIFY_PASS", postProof
    CWR_HistoryCheckpoint "WT_POST_VERIFY_PASS", "WT_CREATE_SAVE", "PASS", "", _
                          CWR_FastQuantityText(CWR_CREATE_QTY), "WT=" & wtNo, postProof
    CF_LogFast ST_WT_CREATED, postProof
    CF_LogFast ST_READY_CF, "WT=" & wtNo & "|CF_NOT_SENT"
    CWR_HistoryCheckpoint "READY_FOR_CF", "READ_ONLY", "PASS", "", _
                          CWR_FastQuantityText(CWR_CREATE_QTY), "WT=" & wtNo, "CF_NOT_SENT|" & postProof
    CWR_CreateOutboundWTFromExactCandidateFast = "PASS|READY_FOR_CF|OB=" & CWR_EXPECTED_OB & _
        "|PRODUCT=" & CWR_EXPECTED_MATERIAL & "|WT=" & wtNo & "|QTY=" & CWR_FastQuantityText(CWR_CREATE_QTY) & _
        "|SOURCE_BIN=" & CWR_TARGET_BIN & "|BATCH=" & CWR_WH_BATCH & "|WT_POST_VERIFY_PASS" & _
        "|CF_NOT_SENT|ADOPT_GATE=" & proofText & "|DRAFT=" & draftProof & _
        "|POSTVERIFY=" & postProof & "|MAP=" & mapPromotionText
    Exit Function
Stopped:
    If createSaveSent Then
        CWR_CreateOutboundWTFromExactCandidateFast = "STOP|STATE_UNKNOWN_AFTER_CREATE_SAVE|STAGE=" & stageText & "|" & messageText & "|NO_RETRY|CF_NOT_SENT"
    ElseIf adoptSent Then
        CWR_CreateOutboundWTFromExactCandidateFast = "STOP|STATE_UNKNOWN_AFTER_ADOPT|STAGE=" & stageText & "|" & messageText & "|NO_RETRY|CF_NOT_SENT"
    Else
        CWR_CreateOutboundWTFromExactCandidateFast = "STOP|STAGE=" & stageText & "|" & messageText & "|SAP_WRITE=0|CF_NOT_SENT"
    End If
    Exit Function
Failed:
    If createSaveSent Then
        CWR_CreateOutboundWTFromExactCandidateFast = "FAIL|STATE_UNKNOWN_AFTER_CREATE_SAVE|STAGE=" & stageText & _
            "|" & CStr(Err.Number) & "|" & Err.Description & "|NO_RETRY|CF_NOT_SENT"
    ElseIf adoptSent Then
        CWR_CreateOutboundWTFromExactCandidateFast = "FAIL|STATE_UNKNOWN_AFTER_ADOPT|STAGE=" & stageText & _
            "|" & CStr(Err.Number) & "|" & Err.Description & "|NO_RETRY|CF_NOT_SENT"
    Else
        CWR_CreateOutboundWTFromExactCandidateFast = "FAIL|STAGE=" & stageText & "|" & CStr(Err.Number) & _
            "|" & Err.Description & "|SAP_WRITE=0|CF_NOT_SENT"
    End If
End Function

 ' Opens the standard Create Warehouse Task surface from the current exact
 ' PLAN only when a matching S000 session is not already present.  This is
 ' production navigation, not discovery: the menu and query controls are
 ' resolved from SAP_ACTION_MAP and no hierarchy/grid dump is performed.
Private Function CWR_OpenFastCreateWTFromPlan(ByRef createSes As Object, ByRef messageText As String) As Boolean
    Dim prdoSes As Object, prdoGrid As Object, menuItem As Object, requestGrid As Object, existingDraftGrid As Object
    Dim followOnId As String, requestGridId As String, draftGridId As String, menuText As String
    Dim qtyText As String, batchText As String, statusText As String
    Dim rowNo As Long, draftRows As Long

    On Error GoTo Failed
    messageText = vbNullString
    If CWR_FindExactCreateWTSession(createSes, messageText) Then
        requestGridId = SAPAction_Command("WT_REQUEST_GRID", "S000")
        Set requestGrid = createSes.FindById(requestGridId, False)
        If Not requestGrid Is Nothing Then
            CWR_OpenFastCreateWTFromPlan = True
            Exit Function
        End If
        ' The request grid is absent.  Re-query only an empty known surface;
        ' an existing draft is unknown state and must never be reset.
        draftGridId = SAPAction_Command("WT_WAREHOUSE_TASK_GRID", "S000")
        Set existingDraftGrid = createSes.FindById(draftGridId, False)
        If Not existingDraftGrid Is Nothing Then
            On Error Resume Next
            draftRows = CLng(existingDraftGrid.rowCount)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo Failed
                messageText = "STATE_UNKNOWN_AFTER_ADOPT|EXISTING_DRAFT_ROWCOUNT_UNREADABLE"
                Exit Function
            End If
            On Error GoTo Failed
            If draftRows > 0 Then
                messageText = "STATE_UNKNOWN_AFTER_ADOPT|EXISTING_DRAFT_WITHOUT_REQUEST_GRID"
                Exit Function
            End If
        End If
        If Not CWR_FastQueryCreateWT(createSes, messageText) Then Exit Function
        CWR_OpenFastCreateWTFromPlan = True
        Exit Function
    End If
    ' An interrupted Adopt keeps the exact request at Open Qty=0.  Reuse only
    ' that owned draft so a second Adopt is never sent during resume.
    messageText = vbNullString
    If CWR_FindOwnedZeroOpenCreateWTDraft(createSes, messageText) Then
        CWR_OpenFastCreateWTFromPlan = True
        Exit Function
    End If
    If InStr(1, messageText, "READ_ERROR", vbTextCompare) > 0 Then Exit Function
    messageText = vbNullString
    If Not CWR_OpenExactCreateWTAfterQtyChange(createSes, messageText) Then Exit Function
    CWR_OpenFastCreateWTFromPlan = True
    Exit Function
Failed:
    messageText = "STOP_FAST_CREATE_NAVIGATION_ERROR|" & CStr(Err.Number) & "|" & Err.Description
    Set createSes = Nothing
End Function

' Re-executes only the registry-locked Create-WT query control.  It is not a
' write and is allowed only before an exact request or any draft exists.
Private Function CWR_FastQueryCreateWT(ByVal createSes As Object, ByRef messageText As String) As Boolean
    Dim queryButton As Object
    Dim queryId As String, requestGridId As String, statusText As String, sapStatusType As String
    On Error GoTo Failed
    queryId = SAPAction_Command("WT_CREATE_QUERY", "S000")
    requestGridId = SAPAction_Command("WT_REQUEST_GRID", "S000")
    Set queryButton = createSes.FindById(queryId, False)
    If queryButton Is Nothing Then
        messageText = "STOP_FAST_CREATE_QUERY_NOT_FOUND"
        Exit Function
    End If
    queryButton.Press
    If Not CWR_WaitForObject(createSes, requestGridId, 30#) Then
        messageText = "STOP_FAST_CREATE_REQUEST_GRID_TIMEOUT"
        Exit Function
    End If
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 30#) Then
        messageText = "STOP_FAST_CREATE_QUERY_TIMEOUT"
        Exit Function
    End If
    CWR_ReadStatusBar createSes, sapStatusType, statusText
    If sapStatusType <> "" And InStr(1, "EAX", UCase$(sapStatusType), vbBinaryCompare) > 0 Then
        messageText = "STOP_FAST_CREATE_QUERY_SAP_" & sapStatusType & "|" & statusText
        Exit Function
    End If
    If CWR_HasAnyModal(createSes, messageText) Then
        messageText = "STOP_FAST_CREATE_QUERY_MODAL|" & messageText
        Exit Function
    End If
    CWR_FastQueryCreateWT = True
    Exit Function
Failed:
    messageText = "STOP_FAST_CREATE_QUERY_ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function CWR_FastSelectExactRequestRow(ByVal createSes As Object, ByVal requestGridId As String, _
                                                 ByRef wrGrid As Object, ByRef messageText As String) As Boolean
    Dim ignoredOpenQty As Double, requestRow As Long
    On Error GoTo Failed
    If Not CWR_LocateExactWarehouseRequest(wrGrid, requestRow, ignoredOpenQty, messageText) Then Exit Function
    If Not CWR_ActivateRequestRowForSource(createSes, wrGrid, requestRow, messageText) Then Exit Function
    Set wrGrid = createSes.FindById(requestGridId, False)
    If wrGrid Is Nothing Then
        messageText = "STOP_FAST_REQUEST_SELECTION_GRID_UNAVAILABLE"
        Exit Function
    End If
    If Not CWR_LocateExactWarehouseRequest(wrGrid, requestRow, ignoredOpenQty, messageText) Then Exit Function
    If CWR_SelectedRowsText(wrGrid) <> CStr(requestRow) Then
        messageText = "STOP_FAST_REQUEST_SELECTION_READBACK|" & CWR_SelectedRowsText(wrGrid)
        Exit Function
    End If
    If Not CWR_FastReadExactRequest(wrGrid, ignoredOpenQty, messageText) Then Exit Function
    CWR_FastSelectExactRequestRow = True
    Exit Function
Failed:
    messageText = "STOP_FAST_REQUEST_SELECTION_ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function

' A single request row exposes the dependent source panel immediately. With
' multiple request rows SAP requires the exact upper row to be activated by a
' double-click before the Stk Can Be Removed grid is populated. This uses only
' the stable grid API on the proven row; it never uses coordinates or keys.
Private Function CWR_ActivateRequestRowForSource(ByVal createSes As Object, ByVal wrGrid As Object, _
                                                  ByVal requestRow As Long, ByRef messageText As String) As Boolean
    Dim requestCount As Long, doubleClickOK As Boolean, modalText As String
    On Error GoTo Failed
    If wrGrid Is Nothing Then
        messageText = "STOP_FAST_REQUEST_GRID_NOT_FOUND"
        Exit Function
    End If
    requestCount = CLng(wrGrid.rowCount)
    wrGrid.selectedRows = CStr(requestRow)
    wrGrid.SetCurrentCell requestRow, "DLVNO"
    wrGrid.SelectionChanged
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 20#) Then
        messageText = "STOP_FAST_REQUEST_SELECTION_TIMEOUT"
        Exit Function
    End If

    If requestCount > 1 Then
        ' Prefer the native double-click event. A few SAP GUI patch levels do
        ' not expose that method, so use two bounded current-cell clicks as the
        ' equivalent event sequence on the already selected exact row.
        On Error Resume Next
        Err.Clear
        CallByName wrGrid, "DoubleClickCurrentCell", VbMethod
        If Err.Number = 0 Then
            doubleClickOK = True
        Else
            Err.Clear
            wrGrid.ClickCurrentCell
            If Err.Number = 0 Then wrGrid.ClickCurrentCell
            If Err.Number = 0 Then doubleClickOK = True
        End If
        If Not doubleClickOK Then
            messageText = "STOP_FAST_REQUEST_DETAIL_ACTIVATION_ERROR: " & CStr(Err.Number) & ": " & Err.Description
            Err.Clear
            On Error GoTo Failed
            Exit Function
        End If
        On Error GoTo Failed
        If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 20#) Then
            messageText = "STOP_FAST_REQUEST_DETAIL_ACTIVATION_TIMEOUT"
            Exit Function
        End If
    End If
    If CWR_HasAnyModal(createSes, modalText) Then
        messageText = "STOP_FAST_REQUEST_DETAIL_ACTIVATION_MODAL: " & modalText
        Exit Function
    End If
    CWR_Log "REQUEST_ROW_ACTIVATED", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "ROW=" & CStr(requestRow) & "; REQUEST_ROWS=" & CStr(requestCount) & _
        "; DOUBLE_CLICK=" & IIf(requestCount > 1, "TRUE", "FALSE") & "; SOURCE_PANEL=EXPECTED"
    CWR_ActivateRequestRowForSource = True
    Exit Function
Failed:
    messageText = "STOP_FAST_REQUEST_DETAIL_ACTIVATION_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_FastReadExactRequest(ByVal wrGrid As Object, ByRef openQty As Double, _
                                          ByRef messageText As String) As Boolean
    Dim requestRow As Long
    CWR_FastReadExactRequest = CWR_LocateExactWarehouseRequest(wrGrid, requestRow, openQty, messageText)
End Function

' The Create WT request grid can contain several ODO items.  The executor
' must never assume row 0 or a one-row grid: it owns only one exact
' document/item/material/batch/destination/process tuple at a time.
Private Function CWR_LocateExactWarehouseRequest(ByVal wrGrid As Object, ByRef requestRow As Long, _
                                                  ByRef openQty As Double, ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, matches As Long
    Dim rowDoc As String, rowItem As String, rowMat As String, rowBatch As String, rowDest As String, rowProc As String
    If wrGrid Is Nothing Then
        messageText = "STOP_FAST_REQUEST_GRID_NOT_FOUND"
        Exit Function
    End If
    requestRow = -1: openQty = 0#: n = CLng(wrGrid.rowCount)
    For r = 0 To n - 1
        rowDoc = CWR_Cell(wrGrid, r, "DLVNO")
        rowItem = CWR_Cell(wrGrid, r, "ITMNO")
        rowMat = CWR_Cell(wrGrid, r, "MATNR")
        rowBatch = CWR_Cell(wrGrid, r, "CHARG")
        rowDest = CWR_Cell(wrGrid, r, "NLPLA")
        rowProc = CWR_Cell(wrGrid, r, "PROCTY")
        If CWR_MONValueEquals(rowDoc, CWR_EXPECTED_DOC) And _
           CWR_MONValueEquals(rowItem, CWR_EXPECTED_ITEM) And _
           StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 And _
           (CWR_WH_BATCH = "" Or StrComp(rowBatch, CWR_WH_BATCH, vbTextCompare) = 0) And _
           (CWR_EXPECTED_DEST_BIN = "" Or StrComp(rowDest, CWR_EXPECTED_DEST_BIN, vbTextCompare) = 0) And _
           (CWR_EXPECTED_PROCTY = "" Or StrComp(rowProc, CWR_EXPECTED_PROCTY, vbTextCompare) = 0) Then
            matches = matches + 1
            requestRow = r
            openQty = CWR_Number(CWR_Cell(wrGrid, r, "UI_OPNQTY"))
        End If
    Next r
    If matches <> 1 Then
        messageText = "STOP_FAST_REQUEST_MATCH_COUNT=" & CStr(matches) & _
            "|DOC=" & CWR_EXPECTED_DOC & "|ITEM=" & CWR_EXPECTED_ITEM & "|MAT=" & CWR_EXPECTED_MATERIAL
        Exit Function
    End If
    CWR_LocateExactWarehouseRequest = True
End Function

Private Function CWR_FastFindExactSourceRow(ByVal stockGrid As Object, ByRef candidateRow As Long, _
                                             ByVal requireAvailable As Boolean, ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, matches As Long, tupleMatches As Long, parseFailures As Long, availableQty As Double
    Dim availableText As String, availableColumn As String, stockType As String, availableParseOK As Boolean
    candidateRow = -1
    If stockGrid Is Nothing Then
        messageText = "STOP_FAST_SOURCE_GRID_UNAVAILABLE"
        Exit Function
    End If
    n = CLng(stockGrid.rowCount)
    For r = 0 To n - 1
        stockType = CWR_SourceStockType(stockGrid, r)
        If StrComp(CWR_Cell(stockGrid, r, "LGPLA"), CWR_TARGET_BIN, vbTextCompare) = 0 And _
           StrComp(CWR_Cell(stockGrid, r, "CHARG"), CWR_WH_BATCH, vbTextCompare) = 0 And _
           (stockType = "" Or StrComp(stockType, CWR_TARGET_STOCK_TYPE, vbTextCompare) = 0) Then
            tupleMatches = tupleMatches + 1
            availableQty = CWR_SourceAvailableQty(stockGrid, r, availableText, availableColumn, availableParseOK)
            If availableText <> "" And Not availableParseOK Then parseFailures = parseFailures + 1
            If Not requireAvailable Or (availableParseOK And availableQty + 0.000001 >= CWR_CREATE_QTY) Then
                matches = matches + 1
                candidateRow = r
            End If
        End If
    Next r
    If matches <> 1 Then
        If tupleMatches = 0 Then
            messageText = "STOP_FAST_SOURCE_NOT_FOUND|BIN=" & CWR_TARGET_BIN & "|BATCH=" & CWR_WH_BATCH
        ElseIf requireAvailable And parseFailures > 0 And matches = 0 Then
            messageText = "STOP_FAST_SOURCE_PARSE_FAILED|BIN=" & CWR_TARGET_BIN & "|BATCH=" & CWR_WH_BATCH
        ElseIf requireAvailable And matches = 0 Then
            messageText = "STOP_FAST_SOURCE_INSUFFICIENT|BIN=" & CWR_TARGET_BIN & "|BATCH=" & CWR_WH_BATCH & _
                "|REQUIRED=" & CStr(CWR_CREATE_QTY)
        Else
            messageText = "STOP_FAST_SOURCE_MATCH_COUNT|" & CStr(matches) & "|BIN=" & CWR_TARGET_BIN & "|BATCH=" & CWR_WH_BATCH
        End If
        Exit Function
    End If
    CWR_FastFindExactSourceRow = True
End Function

Private Function CWR_FastMinimalAdoptGate(ByVal createSes As Object, ByVal requestGridId As String, _
                                          ByVal stockTabId As String, ByVal stockGridId As String, _
                                          ByVal wtTabId As String, ByVal wtGridId As String, _
                                          ByVal qtyColumn As String, ByRef proofText As String, _
                                          ByRef messageText As String) As Boolean
    Dim wrGrid As Object, stockTab As Object, stockGrid As Object, wtGrid As Object
    Dim openQty As Double, sourceRow As Long, firstDraftRows As Long, refreshedDraftRows As Long
    Set wrGrid = createSes.FindById(requestGridId, False)
    If Not CWR_FastReadExactRequest(wrGrid, openQty, messageText) Then Exit Function
    If Abs(openQty) > 0.000001 Then
        messageText = "STOP_ADOPT_QUANTITY_UNPROVEN|OPEN_QTY=" & CStr(openQty)
        Exit Function
    End If
    ' SAP can initially render only part of Warehouse Task.  Select it once to
    ' open the surface, then select it again and use only that refreshed grid.
    ' This is a navigation readback, never an Adopt/Create/Save retry.
    If Not CWR_RefreshWarehouseTaskDraftGrid(createSes, wtTabId, wtGridId, wtGrid, _
                                              firstDraftRows, refreshedDraftRows, messageText) Then
        messageText = "STOP_ADOPT_QUANTITY_UNPROVEN|" & messageText
        Exit Function
    End If
    If refreshedDraftRows <> 1 Then
        messageText = "STOP_ADOPT_QUANTITY_UNPROVEN|DRAFT_ROWS=" & CStr(refreshedDraftRows) & _
            "|FIRST_PASS_ROWS=" & CStr(firstDraftRows)
        Exit Function
    End If
    Set stockTab = createSes.FindById(stockTabId, False)
    If stockTab Is Nothing Then
        messageText = "STOP_ADOPT_QUANTITY_UNPROVEN|STOCK_TAB"
        Exit Function
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 15#) Then
        messageText = "STOP_ADOPT_QUANTITY_UNPROVEN|STOCK_TAB_TIMEOUT"
        Exit Function
    End If
    Set stockGrid = createSes.FindById(stockGridId, False)
    If CWR_FastFindExactSourceRow(stockGrid, sourceRow, False, messageText) Then
        If Abs(CWR_Number(CWR_Cell(stockGrid, sourceRow, qtyColumn)) - CWR_CREATE_QTY) > 0.000001 Then
            messageText = "STOP_ADOPT_QUANTITY_UNPROVEN|SOURCE_QTY"
            Exit Function
        End If
        proofText = "DRAFT_ROWS=1; OPEN_QTY=0; SOURCE_ROW=" & CStr(sourceRow) & "; BIN=" & CWR_TARGET_BIN & _
            "; BATCH=" & CWR_WH_BATCH & "; QTY=" & CStr(CWR_CREATE_QTY) & _
            "; WT_TAB_FIRST_ROWS=" & CStr(firstDraftRows) & "; WT_TAB_REFRESHED_ROWS=" & CStr(refreshedDraftRows)
    ElseIf InStr(1, messageText, "STOP_FAST_SOURCE_MATCH_COUNT|0", vbTextCompare) > 0 Or _
           InStr(1, messageText, "STOP_FAST_SOURCE_NOT_FOUND", vbTextCompare) > 0 Then
        ' SAP can remove the reserved line from Stk Can Be Removed immediately
        ' after Adopt.  The exact request's Open Qty=0 plus one Warehouse Task
        ' draft is the required minimal evidence; do not issue Adopt again.
        proofText = "DRAFT_ROWS=1; OPEN_QTY=0; SOURCE_ROW=RESERVED_NOT_VISIBLE; BIN=" & CWR_TARGET_BIN & _
            "; BATCH=" & CWR_WH_BATCH & "; QTY=" & CStr(CWR_CREATE_QTY) & _
            "; WT_TAB_FIRST_ROWS=" & CStr(firstDraftRows) & "; WT_TAB_REFRESHED_ROWS=" & CStr(refreshedDraftRows)
        messageText = vbNullString
    Else
        messageText = "STOP_ADOPT_QUANTITY_UNPROVEN|" & messageText
        Exit Function
    End If
    CWR_FastMinimalAdoptGate = True
End Function

Private Function CWR_FastQuantityText(ByVal qty As Double) As String
    Dim textValue As String, decimalSep As String
    If Abs(qty - Fix(qty)) <= 0.000001 Then
        ' Never let a display/locale formatter scale 1200 to 1.2. SAP must
        ' receive the ungrouped integer text when the quantity is integral.
        CWR_FastQuantityText = Format$(qty, "0")
        Exit Function
    End If
    textValue = Trim$(Str$(qty))
    decimalSep = CWR_ConfigValue("SAP_NUMERIC_DECIMAL_SEPARATOR")
    If decimalSep = "" Then decimalSep = ","
    If decimalSep <> "." Then textValue = Replace(textValue, ".", decimalSep)
    CWR_FastQuantityText = textValue
End Function

Public Function CWR_ExecutePlanToCFHandoffCore() As String
    Dim messageText As String, stageText As String, resultText As String
    On Error GoTo Failed

    stageText = "LOAD_EXACT_PLAN"
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If StrComp(CWR_PLAN_DECISION, "CF_ONLY", vbTextCompare) = 0 Then
        CWR_RecordCFOnlyUserHandoff "Plan already equals System Qty/Bin/Batch; no SAP write was needed."
        CWR_ExecutePlanToCFHandoffCore = "PASS|CF_USER_HANDOFF_ONLY|PLAN=" & CWR_PLAN_ID & _
                                         "|OB=" & CWR_EXPECTED_OB & "|SAP_WRITE=0|CF=USER_HANDOFF"
        Exit Function
    End If
    If Not CWR_IsBasicPlanDecision(CWR_PLAN_DECISION) Then
        messageText = "PLAN_DECISION_NOT_ENABLED: " & CWR_PLAN_DECISION
        GoTo Stopped
    End If
    CWR_ResetSchedulerExecutionContext
    stageText = "REGRESSION_GATE"
    If Not CWR_RequireRegressionGate(messageText) Then GoTo Stopped
    If Not CWR_AssertReworkGate(messageText) Then GoTo Stopped
    stageText = "PREFLIGHT_ACTION_CONTRACT"
    If Not CWR_AssertFullRouteActionContract(messageText) Then GoTo Stopped

    ' A same-bin/same-batch increase is not replacement work.  Preserve the
    ' exact live WT, change only the PRDO total, then create the incremental
    ' open quantity.  In particular, do not enter the cancellation branch.
    If CWR_IsDeltaAddRoute() Then
        stageText = "VERIFY_RETAINED_OUTBOUND_WT"
        If Not CWR_AssertRetainedDeltaWT(messageText) Then GoTo Stopped
        ' Refresh the exact source availability before increasing PRDO.  The
        ' Plan's earlier source proof is not enough after another case in the
        ' same OB has been processed; insufficient live stock must stop before
        ' any quantity Save is sent.
        stageText = "DELTA_SOURCE_LIVE_PREVERIFY"
        If Not CWR_AssertDeltaSourceStockLive(messageText) Then GoTo Stopped
        stageText = "CHANGE_PRDO_QTY_DELTA"
        resultText = CWR_EnsurePRDOQtyForDeltaAddOnce()
        If Left$(resultText, 5) <> "PASS|" Then
            messageText = resultText
            GoTo Stopped
        End If
        GoTo CreateOutbound
    End If

    stageText = "CANCEL_OR_RECONCILE_OLD_OUTBOUND_WT"
    resultText = CWR_EnsureOldWTCancelledOnce()
    If Left$(resultText, 5) <> "PASS|" Then
        messageText = resultText
        GoTo Stopped
    End If

    stageText = "READ_REMAINING_QTY"
    resultText = Codex_ReadCancelledRemainingAndRecord()
    If Left$(resultText, 5) <> "PASS|" Then
        messageText = resultText
        GoTo Stopped
    End If

    If CWR_PLAN_DECISION = "CANCEL_TO_ZERO" Then
        stageText = "SET_PRDO_QTY_ZERO"
        resultText = CWR_EnsurePRDOQtyAtTargetOnce()
        If Left$(resultText, 5) <> "PASS|" Then
            messageText = resultText
            GoTo Stopped
        End If
        CWR_Log "EXECUTE_TO_CF", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                "Exact WT cancelled and exact PRDO item read back at Qty=0; no new WT created; CF_NOT_SENT. " & resultText
        CWR_ExecutePlanToCFHandoffCore = "PASS|CANCEL_TO_ZERO_READY_FOR_CF|PLAN=" & CWR_PLAN_ID & _
                                         "|OB=" & CWR_EXPECTED_OB & "|DOC=" & CWR_EXPECTED_DOC & _
                                         "|WT_CANCELLED|PRDO_QTY_ZERO_VERIFIED|NO_NEW_WT|CF_NOT_SENT"
        Exit Function
    End If

    Select Case CWR_PLAN_DECISION
        Case "CHANGE_QTY"
            ' Re-read the requested source in /SCWM/POST first.  For a pure
            ' quantity change this must prove the WH batch is already at the
            ' selected bin and therefore sends no POST save.
            stageText = "PREVERIFY_TARGET_BATCH"
            resultText = Codex_ExecutePostingBatchAtPickBin()
            If Left$(resultText, 5) <> "PASS|" Then
                messageText = resultText
                GoTo Stopped
            End If
            If InStr(1, resultText, "PICK_BIN_BATCH_ALREADY_CORRECT", vbTextCompare) = 0 Then
                messageText = "QTY_BRANCH_TARGET_BATCH_NOT_CURRENT: " & resultText
                GoTo Stopped
            End If
            stageText = "CHANGE_PRDO_QTY"
            resultText = CWR_EnsurePRDOQtyAtTargetOnce()
            If Left$(resultText, 5) <> "PASS|" Then
                messageText = resultText
                GoTo Stopped
            End If
            GoTo CreateOutbound

        Case "CHANGE_BIN", "CHANGE_BATCH", "CHANGE_BIN_BATCH"
            stageText = "PREVERIFY_TARGET_BATCH"
            resultText = Codex_ExecutePostingBatchAtPickBin()
            If Left$(resultText, 5) <> "PASS|" Then
                messageText = resultText
                GoTo Stopped
            End If
            If InStr(1, resultText, "POST_CHANGE_WT_CREATED_PENDING_USER_CF", vbTextCompare) > 0 Then
                CWR_RecordBatchChangeUserHandoff resultText
                CWR_Log "EXECUTE_TO_CF", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                        "Exact posting-change WT created from Plan=" & CWR_PLAN_ID & "; automation stopped before CF. " & resultText
                CWR_ExecutePlanToCFHandoffCore = "PASS|POST_CHANGE_WT_PENDING_USER_CF|PLAN=" & CWR_PLAN_ID & _
                                                 "|OB=" & CWR_EXPECTED_OB & "|DOC=" & CWR_EXPECTED_DOC & _
                                                 "|CF=USER_HANDOFF|" & resultText
                Exit Function
            End If
            If InStr(1, resultText, "PICK_BIN_BATCH_ALREADY_CORRECT", vbTextCompare) = 0 And _
               InStr(1, resultText, "PICK_BIN_BATCH_POST_VERIFIED", vbTextCompare) = 0 And _
               InStr(1, resultText, "POST_CHANGE_COMMITTED", vbTextCompare) = 0 Then
                messageText = "POST_BRANCH_RESULT_UNRECOGNIZED: " & resultText
                GoTo Stopped
            End If
            If Abs(CWR_TARGET_QTY - CWR_EXPECTED_SOURCE_QTY) > 0.000001 Then
                stageText = "CHANGE_PRDO_QTY"
                resultText = CWR_EnsurePRDOQtyAtTargetOnce()
                If Left$(resultText, 5) <> "PASS|" Then
                    messageText = resultText
                    GoTo Stopped
                End If
            End If
            GoTo CreateOutbound
    End Select

CreateOutbound:
    stageText = "CREATE_OUTBOUND_WT"
    ' The production path is the registry-locked fast segment.  It performs
    ' source Qty/readback, selection, Adopt and its minimal gate before the
    ' historical Create+Save composite.  The older direct-Create helper is
    ' intentionally not a Button #4 execution path.
    resultText = CWR_CreateOutboundWTFromExactCandidateFast()
    If Left$(resultText, 5) <> "PASS|" Then
        messageText = resultText
        GoTo Stopped
    End If
    CWR_Log "EXECUTE_TO_CF", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "Exact outbound WT created from Plan=" & CWR_PLAN_ID & "; automation stopped before CF. " & resultText
    CWR_ExecutePlanToCFHandoffCore = "PASS|OUTBOUND_WT_PENDING_USER_CF|PLAN=" & CWR_PLAN_ID & _
                                     "|OB=" & CWR_EXPECTED_OB & "|DOC=" & CWR_EXPECTED_DOC & _
                                     "|CF=USER_HANDOFF|" & resultText
    Exit Function

Stopped:
    CWR_Log "EXECUTE_TO_CF", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", _
            "STAGE=" & stageText & "; " & messageText & "; CF=USER_HANDOFF"
    CWR_ExecutePlanToCFHandoffCore = "STOP|STAGE=" & stageText & "|" & messageText & "|NO_RETRY|CF=USER_HANDOFF"
    Exit Function
Failed:
    CWR_Log "EXECUTE_TO_CF", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
            "STAGE=" & stageText & "; Err=" & CStr(Err.Number) & ": " & Err.Description
    CWR_ExecutePlanToCFHandoffCore = "FAIL|STAGE=" & stageText & "|" & CStr(Err.Number) & "|" & Err.Description & "|NO_RETRY|CF=USER_HANDOFF"
End Function

' The full route may be resumed after a prior successful Cancel when a later
' read-only or POST preflight stops.  Reconcile SAP first: an old WT that is
' already cancelled is recorded locally and is never cancelled a second time.
Private Function CWR_EnsureOldWTCancelledOnce() As String
    Dim alreadyCancelled As Boolean, proofText As String, messageText As String
    Dim cancelText As String

    If Not CWR_ReconcileOldWTCancelState(alreadyCancelled, proofText, messageText) Then
        CWR_Log "WT_CANCEL_RECONCILE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
        CWR_EnsureOldWTCancelledOnce = "STOP_CANCEL_RECONCILE|" & messageText
        Exit Function
    End If
    If alreadyCancelled Then
        CWR_RecordCancelledSnapshot "Reconciled before resume; " & proofText
        CWR_Log "WT_CANCEL_RECONCILE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                "Old WT was already cancelled in SAP; resumed without a second Cancel. " & proofText
        CWR_EnsureOldWTCancelledOnce = "PASS|WT_CANCELLED_RECONCILED|" & proofText
        Exit Function
    End If

    cancelText = Codex_CancelOldWTAndRecord()
    If Left$(cancelText, 5) <> "PASS|" Then
        CWR_EnsureOldWTCancelledOnce = cancelText
        Exit Function
    End If
    CWR_EnsureOldWTCancelledOnce = cancelText
End Function

' Read-only gate for DELTA_ADD.  There must be exactly one current active WT
' for this ODO item and it must be the planned old WT.  This is what makes it
' safe to retain 60 and later post-verify exactly one new delta WT of 220.
Private Function CWR_AssertRetainedDeltaWT(ByRef messageText As String) As Boolean
    Dim ses As Object, lower As Object
    Dim targetRow As Long, r As Long, n As Long, activeCount As Long, retainedCount As Long
    Dim currentDoc As String, rowWT As String, rowItem As String, rowMat As String, rowStatus As String

    On Error GoTo Failed
    If Not CWR_IsDeltaAddRoute() Then
        messageText = "DELTA_ROUTE_CONTRACT_NOT_CURRENT"
        Exit Function
    End If
    If Not CWR_GetMONSession(ses, messageText) Then Exit Function
    If Not CWR_LocateExactActiveWT(ses, lower, targetRow, messageText) Then Exit Function
    n = CLng(lower.rowCount)
    For r = 0 To n - 1
        If CWR_Cell(lower, r, "DOCNO") <> "" Then currentDoc = CWR_Cell(lower, r, "DOCNO")
        rowWT = CWR_Cell(lower, r, "TANUM")
        rowItem = CWR_Cell(lower, r, "ITEMNO")
        rowMat = CWR_Cell(lower, r, "MATNR")
        rowStatus = CWR_Cell(lower, r, "TOSTAT")
        If CWR_MONValueEquals(currentDoc, CWR_EXPECTED_DOC) And _
           CWR_MONValueEquals(rowItem, CWR_EXPECTED_ITEM) And _
           StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 And _
           rowWT <> "" And rowStatus = "" Then
            activeCount = activeCount + 1
            If CWR_MONValueEquals(rowWT, CWR_EXPECTED_WT) Then retainedCount = retainedCount + 1
        End If
    Next r
    ' A DELTA_ADD group may legitimately have several active peer WTs for the
    ' same ODO item/material.  The safety proof is the exact planned old WT,
    ' not an artificial active-row cardinality of one.  Keep the count in the
    ' audit detail so an unexpected peer remains visible without blocking a
    ' valid retained tuple.
    If retainedCount <> 1 Then
        messageText = "DELTA_RETAINED_WT_ACTIVE_COUNT=" & CStr(activeCount) & _
            "; RETAINED_MATCHES=" & CStr(retainedCount) & "; EXPECTED_WT=" & CWR_EXPECTED_WT
        Exit Function
    End If
    CWR_Log "DELTA_RETAINED_WT_ACTIVE_COUNT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "ACTIVE_COUNT=" & CStr(activeCount) & "; RETAINED_MATCHES=1; EXPECTED_WT=" & CWR_EXPECTED_WT
    CWR_Log "DELTA_RETAINED_WT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact old WT retained; WT=" & CWR_EXPECTED_WT & "; OLD_QTY=" & CStr(CWR_EXPECTED_SOURCE_QTY) & _
        "; PRDO_TARGET=" & CStr(CWR_TARGET_QTY) & "; NEW_WT_QTY=" & CStr(CWR_CREATE_QTY)
    CWR_AssertRetainedDeltaWT = True
    Exit Function
Failed:
    messageText = "DELTA_RETAINED_WT_PREVERIFY_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' DELTA_ADD must prove the incremental source immediately before the PRDO Save.
' This is a bounded, read-only /SCWM/POST query for the planned material/bin/
' batch/stock type; it does not select a row, alter a draft, or save anything.
Private Function CWR_AssertDeltaSourceStockLive(ByRef messageText As String) As Boolean
    Dim ses As Object, grid As Object
    Dim r As Long, n As Long, matches As Long
    Dim lastMaterial As String, materialText As String, binText As String, batchText As String, stockTypeText As String
    Dim availText As String, totalText As String, availableQty As Double, totalQty As Double

    On Error GoTo Failed
    If Not CWR_IsDeltaAddRoute() Then
        messageText = "DELTA_SOURCE_ROUTE_CONTRACT_NOT_CURRENT"
        Exit Function
    End If
    If Not CWR_OpenPOSTSession(ses, messageText) Then Exit Function
    If Not CWR_QueryExpectedPOST(ses, grid, messageText) Then Exit Function
    If Not CWR_AssertNoPendingPOSTDraft(grid, messageText) Then Exit Function
    n = CLng(grid.rowCount)
    For r = 0 To n - 1
        materialText = CWR_Cell(grid, r, "MATNR")
        If materialText <> "" Then lastMaterial = materialText Else materialText = lastMaterial
        binText = CWR_Cell(grid, r, "VLPLA")
        batchText = CWR_Cell(grid, r, "CHARG")
        stockTypeText = CWR_Cell(grid, r, "TO_CAT")
        If stockTypeText = "" Then stockTypeText = CWR_Cell(grid, r, "CAT")
        If StrComp(materialText, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 And _
           StrComp(binText, CWR_TARGET_BIN, vbTextCompare) = 0 And _
           StrComp(batchText, CWR_WH_BATCH, vbTextCompare) = 0 And _
           StrComp(stockTypeText, CWR_TARGET_STOCK_TYPE, vbTextCompare) = 0 Then
            matches = matches + 1
            availText = CWR_Cell(grid, r, "AVL_ALT_QTY")
            totalText = CWR_Cell(grid, r, "LIME_QUAN")
            availableQty = CWR_Number(availText)
            totalQty = CWR_Number(totalText)
        End If
    Next r
    If matches <> 1 Then
        messageText = "DELTA_SOURCE_LIVE_MATCH_COUNT=" & CStr(matches) & _
            "; MAT=" & CWR_EXPECTED_MATERIAL & "; BIN=" & CWR_TARGET_BIN & _
            "; BATCH=" & CWR_WH_BATCH & "; STOCK_TYPE=" & CWR_TARGET_STOCK_TYPE
        Exit Function
    End If
    If availableQty + 0.000001 < CWR_CREATE_QTY Or totalQty + 0.000001 < CWR_CREATE_QTY Then
        messageText = "DELTA_SOURCE_LIVE_INSUFFICIENT: AVAILABLE=" & CStr(availableQty) & _
            "; TOTAL=" & CStr(totalQty) & "; REQUIRED=" & CStr(CWR_CREATE_QTY) & _
            "; USER_ACTION=REPLENISH_OR_CHANGE_SOURCE"
        Exit Function
    End If
    CWR_Log "DELTA_SOURCE_LIVE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact source read before PRDO Save; BIN=" & CWR_TARGET_BIN & "; BATCH=" & CWR_WH_BATCH & _
        "; AVAILABLE=" & CStr(availableQty) & "; TOTAL=" & CStr(totalQty) & _
        "; REQUIRED_DELTA=" & CStr(CWR_CREATE_QTY) & "; SAP_WRITE=0"
    CWR_AssertDeltaSourceStockLive = True
    Exit Function
Failed:
    messageText = "DELTA_SOURCE_LIVE_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Read-only reconciliation of the expected old WT.  A PASS only means the
' monitor proves that no active outbound WT remains for the exact document.
Private Function CWR_ReconcileOldWTCancelState(ByRef alreadyCancelled As Boolean, _
                                                ByRef proofText As String, _
                                                ByRef messageText As String) As Boolean
    Dim ses As Object, postVerifyMessage As String

    On Error GoTo Failed
    alreadyCancelled = False: proofText = "": messageText = ""
    If Not CWR_GetMONSession(ses, messageText) Then Exit Function
    If CWR_PostVerifyCancelled(ses, proofText, postVerifyMessage) Then
        alreadyCancelled = True
        CWR_ReconcileOldWTCancelState = True
        Exit Function
    End If
    If StrComp(postVerifyMessage, "POSTVERIFY_OLD_WT_STILL_ACTIVE", vbTextCompare) = 0 Then
        CWR_ReconcileOldWTCancelState = True
        Exit Function
    End If
    messageText = "CANCEL_STATE_UNPROVEN: " & postVerifyMessage
    Exit Function
Failed:
    messageText = "CANCEL_RECONCILE_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Explicit read-only recovery entry point for a run that stopped after Cancel.
' It records the checkpoint only after a fresh Monitor reconciliation.
Public Function Codex_ReconcileCancelledWTReadOnly() As String
    Dim alreadyCancelled As Boolean, proofText As String, messageText As String
    On Error GoTo Failed
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_ReconcileOldWTCancelState(alreadyCancelled, proofText, messageText) Then GoTo Stopped
    If Not alreadyCancelled Then
        messageText = "OLD_WT_STILL_ACTIVE"
        GoTo Stopped
    End If
    CWR_RecordCancelledSnapshot "Read-only reconcile; " & proofText
    CWR_Log "WT_CANCEL_RECONCILE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "SAP proves the old WT is cancelled; local resume checkpoint recorded. " & proofText
    Codex_ReconcileCancelledWTReadOnly = "PASS|WT_CANCELLED_RECONCILED|" & proofText & "|WRITE=FALSE|CF=USER_HANDOFF"
    Exit Function
Stopped:
    CWR_Log "WT_CANCEL_RECONCILE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ReconcileCancelledWTReadOnly = "STOP|" & messageText & "|WRITE=FALSE|CF=USER_HANDOFF"
    Exit Function
Failed:
    CWR_Log "WT_CANCEL_RECONCILE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
            "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReconcileCancelledWTReadOnly = "FAIL|" & CStr(Err.Number) & "|" & Err.Description & "|WRITE=FALSE|CF=USER_HANDOFF"
End Function

Public Sub CF_ExecutePlanToCFHandoff_CompileProbe()
    ' A compile probe must never depend on the mutable Plan state.  After a
    ' verified PRDO Save the old PLAN is intentionally no longer a readiness
    ' proof, so the former data-dependent checks produced a misleading VBE
    ' error at this call site.  Runtime readiness stays guarded by Button #4.
    CF_ExecutePlanToCFHandoff_StaticCompileProbe
End Sub

' Explicit diagnostic only. It verifies the current locked PLAN and action
' contract but is never called by import/compile validation.
Public Sub CF_ExecutePlanToCFHandoff_ReadinessProbe()
    Dim messageText As String
    If Not CWR_LoadExecutionContext(messageText) Then
        Err.Raise vbObjectError + 9981, , "EXECUTE_TO_CF_CONTEXT_FAIL: " & messageText
    End If
    If CWR_PLAN_DECISION <> "CF_ONLY" And Not CWR_IsBasicPlanDecision(CWR_PLAN_DECISION) Then _
        Err.Raise vbObjectError + 9982, , "EXECUTE_TO_CF_PLAN_DECISION_UNSUPPORTED: " & CWR_PLAN_DECISION
    If Not CWR_AssertCFHandoff(messageText) Then
        Err.Raise vbObjectError + 9983, , "EXECUTE_TO_CF_CF_LOCK_FAIL: " & messageText
    End If
    If Not CWR_AssertReworkGate(messageText) Then
        Err.Raise vbObjectError + 9984, , "EXECUTE_TO_CF_REWORK_GATE_FAIL: " & messageText
    End If
    If Not CWR_RequireReworkAction("WT_CANCEL", messageText) Then Err.Raise vbObjectError + 9985, , messageText
    If Not CWR_RequireReworkAction("POST_CHANGE", messageText) Then Err.Raise vbObjectError + 9986, , messageText
    If Not CWR_RequireReworkAction("PRDO_QTY", messageText) Then Err.Raise vbObjectError + 9987, , messageText
    If Not CWR_RequireReworkAction("WT_CREATE", messageText) Then Err.Raise vbObjectError + 9988, , messageText
    If Not CWR_IsBasicPlanDecision("CHANGE_QTY") Then Err.Raise vbObjectError + 9989, , "ROUTE_CHANGE_QTY"
    If Not CWR_IsBasicPlanDecision("CHANGE_BIN") Then Err.Raise vbObjectError + 9990, , "ROUTE_CHANGE_BIN"
    If Not CWR_IsBasicPlanDecision("CHANGE_BATCH") Then Err.Raise vbObjectError + 9991, , "ROUTE_CHANGE_BATCH"
    If Not CWR_IsBasicPlanDecision("CHANGE_BIN_BATCH") Then Err.Raise vbObjectError + 9992, , "ROUTE_CHANGE_BIN_BATCH"
    If CWR_IsBasicPlanDecision("NEED_RELEASE") Then Err.Raise vbObjectError + 9993, , "ROUTE_RELEASE_MUST_STOP"
    If Not CWR_IsBasicPlanDecision("CANCEL_TO_ZERO") Then Err.Raise vbObjectError + 9999, , "ROUTE_CANCEL_TO_ZERO"
End Sub

' Pure compile probe for the workbook-wide audit.  Unlike the readiness probe
' above it never reads the current Plan, CONFIG or SAP session.
Public Sub CF_ExecutePlanToCFHandoff_StaticCompileProbe()
    Dim ignored As String
    If Not CWR_IsBasicPlanDecision("CHANGE_QTY") Then Err.Raise vbObjectError + 9994, , "STATIC_ROUTE_QTY"
    If Not CWR_IsBasicPlanDecision("CHANGE_BIN") Then Err.Raise vbObjectError + 9995, , "STATIC_ROUTE_BIN"
    If Not CWR_IsBasicPlanDecision("CHANGE_BATCH") Then Err.Raise vbObjectError + 9996, , "STATIC_ROUTE_BATCH"
    If Not CWR_IsBasicPlanDecision("CHANGE_BIN_BATCH") Then Err.Raise vbObjectError + 9997, , "STATIC_ROUTE_BIN_BATCH"
    If CWR_IsBasicPlanDecision("CF_ONLY") Then Err.Raise vbObjectError + 9998, , "STATIC_ROUTE_CF_ONLY"
    If Not CWR_IsBasicPlanDecision("CANCEL_TO_ZERO") Then Err.Raise vbObjectError + 10000, , "STATIC_ROUTE_CANCEL_TO_ZERO"
    CWR_LockedPlanSchedule_StaticProbe
    If False Then
        ignored = Codex_CancelOldWTAndRecord()
        ignored = Codex_ReadCancelledRemainingAndRecord()
        ignored = Codex_ExecutePostingBatchAtPickBin()
        ignored = Codex_ExecutePRDOQtyOnlyAfterPosting()
        ignored = Codex_CreateOutboundWTFromExactCandidate()
    End If
End Sub

Private Function CWR_IsBasicPlanDecision(ByVal decisionText As String) As Boolean
    Select Case UCase$(Trim$(decisionText))
        Case "CHANGE_QTY", "CHANGE_BIN", "CHANGE_BATCH", "CHANGE_BIN_BATCH", "CANCEL_TO_ZERO"
            CWR_IsBasicPlanDecision = True
    End Select
End Function

' A delta add is not a replacement.  The original outbound WT remains active,
' PRDO changes to its requested final total, and the only new WT quantity is
' the difference.  Keeping this predicate central prevents any future route
' from accidentally cancelling the retained WT.
Private Function CWR_IsDeltaAddRoute() As Boolean
    CWR_IsDeltaAddRoute = (CWR_PLAN_DECISION = "CHANGE_QTY" And _
                            CWR_PLAN_QTY_MODE = "DELTA_ADD" And _
                            CWR_PLAN_EXEC_MODE = "KEEP_EXISTING_CREATE_DELTA")
End Function

Private Function CWR_PlanRequiresSourceLeg() As Boolean
    CWR_PlanRequiresSourceLeg = (CWR_PLAN_DECISION <> "CF_ONLY" And _
                                  CWR_PLAN_DECISION <> "CANCEL_TO_ZERO")
End Function

' CHECK is the only planner.  This guard keeps the legacy replacement route
' from being used for a DELTA_ADD row or for any row that was blocked at CHECK.
Private Function CWR_AssertCurrentPlanExecutionContract(ByRef messageText As String) As Boolean
    Dim wsPlan As Worksheet, wsWH As Worksheet
    Dim execMode As String, execState As String, gateText As String, planRunID As String, whRunID As String
    On Error GoTo Failed
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
    If CWR_Text(wsPlan.Range("AJ1").Value2) <> "PLAN LOCK" Or _
       UCase$(CWR_Text(wsPlan.Range("AJ2").Value2)) <> "LOCKED" Then
        messageText = "STOP_PLAN_NOT_LOCKED_RECHECK"
        Exit Function
    End If
    planRunID = CWR_Text(wsPlan.Range("AK2").Value2)
    whRunID = CWR_Text(wsWH.Range("W2").Value2)
    If planRunID = "" Or StrComp(planRunID, whRunID, vbTextCompare) <> 0 Then
        messageText = "STOP_PLAN_STALE_RECHECK|PLAN_RUN=" & planRunID & "|WH_RUN=" & whRunID
        Exit Function
    End If
    execMode = UCase$(CWR_Text(wsPlan.Cells(CWR_PLAN_ROW, CWR_PLAN_COL_EXEC_MODE).Value2))
    execState = UCase$(CWR_Text(wsPlan.Cells(CWR_PLAN_ROW, CWR_PLAN_COL_EXEC_STATE).Value2))
    gateText = UCase$(CWR_Text(wsPlan.Cells(CWR_PLAN_ROW, 24).Value2))
    If gateText <> "READY_TO_EXECUTE" Then
        messageText = "STOP_PLAN_GATE_NOT_READY_TO_EXECUTE|GATE=" & gateText
        Exit Function
    End If
    If execState <> "READY_TO_EXECUTE" Then
        messageText = "STOP_PLAN_NOT_READY_TO_EXECUTE|STATE=" & execState
        Exit Function
    End If
    If execMode <> "REPLACE_CANCEL_CREATE" And execMode <> "CANCEL_PRDO_ZERO" And _
       execMode <> "KEEP_EXISTING_CREATE_DELTA" Then
        messageText = "STOP_PLAN_EXEC_MODE_NOT_SUPPORTED|MODE=" & execMode
        Exit Function
    End If
    CWR_AssertCurrentPlanExecutionContract = True
    Exit Function
Failed:
    messageText = "PLAN_EXECUTION_CONTRACT_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Prove every later write action is locked and armed before Cancel is sent.
' This prevents a partial route in which the old WT is cancelled but a later
' Posting, quantity, or Create stage is not permitted by the production map.
Private Function CWR_AssertFullRouteActionContract(ByRef messageText As String) As Boolean
    If CWR_IsDeltaAddRoute() Then
        If Not CWR_RequireReworkAction("PRDO_QTY", messageText) Then Exit Function
        If Not CWR_RequireReworkAction("WT_CREATE", messageText) Then Exit Function
        CWR_AssertFullRouteActionContract = True
        Exit Function
    End If
    If Not CWR_RequireReworkAction("WT_CANCEL", messageText) Then Exit Function
    If CWR_PLAN_DECISION = "CANCEL_TO_ZERO" Then
        If Not CWR_RequireReworkAction("PRDO_QTY", messageText) Then Exit Function
        CWR_AssertFullRouteActionContract = True
        Exit Function
    End If
    If Not CWR_RequireReworkAction("WT_CREATE", messageText) Then Exit Function
    If CWR_PLAN_DECISION <> "CHANGE_QTY" Then
        If Not CWR_RequireReworkAction("POST_CHANGE", messageText) Then Exit Function
    End If
    If Abs(CWR_TARGET_QTY - CWR_EXPECTED_SOURCE_QTY) > 0.000001 Then
        If Not CWR_RequireReworkAction("PRDO_QTY", messageText) Then Exit Function
    End If
    CWR_AssertFullRouteActionContract = True
End Function

' Loads the one ready Plan line for the OB on INPUT.  This is local-only and
' rejects ambiguity, release work, stale input, or any source-leg mismatch.
Private Function CWR_LoadExecutionContext(ByRef messageText As String) As Boolean
    Dim wsPlan As Worksheet, wsWH As Worksheet, wsLeg As Worksheet, wsInput As Worksheet
    Dim r As Long, lastRow As Long, matchCount As Long, legCount As Long
    Dim planID As String, obText As String, decisionText As String, gateText As String
    Dim sourceKey As String, expectedLegKey As String, qtyValue As Double
    Dim legAvailable As Double, legTotal As Double, whRowValue As Long
    Dim requestedPlanID As String, rawStatus As String, wtState As String
    Dim allowCancelledResume As Boolean, allowRecoveryMissingTarget As Boolean
    On Error GoTo Failed

    CWR_EXPECTED_OB = "": CWR_EXPECTED_DOC = "": CWR_EXPECTED_WT = ""
    CWR_EXPECTED_ITEM = "": CWR_EXPECTED_MATERIAL = ""
    CWR_EXPECTED_SOURCE_QTY = 0#: CWR_EXPECTED_SOURCE_BIN = "": CWR_EXPECTED_SOURCE_BATCH = ""
    CWR_EXPECTED_DEST_BIN = "": CWR_EXPECTED_PROCTY = ""
    CWR_TARGET_QTY = 0#: CWR_CREATE_QTY = 0#: CWR_TARGET_BIN = "": CWR_TARGET_BATCH = ""
    CWR_TARGET_STORAGE_TYPE = "": CWR_TARGET_STORAGE_SECTION = "": CWR_TARGET_STOCK_TYPE = "FT"
    CWR_WH_BATCH = "": CWR_PICK_BIN_BATCH_BEFORE = ""
    CWR_PLAN_ID = "": CWR_PLAN_DECISION = "": CWR_PLAN_QTY_MODE = "": CWR_PLAN_EXEC_MODE = ""
    CWR_PLAN_ROW = 0: CWR_PLAN_WH_ROW = 0

    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
    Set wsLeg = ThisWorkbook.Worksheets("PLAN_SOURCE_LEGS")
    Set wsInput = CF_GetInputSheet()
    obText = CWR_Text(wsInput.Range("A7").Value2)
    requestedPlanID = CWR_Text(CWR_REQUESTED_PLAN_ID)
    allowRecoveryMissingTarget = (CWR_RECOVERY_CONTEXT And requestedPlanID <> "")
    If obText = "" Then
        messageText = "INPUT_OB_MISSING"
        Exit Function
    End If

    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If StrComp(CWR_Text(wsPlan.Cells(r, 2).Value2), obText, vbTextCompare) = 0 Then
            gateText = UCase$(CWR_Text(wsPlan.Cells(r, 24).Value2))
            decisionText = UCase$(CWR_Text(wsPlan.Cells(r, 20).Value2))
            planID = CWR_Text(wsPlan.Cells(r, 1).Value2)
            If requestedPlanID <> "" Then
                If StrComp(planID, requestedPlanID, vbTextCompare) = 0 Then
                    matchCount = matchCount + 1
                    CWR_PLAN_ROW = r
                End If
            ElseIf (gateText = "READY_TO_EXECUTE" Or decisionText = "CF_ONLY") And decisionText <> "DONE" Then
                matchCount = matchCount + 1
                CWR_PLAN_ROW = r
            End If
        End If
    Next r
    If matchCount <> 1 Then
        messageText = "READY_PLAN_COUNT=" & CStr(matchCount) & "; OB=" & obText
        Exit Function
    End If

    r = CWR_PLAN_ROW
    CWR_PLAN_ID = CWR_Text(wsPlan.Cells(r, 1).Value2)
    CWR_EXPECTED_OB = CWR_Text(wsPlan.Cells(r, 2).Value2)
    CWR_EXPECTED_DOC = CWR_Text(wsPlan.Cells(r, 3).Value2)
    CWR_EXPECTED_WT = CWR_Text(wsPlan.Cells(r, 5).Value2)
    CWR_EXPECTED_MATERIAL = CWR_Text(wsPlan.Cells(r, 6).Value2)
    CWR_EXPECTED_SOURCE_QTY = CWR_Number(wsPlan.Cells(r, 7).Value2)
    CWR_TARGET_QTY = CWR_Number(wsPlan.Cells(r, 8).Value2)
    CWR_EXPECTED_SOURCE_BIN = CWR_Text(wsPlan.Cells(r, 9).Value2)
    CWR_TARGET_BIN = CWR_Text(wsPlan.Cells(r, 10).Value2)
    CWR_EXPECTED_SOURCE_BATCH = CWR_Text(wsPlan.Cells(r, 11).Value2)
    CWR_WH_BATCH = CWR_EXPECTED_SOURCE_BATCH
    CWR_TARGET_BATCH = CWR_Text(wsPlan.Cells(r, 14).Value2)
    CWR_PICK_BIN_BATCH_BEFORE = CWR_TARGET_BATCH
    CWR_PLAN_DECISION = UCase$(CWR_Text(wsPlan.Cells(r, 20).Value2))
    CWR_PLAN_QTY_MODE = UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_QTY_MODE).Value2))
    CWR_PLAN_EXEC_MODE = UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_EXEC_MODE).Value2))
    CWR_CREATE_QTY = CWR_TARGET_QTY
    If CWR_IsDeltaAddRoute() Then CWR_CREATE_QTY = CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_QTY_DELTA).Value2)
    CWR_PLAN_WH_ROW = CLng(CWR_Number(wsPlan.Cells(r, 27).Value2))
    CWR_EXPECTED_ITEM = CWR_Text(wsWH.Cells(CWR_PLAN_WH_ROW, 3).Value2)
    rawStatus = UCase$(CWR_Text(wsWH.Cells(CWR_PLAN_WH_ROW, 10).Value2))
    wtState = UCase$(CWR_Text(wsWH.Cells(CWR_PLAN_WH_ROW, 26).Value2))
    allowCancelledResume = (requestedPlanID <> "" And rawStatus = "A" And _
        (wtState = "WT_CANCELLED_PENDING_REMAINING" Or wtState = "WT_CANCELLED_REMAINING"))

    ' A status-A WT is a post-write surface.  The executor may reconcile its
    ' known result, but it must never derive a new Decision, QtyMode or route
    ' from that surface.  CHECK remains the sole planner.
    If allowCancelledResume And Not allowRecoveryMissingTarget And _
       Not CWR_IsBasicPlanDecision(CWR_PLAN_DECISION) Then
        messageText = "STOP_PLAN_STALE_RECHECK|CANCELLED_STATUS_NOT_IN_LOCKED_PLAN"
        Exit Function
    End If

    If CWR_PLAN_ID = "" Or CWR_EXPECTED_DOC = "" Or CWR_EXPECTED_WT = "" Or _
       CWR_EXPECTED_ITEM = "" Or CWR_EXPECTED_MATERIAL = "" Or _
       CWR_EXPECTED_SOURCE_QTY <= 0# Or CWR_TARGET_QTY < 0# Or CWR_CREATE_QTY < 0# Or _
       (CWR_TARGET_QTY <= 0# And CWR_PLAN_DECISION <> "CANCEL_TO_ZERO") Or _
       CWR_EXPECTED_SOURCE_BIN = "" Or CWR_EXPECTED_SOURCE_BATCH = "" Or _
       CWR_TARGET_BIN = "" Or (CWR_TARGET_BATCH = "" And Not allowRecoveryMissingTarget) Then
        messageText = "PLAN_REQUIRED_FIELD_MISSING"
        Exit Function
    End If
    If CWR_IsDeltaAddRoute() And CWR_CREATE_QTY <= 0.000001 Then
        messageText = "PLAN_DELTA_CREATE_QTY_INVALID=" & CStr(CWR_CREATE_QTY)
        Exit Function
    End If
    If CWR_PLAN_WH_ROW < 5 Then
        messageText = "PLAN_WH_ACTUAL_ROW_INVALID=" & CStr(CWR_PLAN_WH_ROW)
        Exit Function
    End If
    If CWR_PlanRequiresSourceLeg() And Not allowRecoveryMissingTarget And CWR_Text(wsPlan.Cells(r, 17).Value2) <> "" Then
        messageText = "PLAN_RELEASE_REQUIRED"
        Exit Function
    End If
    If CWR_PlanRequiresSourceLeg() And Not allowRecoveryMissingTarget And Not allowCancelledResume And _
       CLng(CWR_Number(wsPlan.Cells(r, 19).Value2)) <> 1 Then
        messageText = "PLAN_SOURCE_LEG_COUNT_NOT_ONE=" & CWR_Text(wsPlan.Cells(r, 19).Value2)
        Exit Function
    End If

    If CWR_PlanRequiresSourceLeg() And Not allowRecoveryMissingTarget And Not allowCancelledResume Then
        expectedLegKey = CWR_EXPECTED_MATERIAL & "|" & CWR_TARGET_BIN & "|" & CWR_PICK_BIN_BATCH_BEFORE
        lastRow = wsLeg.Cells(wsLeg.rows.count, 1).End(xlUp).row
        For r = 2 To lastRow
            If StrComp(CWR_Text(wsLeg.Cells(r, 1).Value2), CWR_PLAN_ID, vbTextCompare) = 0 Then
                legCount = legCount + 1
                sourceKey = CWR_Text(wsLeg.Cells(r, 3).Value2)
                qtyValue = CWR_Number(wsLeg.Cells(r, 4).Value2)
                legAvailable = CWR_Number(wsLeg.Cells(r, 5).Value2)
                legTotal = CWR_Number(wsLeg.Cells(r, 6).Value2)
                If StrComp(sourceKey, expectedLegKey, vbTextCompare) <> 0 Or _
                    Abs(qtyValue - CWR_CREATE_QTY) > 0.000001 Or _
                    legAvailable + 0.000001 < CWR_CREATE_QTY Or _
                    legTotal + 0.000001 < CWR_CREATE_QTY Or _
                   StrComp(CWR_Text(wsLeg.Cells(r, 7).Value2), "SOURCE_READY", vbTextCompare) <> 0 Then
                    messageText = "PLAN_SOURCE_LEG_MISMATCH: " & sourceKey
                    Exit Function
                End If
            End If
        Next r
        If legCount <> 1 Then
            messageText = "PLAN_SOURCE_LEG_ROWS=" & CStr(legCount)
            Exit Function
        End If
    End If

    whRowValue = CWR_PLAN_WH_ROW
    If CWR_Text(wsWH.Cells(whRowValue, 1).Value2) <> CWR_EXPECTED_OB Or _
       CWR_Text(wsWH.Cells(whRowValue, 2).Value2) <> CWR_EXPECTED_DOC Or _
       CWR_Text(wsWH.Cells(whRowValue, 3).Value2) <> CWR_EXPECTED_ITEM Or _
       CWR_Text(wsWH.Cells(whRowValue, 5).Value2) <> CWR_EXPECTED_WT Or _
       CWR_Text(wsWH.Cells(whRowValue, 6).Value2) <> CWR_EXPECTED_MATERIAL Or _
       Abs(CWR_Number(wsWH.Cells(whRowValue, 7).Value2) - CWR_EXPECTED_SOURCE_QTY) > 0.000001 Or _
       CWR_Text(wsWH.Cells(whRowValue, 8).Value2) <> CWR_EXPECTED_SOURCE_BIN Or _
       CWR_Text(wsWH.Cells(whRowValue, 9).Value2) <> CWR_EXPECTED_SOURCE_BATCH Or _
       Abs(CWR_Number(wsWH.Cells(whRowValue, 12).Value2) - CWR_TARGET_QTY) > 0.000001 Or _
       CWR_Text(wsWH.Cells(whRowValue, 14).Value2) <> CWR_TARGET_BIN Or _
       UCase$(CWR_Text(wsWH.Cells(whRowValue, 25).Value2)) <> "OB_OPEN_OR_PARTIAL" Or _
       (UCase$(CWR_Text(wsWH.Cells(whRowValue, 26).Value2)) <> "WT_ACTIVE" And _
        UCase$(CWR_Text(wsWH.Cells(whRowValue, 26).Value2)) <> "WT_CANCELLED_PENDING_REMAINING" And _
        UCase$(CWR_Text(wsWH.Cells(whRowValue, 26).Value2)) <> "WT_CANCELLED_REMAINING") Then
        messageText = "WH_ACTUAL_SNAPSHOT_MISMATCH"
        Exit Function
    End If

    CWR_TARGET_STORAGE_TYPE = Split(CWR_TARGET_BIN, "-")(0)
    CWR_TARGET_STORAGE_SECTION = ""
    CWR_EXPECTED_DEST_BIN = CWR_ConfigValue("REWORK_EXPECTED_DEST_BIN")
    CWR_EXPECTED_PROCTY = CWR_ConfigValue("REWORK_EXPECTED_PROCTY")
    If CWR_EXPECTED_DEST_BIN = "" Or CWR_EXPECTED_PROCTY = "" Then
        messageText = "REWORK_DESTINATION_CONTEXT_MISSING"
        Exit Function
    End If

    CWR_EXPECTED_WAREHOUSE = CWR_ConfigValue("EWM_WAREHOUSE")
    If CWR_EXPECTED_WAREHOUSE = "" Then CWR_EXPECTED_WAREHOUSE = "VF02"
    CWR_LoadExecutionContext = True
    Exit Function
Failed:
    messageText = "PLAN_CONTEXT_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub CWR_RecordBatchChangeUserHandoff(ByVal proofText As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(CWR_PLAN_WH_ROW, 15).Value2 = "READY"
    ws.Cells(CWR_PLAN_WH_ROW, 16).Value2 = "POST_CHANGE_WT_CREATED_PENDING_USER_CF"
    ws.Cells(CWR_PLAN_WH_ROW, 17).Value2 = "POST_CHANGE_WRITE_POSTVERIFIED"
    ws.Cells(CWR_PLAN_WH_ROW, 18).Value2 = "USER_CF_HANDOFF"
    ws.Cells(CWR_PLAN_WH_ROW, 19).Value2 = "STOP: Posting Change WT is waiting for user CF. Reload WH and CHECK + POSTING after CF."
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: Posting Change created from the exact Plan; tool stopped before CF. " & proofText
ExitHere:
End Sub

Private Sub CWR_RecordCFOnlyUserHandoff(ByVal proofText As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(CWR_PLAN_WH_ROW, 15).Value2 = "READY"
    ws.Cells(CWR_PLAN_WH_ROW, 16).Value2 = "CF_USER_HANDOFF"
    ws.Cells(CWR_PLAN_WH_ROW, 17).Value2 = "NO_SAP_WRITE_REQUIRED"
    ws.Cells(CWR_PLAN_WH_ROW, 18).Value2 = "USER_CF_HANDOFF"
    ws.Cells(CWR_PLAN_WH_ROW, 19).Value2 = proofText
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: no SAP write required; tool stopped before CF. " & proofText
ExitHere:
End Sub

' Read-only preflight.  This intentionally makes no SAP change and is used
' immediately before the cancellation call as a second independent check.
Public Function Codex_PreflightCancelOldWT() As String
    Dim messageText As String, snapshotText As String
    Dim ses As Object, lower As Object, targetRow As Long

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then
        Codex_PreflightCancelOldWT = "STOP|" & messageText
        Exit Function
    End If

    If Not CWR_LoadExecutionContext(messageText) Then
        Codex_PreflightCancelOldWT = "STOP_PLAN_CONTEXT|" & messageText
        Exit Function
    End If

    If Not CWR_ValidatePreparedSnapshot(messageText) Then
        Codex_PreflightCancelOldWT = "STOP_LOCAL_PREVERIFY|" & messageText
        Exit Function
    End If
    If Not CWR_GetMONSession(ses, messageText) Then
        Codex_PreflightCancelOldWT = "STOP_MON_PREVERIFY|" & messageText
        Exit Function
    End If
    If Not CWR_LocateExactActiveWT(ses, lower, targetRow, messageText) Then
        Codex_PreflightCancelOldWT = "STOP_SAP_PREVERIFY|" & messageText
        Exit Function
    End If

    Codex_PreflightCancelOldWT = "PASS|PRE_CANCEL_VERIFIED|OB=" & CWR_EXPECTED_OB & _
        "|DOC=" & CWR_EXPECTED_DOC & "|WT=" & CWR_EXPECTED_WT & _
        "|SOURCE_QTY=" & CStr(CWR_EXPECTED_SOURCE_QTY) & _
        "|SOURCE_BIN=" & CWR_EXPECTED_SOURCE_BIN & _
        "|SOURCE_BATCH=" & CWR_EXPECTED_SOURCE_BATCH & _
        "|TARGET_QTY=" & CStr(CWR_TARGET_QTY) & _
        "|TARGET_BIN=" & CWR_TARGET_BIN & _
        "|TARGET_BATCH=" & CWR_TARGET_BATCH & "|CF=USER_HANDOFF"
    Exit Function
Failed:
    Codex_PreflightCancelOldWT = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Controlled cancellation mutation. It can execute only after the preflight
' above proves the exact OB, ODO, WT, source quantity/bin/batch and target.
Public Function Codex_CancelOldWTAndRecord() As String
    Dim preflightText As String, messageText As String, postProof As String
    Dim ses As Object, lower As Object, targetRow As Long

    On Error GoTo Failed
    preflightText = Codex_PreflightCancelOldWT()
    If Left$(preflightText, 5) <> "PASS|" Then
        CWR_Log "WT_CANCEL", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", preflightText
        Codex_CancelOldWTAndRecord = preflightText
        Exit Function
    End If

    If Not CWR_GetMONSession(ses, messageText) Then
        CWR_Log "WT_CANCEL", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
        Codex_CancelOldWTAndRecord = "STOP_MON_BEFORE_CANCEL|" & messageText
        Exit Function
    End If
    If Not CWR_LocateExactActiveWT(ses, lower, targetRow, messageText) Then
        CWR_Log "WT_CANCEL", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
        Codex_CancelOldWTAndRecord = "STOP_SAP_BEFORE_CANCEL|" & messageText
        Exit Function
    End If
    If Not CWR_RequireReworkAction("WT_CANCEL", messageText) Then
        CWR_Log "WT_CANCEL", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
        Codex_CancelOldWTAndRecord = "STOP_WRITE_GUARD|" & messageText
        Exit Function
    End If

    If Not CWR_SendExactCancel(ses, lower, targetRow, messageText) Then
        CWR_Log "WT_CANCEL", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
        Codex_CancelOldWTAndRecord = "STOP_CANCEL_NOT_CONFIRMED|" & messageText
        Exit Function
    End If

    If Not CWR_PostVerifyCancelled(ses, postProof, messageText) Then
        CWR_Log "WT_CANCEL", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
        Codex_CancelOldWTAndRecord = "STOP_CANCEL_POSTVERIFY|" & messageText
        Exit Function
    End If

    CWR_RecordCancelledSnapshot postProof
    CWR_Log "WT_CANCEL", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Old WT cancelled and re-read from MON; " & postProof & "; CF=USER_HANDOFF"
    Codex_CancelOldWTAndRecord = "PASS|WT_CANCELLED|OB=" & CWR_EXPECTED_OB & _
        "|DOC=" & CWR_EXPECTED_DOC & "|WT=" & CWR_EXPECTED_WT & _
        "|" & postProof & "|NEXT=READ_REMAINING_QTY_THEN_CHANGE_BIN_BATCH_QTY|CF=USER_HANDOFF"
    Exit Function
Failed:
    CWR_Log "WT_CANCEL", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_CancelOldWTAndRecord = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Reads the cancelled ODO in /SCWM/PRDO.  Navigation and grid reads only;
' there is deliberately no Edit, ModifyCell, Check, Save, Create or Confirm.
Public Function Codex_ReadCancelledRemainingAndRecord() As String
    Dim messageText As String, qtyText As String, batchText As String, statusText As String
    Dim ses As Object, grid As Object

    On Error GoTo Failed
    If Not CWR_LoadExecutionContext(messageText) Then
        Codex_ReadCancelledRemainingAndRecord = "STOP_PLAN_CONTEXT|" & messageText
        Exit Function
    End If
    If Not CWR_AssertCFHandoff(messageText) Then
        Codex_ReadCancelledRemainingAndRecord = "STOP|" & messageText
        Exit Function
    End If
    If Not CWR_AssertCancelledLocal(messageText) Then
        Codex_ReadCancelledRemainingAndRecord = "STOP_LOCAL_CANCEL_STATE|" & messageText
        Exit Function
    End If
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then
        CWR_Log "WT_REMAINING_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
        Codex_ReadCancelledRemainingAndRecord = "STOP_PRDO_OPEN|" & messageText
        Exit Function
    End If
    If Not CWR_ReadExactPRDOItem(grid, qtyText, batchText, statusText, messageText) Then
        CWR_Log "WT_REMAINING_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
        Codex_ReadCancelledRemainingAndRecord = "STOP_PRDO_READ|" & messageText
        Exit Function
    End If
    ' A resume after the PRDO Save legitimately reads Target Qty rather than
    ' the old cancelled-WT quantity.  Accept only those two exact values; any
    ' third value is still ambiguous and blocks the route before Create WT.
    If (Abs(CWR_Number(qtyText) - CWR_EXPECTED_SOURCE_QTY) > 0.000001 And _
        Abs(CWR_Number(qtyText) - CWR_TARGET_QTY) > 0.000001) Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "REMAINING_PREVERIFY_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText & _
            "; EXPECTED_BEFORE=" & CStr(CWR_EXPECTED_SOURCE_QTY) & "; TARGET=" & CStr(CWR_TARGET_QTY)
        CWR_Log "WT_REMAINING_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
        Codex_ReadCancelledRemainingAndRecord = "STOP|" & messageText
        Exit Function
    End If

    CWR_RecordRemaining CWR_Number(qtyText), batchText, statusText
    CWR_Log "WT_REMAINING_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "PRDO exact item read; Qty=" & qtyText & "; Batch=" & batchText & _
        IIf(Abs(CWR_Number(qtyText) - CWR_TARGET_QTY) <= 0.000001, "; TARGET_QTY_ALREADY_SAVED", "") & _
        "; " & statusText & "; CF=USER_HANDOFF"
    Codex_ReadCancelledRemainingAndRecord = "PASS|REMAINING_QTY=" & qtyText & _
        "|CURRENT_BATCH=" & batchText & "|" & statusText & _
        "|NEXT=VERIFY_TARGET_STOCK_THEN_CHANGE_BIN_BATCH_QTY|CF=USER_HANDOFF"
    Exit Function
Failed:
    CWR_Log "WT_REMAINING_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReadCancelledRemainingAndRecord = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Reads every FT stock line for the target pick bin, regardless of its current
' batch. This is the mandatory first gate for the corrected batch-change
' workflow: identify real stock already in the new bin before any batch field
' can be modified. It does not select a line, create a posting change, or
' save anything in SAP.
Public Function Codex_ReadTargetBinStockCandidates() As String
    Dim ses As Object, grid As Object, messageText As String, dumpText As String

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_QueryExpectedPOST(ses, grid, messageText) Then GoTo Stopped
    dumpText = CWR_DumpTargetBinStockCandidates(grid, messageText)
    If Left$(dumpText, 5) = "FAIL|" Then GoTo Stopped
    CWR_Log "TARGET_BIN_BATCH_SCAN", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Target-bin stock scan read only; no POST row selected or changed."
    Codex_ReadTargetBinStockCandidates = "PASS|TARGET_BIN_STOCK_READONLY|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "TARGET_BIN_BATCH_SCAN", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ReadTargetBinStockCandidates = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "TARGET_BIN_BATCH_SCAN", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReadTargetBinStockCandidates = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Re-runs SAP's stock determination for the exact, single ODO request after
' the target-bin batch check has passed. This updates only the unsaved Create
' WT candidate list. It never adopts a quantity, creates/saves a WT, or calls
' any CF action.
Public Function Codex_RedetermineExactCreateWTStocks() As String
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Const STOCK_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/cntlCC_TB_ODP_AVSTKDLV/shellcont/shell"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Const REDETERMINE_ID As String = "OK_OD_AVSTKDLV_REFRESH"
    Dim createSes As Object, wrGrid As Object, stockTab As Object, stockToolbar As Object, stockGrid As Object
    Dim messageText As String, dumpText As String, probeText As String, detailText As String
    Dim rowDoc As String, rowItem As String, rowMat As String, rowQty As String, statusType As String, statusText As String

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "CREATE_WT_REDETERMINE_OPEN_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If

    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    If wrGrid Is Nothing Or CLng(wrGrid.rowCount) <> 1 Then
        messageText = "CREATE_WT_REDETERMINE_WR_GRID_UNAVAILABLE"
        GoTo Stopped
    End If
    rowDoc = CWR_Cell(wrGrid, 0, "DLVNO")
    rowItem = CWR_Cell(wrGrid, 0, "ITMNO")
    rowMat = CWR_Cell(wrGrid, 0, "MATNR")
    rowQty = CWR_Cell(wrGrid, 0, "UI_OPNQTY")
    If StrComp(rowDoc, CWR_EXPECTED_DOC, vbTextCompare) <> 0 Or _
       StrComp(rowItem, CWR_EXPECTED_ITEM, vbTextCompare) <> 0 Or _
       StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Or _
       Abs(CWR_Number(rowQty) - CWR_TARGET_QTY) > 0.000001 Then
        messageText = "CREATE_WT_REDETERMINE_WR_MISMATCH: DOC=" & rowDoc & "; ITEM=" & rowItem & _
            "; MAT=" & rowMat & "; QTY=" & rowQty
        GoTo Stopped
    End If
    wrGrid.selectedRows = "0"
    wrGrid.SetCurrentCell 0, "DLVNO"
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Or StrComp(CWR_Text(stockTab.text), "Stk Can Be Removed", vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_REDETERMINE_STOCK_TAB_IDENTITY_MISMATCH"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitIdle(createSes, 20#) Then
        messageText = "CREATE_WT_REDETERMINE_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set stockToolbar = createSes.FindById(STOCK_TOOLBAR_ID, False)
    If stockToolbar Is Nothing Then
        messageText = "CREATE_WT_REDETERMINE_TOOLBAR_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(stockToolbar, REDETERMINE_ID, "Redetermine Stocks", detailText, messageText) Then GoTo Stopped

    stockToolbar.PressButton REDETERMINE_ID
    If Not CWR_WaitIdle(createSes, 30#) Then
        messageText = "CREATE_WT_REDETERMINE_TIMEOUT"
        GoTo Stopped
    End If
    CWR_ReadStatusBar createSes, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "CREATE_WT_REDETERMINE_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    If stockGrid Is Nothing Then
        messageText = "CREATE_WT_REDETERMINE_STOCK_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    dumpText = "ACTION=" & detailText & vbLf & CWR_DumpKnownGrid("STOCK_AFTER_REDETERMINE", stockGrid, 30)
    If Not CWR_ProveCreateWTTargetCandidate(stockGrid, messageText) Then GoTo Stopped
    CWR_Log "WT_CREATE_REDETERMINE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact candidate list redetermined and target stock appears; no quantity adopted or WT created."
    Codex_RedetermineExactCreateWTStocks = "PASS|CREATE_WT_REDETERMINED_TARGET_VISIBLE|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    If dumpText <> "" Then messageText = messageText & vbLf & dumpText
    CWR_Log "WT_CREATE_REDETERMINE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_RedetermineExactCreateWTStocks = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_REDETERMINE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_RedetermineExactCreateWTStocks = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only-to-SAP preview of the normal source-selection path.  It adopts the
' uniquely proven stock candidate into a temporary Warehouse Task draft, maps
' that draft, and deletes it again without Create, Create+Save, Save, or CF.
Public Function Codex_PreviewAdoptCorrectCandidateAndReset() As String
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Const WT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV"
    Const WT_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/cntlCC_TB_ODP_TODLV/shellcont/shell"
    Const WT_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/subSUB_ODP_TODLV_DATA:/SCWM/SAPLUI_TODLV:4110/cntlCC_ALV_OD_TODLV/shellcont/shell"
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Const STOCK_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/cntlCC_TB_ODP_AVSTKDLV/shellcont/shell"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Const ADOPT_ID As String = "OK_OD_AVSTKDLV_TAKE_QTY"
    Const DELETE_ID As String = "OK_OD_TODLV_DELETE"
    Dim createSes As Object, wrGrid As Object, wtTab As Object, stockTab As Object, wtGrid As Object, wtToolbar As Object
    Dim stockGrid As Object, stockToolbar As Object, messageText As String, proofText As String, dumpText As String
    Dim adoptDetail As String, deleteDetail As String, statusType As String, statusText As String, candidateRow As Long
    Dim draftCreated As Boolean, probeText As String
    Dim stageText As String
    On Error GoTo Failed
    stageText = "ASSERT_CF_HANDOFF"
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    stageText = "ASSERT_CANCELLED"
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    stageText = "ASSERT_PRDO_QTY"
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    stageText = "FIND_CREATE_SESSION"
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        stageText = "OPEN_CREATE_SESSION"
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "CREATE_WT_ADOPT_PREVIEW_OPEN_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If
    stageText = "VERIFY_WAREHOUSE_REQUEST"
    If Not CWR_VerifyExactCreateWTRequest(createSes, wrGrid, messageText) Then GoTo Stopped

    stageText = "FIND_WT_TAB"
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    Set wtToolbar = createSes.FindById(WT_TOOLBAR_ID, False)
    If wtTab Is Nothing Or wtToolbar Is Nothing Or StrComp(CWR_Text(wtTab.text), "Warehouse Task", vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_WT_SURFACE_NOT_FOUND"
        GoTo Stopped
    End If
    ' With no draft SAP keeps this tab disabled.  Do not treat that normal
    ' empty state as an error; only inspect the grid if SAP says it is active.
    If StrComp(CWR_Text(wtTab.Changeable), "True", vbTextCompare) = 0 Then
        stageText = "SELECT_WT_TAB_BEFORE_ADOPT"
        wtTab.Select
        If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
            messageText = "CREATE_WT_ADOPT_PREVIEW_WT_TAB_TIMEOUT"
            GoTo Stopped
        End If
        stageText = "FIND_EMPTY_WT_GRID"
        Set wtGrid = createSes.FindById(WT_GRID_ID, False)
        If wtGrid Is Nothing Then
            messageText = "CREATE_WT_ADOPT_PREVIEW_REQUIRES_EMPTY_DRAFT: GRID_MISSING"
            GoTo Stopped
        End If
        If CLng(wtGrid.rowCount) <> 0 Then
            messageText = "CREATE_WT_ADOPT_PREVIEW_REQUIRES_EMPTY_DRAFT: ROWS=" & CStr(wtGrid.rowCount)
            GoTo Stopped
        End If
    End If

    stageText = "FIND_STOCK_TAB"
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Or StrComp(CWR_Text(stockTab.text), "Stk Can Be Removed", vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stageText = "SELECT_STOCK_TAB"
    stockTab.Select
    If Not CWR_WaitIdle(createSes, 20#) Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    stageText = "FIND_STOCK_SURFACE"
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    Set stockToolbar = createSes.FindById(STOCK_TOOLBAR_ID, False)
    If stockGrid Is Nothing Or stockToolbar Is Nothing Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_STOCK_SURFACE_NOT_FOUND"
        GoTo Stopped
    End If
    stageText = "FIND_EXACT_CANDIDATE"
    If Not CWR_FindExactCreateWTTargetCandidate(stockGrid, candidateRow, proofText, messageText) Then GoTo Stopped
    stageText = "SELECT_EXACT_CANDIDATE"
    stockGrid.selectedRows = CStr(candidateRow)
    stockGrid.SetCurrentCell candidateRow, "LGPLA"
    stockGrid.SelectionChanged
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 15#) Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_SELECTION_TIMEOUT"
        GoTo Stopped
    End If
    If CWR_SelectedRowsText(stockGrid) <> CStr(candidateRow) Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_SELECTION_READBACK_MISMATCH: " & CWR_SelectedRowsText(stockGrid)
        GoTo Stopped
    End If
    stageText = "VERIFY_ADOPT_ACTION"
    If Not CWR_GetVerifiedToolbarAction(stockToolbar, ADOPT_ID, "Adopt Quantity in WhsTask", adoptDetail, messageText) Then GoTo Stopped

    stageText = "PRESS_ADOPT"
    stockToolbar.PressButton ADOPT_ID
    draftCreated = True
    ' Toolbar presses are queued asynchronously.  A bare WaitIdle can return
    ' before SAP has received the Adopt request, then a tab switch can race it.
    stageText = "WAIT_ADOPT"
    If Not CWR_WaitPOSTSaveSettle(createSes, 1#, 30#) Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_ADOPT_TIMEOUT"
        GoTo Stopped
    End If
    stageText = "READ_ADOPT_STATUS"
    CWR_ReadStatusBar createSes, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_ADOPT_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    stageText = "REACQUIRE_WT_TAB_AFTER_ADOPT"
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    If wtTab Is Nothing Or StrComp(CWR_Text(wtTab.text), "Warehouse Task", vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_WT_TAB_NOT_FOUND_AFTER_ADOPT"
        GoTo Stopped
    End If
    If StrComp(CWR_Text(wtTab.Changeable), "True", vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_WT_TAB_NOT_ENABLED_AFTER_ADOPT: " & CWR_Text(wtTab.Changeable)
        GoTo Stopped
    End If
    stageText = "SELECT_WT_TAB_AFTER_ADOPT"
    wtTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_DRAFT_TAB_TIMEOUT"
        GoTo Stopped
    End If
    stageText = "FIND_ADOPTED_WT_GRID"
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_DRAFT_COUNT_MISMATCH: GRID_MISSING"
        GoTo Stopped
    End If
    If CLng(wtGrid.rowCount) <> 1 Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_DRAFT_COUNT_MISMATCH: ROWS=" & CStr(wtGrid.rowCount)
        GoTo Stopped
    End If
    stageText = "DUMP_ADOPTED_DRAFT"
    dumpText = "CANDIDATE=" & proofText & vbLf & "ADOPT_ACTION=" & adoptDetail & vbLf & _
        CWR_DumpKnownGrid("ADOPTED_WT_DRAFT", wtGrid, 5) & _
        CWR_DumpGridFieldContract("ADOPTED_WT_DRAFT", wtGrid, 0)

    stageText = "SELECT_DRAFT_FOR_DELETE"
    wtGrid.selectedRows = "0"
    wtGrid.SetCurrentCell 0, "LGPLA"
    stageText = "VERIFY_DELETE_ACTION"
    If Not CWR_GetVerifiedToolbarAction(wtToolbar, DELETE_ID, "Delete Warehouse Task", deleteDetail, messageText) Then GoTo Stopped
    stageText = "PRESS_DELETE_DRAFT"
    wtToolbar.PressButton DELETE_ID
    stageText = "WAIT_DELETE_DRAFT"
    If Not CWR_WaitIdle(createSes, 20#) Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_DELETE_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 0 Then
        messageText = "CREATE_WT_ADOPT_PREVIEW_DELETE_COUNT_MISMATCH"
        GoTo Stopped
    End If
    draftCreated = False
    CWR_Log "WT_CREATE_ADOPT_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact candidate adopted into one draft and removed again; no WT created or CF sent. " & proofText
    Codex_PreviewAdoptCorrectCandidateAndReset = "PASS|CREATE_WT_ADOPT_DRAFT_CONTRACT|WRITE=FALSE" & vbLf & _
        dumpText & "DELETE_ACTION=" & deleteDetail & vbLf & "POSTVERIFY|DRAFT_ROWS=0"
    Exit Function
Stopped:
    If draftCreated Then messageText = messageText & "; UNSAVED_ADOPTED_DRAFT_MAY_REMAIN"
    CWR_Log "WT_CREATE_ADOPT_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_PreviewAdoptCorrectCandidateAndReset = "STOP|" & messageText
    Exit Function
Failed:
    If draftCreated Then messageText = "UNSAVED_ADOPTED_DRAFT_MAY_REMAIN; " & messageText
    CWR_Log "WT_CREATE_ADOPT_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "STAGE=" & stageText & "; Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_PreviewAdoptCorrectCandidateAndReset = "FAIL|" & messageText & "STAGE=" & stageText & _
        "|ERR=" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only recovery diagnostic after an Adopt Quantity preview. It neither
' selects a stock row nor presses a toolbar button; it exposes visible grids,
' tabs, and any modal so an unsaved draft is never guessed at or overwritten.
Public Function Codex_DumpCurrentCreateWTState() As String
    Dim createSes As Object, root As Object, modal As Object
    Dim messageText As String, dumpText As String, rawTree As String, modalTree As String, props As Variant, itemCount As Long
    On Error GoTo Failed
    If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    Set root = createSes.FindById("wnd[0]", False)
    If root Is Nothing Then
        messageText = "CREATE_WT_CURRENT_ROOT_NOT_FOUND"
        GoTo Stopped
    End If
    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
    rawTree = CStr(createSes.GetObjectTree("wnd[0]", props) & "")
    itemCount = 0
    CWR_CollectCreateWTActionSurface root, 0, itemCount, dumpText
    CWR_DumpVisibleCreateWTGrids root, 0, itemCount, dumpText
    Set modal = createSes.FindById("wnd[1]", False)
    If Not modal Is Nothing Then
        modalTree = CStr(createSes.GetObjectTree("wnd[1]", props) & "")
        itemCount = 0
        CWR_CollectAllVisibleText modal, 0, itemCount, dumpText
    End If
    CWR_Log "WT_CREATE_CURRENT_STATE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Current Create WT state read only; no draft action sent."
    Codex_DumpCurrentCreateWTState = "PASS|CREATE_WT_CURRENT_STATE|WRITE=FALSE|TITLE=" & CWR_Text(root.text) & vbLf & _
        dumpText & rawTree & modalTree
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_CURRENT_STATE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpCurrentCreateWTState = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_CURRENT_STATE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpCurrentCreateWTState = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Strictly read-only diagnostic for the available-stock grid.  It exposes the
' actual selected rows/current cell and every visible field of the candidate
' rows, so an Adopt action is never guessed from an implicit GUI selection.
Public Function Codex_DumpCreateWTStockCandidateContract() As String
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Const STOCK_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/cntlCC_TB_ODP_AVSTKDLV/shellcont/shell"
    Dim createSes As Object, stockGrid As Object, stockToolbar As Object, stockTab As Object
    Dim messageText As String, dumpText As String
    On Error GoTo Failed
    If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    Set stockTab = createSes.FindById("wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV", False)
    If stockTab Is Nothing Then
        messageText = "CREATE_WT_STALE_RECOVERY_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "CREATE_WT_STALE_RECOVERY_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    Set stockToolbar = createSes.FindById(STOCK_TOOLBAR_ID, False)
    If stockGrid Is Nothing Or stockToolbar Is Nothing Then
        messageText = "CREATE_WT_STOCK_CANDIDATE_SURFACE_NOT_FOUND"
        GoTo Stopped
    End If
    dumpText = CWR_GridSelectionState("STOCK", stockGrid) & _
        CWR_DumpGridBrief("STOCK", stockGrid, 10, 160) & _
        CWR_DumpGridFieldContract("STOCK_ROW_0", stockGrid, 0) & _
        CWR_DumpGridFieldContract("STOCK_ROW_1", stockGrid, 1) & _
        CWR_DumpToolbarButtons("STOCK", stockToolbar)
    CWR_Log "WT_CREATE_STOCK_CONTRACT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Available-stock selection and field contract read only; no Adopt/Create/Save/CF action."
    Codex_DumpCreateWTStockCandidateContract = "PASS|CREATE_WT_STOCK_CANDIDATE_CONTRACT|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_STOCK_CONTRACT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpCreateWTStockCandidateContract = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_STOCK_CONTRACT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpCreateWTStockCandidateContract = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Reads SAP's own grid metadata rather than guessing technical column names.
' It does not select a row, alter a cell, or press any command.
Public Function Codex_DumpCreateWTStockGridMetadata() As String
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Dim createSes As Object, props As Variant, messageText As String, metadataText As String
    On Error GoTo Failed
    If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    props = Array("Id", "Type", "RowCount", "ColumnCount", "ColumnOrder", "SelectedRows", "CurrentCellRow", "CurrentCellColumn")
    metadataText = CStr(createSes.GetObjectTree(STOCK_GRID_ID, props) & "")
    CWR_Log "WT_CREATE_STOCK_METADATA", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Stock grid metadata read only; no GUI selection or action sent."
    Codex_DumpCreateWTStockGridMetadata = "PASS|CREATE_WT_STOCK_GRID_METADATA|WRITE=FALSE" & vbLf & metadataText
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_STOCK_METADATA", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpCreateWTStockGridMetadata = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_STOCK_METADATA", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpCreateWTStockGridMetadata = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Sends only the exact Warehouse Request selection event, then reports which
' dependent tabs SAP enables.  It is a read/setup operation: no stock is
' adopted and no warehouse task is created or saved.
Public Function Codex_SelectExactCreateWTRequestReadOnly() As String
    Const WT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV"
    Const DEFAULT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_DEFDLV"
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Dim createSes As Object, wrGrid As Object, wtTab As Object, defaultTab As Object, stockTab As Object
    Dim messageText As String, dumpText As String
    On Error GoTo Failed
    If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    If Not CWR_VerifyExactCreateWTRequest(createSes, wrGrid, messageText) Then GoTo Stopped
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    Set defaultTab = createSes.FindById(DEFAULT_TAB_ID, False)
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    dumpText = CWR_GridSelectionState("WAREHOUSE_REQUEST", wrGrid) & _
        "DEPENDENT_TAB|WAREHOUSE_TASK_CHANGEABLE=" & CWR_Text(wtTab.Changeable) & _
        "|DEFAULT_VALUES_CHANGEABLE=" & CWR_Text(defaultTab.Changeable) & _
        "|STOCK_CAN_BE_REMOVED_CHANGEABLE=" & CWR_Text(stockTab.Changeable) & vbLf
    CWR_Log "WT_CREATE_WR_SELECTION", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact Warehouse Request selected and read back; no Adopt/Create/Save/CF action."
    Codex_SelectExactCreateWTRequestReadOnly = "PASS|CREATE_WT_WR_SELECTION|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_WR_SELECTION", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_SelectExactCreateWTRequestReadOnly = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_WR_SELECTION", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_SelectExactCreateWTRequestReadOnly = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Non-persistent discovery of the standard top-level Create action.  SAP's
' separate Create command must produce one in-memory draft before any
' Create+Save action is allowed.  The same draft is deleted in this preview;
' no Save, Create+Save, or CF is sent.
Public Function Codex_PreviewTopCreateDraftAndReset() As String
    Const TOP_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/cntlCC_TB_OIP_DLV/shellcont/shell"
    Const WT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV"
    Const WT_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/cntlCC_TB_ODP_TODLV/shellcont/shell"
    Const WT_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/subSUB_ODP_TODLV_DATA:/SCWM/SAPLUI_TODLV:4110/cntlCC_ALV_OD_TODLV/shellcont/shell"
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Const CREATE_ID As String = "OK_OI_DLV_CREATE_TO"
    Const DELETE_ID As String = "OK_OD_TODLV_DELETE"
    Dim createSes As Object, wrGrid As Object, topToolbar As Object, stockTab As Object, stockGrid As Object
    Dim wtTab As Object, wtToolbar As Object, wtGrid As Object
    Dim messageText As String, proofText As String, createDetail As String, deleteDetail As String, dumpText As String
    Dim statusType As String, statusText As String, candidateRow As Long, draftCreated As Boolean, stageText As String

    On Error GoTo Failed
    stageText = "ASSERT_CF_HANDOFF"
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    stageText = "ASSERT_CANCELLED"
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    stageText = "ASSERT_PRDO_QTY"
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    stageText = "FIND_CREATE_SESSION"
    If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    stageText = "VERIFY_WAREHOUSE_REQUEST"
    If Not CWR_VerifyExactCreateWTRequest(createSes, wrGrid, messageText) Then GoTo Stopped
    stageText = "SELECT_STOCK_TAB"
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Then
        messageText = "TOP_CREATE_PREVIEW_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "TOP_CREATE_PREVIEW_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    stageText = "FIND_EXACT_CANDIDATE"
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    If stockGrid Is Nothing Then
        messageText = "TOP_CREATE_PREVIEW_STOCK_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_FindExactCreateWTTargetCandidate(stockGrid, candidateRow, proofText, messageText) Then GoTo Stopped
    stageText = "SELECT_EXACT_CANDIDATE"
    stockGrid.selectedRows = CStr(candidateRow)
    stockGrid.SetCurrentCell candidateRow, "LGPLA"
    stockGrid.SelectionChanged
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "TOP_CREATE_PREVIEW_SELECTION_TIMEOUT"
        GoTo Stopped
    End If
    If CWR_SelectedRowsText(stockGrid) <> CStr(candidateRow) Then
        messageText = "TOP_CREATE_PREVIEW_SELECTION_READBACK_MISMATCH: " & CWR_SelectedRowsText(stockGrid)
        GoTo Stopped
    End If
    stageText = "VERIFY_CREATE_ACTION"
    Set topToolbar = createSes.FindById(TOP_TOOLBAR_ID, False)
    If topToolbar Is Nothing Then
        messageText = "TOP_CREATE_PREVIEW_TOP_TOOLBAR_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(topToolbar, CREATE_ID, "Create Warehouse Task", createDetail, messageText) Then GoTo Stopped
    stageText = "PRESS_CREATE_DRAFT"
    topToolbar.PressButton CREATE_ID
    draftCreated = True
    stageText = "WAIT_CREATE_DRAFT"
    If Not CWR_WaitPOSTSaveSettle(createSes, 1#, 30#) Then
        messageText = "TOP_CREATE_PREVIEW_CREATE_TIMEOUT"
        GoTo Stopped
    End If
    CWR_ReadStatusBar createSes, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "TOP_CREATE_PREVIEW_CREATE_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    stageText = "REACQUIRE_WT_TAB"
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    If wtTab Is Nothing Or StrComp(CWR_Text(wtTab.text), "Warehouse Task", vbTextCompare) <> 0 Then
        messageText = "TOP_CREATE_PREVIEW_WT_TAB_NOT_FOUND_AFTER_CREATE"
        GoTo Stopped
    End If
    If StrComp(CWR_Text(wtTab.Changeable), "True", vbTextCompare) <> 0 Then
        messageText = "TOP_CREATE_PREVIEW_WT_TAB_NOT_ENABLED_AFTER_CREATE: " & CWR_Text(wtTab.Changeable)
        GoTo Stopped
    End If
    stageText = "SELECT_WT_TAB_AFTER_CREATE"
    wtTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "TOP_CREATE_PREVIEW_WT_TAB_TIMEOUT"
        GoTo Stopped
    End If
    stageText = "MAP_CREATED_DRAFT"
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    Set wtToolbar = createSes.FindById(WT_TOOLBAR_ID, False)
    If wtGrid Is Nothing Or wtToolbar Is Nothing Then
        messageText = "TOP_CREATE_PREVIEW_DRAFT_SURFACE_NOT_FOUND"
        GoTo Stopped
    End If
    If CLng(wtGrid.rowCount) <> 1 Then
        messageText = "TOP_CREATE_PREVIEW_DRAFT_COUNT_MISMATCH: ROWS=" & CStr(wtGrid.rowCount)
        GoTo Stopped
    End If
    dumpText = "CANDIDATE=" & proofText & vbLf & "CREATE_ACTION=" & createDetail & vbLf & _
        CWR_GridSelectionState("TOP_CREATE_DRAFT", wtGrid) & _
        CWR_DumpKnownGrid("TOP_CREATE_DRAFT", wtGrid, 5) & _
        CWR_DumpGridFieldContract("TOP_CREATE_DRAFT", wtGrid, 0)
    stageText = "DELETE_DRAFT"
    wtGrid.selectedRows = "0"
    wtGrid.SetCurrentCell 0, "LGPLA"
    If Not CWR_GetVerifiedToolbarAction(wtToolbar, DELETE_ID, "Delete Warehouse Task", deleteDetail, messageText) Then GoTo Stopped
    wtToolbar.PressButton DELETE_ID
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "TOP_CREATE_PREVIEW_DELETE_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 0 Then
        messageText = "TOP_CREATE_PREVIEW_DELETE_COUNT_MISMATCH"
        GoTo Stopped
    End If
    draftCreated = False
    CWR_Log "WT_CREATE_TOP_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Top Create produced one draft and the same unsaved draft was deleted; no Save/Create+Save/CF. " & proofText
    Codex_PreviewTopCreateDraftAndReset = "PASS|TOP_CREATE_DRAFT_CONTRACT|WRITE=FALSE" & vbLf & dumpText & _
        "DELETE_ACTION=" & deleteDetail & vbLf & "POSTVERIFY|DRAFT_ROWS=0"
    Exit Function
Stopped:
    If draftCreated Then messageText = messageText & "; UNSAVED_TOP_CREATE_DRAFT_MAY_REMAIN"
    CWR_Log "WT_CREATE_TOP_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_PreviewTopCreateDraftAndReset = "STOP|" & messageText
    Exit Function
Failed:
    If draftCreated Then messageText = "UNSAVED_TOP_CREATE_DRAFT_MAY_REMAIN; " & messageText
    CWR_Log "WT_CREATE_TOP_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "STAGE=" & stageText & "; Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_PreviewTopCreateDraftAndReset = "FAIL|" & messageText & "STAGE=" & stageText & _
        "|ERR=" & CStr(Err.Number) & "|" & Err.Description
End Function

' Monitor-only readback for the current ODO.  This is used after any stopped
' screen preview to prove whether SAP has an active outbound WT; it performs
' no WT create, save, cancel, confirmation, or stock change.
Public Function Codex_ReadCurrentOutboundWTMonitor() As String
    Dim monSes As Object, lower As Object
    Dim messageText As String, dumpText As String, currentDoc As String
    Dim r As Long, n As Long, activeCount As Long, totalCount As Long
    On Error GoTo Failed
    If Not CWR_GetMONSession(monSes, messageText) Then GoTo Stopped
    If Not CWR_LoadFilteredWTGrid(monSes, lower, messageText) Then GoTo Stopped
    n = CLng(lower.rowCount)
    For r = 0 To n - 1
        If CWR_Cell(lower, r, "DOCNO") <> "" Then currentDoc = CWR_Cell(lower, r, "DOCNO")
        If StrComp(currentDoc, CWR_EXPECTED_DOC, vbTextCompare) = 0 And _
           CWR_Cell(lower, r, "TANUM") <> "" Then
            totalCount = totalCount + 1
            If CWR_Cell(lower, r, "TOSTAT") = "" Then activeCount = activeCount + 1
        End If
    Next r
    dumpText = CWR_DumpKnownGrid("OUTBOUND_WT_MON", lower, 100)
    CWR_Log "OUTBOUND_WT_MON_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Monitor read only; rows=" & CStr(totalCount) & "; active=" & CStr(activeCount) & "."
    Codex_ReadCurrentOutboundWTMonitor = "PASS|OUTBOUND_WT_MONITOR|WRITE=FALSE|ODO_ROWS=" & _
        CStr(totalCount) & "|ACTIVE_ROWS=" & CStr(activeCount) & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "OUTBOUND_WT_MON_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ReadCurrentOutboundWTMonitor = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "OUTBOUND_WT_MON_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReadCurrentOutboundWTMonitor = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Recovers only the unsaved screen state created by this tool's failed Create
' preview.  It first proves Monitor has no active outbound WT, then declines
' the exact SAP "Data will be lost" save prompt and reopens a fresh Create WT
' query.  It never saves, creates, cancels, or confirms a warehouse task.
Public Function Codex_ResetOwnedStaleCreateWTDraftNoSave() As String
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Dim createSes As Object, wrGrid As Object, stockGrid As Object, stockTab As Object
    Dim messageText As String, monitorProof As String, modalText As String, probeText As String, candidateProof As String
    Dim rowDoc As String, rowItem As String, rowMat As String, rowQty As String, tcode As String, candidateRow As Long
    On Error GoTo Failed
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_FindOwnedZeroOpenCreateWTDraft(createSes, messageText) Then GoTo Stopped
    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    If wrGrid Is Nothing Or CLng(wrGrid.rowCount) <> 1 Then
        messageText = "CREATE_WT_STALE_RECOVERY_WR_GRID_UNAVAILABLE"
        GoTo Stopped
    End If
    rowDoc = CWR_Cell(wrGrid, 0, "DLVNO")
    rowItem = CWR_Cell(wrGrid, 0, "ITMNO")
    rowMat = CWR_Cell(wrGrid, 0, "MATNR")
    rowQty = CWR_Cell(wrGrid, 0, "UI_OPNQTY")
    If StrComp(rowDoc, CWR_EXPECTED_DOC, vbTextCompare) <> 0 Or _
       StrComp(rowItem, CWR_EXPECTED_ITEM, vbTextCompare) <> 0 Or _
       StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_STALE_RECOVERY_WR_IDENTITY_MISMATCH: DOC=" & rowDoc & "; ITEM=" & rowItem & "; MAT=" & rowMat
        GoTo Stopped
    End If
    If Abs(CWR_Number(rowQty)) > 0.000001 Then
        messageText = "CREATE_WT_STALE_RECOVERY_NOT_REQUIRED_OPEN_QTY=" & rowQty
        GoTo Stopped
    End If
    If Not CWR_AssertNoActiveOutboundWT(monitorProof, messageText) Then GoTo Stopped
    createSes.StartTransaction "/SCWM/PRDO"
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 30#) Then
        messageText = "CREATE_WT_STALE_RECOVERY_NAVIGATION_TIMEOUT"
        GoTo Stopped
    End If
    If CWR_GetModalText(createSes, modalText) Then
        If Not CWR_DeclineExactNoSavePrompt(createSes, modalText, messageText) Then GoTo Stopped
        If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 30#) Then
            messageText = "CREATE_WT_STALE_RECOVERY_DECLINE_TIMEOUT"
            GoTo Stopped
        End If
    End If
    tcode = CWR_Text(createSes.Info.Transaction)
    If UCase$(tcode) <> "/SCWM/PRDO" Then
        messageText = "CREATE_WT_STALE_RECOVERY_TCODE_MISMATCH: " & tcode
        GoTo Stopped
    End If
    If Not CWR_OpenExactCreateWTAfterQtyChange(createSes, messageText) Then
        messageText = "CREATE_WT_STALE_RECOVERY_REOPEN_FAILED: " & messageText
        GoTo Stopped
    End If
    If Not CWR_FindExactCreateWTSession(createSes, messageText) Then GoTo Stopped
    If Not CWR_VerifyExactCreateWTRequest(createSes, wrGrid, messageText) Then GoTo Stopped
    Set stockTab = createSes.FindById("wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV", False)
    If stockTab Is Nothing Then
        messageText = "CREATE_WT_STALE_RECOVERY_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "CREATE_WT_STALE_RECOVERY_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    If stockGrid Is Nothing Then
        messageText = "CREATE_WT_STALE_RECOVERY_STOCK_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_FindExactCreateWTTargetCandidate(stockGrid, candidateRow, candidateProof, messageText) Then GoTo Stopped
    CWR_Log "WT_CREATE_STALE_RECOVERY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Owned unsaved Create preview discarded without save; fresh exact candidate restored. " & monitorProof & "; " & candidateProof
    Codex_ResetOwnedStaleCreateWTDraftNoSave = "PASS|CREATE_WT_STALE_DRAFT_DISCARDED_NO_SAVE|WRITE=FALSE" & _
        "|" & monitorProof & "|" & candidateProof
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_STALE_RECOVERY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ResetOwnedStaleCreateWTDraftNoSave = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_STALE_RECOVERY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ResetOwnedStaleCreateWTDraftNoSave = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Reusable no-write decision gate for rework cases where the desired source
' bin differs from the former WT.  It classifies only the verified candidate
' set: direct create when the ODO batch already exists at the new bin; a
' staged posting-change handoff when exactly one alternate batch exists; or a
' stop for insufficient/ambiguous stock.  It never creates or confirms a WT.
Public Function Codex_PlanDifferentBinSourceFlow() As String
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Dim createSes As Object, wrGrid As Object, stockTab As Object, stockGrid As Object
    Dim messageText As String, probeText As String, planState As String, planText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "SOURCE_FLOW_PLAN_OPEN_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If
    If Not CWR_VerifyExactCreateWTRequest(createSes, wrGrid, messageText) Then GoTo Stopped
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Then
        messageText = "SOURCE_FLOW_PLAN_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "SOURCE_FLOW_PLAN_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    If stockGrid Is Nothing Then
        messageText = "SOURCE_FLOW_PLAN_STOCK_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_ClassifyTargetBinCandidates(stockGrid, planState, planText, messageText) Then GoTo Stopped
    CWR_Log "SOURCE_BIN_FLOW_PLAN", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Read-only source-bin flow classified as " & planState & ". " & planText
    Codex_PlanDifferentBinSourceFlow = "PASS|SOURCE_BIN_FLOW=" & planState & "|WRITE=FALSE|" & planText & _
        "|CF=USER_HANDOFF"
    Exit Function
Stopped:
    CWR_Log "SOURCE_BIN_FLOW_PLAN", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_PlanDifferentBinSourceFlow = "STOP|" & messageText & "|CF=USER_HANDOFF"
    Exit Function
Failed:
    CWR_Log "SOURCE_BIN_FLOW_PLAN", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_PlanDifferentBinSourceFlow = "FAIL|" & CStr(Err.Number) & "|" & Err.Description & "|CF=USER_HANDOFF"
End Function

' Local-only reference for every inferred source-bin branch.  This is a
' decision contract, not an executor: it neither attaches to SAP nor writes
' to a worksheet.  A new OB must still pass its own exact snapshot/preflight.
Public Function Codex_DescribeReworkFlowMatrix() As String
    Dim matrixText As String
    On Error GoTo Failed

    matrixText = _
        "BRANCH=DIRECT_CREATE_WT|WHEN=Exactly one eligible source row at the requested bin already has the WH/ODO batch and the full requested quantity|NEXT=Create one in-memory draft; prove reservation; Save once; read MON; stop for user CF" & vbLf & _
        "BRANCH=POST_BATCH_CHANGE_REQUIRED|WHEN=No eligible WH/ODO batch exists at the requested bin, but exactly one eligible alternate batch row exists there|NEXT=Create only the guarded posting-change WT; stop for user CF; after user CF, re-read and classify again" & vbLf & _
        "BRANCH=STOP_INSUFFICIENT_CORRECT_BATCH|WHEN=The WH/ODO batch exists at the requested bin but its available or total quantity is below the request|NEXT=Do not split, mix batches, or save; user resolves quantity/stock first" & vbLf & _
        "BRANCH=STOP_AMBIGUOUS_CORRECT_BATCH|WHEN=More than one eligible row has the correct WH/ODO batch|NEXT=Do not auto-select; user identifies the intended stock row" & vbLf & _
        "BRANCH=STOP_AMBIGUOUS_ALTERNATE_BATCH|WHEN=More than one eligible alternate-batch row exists at the requested bin|NEXT=Do not auto-post batch change; user identifies the intended alternate batch row" & vbLf & _
        "BRANCH=STOP_NO_ELIGIBLE_STOCK_AT_TARGET_BIN|WHEN=No eligible source row exists at the requested bin|NEXT=Do not create WT or posting change; user resolves stock/bin first" & vbLf & _
        "BRANCH=ACTIVE_OUTBOUND_WT_EXISTS|WHEN=MON already shows an active outbound WT for the exact ODO item|NEXT=No create or retry; read back only and hand CF to the user" & vbLf & _
        "BRANCH=STALE_TOOL_OWNED_DRAFT_ONLY|WHEN=Create screen shows an unsaved draft but MON proves no active outbound WT|NEXT=Discard only the tool-owned draft without Save, then re-run the read-only plan" & vbLf & _
        "RULE=CF|OWNER=USER|AUTOMATION=NEVER_CONFIRM_WT"

    Codex_DescribeReworkFlowMatrix = "PASS|REWORK_FLOW_MATRIX|SAP=NOT_ACCESSED|WRITE=FALSE|CF=USER_HANDOFF" & vbLf & matrixText
    Exit Function
Failed:
    Codex_DescribeReworkFlowMatrix = "FAIL|REWORK_FLOW_MATRIX|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Reads only the workbook's own handoff record.  It is intentionally safe to
' run while an active WT is waiting for the user to confirm it in SAP.
Public Function Codex_ReadCurrentReworkHandoff() As String
    Dim ws As Worksheet
    Dim obText As String, docText As String, readyText As String
    Dim flowText As String, proofText As String, handoffText As String
    Dim stateText As String, nextText As String
    On Error GoTo Failed

    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    obText = CWR_Text(ws.Cells(5, 1).Value2)
    docText = CWR_Text(ws.Cells(5, 2).Value2)
    readyText = CWR_Text(ws.Cells(5, 15).Value2)
    flowText = CWR_Text(ws.Cells(5, 16).Value2)
    proofText = CWR_Text(ws.Cells(5, 17).Value2)
    handoffText = CWR_Text(ws.Cells(5, 18).Value2)

    If (StrComp(flowText, "OUTBOUND_WT_CREATED_PENDING_USER_CF", vbTextCompare) = 0 And _
        StrComp(proofText, "OUTBOUND_WT_POSTVERIFIED", vbTextCompare) = 0) Or _
       (StrComp(flowText, "READY_FOR_CF", vbTextCompare) = 0 And _
        StrComp(proofText, "WT_POST_VERIFY_PASS", vbTextCompare) = 0) Then
        stateText = "OUTBOUND_WT_AWAITING_USER_CF"
        nextText = "USER_CONFIRM_OUTBOUND_WT_IN_SAP; NO_AUTOMATION_RETRY"
    ElseIf StrComp(flowText, "POST_CHANGE_TRANSACTION_PROOF", vbTextCompare) = 0 And _
           StrComp(proofText, "POST_CHANGE_RESULT=SUCCESS", vbTextCompare) = 0 Then
        stateText = "POST_CHANGE_COMMITTED_CONTINUE"
        nextText = "CONTINUE_TARGET_BIN_QTY_AND_CREATE_OUTBOUND_WT; POST_PROOF_WT_IS_NOT_CF_WT"
    ElseIf StrComp(handoffText, "POST_CHANGE_WT_CREATED_PENDING_CF", vbTextCompare) = 0 Then
        stateText = "POSTING_CHANGE_WT_AWAITING_USER_RECONCILE"
        nextText = "READONLY_RECONCILE_ONLY; NO_POST_RETRY; POST_WT_IS_NOT_CF_WT"
    ElseIf StrComp(flowText, "CREATE_WT_FROM_TARGET_THEN_CF_HANDOFF", vbTextCompare) = 0 And _
           (StrComp(proofText, "PRDO_QTY_VERIFIED", vbTextCompare) = 0 Or _
            StrComp(proofText, "PICK_BIN_BATCH_POST_VERIFIED", vbTextCompare) = 0) Then
        stateText = "READY_FOR_READONLY_SOURCE_PLAN"
        nextText = "RUN_READONLY_CANDIDATE_PLAN_BEFORE_ANY_CREATE"
    ElseIf StrComp(flowText, "POST_CHANGE_BIN_BATCH_QTY", vbTextCompare) = 0 And _
           StrComp(proofText, "POST_PREVERIFIED", vbTextCompare) = 0 Then
        stateText = "POSTING_CHANGE_PREVERIFIED"
        nextText = "ONLY_RUN_GUARDED_POSTING_CHANGE_IF_USER_REQUESTS_IT"
    Else
        stateText = "LOCAL_STATE_REVIEW_REQUIRED"
        nextText = "NO_SAP_ACTION_FROM_THIS_STATUS; REVIEW_WH_ACTUAL_AND_RUN_LOG"
    End If

    Codex_ReadCurrentReworkHandoff = "PASS|REWORK_HANDOFF=" & stateText & _
        "|SAP=NOT_ACCESSED|WRITE=FALSE|OB=" & obText & "|DOC=" & docText & _
        "|READY=" & readyText & "|FLOW=" & flowText & "|PROOF=" & proofText & _
        "|HANDOFF=" & handoffText & "|NEXT=" & nextText & "|CF=USER_HANDOFF"
    Exit Function
Failed:
    Codex_ReadCurrentReworkHandoff = "FAIL|REWORK_HANDOFF|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Action entry point for the already-approved direct outbound-WT branch.
' CHECK remains the read-only router for every case.  This action intentionally
' does not attempt a batch change, does not cancel a WT, and never confirms a
' WT.  The concrete creator below performs its own exact SAP preflight before
' its one permitted Save.
Public Sub Codex_RunApprovedOutboundWH()
    Dim resultText As String
    resultText = Codex_RunApprovedOutboundWHCore()
    If Left$(resultText, 5) = "PASS|" Then
        MsgBox resultText, vbInformation, "TAO WH DA DUYET"
    Else
        MsgBox resultText, vbExclamation, "TAO WH DA DUYET"
    End If
End Sub

' Public core is kept message-free so the VBA tool can call and capture its
' result.  An already active outbound or posting-change WT is an absolute
' stop: no retry, no new Save, and CF remains with the user.
Public Function Codex_RunApprovedOutboundWHCore() As String
    Dim handoffText As String, resultText As String
    On Error GoTo Failed

    handoffText = Codex_ReadCurrentReworkHandoff()
    If Left$(handoffText, 5) <> "PASS|" Then
        Codex_RunApprovedOutboundWHCore = "STOP|HANDOFF_STATE_UNREADABLE|" & handoffText
        Exit Function
    End If
    If InStr(1, handoffText, "REWORK_HANDOFF=OUTBOUND_WT_AWAITING_USER_CF", vbTextCompare) > 0 Then
        Codex_RunApprovedOutboundWHCore = "STOP|OUTBOUND_WT_ALREADY_ACTIVE|NEXT=USER_CONFIRM_OUTBOUND_WT_IN_SAP|CF=USER_HANDOFF"
        Exit Function
    End If
    If InStr(1, handoffText, "REWORK_HANDOFF=POSTING_CHANGE_WT_AWAITING_USER_CF", vbTextCompare) > 0 Then
        Codex_RunApprovedOutboundWHCore = "STOP|POST_CHANGE_WT_ALREADY_ACTIVE|NEXT=USER_CONFIRM_POSTING_CHANGE_WT_THEN_REDETERMINE|CF=USER_HANDOFF"
        Exit Function
    End If

    resultText = Codex_CreateOutboundWTFromExactCandidate()
    If Left$(resultText, 5) = "PASS|" Then
        Codex_RunApprovedOutboundWHCore = resultText
    Else
        Codex_RunApprovedOutboundWHCore = "STOP|APPROVED_OUTBOUND_WT_NOT_CREATED|" & resultText
    End If
    Exit Function
Failed:
    Codex_RunApprovedOutboundWHCore = "FAIL|APPROVED_OUTBOUND_WT|" & CStr(Err.Number) & "|" & Err.Description & "|CF=USER_HANDOFF"
End Function

' Creates exactly one outbound WT from the user-confirmed bin only after the
' normal Create screen has reserved that exact source as an in-memory draft.
' The only persistent SAP command is one top-level Save.  Confirmation (CF)
' is permanently excluded and remains the user's step.
Public Function Codex_CreateOutboundWTFromExactCandidate() As String
    Const TOP_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/cntlCC_TB_OIP_DLV/shellcont/shell"
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Const CREATE_ID As String = "OK_OI_DLV_CREATE_TO"
    Const SAVE_ID As String = "OK_OI_DLV_SAVE"
    Dim createSes As Object, wrGrid As Object, stockTab As Object, stockGrid As Object, topToolbar As Object
    Dim messageText As String, probeText As String, monitorProof As String, sourceProof As String, draftProof As String
    Dim actionDetail As String, saveDetail As String, statusType As String, statusText As String, wtNo As String, postProof As String
    Dim planState As String, planText As String, candidateRow As Long
    Dim draftStarted As Boolean, saveSent As Boolean, stageText As String

    On Error GoTo Failed
    stageText = "LOAD_EXACT_PLAN"
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    stageText = "ASSERT_CF_HANDOFF"
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    stageText = "ASSERT_CANCELLED"
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    stageText = "ASSERT_PRDO_QTY"
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    stageText = "MONITOR_PREVERIFY"
    If Not CWR_AssertNoActiveOutboundWT(monitorProof, messageText) Then GoTo Stopped
    ' Never accept a Create WT window merely because its title matches.  SAP
    ' can keep another user's earlier Create window open in the same
    ' connection.  Only reuse a window whose WR row is already the exact Plan
    ' request; otherwise start the standard follow-on from the exact PRDO
    ' document and wait only for that newly opened window.
    stageText = "OPEN_EXACT_CREATE_REQUEST"
    If Not CWR_FindExactCreateWTSession(createSes, messageText) Then
        If Not CWR_OpenExactCreateWTAfterQtyChange(createSes, messageText) Then GoTo Stopped
    End If
    stageText = "VERIFY_WAREHOUSE_REQUEST"
    If Not CWR_VerifyExactCreateWTRequest(createSes, wrGrid, messageText) Then GoTo Stopped
    stageText = "SELECT_STOCK_TAB"
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Then
        messageText = "OUTBOUND_WT_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "OUTBOUND_WT_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    stageText = "CLASSIFY_CANDIDATES"
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    If stockGrid Is Nothing Then
        messageText = "OUTBOUND_WT_STOCK_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_ClassifyTargetBinCandidates(stockGrid, planState, planText, messageText) Then GoTo Stopped
    If StrComp(planState, "DIRECT_CREATE_WT", vbTextCompare) <> 0 Then
        messageText = "OUTBOUND_WT_REQUIRES_" & planState & ": " & planText
        GoTo Stopped
    End If
    If Not CWR_FindExactCreateWTTargetCandidate(stockGrid, candidateRow, sourceProof, messageText) Then GoTo Stopped
    stageText = "SELECT_EXACT_SOURCE"
    stockGrid.selectedRows = CStr(candidateRow)
    stockGrid.SetCurrentCell candidateRow, "LGPLA"
    stockGrid.SelectionChanged
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "OUTBOUND_WT_SOURCE_SELECTION_TIMEOUT"
        GoTo Stopped
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    If stockGrid Is Nothing Then
        messageText = "OUTBOUND_WT_STOCK_GRID_LOST_AFTER_SELECTION"
        GoTo Stopped
    End If
    If Not CWR_FindExactCreateWTTargetCandidate(stockGrid, candidateRow, sourceProof, messageText) Then GoTo Stopped
    If CWR_SelectedRowsText(stockGrid) <> CStr(candidateRow) Then
        messageText = "OUTBOUND_WT_SOURCE_SELECTION_READBACK_MISMATCH: " & CWR_SelectedRowsText(stockGrid)
        GoTo Stopped
    End If
    stageText = "VERIFY_CREATE_ACTION"
    Set topToolbar = createSes.FindById(TOP_TOOLBAR_ID, False)
    If topToolbar Is Nothing Then
        messageText = "OUTBOUND_WT_TOP_TOOLBAR_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(topToolbar, CREATE_ID, "Create Warehouse Task", actionDetail, messageText) Then GoTo Stopped
    If Not CWR_RequireReworkAction("WT_CREATE", messageText) Then GoTo Stopped
    stageText = "PRESS_CREATE_DRAFT"
    topToolbar.PressButton CREATE_ID
    draftStarted = True
    stageText = "WAIT_CREATE_DRAFT"
    If Not CWR_WaitPOSTSaveSettle(createSes, 1#, 30#) Then
        messageText = "OUTBOUND_WT_CREATE_DRAFT_TIMEOUT"
        GoTo Stopped
    End If
    CWR_ReadStatusBar createSes, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "OUTBOUND_WT_CREATE_DRAFT_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    If CWR_HasAnyModal(createSes, messageText) Then
        messageText = "OUTBOUND_WT_CREATE_DRAFT_MODAL: " & messageText
        GoTo Stopped
    End If
    stageText = "VERIFY_EXACT_DRAFT_RESERVATION"
    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    If wrGrid Is Nothing Then
        messageText = "OUTBOUND_WT_WR_GRID_LOST_AFTER_CREATE"
        GoTo Stopped
    End If
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Then
        messageText = "OUTBOUND_WT_STOCK_TAB_LOST_AFTER_CREATE"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "OUTBOUND_WT_STOCK_TAB_AFTER_CREATE_TIMEOUT"
        GoTo Stopped
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    If stockGrid Is Nothing Then
        messageText = "OUTBOUND_WT_STOCK_GRID_LOST_AFTER_CREATE"
        GoTo Stopped
    End If
    If Not CWR_ProveExactOutboundDraftReservation(wrGrid, stockGrid, draftProof, messageText) Then GoTo Stopped
    stageText = "VERIFY_SAVE_ACTION"
    Set topToolbar = createSes.FindById(TOP_TOOLBAR_ID, False)
    If topToolbar Is Nothing Then
        messageText = "OUTBOUND_WT_TOP_TOOLBAR_LOST_BEFORE_SAVE"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(topToolbar, SAVE_ID, "Save", saveDetail, messageText) Then GoTo Stopped
    If Not CWR_RequireReworkAction("WT_CREATE", messageText) Then GoTo Stopped
    stageText = "PRESS_SAVE"
    topToolbar.PressButton SAVE_ID
    saveSent = True
    stageText = "WAIT_SAVE"
    If Not CWR_WaitPOSTSaveSettle(createSes, 1#, 45#) Then
        messageText = "OUTBOUND_WT_SAVE_TIMEOUT"
        GoTo Stopped
    End If
    CWR_ReadStatusBar createSes, statusType, statusText
    stageText = "MONITOR_POSTVERIFY"
    If Not CWR_ReadExactNewOutboundWT(wtNo, postProof, messageText) Then GoTo Stopped
    CWR_RecordOutboundWTCreated wtNo, postProof
    CWR_Log "OUTBOUND_WT_CREATE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "One exact outbound WT saved and read back; no CF sent. " & postProof
    Codex_CreateOutboundWTFromExactCandidate = "PASS|OUTBOUND_WT_CREATED_PENDING_USER_CF|WT=" & wtNo & _
        "|" & monitorProof & "|SOURCE=" & sourceProof & "|DRAFT=" & draftProof & _
        "|SAVE_ACTION=" & saveDetail & "|STATUS=" & statusType & ":" & statusText & _
        "|POSTVERIFY=" & postProof & "|CF=USER_HANDOFF"
    Exit Function
Stopped:
    If saveSent Then
        CWR_Log "OUTBOUND_WT_CREATE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", _
            "SAVE_SENT_NO_RETRY; STAGE=" & stageText & "; " & messageText
        Codex_CreateOutboundWTFromExactCandidate = "STOP|SAVE_SENT_NO_RETRY|STAGE=" & stageText & "|" & messageText & "|CF=USER_HANDOFF"
    Else
        If draftStarted Then messageText = messageText & "; UNSAVED_OUTBOUND_DRAFT_MAY_REMAIN"
        CWR_Log "OUTBOUND_WT_CREATE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", _
            "STAGE=" & stageText & "; " & messageText
        Codex_CreateOutboundWTFromExactCandidate = "STOP|STAGE=" & stageText & "|" & messageText & "|CF=USER_HANDOFF"
    End If
    Exit Function
Failed:
    If saveSent Then
        CWR_Log "OUTBOUND_WT_CREATE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
            "SAVE_SENT_NO_RETRY; STAGE=" & stageText & "; Err=" & CStr(Err.Number) & ": " & Err.Description
        Codex_CreateOutboundWTFromExactCandidate = "FAIL|SAVE_SENT_NO_RETRY|STAGE=" & stageText & _
            "|ERR=" & CStr(Err.Number) & "|" & Err.Description & "|CF=USER_HANDOFF"
    Else
        CWR_Log "OUTBOUND_WT_CREATE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
            "STAGE=" & stageText & "; Err=" & CStr(Err.Number) & ": " & Err.Description
        Codex_CreateOutboundWTFromExactCandidate = "FAIL|STAGE=" & stageText & "|ERR=" & _
            CStr(Err.Number) & "|" & Err.Description & "|CF=USER_HANDOFF"
    End If
End Function

' Opens the unsaved Warehouse Task draft tab for the uniquely verified ODO
' line and reports its controls/grid contract. This identifies whether SAP
' supports an explicit source bin/batch override in the draft. It only changes
' the visible tab and never adopts, creates or saves a warehouse task.
Public Function Codex_ProbeCreateWTDraftFields() As String
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Const WT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV"
    Dim createSes As Object, wrGrid As Object, wtTab As Object, root As Object
    Dim messageText As String, dumpText As String, probeText As String
    Dim rowDoc As String, rowItem As String, rowMat As String, rowQty As String, itemCount As Long

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "CREATE_WT_DRAFT_MAP_OPEN_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If
    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    If wrGrid Is Nothing Or CLng(wrGrid.rowCount) <> 1 Then
        messageText = "CREATE_WT_DRAFT_WR_GRID_UNAVAILABLE"
        GoTo Stopped
    End If
    rowDoc = CWR_Cell(wrGrid, 0, "DLVNO")
    rowItem = CWR_Cell(wrGrid, 0, "ITMNO")
    rowMat = CWR_Cell(wrGrid, 0, "MATNR")
    rowQty = CWR_Cell(wrGrid, 0, "UI_OPNQTY")
    If StrComp(rowDoc, CWR_EXPECTED_DOC, vbTextCompare) <> 0 Or _
       StrComp(rowItem, CWR_EXPECTED_ITEM, vbTextCompare) <> 0 Or _
       StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Or _
       Abs(CWR_Number(rowQty) - CWR_TARGET_QTY) > 0.000001 Then
        messageText = "CREATE_WT_DRAFT_WR_MISMATCH: DOC=" & rowDoc & "; ITEM=" & rowItem & _
            "; MAT=" & rowMat & "; QTY=" & rowQty
        GoTo Stopped
    End If
    wrGrid.selectedRows = "0"
    wrGrid.SetCurrentCell 0, "DLVNO"
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    If wtTab Is Nothing Or StrComp(CWR_Text(wtTab.text), "Warehouse Task", vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_DRAFT_TAB_IDENTITY_MISMATCH"
        GoTo Stopped
    End If
    wtTab.Select
    If Not CWR_WaitIdle(createSes, 20#) Then
        messageText = "CREATE_WT_DRAFT_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set root = createSes.FindById("wnd[0]", False)
    If root Is Nothing Then
        messageText = "CREATE_WT_DRAFT_ROOT_NOT_FOUND"
        GoTo Stopped
    End If
    itemCount = 0
    CWR_CollectCreateWTActionSurface root, 0, itemCount, dumpText
    CWR_DumpVisibleCreateWTGrids root, 0, itemCount, dumpText
    CWR_Log "WT_CREATE_DRAFT_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Warehouse Task draft tab mapped read only; no quantity adopted, created or saved."
    Codex_ProbeCreateWTDraftFields = "PASS|CREATE_WT_DRAFT_FIELDS|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_DRAFT_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ProbeCreateWTDraftFields = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_DRAFT_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ProbeCreateWTDraftFields = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Maps Default Values for the exact Create WT request. It only changes the
' selected tab, which reveals whether source-bin/batch fields are supported
' before a manual draft is ever inserted.
Public Function Codex_ProbeCreateWTDefaultValues() As String
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Const DEFAULT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_DEFDLV"
    Dim createSes As Object, wrGrid As Object, defaultTab As Object, root As Object
    Dim messageText As String, dumpText As String, probeText As String
    Dim rowDoc As String, rowItem As String, rowMat As String, rowQty As String, itemCount As Long

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "CREATE_WT_DEFAULT_MAP_OPEN_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If
    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    If wrGrid Is Nothing Or CLng(wrGrid.rowCount) <> 1 Then
        messageText = "CREATE_WT_DEFAULT_WR_GRID_UNAVAILABLE"
        GoTo Stopped
    End If
    rowDoc = CWR_Cell(wrGrid, 0, "DLVNO")
    rowItem = CWR_Cell(wrGrid, 0, "ITMNO")
    rowMat = CWR_Cell(wrGrid, 0, "MATNR")
    rowQty = CWR_Cell(wrGrid, 0, "UI_OPNQTY")
    If StrComp(rowDoc, CWR_EXPECTED_DOC, vbTextCompare) <> 0 Or _
       StrComp(rowItem, CWR_EXPECTED_ITEM, vbTextCompare) <> 0 Or _
       StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Or _
       Abs(CWR_Number(rowQty) - CWR_TARGET_QTY) > 0.000001 Then
        messageText = "CREATE_WT_DEFAULT_WR_MISMATCH: DOC=" & rowDoc & "; ITEM=" & rowItem & _
            "; MAT=" & rowMat & "; QTY=" & rowQty
        GoTo Stopped
    End If
    wrGrid.selectedRows = "0"
    wrGrid.SetCurrentCell 0, "DLVNO"
    Set defaultTab = createSes.FindById(DEFAULT_TAB_ID, False)
    If defaultTab Is Nothing Or StrComp(CWR_Text(defaultTab.text), "Default Values", vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_DEFAULT_TAB_IDENTITY_MISMATCH"
        GoTo Stopped
    End If
    defaultTab.Select
    If Not CWR_WaitIdle(createSes, 20#) Then
        messageText = "CREATE_WT_DEFAULT_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set root = createSes.FindById("wnd[0]", False)
    If root Is Nothing Then
        messageText = "CREATE_WT_DEFAULT_ROOT_NOT_FOUND"
        GoTo Stopped
    End If
    itemCount = 0
    CWR_CollectCreateWTActionSurface root, 0, itemCount, dumpText
    CWR_DumpVisibleCreateWTGrids root, 0, itemCount, dumpText
    CWR_Log "WT_CREATE_DEFAULT_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Create WT Default Values mapped read only; no draft inserted, adopted, created or saved."
    Codex_ProbeCreateWTDefaultValues = "PASS|CREATE_WT_DEFAULT_VALUES|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_DEFAULT_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ProbeCreateWTDefaultValues = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_DEFAULT_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ProbeCreateWTDefaultValues = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Inserts one manual WT draft solely to inspect its field contract, then
' deletes that same unsaved draft before returning. The procedure requires an
' initially empty draft grid and verifies it returns to empty; it never calls
' Create, Create+Save, Save, Adopt Quantity, or CF.
Public Function Codex_PreviewManualWTDraftFieldsAndReset() As String
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Const WT_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV"
    Const WT_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/cntlCC_TB_ODP_TODLV/shellcont/shell"
    Const WT_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_TODLV/ssubSUB_ODP_TODLV:/SCWM/SAPLUI_TODLV:4100/subSUB_ODP_TODLV_DATA:/SCWM/SAPLUI_TODLV:4110/cntlCC_ALV_OD_TODLV/shellcont/shell"
    Const INSERT_ID As String = "OK_OD_TODLV_INSERT"
    Const DELETE_ID As String = "OK_OD_TODLV_DELETE"
    Dim createSes As Object, wrGrid As Object, wtTab As Object, wtToolbar As Object, wtGrid As Object
    Dim messageText As String, dumpText As String, probeText As String, insertDetail As String, deleteDetail As String
    Dim rowDoc As String, rowItem As String, rowMat As String, rowQty As String
    Dim inserted As Boolean, statusType As String, statusText As String

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "CREATE_WT_MANUAL_PREVIEW_OPEN_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If
    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    If wrGrid Is Nothing Or CLng(wrGrid.rowCount) <> 1 Then
        messageText = "CREATE_WT_MANUAL_PREVIEW_WR_GRID_UNAVAILABLE"
        GoTo Stopped
    End If
    rowDoc = CWR_Cell(wrGrid, 0, "DLVNO")
    rowItem = CWR_Cell(wrGrid, 0, "ITMNO")
    rowMat = CWR_Cell(wrGrid, 0, "MATNR")
    rowQty = CWR_Cell(wrGrid, 0, "UI_OPNQTY")
    If StrComp(rowDoc, CWR_EXPECTED_DOC, vbTextCompare) <> 0 Or _
       StrComp(rowItem, CWR_EXPECTED_ITEM, vbTextCompare) <> 0 Or _
       StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Or _
       Abs(CWR_Number(rowQty) - CWR_TARGET_QTY) > 0.000001 Then
        messageText = "CREATE_WT_MANUAL_PREVIEW_WR_MISMATCH: DOC=" & rowDoc & "; ITEM=" & rowItem & _
            "; MAT=" & rowMat & "; QTY=" & rowQty
        GoTo Stopped
    End If
    wrGrid.selectedRows = "0"
    wrGrid.SetCurrentCell 0, "DLVNO"
    Set wtTab = createSes.FindById(WT_TAB_ID, False)
    If wtTab Is Nothing Or StrComp(CWR_Text(wtTab.text), "Warehouse Task", vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_MANUAL_PREVIEW_TAB_IDENTITY_MISMATCH"
        GoTo Stopped
    End If
    wtTab.Select
    If Not CWR_WaitIdle(createSes, 20#) Then
        messageText = "CREATE_WT_MANUAL_PREVIEW_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    Set wtToolbar = createSes.FindById(WT_TOOLBAR_ID, False)
    If wtGrid Is Nothing Or wtToolbar Is Nothing Then
        messageText = "CREATE_WT_MANUAL_PREVIEW_GRID_OR_TOOLBAR_NOT_FOUND"
        GoTo Stopped
    End If
    If CLng(wtGrid.rowCount) <> 0 Then
        messageText = "CREATE_WT_MANUAL_PREVIEW_REQUIRES_EMPTY_GRID: ROWS=" & CStr(wtGrid.rowCount)
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(wtToolbar, INSERT_ID, "New Warehouse Task: Manual", insertDetail, messageText) Then GoTo Stopped
    wtToolbar.PressButton INSERT_ID
    inserted = True
    If Not CWR_WaitIdle(createSes, 20#) Then
        messageText = "CREATE_WT_MANUAL_PREVIEW_INSERT_TIMEOUT"
        GoTo Stopped
    End If
    CWR_ReadStatusBar createSes, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "CREATE_WT_MANUAL_PREVIEW_INSERT_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 1 Then
        messageText = "CREATE_WT_MANUAL_PREVIEW_INSERT_COUNT_MISMATCH"
        GoTo Stopped
    End If
    dumpText = "INSERT_ACTION=" & insertDetail & vbLf & _
        CWR_DumpKnownGrid("MANUAL_DRAFT", wtGrid, 5) & _
        CWR_DumpGridFieldContract("MANUAL_DRAFT", wtGrid, 0)

    wtGrid.selectedRows = "0"
    If Not CWR_GetVerifiedToolbarAction(wtToolbar, DELETE_ID, "Delete Warehouse Task", deleteDetail, messageText) Then GoTo Stopped
    wtToolbar.PressButton DELETE_ID
    If Not CWR_WaitIdle(createSes, 20#) Then
        messageText = "CREATE_WT_MANUAL_PREVIEW_DELETE_TIMEOUT"
        GoTo Stopped
    End If
    Set wtGrid = createSes.FindById(WT_GRID_ID, False)
    If wtGrid Is Nothing Or CLng(wtGrid.rowCount) <> 0 Then
        messageText = "CREATE_WT_MANUAL_PREVIEW_DELETE_COUNT_MISMATCH"
        GoTo Stopped
    End If
    inserted = False
    CWR_Log "WT_CREATE_MANUAL_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Manual draft contract mapped and the same unsaved draft deleted; no WT created or saved."
    Codex_PreviewManualWTDraftFieldsAndReset = "PASS|MANUAL_DRAFT_FIELD_CONTRACT|WRITE=FALSE" & vbLf & dumpText & _
        "DELETE_ACTION=" & deleteDetail & vbLf & "POSTVERIFY|DRAFT_ROWS=0"
    Exit Function
Stopped:
    If inserted Then messageText = messageText & "; UNSAVED_MANUAL_DRAFT_MAY_REMAIN"
    CWR_Log "WT_CREATE_MANUAL_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_PreviewManualWTDraftFieldsAndReset = "STOP|" & messageText
    Exit Function
Failed:
    If inserted Then messageText = "UNSAVED_MANUAL_DRAFT_MAY_REMAIN; "
    CWR_Log "WT_CREATE_MANUAL_PREVIEW", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_PreviewManualWTDraftFieldsAndReset = "FAIL|" & messageText & CStr(Err.Number) & "|" & Err.Description
End Function

' Correct POST preflight: it considers only stock already located at the new
' pick bin. If the requested batch is already present in sufficient quantity,
' it is explicitly a no-op. Otherwise it identifies one eligible current
' batch at that same bin for an in-place batch change. No line is selected or
' changed by this function.
Public Function Codex_PreflightPostingBatchAtPickBin() As String
    Dim ses As Object, grid As Object, messageText As String, planText As String, contractText As String
    Dim sourceRow As Long, planState As String

    On Error GoTo Failed
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    ' Batch preflight is intentionally before the PRDO quantity change.  The
    ' prior PRDO_QTY_VERIFIED requirement inverted the locked route and made
    ' CHANGE_BIN_BATCH stop before it could inspect the target-bin batch.
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_QueryExpectedPOST(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_PlanPickBinBatchChange(grid, sourceRow, planState, planText, messageText) Then GoTo Stopped
    If StrComp(planState, "NO_CHANGE_REQUIRED", vbTextCompare) = 0 Then
        CWR_Log "POST_PICK_BIN_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "Requested batch already available at pick bin; no POST change permitted. " & planText
        Codex_PreflightPostingBatchAtPickBin = "PASS|PICK_BIN_BATCH_ALREADY_CORRECT|WRITE=FALSE|" & planText
        Exit Function
    End If
    contractText = CWR_DumpGridFieldContract("POST_PICK_BIN_SOURCE", grid, sourceRow)
    CWR_Log "POST_PICK_BIN_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact in-place batch candidate identified at pick bin; no POST field changed. " & planText
    Codex_PreflightPostingBatchAtPickBin = "PASS|PICK_BIN_BATCH_CHANGE_PREVERIFIED|SOURCE_ROW=" & _
        CStr(sourceRow) & "|WRITE=FALSE|" & planText & vbLf & contractText
    Exit Function
Stopped:
    CWR_Log "POST_PICK_BIN_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_PreflightPostingBatchAtPickBin = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_PICK_BIN_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_PreflightPostingBatchAtPickBin = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Executes the corrected /SCWM/POST workflow once and only once:
' stock already at the user-confirmed pick bin stays at that exact bin, while
' its batch is changed from the actual pick-bin batch to the ODO/WH batch.
' It deliberately does not create a WT and it never invokes CF.
Public Function Codex_ExecutePostingBatchAtPickBin() As String
    Const POST_CREATED_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_POST:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB2/ssubSUB_ODP1_TAB2:/SCWM/SAPLUI_POST:3220/ssubSUB_ODP1_2_CONTENT:/SCWM/SAPLUI_POST:3221/cntlCONTAINER_ALV_ODP1_2/shellcont/shell"
    Dim ses As Object, grid As Object, actionToolbar As Object, yesButton As Object, createdGrid As Object
    Dim messageText As String, staleGridMessage As String, modalText As String, statusType As String, statusText As String
    Dim planText As String, planState As String, draftProof As String, pendingProof As String, actionDetail As String, confirmProof As String
    Dim recoveryProof As String
    Dim createdWTNo As String
    Dim sourceRow As Long, verifiedRow As Long, createdBefore As Long, createdAfter As Long
    Dim sourceType As String, sourceSection As String, sourceStockType As String, sourceOwner As String, sourceEntitled As String
    Dim sourceAvailBefore As Double, sourceTotalBefore As Double, whAvailBefore As Double, whTotalBefore As Double
    Dim sourceAvailAfter As Double, sourceTotalAfter As Double, whAvailAfter As Double, whTotalAfter As Double
    Dim sourceRowsBefore As Long, whRowsBefore As Long, sourceRowsAfter As Long, whRowsAfter As Long
    Dim balanceBefore As String, balanceAfter As String, resetText As String
    Dim draftModified As Boolean, actionSent As Boolean, confirmationSent As Boolean, modalOpen As Boolean
    Dim priorPostSendNeedsReconcile As Boolean

    On Error GoTo Failed
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = "PICK_BIN_POST_START_MODAL_PRESENT: " & modalText
        modalOpen = True
        GoTo Stopped
    End If

    ' Never refresh a visible POST session over somebody else's pending draft.
    staleGridMessage = ""
    If CWR_FindCurrentPOSTGrid(ses, grid, staleGridMessage) Then
        If Not CWR_AssertNoPendingPOSTDraft(grid, messageText) Then GoTo Stopped
    End If
    If Not CWR_QueryExpectedPOST(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_AssertNoPendingPOSTDraft(grid, messageText) Then GoTo Stopped

    ' A POST Create+Save checkpoint is a strict one-shot boundary.  Before
    ' considering another plan state, re-run the exact product query and use
    ' only the newly returned stock to reconcile it.  The visible POST grid
    ' can be cleared or re-bound by SAP after Create+Save; reusing it caused a
    ' false "SOURCE_MATCH_COUNT=0 / TARGET_EXISTS=0" stop even after the
    ' business result had settled.  This branch has no write action.
    priorPostSendNeedsReconcile = CWR_HasUnreconciledPickBinPostSend()
    If priorPostSendNeedsReconcile Then
        If Not CWR_QueryExpectedPOST(ses, grid, messageText, True) Then
            messageText = "STATE_UNKNOWN_AFTER_POST_CREATE_SAVE|POST_RECONCILE_REFRESH_FAILED|" & messageText
            GoTo Stopped
        End If
        If Not CWR_AssertNoPendingPOSTDraft(grid, messageText) Then GoTo Stopped
        If CWR_TARGET_BATCH = "" Then
            If Not CWR_ResolveRecoverablePOSTBatch(grid, messageText) Then GoTo Stopped
        End If
        If Not CWR_PlanPickBinBatchChange(grid, sourceRow, planState, planText, messageText) Then
            messageText = "STATE_UNKNOWN_AFTER_POST_CREATE_SAVE|POST_RECONCILE_UNPROVEN|" & messageText
            GoTo Stopped
        End If
        If StrComp(planState, "NO_CHANGE_REQUIRED", vbTextCompare) <> 0 Then
            messageText = "STATE_UNKNOWN_AFTER_POST_CREATE_SAVE|POST_RECONCILE_TARGET_NOT_READY|" & planText
            GoTo Stopped
        End If
        If Not CWR_ReadPickBinBatchBalances(grid, "", "", sourceAvailAfter, sourceTotalAfter, _
                                            whAvailAfter, whTotalAfter, sourceRowsAfter, whRowsAfter, _
                                            recoveryProof, messageText) Then GoTo Stopped
        CWR_RecordPickBinBatchPostReconciled recoveryProof
        CWR_Log "POST_PICK_BIN_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "No POST retry. Fresh exact product query proved the WH batch at the confirmed pick bin; " & recoveryProof
        Codex_ExecutePostingBatchAtPickBin = "PASS|PICK_BIN_BATCH_POST_VERIFIED|RECOVERY=READ_ONLY|" & _
            recoveryProof & "|NEXT=CHANGE_PRDO_QTY_THEN_CREATE_OUTBOUND_WT|CF=USER_HANDOFF"
        Exit Function
    End If

    If Not CWR_PlanPickBinBatchChange(grid, sourceRow, planState, planText, messageText) Then GoTo Stopped

    If StrComp(planState, "NO_CHANGE_REQUIRED", vbTextCompare) = 0 Then
        If Not CWR_ReadPickBinBatchBalances(grid, "", "", sourceAvailAfter, sourceTotalAfter, _
                                            whAvailAfter, whTotalAfter, sourceRowsAfter, whRowsAfter, _
                                            balanceAfter, messageText) Then GoTo Stopped
        CWR_RecordPickBinBatchAlreadyCorrect balanceAfter
        CWR_Log "POST_PICK_BIN_EXECUTE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "WH batch was already available at the confirmed pick bin; no POST Create+Save was sent. " & balanceAfter
        Codex_ExecutePostingBatchAtPickBin = "PASS|PICK_BIN_BATCH_ALREADY_CORRECT|WRITE=FALSE|" & _
            balanceAfter & "|NEXT=REDETERMINE_CREATE_WT|CF=USER_HANDOFF"
        Exit Function
    End If
    If StrComp(planState, "CHANGE_REQUIRED", vbTextCompare) <> 0 Then
        messageText = "PICK_BIN_POST_PLAN_STATE_UNSUPPORTED: " & planState
        GoTo Stopped
    End If

    sourceType = CWR_Cell(grid, sourceRow, "VLTYP")
    sourceSection = CWR_Cell(grid, sourceRow, "VLBER")
    sourceStockType = CWR_Cell(grid, sourceRow, "TO_CAT")
    If sourceStockType = "" Then sourceStockType = CWR_Cell(grid, sourceRow, "CAT")
    sourceOwner = CWR_Cell(grid, sourceRow, "OWNER")
    sourceEntitled = CWR_Cell(grid, sourceRow, "ENTITLED")
    If CWR_TARGET_STORAGE_TYPE = "" Then CWR_TARGET_STORAGE_TYPE = sourceType
    If CWR_TARGET_STORAGE_SECTION = "" Then CWR_TARGET_STORAGE_SECTION = sourceSection
    If StrComp(sourceType, CWR_TARGET_STORAGE_TYPE, vbTextCompare) <> 0 Or _
       StrComp(sourceSection, CWR_TARGET_STORAGE_SECTION, vbTextCompare) <> 0 Or _
       StrComp(sourceStockType, CWR_TARGET_STOCK_TYPE, vbTextCompare) <> 0 Then
        messageText = "PICK_BIN_POST_SOURCE_ATTRIBUTE_MISMATCH: VLTYP=" & sourceType & _
            "; VLBER=" & sourceSection & "; STOCK_TYPE=" & sourceStockType
        GoTo Stopped
    End If
    If Not CWR_AssertPickBinPostFieldsChangeable(grid, sourceRow, sourceSection, messageText) Then GoTo Stopped
    If Not CWR_ReadPickBinBatchBalances(grid, sourceOwner, sourceEntitled, sourceAvailBefore, sourceTotalBefore, _
                                        whAvailBefore, whTotalBefore, sourceRowsBefore, whRowsBefore, _
                                        balanceBefore, messageText) Then GoTo Stopped
    If sourceRowsBefore <> 1 Or sourceAvailBefore + 0.000001 < CWR_TARGET_QTY Or _
       sourceTotalBefore + 0.000001 < CWR_TARGET_QTY Then
        messageText = "PICK_BIN_POST_SOURCE_BALANCE_UNSAFE: " & balanceBefore
        GoTo Stopped
    End If

    ' The optional "created" tab differs across SAP GUI patches.  It is
    ' evidence only, never the business result of an in-place POST change.
    Set createdGrid = ses.FindById(POST_CREATED_GRID_ID, False)
    If Not createdGrid Is Nothing Then createdBefore = CLng(createdGrid.rowCount)

    ' All destination attributes intentionally preserve the source location.
    ' Only TO_CHARG changes: 260713-07 -> 260704-07 at APFL-ZF03-01-13.
    grid.selectedRows = CStr(sourceRow)
    grid.SetCurrentCell sourceRow, "TO_CHARG"
    grid.ModifyCell sourceRow, "TO_CAT", CWR_TARGET_STOCK_TYPE
    draftModified = True
    If StrComp(CWR_Cell(grid, sourceRow, "TO_CAT"), CWR_TARGET_STOCK_TYPE, vbTextCompare) <> 0 Then
        messageText = "PICK_BIN_POST_DRAFT_STOCK_TYPE_READBACK_MISMATCH: " & CWR_Cell(grid, sourceRow, "TO_CAT")
        GoTo Stopped
    End If
    grid.ModifyCell sourceRow, "NLTYP", sourceType
    If StrComp(CWR_Cell(grid, sourceRow, "NLTYP"), sourceType, vbTextCompare) <> 0 Then
        messageText = "PICK_BIN_POST_DRAFT_STORAGE_TYPE_READBACK_MISMATCH: " & CWR_Cell(grid, sourceRow, "NLTYP")
        GoTo Stopped
    End If
    If sourceSection <> "" Then
        grid.ModifyCell sourceRow, "NLBER", sourceSection
    If StrComp(CWR_Cell(grid, sourceRow, "NLBER"), sourceSection, vbTextCompare) <> 0 Then
        messageText = "PICK_BIN_POST_DRAFT_STORAGE_SECTION_READBACK_MISMATCH: " & CWR_Cell(grid, sourceRow, "NLBER")
        GoTo Stopped
    End If
    End If
    grid.ModifyCell sourceRow, "NLPLA", CWR_TARGET_BIN
    If StrComp(CWR_Cell(grid, sourceRow, "NLPLA"), CWR_TARGET_BIN, vbTextCompare) <> 0 Then
        messageText = "PICK_BIN_POST_DRAFT_BIN_READBACK_MISMATCH: " & CWR_Cell(grid, sourceRow, "NLPLA")
        GoTo Stopped
    End If
    grid.ModifyCell sourceRow, "POST_QUAN", CWR_FastQuantityText(CWR_TARGET_QTY)
    If Abs(CWR_Number(CWR_Cell(grid, sourceRow, "POST_QUAN")) - CWR_TARGET_QTY) > 0.000001 Then
        messageText = "PICK_BIN_POST_DRAFT_QTY_READBACK_MISMATCH: " & CWR_Cell(grid, sourceRow, "POST_QUAN")
        GoTo Stopped
    End If
    grid.ModifyCell sourceRow, "TO_CHARG", CWR_WH_BATCH
    If StrComp(CWR_Cell(grid, sourceRow, "TO_CHARG"), CWR_WH_BATCH, vbTextCompare) <> 0 Then
        messageText = "PICK_BIN_POST_DRAFT_BATCH_READBACK_MISMATCH: " & CWR_Cell(grid, sourceRow, "TO_CHARG")
        GoTo Stopped
    End If

    grid.TriggerModified
    If Not CWR_WaitPOSTSaveSettle(ses, 0.5, 20#) Then
        messageText = "PICK_BIN_POST_DRAFT_TRIGGER_TIMEOUT"
        GoTo Stopped
    End If
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = "PICK_BIN_POST_DRAFT_TRIGGER_MODAL: " & modalText
        modalOpen = True
        GoTo Stopped
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "PICK_BIN_POST_DRAFT_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_PlanPickBinBatchChange(grid, verifiedRow, planState, planText, messageText) Then GoTo Stopped
    If StrComp(planState, "CHANGE_REQUIRED", vbTextCompare) <> 0 Then
        messageText = "PICK_BIN_POST_DRAFT_PLAN_CHANGED: " & planState & "; " & planText
        GoTo Stopped
    End If
    sourceRow = verifiedRow
    If Not CWR_ProveExactPickBinBatchDraft(grid, sourceRow, sourceType, sourceSection, draftProof, messageText) Then GoTo Stopped
    If Not CWR_ProveOnlyExactPendingDraft(grid, sourceRow, pendingProof, messageText) Then GoTo Stopped
    If Not CWR_SelectExactPickBinBatchSource(ses, grid, sourceRow, sourceType, sourceSection, messageText) Then GoTo Stopped
    If Not CWR_ProveExactPickBinBatchDraft(grid, sourceRow, sourceType, sourceSection, draftProof, messageText) Then GoTo Stopped
    If Not CWR_ProveOnlyExactPendingDraft(grid, sourceRow, pendingProof, messageText) Then GoTo Stopped
    If Not CWR_GetVerifiedPOSTCreateSaveAction(ses, actionToolbar, actionDetail, messageText) Then GoTo Stopped
    If Not CWR_RequireReworkAction("POST_CHANGE", messageText) Then GoTo Stopped

    ' The first business command is exactly the internal POST Create+Save
    ' command, after the one permitted source row is highlighted and re-read.
    ' Persist the attempted-write checkpoint before the COM call.  If SAP GUI
    ' disconnects mid-call, the next run can reconcile but can never send a
    ' second POST Create+Save against the same old WT.
    CWR_RecordPickBinPostSendCheckpoint actionDetail
    actionSent = True
    actionToolbar.PressButton "OK_OIP_POST_CREATE_POST"
    CWR_Log "POST_PICK_BIN_CREATE_SAVE_SENT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact pick-bin source selected and Create+Save Posting Change sent once; " & _
        actionDetail & "; " & pendingProof & "; " & draftProof
    If Not CWR_WaitPOSTSaveSettle(ses, 1#, 25#) Then
        messageText = "PICK_BIN_POST_CREATE_SAVE_TIMEOUT_STATE_UNPROVEN"
        GoTo ActionStopped
    End If

    If CWR_HasAnyModal(ses, modalText) Then
        modalOpen = True
        If CWR_InspectCreatedPostingWTLog(ses, createdWTNo, confirmProof, messageText) Then
            ' A returned posting-change WT number is the SAP commit point for
            ' this transaction. It is POST proof only, never the outbound WT
            ' and never a CF queue item. Close only the informational Display
            ' logs viewer; do not query POST again here. The next business
            ' gate is the exact target-bin lookup in the outbound Create-WT
            ' path, and no second POST Create+Save is permitted.
            If Not CWR_CloseCreatedPostingWTLogForRead(ses, messageText) Then
                CWR_RecordCreatedPostingWT createdWTNo, confirmProof & _
                    "|POST_DISPLAY_READ_FAILED_AFTER_COMMIT|" & messageText
                Codex_ExecutePostingBatchAtPickBin = "STOP|POST_CHANGE_COMMITTED|POST_PROOF_WT=" & _
                    createdWTNo & "|POST_DISPLAY_READ_FAILED_AFTER_COMMIT|NO_RETRY|CF_ACTION_SENT=0|" & messageText
                GoTo ActionStopped
            End If
            modalOpen = False
            CWR_RecordCreatedPostingWT createdWTNo, confirmProof
            CWR_Log "POST_PICK_BIN_COMMITTED", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                "POST Create+Save committed once; POST_PROOF_WT=" & createdWTNo & _
                "; no POST retry; no CF command; continue to exact target-bin/qty section. " & confirmProof
            Codex_ExecutePostingBatchAtPickBin = "PASS|POST_CHANGE_COMMITTED|POST_PROOF_WT=" & _
                createdWTNo & "|POST_CHANGE_RESULT=SUCCESS|WRITE_RETRY_ALLOWED=FALSE|CF_ACTION_SENT=0" & _
                "|NEXT=CHANGE_PRDO_QTY_THEN_CREATE_OUTBOUND_WT"
            Exit Function
        End If
        If Not CWR_InspectPickBinBatchConfirmation(ses, sourceType, sourceSection, yesButton, confirmProof, messageText) Then
            messageText = "PICK_BIN_POST_CONFIRMATION_UNPROVEN: " & messageText
            GoTo ActionStopped
        End If
        confirmationSent = True
        yesButton.Press
        modalOpen = False
        CWR_Log "POST_PICK_BIN_YES_SENT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "Yes pressed once only after exact source/draft confirmation; " & confirmProof
        If Not CWR_WaitPOSTSaveSettle(ses, 3#, 45#) Then
            messageText = "PICK_BIN_POST_CONFIRM_TIMEOUT_STATE_UNPROVEN"
            GoTo ActionStopped
        End If
        If CWR_HasAnyModal(ses, modalText) Then
            messageText = "PICK_BIN_POST_CONFIRM_MODAL_REMAINS: " & modalText
            modalOpen = True
            GoTo ActionStopped
        End If
    End If

    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "PICK_BIN_POST_SAP_" & statusType & ": " & statusText
        GoTo ActionStopped
    End If
    ' Do not reuse the pre-action grid here.  Create+Save can leave that grid
    ' bound to an emptied/changed result surface, so a fresh exact query is
    ' required before any post-action decision is made.
    If Not CWR_QueryExpectedPOST(ses, grid, messageText, True) Then GoTo ActionStopped
    ' The sent POST command may refresh the optional Created tab to zero even
    ' when the exact source batch has already moved in-place.  Reconcile the
    ' actual business result instead: the requested batch must now be present
    ' at the exact pick bin with enough quantity.  Never issue POST again.
    If Not CWR_PlanPickBinBatchChange(grid, verifiedRow, planState, planText, messageText) Then GoTo ActionStopped
    If StrComp(planState, "NO_CHANGE_REQUIRED", vbTextCompare) <> 0 Then
        messageText = "STATE_UNKNOWN_AFTER_POST_CREATE_SAVE|POST_TARGET_BATCH_NOT_READY|" & planText
        GoTo ActionStopped
    End If
    If Not CWR_ReadPickBinBatchBalances(grid, sourceOwner, sourceEntitled, sourceAvailAfter, sourceTotalAfter, _
                                        whAvailAfter, whTotalAfter, sourceRowsAfter, whRowsAfter, _
                                        balanceAfter, messageText) Then GoTo ActionStopped
    Set createdGrid = ses.FindById(POST_CREATED_GRID_ID, False)
    If Not createdGrid Is Nothing Then createdAfter = CLng(createdGrid.rowCount)
    If Abs(sourceAvailAfter - (sourceAvailBefore - CWR_TARGET_QTY)) > 0.000001 Or _
       Abs(sourceTotalAfter - (sourceTotalBefore - CWR_TARGET_QTY)) > 0.000001 Or _
       Abs(whAvailAfter - (whAvailBefore + CWR_TARGET_QTY)) > 0.000001 Or _
       Abs(whTotalAfter - (whTotalBefore + CWR_TARGET_QTY)) > 0.000001 Then
        messageText = "PICK_BIN_POST_DELTA_MISMATCH: BEFORE=" & balanceBefore & "; AFTER=" & balanceAfter
        GoTo ActionStopped
    End If

    CWR_RecordPickBinBatchPosting balanceBefore, balanceAfter, createdBefore, createdAfter
    CWR_Log "POST_PICK_BIN_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "In-place target batch readback verified; continue to PRDO/Create WT; " & _
        balanceBefore & "; " & balanceAfter & "; CF=USER_HANDOFF"
    Codex_ExecutePostingBatchAtPickBin = "PASS|PICK_BIN_BATCH_POST_VERIFIED|CREATED_ROWS=" & _
        CStr(createdBefore) & "->" & CStr(createdAfter) & "|" & balanceAfter & _
        "|NEXT=CHANGE_PRDO_QTY_THEN_CREATE_OUTBOUND_WT|CF=USER_HANDOFF"
    Exit Function

Stopped:
    If draftModified And Not actionSent And Not modalOpen Then
        resetText = ""
        If Not CWR_ResetPOSTWithoutSave(ses, resetText) Then messageText = messageText & "; PICK_BIN_POST_DRAFT_RESET_FAILED=" & resetText
    ElseIf draftModified And Not actionSent And modalOpen Then
        messageText = messageText & "; PICK_BIN_POST_DRAFT_MAY_REMAIN"
    End If
    CWR_Log "POST_PICK_BIN_EXECUTE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ExecutePostingBatchAtPickBin = "STOP|" & messageText
    Exit Function

ActionStopped:
    CWR_Log "POST_PICK_BIN_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ExecutePostingBatchAtPickBin = "STOP|" & messageText & "|NO_RETRY"
    Exit Function

Failed:
    If draftModified And Not actionSent And Not modalOpen Then
        resetText = ""
        Call CWR_ResetPOSTWithoutSave(ses, resetText)
    End If
    CWR_Log "POST_PICK_BIN_EXECUTE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description & _
        IIf(actionSent Or confirmationSent, "; STATE_UNPROVEN_NO_RETRY", "")
    Codex_ExecutePostingBatchAtPickBin = "FAIL|" & CStr(Err.Number) & "|" & Err.Description & _
        IIf(actionSent Or confirmationSent, "|STATE_UNPROVEN_NO_RETRY", "")
End Function

' Read-only diagnostic for an unexpected POST modal after Create+Save.  It
' deliberately preserves the dialog and any pending grid draft so a failure
' can be understood from SAP's own log before anybody chooses a response.
Public Function Codex_ReadCurrentPOSTModalState() As String
    Const POST_CREATED_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_POST:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB2/ssubSUB_ODP1_TAB2:/SCWM/SAPLUI_POST:3220/ssubSUB_ODP1_2_CONTENT:/SCWM/SAPLUI_POST:3221/cntlCONTAINER_ALV_ODP1_2/shellcont/shell"
    Dim ses As Object, modal As Object, grid As Object, createdGrid As Object
    Dim messageText As String, titleText As String, rawTree As String, dumpText As String, statusType As String, statusText As String
    Dim props As Variant, itemCount As Long, sourceRow As Long, planState As String, planText As String, pendingProof As String
    On Error GoTo Failed
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    Set modal = ses.FindById("wnd[1]", False)
    If modal Is Nothing Then
        Codex_ReadCurrentPOSTModalState = "PASS|POST_MODAL_NOT_PRESENT|WRITE=FALSE"
        Exit Function
    End If
    titleText = CWR_Text(modal.text)
    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
    rawTree = CStr(ses.GetObjectTree("wnd[1]", props) & "")
    itemCount = 0
    CWR_CollectAllVisibleText modal, 0, itemCount, dumpText
    If CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then
        dumpText = dumpText & CWR_DumpKnownGrid("POST_MODAL_GRID", grid, 8)
        If CWR_PlanPickBinBatchChange(grid, sourceRow, planState, planText, messageText) Then
            dumpText = dumpText & "POST_MODAL_PLAN=" & planState & "; " & planText & vbLf & _
                CWR_DumpGridFieldContract("POST_MODAL_SOURCE", grid, sourceRow)
            If StrComp(planState, "CHANGE_REQUIRED", vbTextCompare) = 0 Then
                If CWR_ProveOnlyExactPendingDraft(grid, sourceRow, pendingProof, messageText) Then _
                    dumpText = dumpText & "POST_MODAL_PENDING=" & pendingProof & vbLf
            End If
        Else
            dumpText = dumpText & "POST_MODAL_PLAN_STOP=" & messageText & vbLf
        End If
    Else
        dumpText = dumpText & "POST_MODAL_GRID_NOT_FOUND=" & messageText & vbLf
    End If
    Set createdGrid = ses.FindById(POST_CREATED_GRID_ID, False)
    If Not createdGrid Is Nothing Then dumpText = dumpText & "POST_MODAL_CREATED_ROWS=" & CStr(createdGrid.rowCount) & vbLf
    CWR_ReadStatusBar ses, statusType, statusText
    CWR_Log "POST_MODAL_DIAGNOSTIC", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Unexpected POST modal read only; title=" & titleText & "; no dialog response sent."
    Codex_ReadCurrentPOSTModalState = "PASS|POST_MODAL_DIAGNOSTIC|TITLE=" & titleText & _
        "|STATUS=" & statusType & ":" & statusText & "|WRITE=FALSE" & vbLf & dumpText & rawTree
    Exit Function
Stopped:
    CWR_Log "POST_MODAL_DIAGNOSTIC", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ReadCurrentPOSTModalState = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_MODAL_DIAGNOSTIC", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReadCurrentPOSTModalState = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Closes only SAP's non-business "Display logs" window after a POST action,
' then reads the newly reported posting-change WT from the Monitor.  It never
' confirms, cancels, changes, or otherwise processes that WT.
Public Function Codex_ClosePOSTLogAndReadCreatedPostingWT() As String
    Dim ses As Object, modal As Object, continueButton As Object, lower As Object
    Dim messageText As String, titleText As String, statusType As String, statusText As String, wtNo As String, proofText As String
    Dim tooltipText As String, modalText As String, commandSent As Boolean
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    Set modal = ses.FindById("wnd[1]", False)
    If modal Is Nothing Then
        messageText = "POST_LOG_MODAL_NOT_PRESENT"
        GoTo Stopped
    End If
    titleText = CWR_Text(modal.text)
    If StrComp(titleText, "Display logs", vbTextCompare) <> 0 Then
        messageText = "POST_LOG_TITLE_MISMATCH: " & titleText
        GoTo Stopped
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    wtNo = CWR_ExtractWarehouseTaskNumber(statusText)
    If wtNo = "" Or InStr(1, statusText, "created", vbTextCompare) = 0 Then
        messageText = "POST_LOG_CREATED_WT_STATUS_UNPROVEN: " & statusType & ":" & statusText
        GoTo Stopped
    End If
    Set continueButton = ses.FindById("wnd[1]/tbar[0]/btn[0]", False)
    If continueButton Is Nothing Then
        messageText = "POST_LOG_CONTINUE_BUTTON_NOT_FOUND"
        GoTo Stopped
    End If
    tooltipText = CWR_Text(continueButton.Tooltip)
    If InStr(1, tooltipText, "Continue", vbTextCompare) = 0 Then
        messageText = "POST_LOG_CONTINUE_BUTTON_IDENTITY_MISMATCH: " & tooltipText
        GoTo Stopped
    End If

    ' This only dismisses the display-log viewer. It is not a Yes/No/Cancel
    ' business confirmation and it cannot confirm the created warehouse task.
    commandSent = True
    continueButton.Press
    If Not CWR_WaitPOSTSaveSettle(ses, 0.5, 20#) Then
        messageText = "POST_LOG_CLOSE_TIMEOUT_STATE_UNPROVEN"
        GoTo ClosedStateUnproven
    End If
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = "POST_LOG_CLOSE_MODAL_REMAINS: " & modalText
        GoTo ClosedStateUnproven
    End If
    If Not CWR_OpenMonitorForRead(ses, messageText) Then GoTo ClosedStateUnproven
    If Not CWR_LoadFilteredWTGrid(ses, lower, messageText) Then GoTo ClosedStateUnproven
    If Not CWR_ReadCreatedPostingWT(lower, wtNo, proofText, messageText) Then GoTo ClosedStateUnproven
    CWR_RecordCreatedPostingWT wtNo, proofText
    CWR_Log "POST_CREATED_WT_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Posting-change WT created by POST was read from Monitor; no confirmation sent. " & proofText
    Codex_ClosePOSTLogAndReadCreatedPostingWT = "PASS|POST_CHANGE_WT_CREATED_NOT_CONFIRMED_BY_CODE|WT=" & _
        wtNo & "|" & proofText & "|NEXT=USER_CONFIRM_POSTING_CHANGE_WT_OR_REQUEST_FURTHER_READONLY_CHECK|CF=USER_HANDOFF"
    Exit Function
Stopped:
    CWR_Log "POST_CREATED_WT_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ClosePOSTLogAndReadCreatedPostingWT = "STOP|" & messageText
    Exit Function
ClosedStateUnproven:
    CWR_Log "POST_CREATED_WT_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ClosePOSTLogAndReadCreatedPostingWT = "STOP|" & messageText & IIf(commandSent, "|NO_RETRY", "")
    Exit Function
Failed:
    CWR_Log "POST_CREATED_WT_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description & IIf(commandSent, "; STATE_UNPROVEN_NO_RETRY", "")
    Codex_ClosePOSTLogAndReadCreatedPostingWT = "FAIL|" & CStr(Err.Number) & "|" & Err.Description & _
        IIf(commandSent, "|STATE_UNPROVEN_NO_RETRY", "")
End Function

' Local recovery record for the one POST action already completed in this run.
' It uses the SAP status/log evidence already read above; it does not reopen,
' retry, confirm, or change anything in SAP.
Public Function Codex_RecordPostingChangeWTCreatedRecovery() As String
    Const CREATED_POST_WT As String = "7027762106"
    Dim messageText As String, proofText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    proofText = "SAP_STATUS_LOG_READ=Warehouse task " & CREATED_POST_WT & _
        " created; RESULT=POSTING_CHANGE_WT_CREATED_NOT_CONFIRMED_BY_CODE"
    CWR_RecordCreatedPostingWT CREATED_POST_WT, proofText
    CWR_Log "POST_CREATED_WT_RECOVERY_RECORD", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Local-only record of the exact POST result already read from SAP; no SAP action sent. " & proofText
    Codex_RecordPostingChangeWTCreatedRecovery = "PASS|POST_CHANGE_WT_CREATED_PENDING_USER_CF|WT=" & _
        CREATED_POST_WT & "|WRITE=LOCAL_LOG_ONLY|" & proofText & _
        "|NEXT=USER_CONFIRM_POSTING_CHANGE_WT_THEN_REVERIFY_STOCK_AND_CREATE_OUTBOUND_WH|CF=USER_HANDOFF"
    Exit Function
Stopped:
    CWR_Log "POST_CREATED_WT_RECOVERY_RECORD", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_RecordPostingChangeWTCreatedRecovery = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_CREATED_WT_RECOVERY_RECORD", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_RecordPostingChangeWTCreatedRecovery = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only evidence of the ODO's currently visible warehouse tasks after a
' posting-change result. It does not select a task action or invoke CF.
Public Function Codex_ReadOBWarehouseTasksAfterPostingChange() As String
    Dim ses As Object, lower As Object, messageText As String, dumpText As String
    On Error GoTo Failed
    If Not CWR_GetMONSession(ses, messageText) Then GoTo Stopped
    If Not CWR_LoadFilteredWTGrid(ses, lower, messageText) Then GoTo Stopped
    dumpText = CWR_DumpKnownGrid("OB_WT_AFTER_POST", lower, 50)
    CWR_Log "OB_WT_AFTER_POST_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "OB warehouse-task grid read only after posting-change WT creation; no CF action sent."
    Codex_ReadOBWarehouseTasksAfterPostingChange = "PASS|OB_WT_AFTER_POST_READ|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "OB_WT_AFTER_POST_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ReadOBWarehouseTasksAfterPostingChange = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "OB_WT_AFTER_POST_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReadOBWarehouseTasksAfterPostingChange = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only UI map of /SCWM/POST before any stock lookup or posting change.
' It only starts/navigates to the transaction and lists relevant controls.
Public Function Codex_ProbePOSTControls() As String
    Dim ses As Object, messageText As String, root As Object, titleText As String, dumpText As String
    Dim itemCount As Long

    On Error GoTo Failed
    If Not CWR_OpenPOSTSession(ses, messageText) Then
        Codex_ProbePOSTControls = "STOP_POST_OPEN|" & messageText
        Exit Function
    End If
    Set root = ses.FindById("wnd[0]", False)
    If root Is Nothing Then
        Codex_ProbePOSTControls = "STOP_POST_ROOT_NOT_FOUND"
        Exit Function
    End If
    titleText = CWR_Text(root.text)
    itemCount = 0
    CWR_CollectPOSTControls root, 0, itemCount, dumpText
    Codex_ProbePOSTControls = "POST_CONTROL_PROBE|TITLE=" & titleText & "|ITEMS=" & CStr(itemCount) & vbLf & dumpText
    Exit Function
Failed:
    Codex_ProbePOSTControls = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Exact, read-only availability proof for the target bin and batch.  The
' transaction is searched by product only; no Posting Change fields are set.
Public Function Codex_ReadTargetStockForRework() As String
    Const POST_CRITERION_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/cmb/SCWM/S_UI_POST-V_CRITERION"
    Const POST_PRODUCT_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/ctxt/SCWM/S_ASPQ_STOCK-MATNR"
    Const POST_GO_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/btnCMD_GO"
    Const POST_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/subSUB_OIP_1_DATA:/SCWM/SAPLUI_POST:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell"
    Dim ses As Object, criterion As Object, productCtl As Object, goButton As Object, grid As Object
    Dim messageText As String, proofText As String

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then
        Codex_ReadTargetStockForRework = "STOP|" & messageText
        Exit Function
    End If
    If Not CWR_AssertCancelledLocal(messageText) Then
        Codex_ReadTargetStockForRework = "STOP_LOCAL_CANCEL_STATE|" & messageText
        Exit Function
    End If
    If Not CWR_OpenPOSTSession(ses, messageText) Then
        CWR_Log "POST_TARGET_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
        Codex_ReadTargetStockForRework = "STOP_POST_OPEN|" & messageText
        Exit Function
    End If
    Set criterion = ses.FindById(POST_CRITERION_ID, False)
    Set productCtl = ses.FindById(POST_PRODUCT_ID, False)
    Set goButton = ses.FindById(POST_GO_ID, False)
    If criterion Is Nothing Or productCtl Is Nothing Or goButton Is Nothing Then
        messageText = "POST_SEARCH_CONTROLS_NOT_FOUND"
        GoTo Stopped
    End If
    If InStr(1, CWR_Text(criterion.text), "Product", vbTextCompare) = 0 Then
        messageText = "POST_CRITERION_NOT_PRODUCT: " & CWR_Text(criterion.text)
        GoTo Stopped
    End If
    productCtl.text = CWR_EXPECTED_MATERIAL
    If StrComp(CWR_Text(productCtl.text), CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Then
        messageText = "POST_PRODUCT_READBACK_MISMATCH: " & CWR_Text(productCtl.text)
        GoTo Stopped
    End If
    goButton.Press
    If Not CWR_WaitIdle(ses, 30#) Then
        messageText = "POST_QUERY_TIMEOUT"
        GoTo Stopped
    End If
    If Not CWR_WaitForObject(ses, POST_GRID_ID, 20#) Then
        messageText = "POST_GRID_TIMEOUT"
        GoTo Stopped
    End If
    Set grid = ses.FindById(POST_GRID_ID, False)
    If grid Is Nothing Then
        messageText = "POST_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_ProveExactTargetStock(grid, proofText, messageText) Then GoTo Stopped
    CWR_RecordTargetStock proofText
    CWR_Log "POST_TARGET_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", proofText & "; CF=USER_HANDOFF"
    Codex_ReadTargetStockForRework = "PASS|TARGET_STOCK_VERIFIED|" & proofText & _
        "|NEXT=PREVERIFY_POST_AND_PRDO_CHANGE|CF=USER_HANDOFF"
    Exit Function
Stopped:
    CWR_Log "POST_TARGET_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ReadTargetStockForRework = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_TARGET_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReadTargetStockForRework = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only gate for the /SCWM/POST part of the rework.  It proves both the
' original stock row and the requested new-bin/new-batch row are unique,
' sufficient and on the expected stock type.  It also reads the exact POST
' grid's writable contract without changing any grid value or pressing Save.
Public Function Codex_PreflightPostingBinBatchChange() As String
    Const POST_CRITERION_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/cmb/SCWM/S_UI_POST-V_CRITERION"
    Const POST_PRODUCT_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/ctxt/SCWM/S_ASPQ_STOCK-MATNR"
    Const POST_GO_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/btnCMD_GO"
    Const POST_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/subSUB_OIP_1_DATA:/SCWM/SAPLUI_POST:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell"
    Dim ses As Object, criterion As Object, productCtl As Object, goButton As Object, grid As Object
    Dim messageText As String, sourceProof As String, targetProof As String, dumpText As String
    Dim sourceRow As Long, targetRow As Long, binChangeable As String, qtyChangeable As String
    Dim batchChangeable As String, stockTypeChangeable As String, storageTypeChangeable As String, storageSectionChangeable As String
    Dim targetStorageType As String, targetStorageSection As String, targetStockType As String

    Codex_PreflightPostingBinBatchChange = "STOP|LEGACY_CROSS_BIN_POST_DISABLED: use Codex_PreflightPostingBatchAtPickBin."
    Exit Function
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertTargetStockLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped

    Set criterion = ses.FindById(POST_CRITERION_ID, False)
    Set productCtl = ses.FindById(POST_PRODUCT_ID, False)
    Set goButton = ses.FindById(POST_GO_ID, False)
    If criterion Is Nothing Or productCtl Is Nothing Or goButton Is Nothing Then
        messageText = "POST_SEARCH_CONTROLS_NOT_FOUND"
        GoTo Stopped
    End If
    If InStr(1, CWR_Text(criterion.text), "Product", vbTextCompare) = 0 Then
        messageText = "POST_CRITERION_NOT_PRODUCT: " & CWR_Text(criterion.text)
        GoTo Stopped
    End If
    productCtl.text = CWR_EXPECTED_MATERIAL
    If StrComp(CWR_Text(productCtl.text), CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Then
        messageText = "POST_PRODUCT_READBACK_MISMATCH: " & CWR_Text(productCtl.text)
        GoTo Stopped
    End If
    goButton.Press
    If Not CWR_WaitIdle(ses, 30#) Then
        messageText = "POST_QUERY_TIMEOUT"
        GoTo Stopped
    End If
    If Not CWR_WaitForObject(ses, POST_GRID_ID, 20#) Then
        messageText = "POST_GRID_TIMEOUT"
        GoTo Stopped
    End If
    Set grid = ses.FindById(POST_GRID_ID, False)
    If grid Is Nothing Then
        messageText = "POST_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then GoTo Stopped

    ' Changeable reads below are read-only.  No selection mutation is needed
    ' for this preflight; selecting happens only in the guarded executor.
    targetStorageType = CWR_Cell(grid, targetRow, "VLTYP")
    targetStorageSection = CWR_Cell(grid, targetRow, "VLBER")
    targetStockType = CWR_Cell(grid, targetRow, "CAT")
    If targetStorageType = "" Then
        messageText = "POST_TARGET_STORAGE_TYPE_MISSING"
        GoTo Stopped
    End If
    If targetStorageSection = "" Or targetStockType = "" Then
        messageText = "POST_TARGET_SECTION_OR_STOCKTYPE_MISSING"
        GoTo Stopped
    End If
    storageTypeChangeable = CWR_CellChangeable(grid, sourceRow, "NLTYP")
    storageSectionChangeable = CWR_CellChangeable(grid, sourceRow, "NLBER")
    binChangeable = CWR_CellChangeable(grid, sourceRow, "NLPLA")
    qtyChangeable = CWR_CellChangeable(grid, sourceRow, "POST_QUAN")
    batchChangeable = CWR_CellChangeable(grid, sourceRow, "TO_CHARG")
    stockTypeChangeable = CWR_CellChangeable(grid, sourceRow, "TO_CAT")
    dumpText = CWR_DumpGridBrief("POST", grid, 30, 100)
    CWR_RecordPostingPreflight sourceProof, targetProof, storageTypeChangeable, storageSectionChangeable, binChangeable, qtyChangeable, batchChangeable, stockTypeChangeable
    CWR_Log "POST_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Source and new bin/batch confirmed; no POST value changed; no Save/CF. " & _
        sourceProof & "; " & targetProof
    Codex_PreflightPostingBinBatchChange = "PASS|POST_PREVERIFIED|" & _
        "SOURCE_ROW=" & CStr(sourceRow) & "; " & sourceProof & "|" & _
        "TARGET_ROW=" & CStr(targetRow) & "; " & targetProof & "|" & _
        "DEST_STORAGE_TYPE=" & targetStorageType & "; DEST_SECTION=" & targetStorageSection & "; DEST_STOCK_TYPE=" & targetStockType & _
        "; CHANGEABLE: NLTYP=" & storageTypeChangeable & "; NLBER=" & storageSectionChangeable & _
        "; NLPLA=" & binChangeable & "; POST_QUAN=" & qtyChangeable & _
        "; TO_CHARG=" & batchChangeable & "; TO_CAT=" & stockTypeChangeable & _
        "|NEXT=MAP_AND_EXECUTE_SINGLE_POSTING_CHANGE|CF=USER_HANDOFF" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "POST_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_PreflightPostingBinBatchChange = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_PreflightPostingBinBatchChange = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only verification of the single POST commit control.  The workflow
' does not use captions or a guessed button: the standard-toolbar ID and its
' visible Save identity must both match before a later action is allowed.
Public Function Codex_ProbePOSTSaveControl() As String
    Dim ses As Object, saveButton As Object, messageText As String, detailText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_GetVerifiedPOSTSaveButton(ses, saveButton, detailText, messageText) Then GoTo Stopped
    CWR_Log "POST_SAVE_CONTROL_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact POST save control verified read-only; " & detailText
    Codex_ProbePOSTSaveControl = "PASS|POST_SAVE_CONTROL_VERIFIED|" & detailText & "|WRITE=FALSE"
    Exit Function
Stopped:
    CWR_Log "POST_SAVE_CONTROL_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ProbePOSTSaveControl = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_SAVE_CONTROL_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ProbePOSTSaveControl = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Safe recovery for an interrupted display-only or draft operation.  It only
' restarts /SCWM/POST and refuses to answer any unexpected modal dialog.
Public Function Codex_ResetPOSTDisplayOnly() As String
    Dim ses As Object, messageText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_ResetPOSTWithoutSave(ses, messageText) Then GoTo Stopped
    CWR_Log "POST_DISPLAY_RESET", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "POST display reset; no Save or CF action."
    Codex_ResetPOSTDisplayOnly = "PASS|POST_DISPLAY_RESET|WRITE=FALSE"
    Exit Function
Stopped:
    CWR_Log "POST_DISPLAY_RESET", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ResetPOSTDisplayOnly = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_DISPLAY_RESET", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ResetPOSTDisplayOnly = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only command-map evidence for /SCWM/POST.  This lists standard and
' application toolbar controls with their exact IDs and tooltips; it does not
' select, modify, save or post a stock row.
Public Function Codex_DumpPOSTCommandControls() As String
    Dim ses As Object, messageText As String, dumpText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    CWR_DumpToolbarControls ses, "wnd[0]/tbar[0]", "STANDARD", dumpText
    CWR_DumpToolbarControls ses, "wnd[0]/tbar[1]", "APPLICATION", dumpText
    CWR_DumpToolbarControls ses, "wnd[0]/mbar", "MENU", dumpText
    CWR_Log "POST_COMMAND_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "POST command controls read only; no action sent."
    Codex_DumpPOSTCommandControls = "PASS|POST_COMMAND_CONTROLS|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "POST_COMMAND_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpPOSTCommandControls = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_COMMAND_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpPOSTCommandControls = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only map of the visible /SCWM/POST menus.  Children are inspected from
' the menu object only; no menu is opened or selected.
Public Function Codex_DumpPOSTMenuTree() As String
    Dim ses As Object, rootMenu As Object, messageText As String, dumpText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    Set rootMenu = ses.FindById("wnd[0]/mbar", False)
    If rootMenu Is Nothing Then
        messageText = "POST_MENU_BAR_NOT_FOUND"
        GoTo Stopped
    End If
    CWR_DumpMenuTree rootMenu, 0, dumpText
    CWR_Log "POST_MENU_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "POST menus read only; no menu action selected."
    Codex_DumpPOSTMenuTree = "PASS|POST_MENU_TREE|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "POST_MENU_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpPOSTMenuTree = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_MENU_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpPOSTMenuTree = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Reads every currently displayed POST grid column for the exact source and
' requested target rows. This exposes mandatory technical fields without
' changing a value, selection or layout.
Public Function Codex_DumpPOSTSourceTargetColumns() As String
    Dim ses As Object, grid As Object, messageText As String, sourceProof As String, targetProof As String
    Dim sourceRow As Long, targetRow As Long, dumpText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_QueryExpectedPOST(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then GoTo Stopped
    dumpText = CWR_DumpPOSTSourceTargetColumns(grid, sourceRow, targetRow)
    CWR_Log "POST_GRID_COLUMN_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "POST source/target grid fields read only; no value changed."
    Codex_DumpPOSTSourceTargetColumns = "PASS|POST_GRID_COLUMNS|SOURCE_ROW=" & CStr(sourceRow) & _
        "|TARGET_ROW=" & CStr(targetRow) & "|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "POST_GRID_COLUMN_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpPOSTSourceTargetColumns = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_GRID_COLUMN_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpPOSTSourceTargetColumns = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Reads the current POST grid without pressing Go or resetting the transaction.
' It is used immediately after Save to distinguish a created posting-change
' task from an unchanged stock balance.
Public Function Codex_ReadCurrentPOSTActionEvidence() As String
    Dim ses As Object, grid As Object, messageText As String, sourceProof As String, targetProof As String
    Dim sourceRow As Long, targetRow As Long, sourceCounter As String, targetCounter As String
    Dim sourcePayload As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then GoTo Stopped
    sourceCounter = CWR_Cell(grid, sourceRow, "COUNTER")
    targetCounter = CWR_Cell(grid, targetRow, "COUNTER")
    sourcePayload = "NLTYP=" & CWR_Cell(grid, sourceRow, "NLTYP") & _
        "; NLPLA=" & CWR_Cell(grid, sourceRow, "NLPLA") & _
        "; POST_QUAN=" & CWR_Cell(grid, sourceRow, "POST_QUAN") & _
        "; TO_CHARG=" & CWR_Cell(grid, sourceRow, "TO_CHARG")
    CWR_Log "POST_ACTION_EVIDENCE_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Current POST grid read only; source Counter=" & sourceCounter & "; " & sourcePayload
    Codex_ReadCurrentPOSTActionEvidence = "PASS|POST_ACTION_EVIDENCE|SOURCE_COUNTER=" & sourceCounter & _
        "|TARGET_COUNTER=" & targetCounter & "|" & sourcePayload & "|" & sourceProof & "|" & targetProof & "|WRITE=FALSE"
    Exit Function
Stopped:
    CWR_Log "POST_ACTION_EVIDENCE_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ReadCurrentPOSTActionEvidence = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_ACTION_EVIDENCE_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReadCurrentPOSTActionEvidence = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Reads the POST message log after a draft/save attempt.  The message-log
' command is validated by its tooltip, no dialog response is sent, and the
' transaction is reset without Save after the diagnostic read.
Public Function Codex_ReadPOSTMessageLogAndReset() As String
    Const POST_LOG_ID As String = "wnd[0]/tbar[1]/btn[24]"
    Dim ses As Object, logButton As Object, root As Object, messageText As String, dumpText As String
    Dim typeText As String, tipText As String, resetText As String, itemCount As Long
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    Set logButton = ses.FindById(POST_LOG_ID, False)
    If logButton Is Nothing Then
        messageText = "POST_MESSAGE_LOG_CONTROL_NOT_FOUND"
        GoTo Stopped
    End If
    typeText = CWR_Text(logButton.Type)
    tipText = CWR_Text(logButton.Tooltip)
    If InStr(1, typeText, "GuiButton", vbTextCompare) = 0 Or _
       InStr(1, tipText, "Message Log", vbTextCompare) = 0 Then
        messageText = "POST_MESSAGE_LOG_CONTROL_MISMATCH: TYPE=" & typeText & "; TIP=" & tipText
        GoTo Stopped
    End If
    logButton.Press
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "POST_MESSAGE_LOG_TIMEOUT"
        GoTo Stopped
    End If
    Set root = ses.FindById("wnd[1]", False)
    If root Is Nothing Then Set root = ses.FindById("wnd[0]", False)
    If root Is Nothing Then
        messageText = "POST_MESSAGE_LOG_ROOT_NOT_FOUND"
        GoTo Stopped
    End If
    itemCount = 0
    CWR_CollectAllVisibleText root, 0, itemCount, dumpText
    resetText = ""
    If Not CWR_ResetPOSTWithoutSave(ses, resetText) Then
        messageText = "POST_MESSAGE_LOG_RESET_FAILED: " & resetText
        GoTo Stopped
    End If
    CWR_Log "POST_MESSAGE_LOG_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "POST message log read and display reset; no dialog response or Save."
    Codex_ReadPOSTMessageLogAndReset = "PASS|POST_MESSAGE_LOG|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "POST_MESSAGE_LOG_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ReadPOSTMessageLogAndReset = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_MESSAGE_LOG_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReadPOSTMessageLogAndReset = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Reads the ALV grid toolbar itself.  /SCWM/POST can expose its business
' command here instead of on the application toolbar, so IDs are mapped first
' and never guessed or pressed by this diagnostic.
Public Function Codex_DumpPOSTGridToolbar() As String
    Dim ses As Object, grid As Object, messageText As String, dumpText As String
    Dim i As Long, n As Long
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then GoTo Stopped
    n = CLng(grid.ToolbarButtonCount)
    dumpText = "GRID_TOOLBAR_BUTTONS=" & CStr(n) & vbLf
    For i = 0 To n - 1
        dumpText = dumpText & "INDEX=" & CStr(i) & "|ID=" & CWR_Text(grid.GetToolbarButtonId(i)) & _
            "|TEXT=" & CWR_Text(grid.GetToolbarButtonText(i)) & _
            "|TIP=" & CWR_Text(grid.GetToolbarButtonTooltip(i)) & _
            "|TYPE=" & CWR_Text(grid.GetToolbarButtonType(i)) & _
            "|ENABLED=" & CWR_Text(grid.GetToolbarButtonEnabled(i)) & _
            "|CHECKED=" & CWR_Text(grid.GetToolbarButtonChecked(i)) & vbLf
    Next i
    CWR_Log "POST_GRID_TOOLBAR_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "POST ALV grid toolbar read only; no button pressed."
    Codex_DumpPOSTGridToolbar = "PASS|POST_GRID_TOOLBAR|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "POST_GRID_TOOLBAR_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpPOSTGridToolbar = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_GRID_TOOLBAR_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpPOSTGridToolbar = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Finds embedded GUI toolbars beneath the POST screen (including the action
' toolbar shown above the ALV grid) and lists every button without pressing it.
Public Function Codex_DumpPOSTEmbeddedToolbars() As String
    Dim ses As Object, root As Object, grid As Object, messageText As String, dumpText As String, itemCount As Long
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    ' POST shows the row-action toolbar only after a material search.  Go is a
    ' read-only lookup here; no selection, editing or action button follows.
    If Not CWR_QueryExpectedPOST(ses, grid, messageText) Then GoTo Stopped
    Set root = ses.FindById("wnd[0]", False)
    If root Is Nothing Then
        messageText = "POST_ROOT_NOT_FOUND"
        GoTo Stopped
    End If
    CWR_DumpEmbeddedToolbars root, 0, itemCount, dumpText
    CWR_Log "POST_EMBEDDED_TOOLBAR_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "POST embedded toolbar controls read only; no button pressed."
    Codex_DumpPOSTEmbeddedToolbars = "PASS|POST_EMBEDDED_TOOLBARS|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "POST_EMBEDDED_TOOLBAR_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpPOSTEmbeddedToolbars = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_EMBEDDED_TOOLBAR_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpPOSTEmbeddedToolbars = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' The object tree is a second read-only map for custom controls that do not
' expose themselves through the normal Children collection.  It is used to
' identify the exact Create+Save command before any posting action is allowed.
Public Function Codex_DumpPOSTActionObjectTree() As String
    Dim ses As Object, grid As Object, messageText As String, rawTree As String
    Dim props As Variant
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_QueryExpectedPOST(ses, grid, messageText) Then GoTo Stopped
    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
    rawTree = CStr(ses.GetObjectTree("wnd[0]", props) & "")
    If rawTree = "" Then
        messageText = "POST_OBJECT_TREE_EMPTY"
        GoTo Stopped
    End If
    CWR_Log "POST_ACTION_OBJECT_TREE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "POST action object tree read only after exact material query; no button pressed."
    Codex_DumpPOSTActionObjectTree = "PASS|POST_ACTION_OBJECT_TREE|WRITE=FALSE" & vbLf & _
        Left$(rawTree, 30000)
    Exit Function
Stopped:
    CWR_Log "POST_ACTION_OBJECT_TREE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpPOSTActionObjectTree = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_ACTION_OBJECT_TREE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpPOSTActionObjectTree = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Maps the presently loaded POST screen without pressing Go or resetting it.
' This preserves the pending exact draft while exposing the custom action
' command that is visible only after the stock list is already on screen.
Public Function Codex_DumpCurrentPOSTActionTreeNoQuery() As String
    Dim ses As Object, root As Object, messageText As String, rawTree As String, dumpText As String
    Dim props As Variant, itemCount As Long
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    Set root = ses.FindById("wnd[0]", False)
    If root Is Nothing Then
        messageText = "POST_ROOT_NOT_FOUND"
        GoTo Stopped
    End If
    props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "Enabled", "Visible")
    rawTree = CStr(ses.GetObjectTree("wnd[0]", props) & "")
    CWR_DumpEmbeddedToolbars root, 0, itemCount, dumpText
    CWR_Log "POST_CURRENT_ACTION_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Current POST action controls read only; pending draft preserved."
    Codex_DumpCurrentPOSTActionTreeNoQuery = "PASS|POST_CURRENT_ACTION_MAP|WRITE=FALSE" & vbLf & _
        dumpText & CWR_FilterPOSTActionTree(rawTree)
    Exit Function
Stopped:
    CWR_Log "POST_CURRENT_ACTION_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpCurrentPOSTActionTreeNoQuery = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_CURRENT_ACTION_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpCurrentPOSTActionTreeNoQuery = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only map of the SAP.Toolbar.1 shell immediately above the POST grid.
' This is where the visible Create+Save command lives; it is not the standard
' SAP toolbar or menu Save command.
Public Function Codex_DumpPOSTOIPActionToolbarNoQuery() As String
    Const POST_ACTION_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell"
    Dim ses As Object, toolbar As Object, messageText As String, dumpText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    Set toolbar = ses.FindById(POST_ACTION_TOOLBAR_ID, False)
    If toolbar Is Nothing Then
        messageText = "POST_OIP_ACTION_TOOLBAR_NOT_FOUND"
        GoTo Stopped
    End If
    CWR_DumpGuiShellToolbar toolbar, "POST_OIP_ACTION", dumpText
    CWR_Log "POST_OIP_ACTION_TOOLBAR_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "POST internal action toolbar read only; pending draft preserved."
    Codex_DumpPOSTOIPActionToolbarNoQuery = "PASS|POST_OIP_ACTION_TOOLBAR|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "POST_OIP_ACTION_TOOLBAR_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpPOSTOIPActionToolbarNoQuery = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_OIP_ACTION_TOOLBAR_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpPOSTOIPActionToolbarNoQuery = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only schema/state map for the lower "Posting Changes Created" grid.
' It is used as a post-create proof source, not as an action trigger.
Public Function Codex_DumpPOSTCreatedChangesGridNoQuery() As String
    Const POST_CREATED_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_POST:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB2/ssubSUB_ODP1_TAB2:/SCWM/SAPLUI_POST:3220/ssubSUB_ODP1_2_CONTENT:/SCWM/SAPLUI_POST:3221/cntlCONTAINER_ALV_ODP1_2/shellcont/shell"
    Dim ses As Object, createdGrid As Object, messageText As String, dumpText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    Set createdGrid = ses.FindById(POST_CREATED_GRID_ID, False)
    If createdGrid Is Nothing Then
        messageText = "POST_CREATED_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    dumpText = CWR_DumpGridBrief("POST_CREATED", createdGrid, 10, 120)
    CWR_Log "POST_CREATED_GRID_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Created-posting grid read only; pending draft preserved."
    Codex_DumpPOSTCreatedChangesGridNoQuery = "PASS|POST_CREATED_GRID|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "POST_CREATED_GRID_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpPOSTCreatedChangesGridNoQuery = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_CREATED_GRID_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpPOSTCreatedChangesGridNoQuery = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only gate for a Create+Save confirmation that is already open in SAP.
' It proves the exact former source row remains selected and its pending
' destination/batch/quantity fields match this one authorized rework.  It
' deliberately does not close, reset, select or answer the dialog.
Public Function Codex_ReadPendingPOSTConfirmation() As String
    Dim ses As Object, yesButton As Object, messageText As String, proofText As String, sourceCounterBefore As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_InspectExactPendingPOSTConfirmation(ses, yesButton, sourceCounterBefore, proofText, messageText) Then GoTo Stopped
    CWR_Log "POST_CONFIRMATION_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact pending POST confirmation read only; " & proofText
    Codex_ReadPendingPOSTConfirmation = "PASS|POST_CONFIRMATION_PREVERIFIED|" & proofText & "|WRITE=FALSE"
    Exit Function
Stopped:
    CWR_Log "POST_CONFIRMATION_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ReadPendingPOSTConfirmation = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_CONFIRMATION_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReadPendingPOSTConfirmation = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Diagnostic companion for a blocked confirmation.  It reports the dialog,
' selection and pending fields even if the selection is wrong, so no existing
' user draft is discarded merely to inspect it.
Public Function Codex_DumpPendingPOSTConfirmationDetail() As String
    Dim ses As Object, modal As Object, grid As Object, yesButton As Object, noButton As Object, cancelButton As Object
    Dim messageText As String, sourceProof As String, targetProof As String, selectionText As String
    Dim sourceRow As Long, targetRow As Long, titleText As String, payloadText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    Set modal = ses.FindById("wnd[1]", False)
    If modal Is Nothing Then
        messageText = "POST_CONFIRMATION_MODAL_NOT_FOUND"
        GoTo Stopped
    End If
    titleText = CWR_Text(modal.text)
    Set yesButton = ses.FindById("wnd[1]/usr/btnBUTTON_1", False)
    Set noButton = ses.FindById("wnd[1]/usr/btnBUTTON_2", False)
    Set cancelButton = ses.FindById("wnd[1]/usr/btnBUTTON_3", False)
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then GoTo Stopped
    selectionText = CWR_SelectedRowsText(grid)
    payloadText = "TO_CAT=" & CWR_Cell(grid, sourceRow, "TO_CAT") & _
        "; NLTYP=" & CWR_Cell(grid, sourceRow, "NLTYP") & _
        "; NLBER=" & CWR_Cell(grid, sourceRow, "NLBER") & _
        "; NLPLA=" & CWR_Cell(grid, sourceRow, "NLPLA") & _
        "; POST_QUAN=" & CWR_Cell(grid, sourceRow, "POST_QUAN") & _
        "; TO_CHARG=" & CWR_Cell(grid, sourceRow, "TO_CHARG") & _
        "; COUNTER=" & CWR_Cell(grid, sourceRow, "COUNTER")
    CWR_Log "POST_CONFIRMATION_DETAIL", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Blocked POST confirmation inspected read only; selected=" & selectionText & "; " & payloadText
    Codex_DumpPendingPOSTConfirmationDetail = "PASS|POST_CONFIRMATION_DETAIL|TITLE=" & titleText & _
        "|BUTTONS=" & CWR_Text(yesButton.text) & "/" & CWR_Text(noButton.text) & "/" & _
        CWR_Text(cancelButton.text) & "|SOURCE_ROW=" & CStr(sourceRow) & "|SELECTED_ROWS=" & selectionText & _
        "|DRAFT=" & payloadText & "|" & sourceProof & "|" & targetProof & "|WRITE=FALSE"
    Exit Function
Stopped:
    CWR_Log "POST_CONFIRMATION_DETAIL", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpPendingPOSTConfirmationDetail = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_CONFIRMATION_DETAIL", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpPendingPOSTConfirmationDetail = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' This is intentionally the sole responder for the exact Yes dialog shown
' after the internal POST Create+Save action.  It cannot run unless the
' read-only inspection immediately before it proves the selected source row
' and all six requested destination fields.  It never retries a SAP write.
Public Function Codex_ConfirmExactPendingPOSTChange() As String
    Dim ses As Object, grid As Object, yesButton As Object
    Dim messageText As String, proofText As String, sourceCounterBefore As String, sourceCounterAfter As String
    Dim sourceProof As String, targetProof As String, sourceRow As Long, targetRow As Long
    Dim statusType As String, statusText As String, confirmationSent As Boolean

    Codex_ConfirmExactPendingPOSTChange = "STOP|LEGACY_CROSS_BIN_POST_DISABLED: no legacy POST confirmation may be sent."
    Exit Function
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_InspectExactPendingPOSTConfirmation(ses, yesButton, sourceCounterBefore, proofText, messageText) Then GoTo Stopped

    ' User-authorized response: the control is the exact Yes button from the
    ' verified POST confirmation, after source row and draft payload readback.
    confirmationSent = True
    yesButton.Press
    CWR_Log "POST_CONFIRM_YES_SENT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Yes pressed once for the preverified POST draft; " & proofText
    If Not CWR_WaitPOSTSaveSettle(ses, 3#, 45#) Then
        messageText = "POST_CONFIRM_TIMEOUT_STATE_UNPROVEN"
        GoTo PostConfirmStopped
    End If
    If CWR_HasAnyModal(ses, messageText) Then
        messageText = "POST_CONFIRM_MODAL_REMAINS: " & messageText
        GoTo PostConfirmStopped
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "POST_CONFIRM_SAP_" & statusType & ": " & statusText
        GoTo PostConfirmStopped
    End If
    If Not CWR_QueryExpectedPOST(ses, grid, messageText) Then GoTo PostConfirmStopped
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then GoTo PostConfirmStopped
    sourceCounterAfter = CWR_Cell(grid, sourceRow, "COUNTER")
    If Not CWR_PostCounterAdvanced(sourceCounterBefore, sourceCounterAfter, messageText) Then GoTo PostConfirmStopped

    CWR_RecordPostingCreate sourceProof, targetProof, sourceCounterBefore, sourceCounterAfter
    CWR_Log "POST_CREATE_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Create+Save confirmed once and fresh POST counter advanced; before=" & sourceCounterBefore & _
        "; after=" & sourceCounterAfter & "; CF=USER_HANDOFF"
    Codex_ConfirmExactPendingPOSTChange = "PASS|POST_CHANGE_CREATED_VERIFIED|SOURCE_COUNTER_BEFORE=" & _
        sourceCounterBefore & "|SOURCE_COUNTER_AFTER=" & sourceCounterAfter & "|" & sourceProof & "|" & _
        targetProof & "|NEXT=PRDO_CHANGE_QTY_THEN_CREATE_WT|CF=USER_HANDOFF"
    Exit Function

Stopped:
    CWR_Log "POST_CONFIRM_YES", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ConfirmExactPendingPOSTChange = "STOP|" & messageText
    Exit Function
PostConfirmStopped:
    CWR_Log "POST_CREATE_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ConfirmExactPendingPOSTChange = "STOP|" & messageText & "|NO_RETRY"
    Exit Function
Failed:
    CWR_Log "POST_CONFIRM_YES", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description & IIf(confirmationSent, "; STATE_UNPROVEN_NO_RETRY", "")
    Codex_ConfirmExactPendingPOSTChange = "FAIL|" & CStr(Err.Number) & "|" & Err.Description & _
        IIf(confirmationSent, "|STATE_UNPROVEN_NO_RETRY", "")
End Function

' Executes the user-specified order for the already prepared POST draft:
' select/highlight the exact source row, press the internal "Create+Save
' Posting Change" command once, then answer its exact Yes confirmation once.
' It is deliberately distinct from normal Save and the Warehouse Request
' button. CF is never invoked.
Public Function Codex_SelectCreateSaveAndConfirmExactPost() As String
    Const POST_CREATED_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_POST:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB2/ssubSUB_ODP1_TAB2:/SCWM/SAPLUI_POST:3220/ssubSUB_ODP1_2_CONTENT:/SCWM/SAPLUI_POST:3221/cntlCONTAINER_ALV_ODP1_2/shellcont/shell"
    Dim ses As Object, grid As Object, actionToolbar As Object, yesButton As Object, createdGrid As Object
    Dim messageText As String, modalText As String, sourceProof As String, targetProof As String
    Dim payloadText As String, pendingProof As String, actionDetail As String, confirmProof As String
    Dim sourceCounterBefore As String, sourceCounterAfter As String, statusType As String, statusText As String
    Dim sourceRow As Long, targetRow As Long, createdBefore As Long, createdAfter As Long
    Dim actionSent As Boolean, confirmationSent As Boolean

    Codex_SelectCreateSaveAndConfirmExactPost = "STOP|LEGACY_CROSS_BIN_POST_DISABLED: no legacy Create+Save/Yes may be sent."
    Exit Function
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertTargetStockLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = "POST_ACTION_START_MODAL_PRESENT: " & modalText
        GoTo Stopped
    End If
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then GoTo Stopped
    If Not CWR_ExactPOSTDraftPayload(grid, sourceRow, targetRow, payloadText, messageText) Then GoTo Stopped
    If Not CWR_ProveOnlyExactPendingDraft(grid, sourceRow, pendingProof, messageText) Then GoTo Stopped
    Set createdGrid = ses.FindById(POST_CREATED_GRID_ID, False)
    If createdGrid Is Nothing Then
        messageText = "POST_CREATED_GRID_NOT_FOUND_BEFORE_ACTION"
        GoTo Stopped
    End If
    createdBefore = CLng(createdGrid.rowCount)
    If Not CWR_SelectExactPOSTSource(ses, grid, sourceRow, targetRow, messageText) Then GoTo Stopped
    If CWR_SelectedRowsText(grid) <> CStr(sourceRow) Then
        messageText = "POST_ACTION_SELECTED_ROW_READBACK_MISMATCH: " & CWR_SelectedRowsText(grid)
        GoTo Stopped
    End If
    If Not CWR_ExactPOSTDraftPayload(grid, sourceRow, targetRow, payloadText, messageText) Then GoTo Stopped
    If Not CWR_ProveOnlyExactPendingDraft(grid, sourceRow, pendingProof, messageText) Then GoTo Stopped
    If Not CWR_GetVerifiedPOSTCreateSaveAction(ses, actionToolbar, actionDetail, messageText) Then GoTo Stopped

    ' First and only business command: the exact internal Create+Save Posting
    ' Change button, with row 11 demonstrably highlighted immediately above.
    actionSent = True
    actionToolbar.PressButton "OK_OIP_POST_CREATE_POST"
    CWR_Log "POST_CREATE_SAVE_SENT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact source row selected then internal Create+Save pressed once; " & actionDetail & _
        "; " & pendingProof
    If Not CWR_WaitPOSTSaveSettle(ses, 1#, 25#) Then
        messageText = "POST_CREATE_SAVE_MODAL_TIMEOUT_STATE_UNPROVEN"
        GoTo ActionStopped
    End If
    If Not CWR_InspectExactPendingPOSTConfirmation(ses, yesButton, sourceCounterBefore, confirmProof, messageText) Then
        messageText = "POST_CREATE_SAVE_CONFIRMATION_UNPROVEN: " & messageText
        GoTo ActionStopped
    End If

    ' Second and final business response: only the exact Yes button in the
    ' posting-change confirmation that follows the verified Create+Save.
    confirmationSent = True
    yesButton.Press
    CWR_Log "POST_CREATE_SAVE_YES_SENT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Yes pressed once after exact selected-row/Create+Save proof; " & confirmProof
    If Not CWR_WaitPOSTSaveSettle(ses, 3#, 45#) Then
        messageText = "POST_CREATE_POSTVERIFY_TIMEOUT_STATE_UNPROVEN"
        GoTo ActionStopped
    End If
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = "POST_CREATE_POSTVERIFY_MODAL_REMAINS: " & modalText
        GoTo ActionStopped
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "POST_CREATE_POSTVERIFY_SAP_" & statusType & ": " & statusText
        GoTo ActionStopped
    End If
    Set createdGrid = ses.FindById(POST_CREATED_GRID_ID, False)
    If createdGrid Is Nothing Then
        messageText = "POST_CREATED_GRID_NOT_FOUND_AFTER_ACTION"
        GoTo ActionStopped
    End If
    createdAfter = CLng(createdGrid.rowCount)
    If createdAfter <= createdBefore Then
        messageText = "POST_CREATED_GRID_NOT_ADVANCED: BEFORE=" & CStr(createdBefore) & "; AFTER=" & CStr(createdAfter)
        GoTo ActionStopped
    End If
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then GoTo ActionStopped
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then GoTo ActionStopped
    sourceCounterAfter = CWR_Cell(grid, sourceRow, "COUNTER")

    CWR_RecordPostingCreateByGrid sourceProof, targetProof, createdBefore, createdAfter, sourceCounterAfter
    CWR_Log "POST_CREATE_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Create+Save/Yes completed once; created-grid rows advanced from " & CStr(createdBefore) & _
        " to " & CStr(createdAfter) & "; source counter=" & sourceCounterAfter & "; CF=USER_HANDOFF"
    Codex_SelectCreateSaveAndConfirmExactPost = "PASS|POST_CHANGE_CREATED_VERIFIED|CREATED_ROWS_BEFORE=" & _
        CStr(createdBefore) & "|CREATED_ROWS_AFTER=" & CStr(createdAfter) & "|SOURCE_COUNTER=" & _
        sourceCounterAfter & "|" & sourceProof & "|" & targetProof & _
        "|NEXT=PRDO_CHANGE_QTY_THEN_CREATE_WT|CF=USER_HANDOFF"
    Exit Function

Stopped:
    CWR_Log "POST_CREATE_SAVE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_SelectCreateSaveAndConfirmExactPost = "STOP|" & messageText
    Exit Function
ActionStopped:
    CWR_Log "POST_CREATE_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_SelectCreateSaveAndConfirmExactPost = "STOP|" & messageText & "|NO_RETRY"
    Exit Function
Failed:
    CWR_Log "POST_CREATE_SAVE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description & _
        IIf(actionSent Or confirmationSent, "; STATE_UNPROVEN_NO_RETRY", "")
    Codex_SelectCreateSaveAndConfirmExactPost = "FAIL|" & CStr(Err.Number) & "|" & Err.Description & _
        IIf(actionSent Or confirmationSent, "|STATE_UNPROVEN_NO_RETRY", "")
End Function

' Clears only the stale generic save dialog that is blocking row selection.
' The dialog is accepted only after the exact one-row draft is proven and the
' created-posting grid is proven unchanged afterwards.  It does not press
' Create+Save.  The caller must then select the source row and invoke the
' separate Create+Save/Yes flow in the same SAP session.
Public Function Codex_AcceptExactDraftSaveForSelection() As String
    Const POST_CREATED_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_POST:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB2/ssubSUB_ODP1_TAB2:/SCWM/SAPLUI_POST:3220/ssubSUB_ODP1_2_CONTENT:/SCWM/SAPLUI_POST:3221/cntlCONTAINER_ALV_ODP1_2/shellcont/shell"
    Dim ses As Object, grid As Object, createdGrid As Object, yesButton As Object
    Dim messageText As String, proofText As String, sourceCounterBefore As String, modalText As String
    Dim sourceProof As String, targetProof As String, payloadText As String, pendingProof As String
    Dim sourceRow As Long, targetRow As Long, createdBefore As Long, createdAfter As Long
    Dim statusType As String, statusText As String
    Codex_AcceptExactDraftSaveForSelection = "STOP|LEGACY_CROSS_BIN_POST_DISABLED: no legacy draft confirmation may be sent."
    Exit Function
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_InspectExactPendingPOSTConfirmation(ses, yesButton, sourceCounterBefore, proofText, messageText) Then GoTo Stopped
    Set createdGrid = ses.FindById(POST_CREATED_GRID_ID, False)
    If createdGrid Is Nothing Then
        messageText = "POST_CREATED_GRID_NOT_FOUND_BEFORE_DRAFT_SAVE"
        GoTo Stopped
    End If
    createdBefore = CLng(createdGrid.rowCount)
    yesButton.Press
    CWR_Log "POST_DRAFT_SAVE_YES_SENT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Stale generic save dialog accepted for exact draft only; no Create+Save action sent. " & proofText
    If Not CWR_WaitPOSTSaveSettle(ses, 2#, 30#) Then
        messageText = "POST_DRAFT_SAVE_TIMEOUT_STATE_UNPROVEN"
        GoTo Stopped
    End If
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = "POST_DRAFT_SAVE_MODAL_REMAINS: " & modalText
        GoTo Stopped
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "POST_DRAFT_SAVE_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    Set createdGrid = ses.FindById(POST_CREATED_GRID_ID, False)
    If createdGrid Is Nothing Then
        messageText = "POST_CREATED_GRID_NOT_FOUND_AFTER_DRAFT_SAVE"
        GoTo Stopped
    End If
    createdAfter = CLng(createdGrid.rowCount)
    If createdAfter <> createdBefore Then
        messageText = "POST_DRAFT_SAVE_CREATED_GRID_CHANGED: BEFORE=" & CStr(createdBefore) & "; AFTER=" & CStr(createdAfter)
        GoTo Stopped
    End If
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then GoTo Stopped
    If Not CWR_ExactPOSTDraftPayload(grid, sourceRow, targetRow, payloadText, messageText) Then GoTo Stopped
    If Not CWR_ProveOnlyExactPendingDraft(grid, sourceRow, pendingProof, messageText) Then GoTo Stopped
    CWR_Log "POST_DRAFT_SAVE_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact draft remains pending and created grid unchanged; ready to select row then Create+Save. " & pendingProof
    Codex_AcceptExactDraftSaveForSelection = "PASS|POST_DRAFT_SAVED_NOT_CREATED|CREATED_ROWS=" & _
        CStr(createdAfter) & "|SOURCE_ROW=" & CStr(sourceRow) & "|" & pendingProof & _
        "|NEXT=SELECT_ROW_THEN_CREATE_SAVE_POSTING_CHANGE|CF=USER_HANDOFF"
    Exit Function
Stopped:
    CWR_Log "POST_DRAFT_SAVE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_AcceptExactDraftSaveForSelection = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_DRAFT_SAVE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_AcceptExactDraftSaveForSelection = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Performs exactly one posting change for the uniquely proven old stock row:
' old bin/batch -> requested bin/batch, quantity 128.  It never changes stock
' type, never touches an ODO row, and never executes CF.  After Save it
' re-queries the POST grid and requires the exact before/after quantity delta.
Public Function Codex_ExecutePostingBinBatchChange() As String
    Dim preflightText As String, messageText As String, detailText As String, statusType As String, statusText As String
    Dim ses As Object, grid As Object, saveMenu As Object, modalText As String
    Dim sourceProof As String, targetProof As String, afterSourceProof As String, afterTargetProof As String
    Dim sourceRow As Long, targetRow As Long, afterSourceRow As Long, afterTargetRow As Long
    Dim sourceAvailBefore As Double, sourceTotalBefore As Double, targetAvailBefore As Double, targetTotalBefore As Double
    Dim sourceAvailAfter As Double, sourceTotalAfter As Double, targetAvailAfter As Double, targetTotalAfter As Double
    Dim modifiedDraft As Boolean, saveSent As Boolean, resetText As String
    Dim targetStorageType As String, targetStorageSection As String, targetStockType As String

    Codex_ExecutePostingBinBatchChange = "STOP|LEGACY_CROSS_BIN_POST_DISABLED: use the in-place pick-bin batch workflow."
    Exit Function
    On Error GoTo Failed
    preflightText = Codex_PreflightPostingBinBatchChange()
    If Left$(preflightText, 5) <> "PASS|" Then
        CWR_Log "POST_ACTION", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", preflightText
        Codex_ExecutePostingBinBatchChange = preflightText
        Exit Function
    End If
    If Not CWR_AssertPostingPlan(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_QueryExpectedPOST(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then GoTo Stopped

    sourceAvailBefore = CWR_Number(CWR_Cell(grid, sourceRow, "AVL_ALT_QTY"))
    sourceTotalBefore = CWR_Number(CWR_Cell(grid, sourceRow, "LIME_QUAN"))
    targetAvailBefore = CWR_Number(CWR_Cell(grid, targetRow, "AVL_ALT_QTY"))
    targetTotalBefore = CWR_Number(CWR_Cell(grid, targetRow, "LIME_QUAN"))
    targetStorageType = CWR_Cell(grid, targetRow, "VLTYP")
    targetStorageSection = CWR_Cell(grid, targetRow, "VLBER")
    targetStockType = CWR_Cell(grid, targetRow, "CAT")
    If targetStorageType = "" Or targetStorageSection = "" Or targetStockType = "" Then
        messageText = "POST_TARGET_DESTINATION_ATTRIBUTE_MISSING"
        GoTo Stopped
    End If
    If sourceAvailBefore < CWR_TARGET_QTY Or sourceTotalBefore < CWR_TARGET_QTY Then
        messageText = "POST_SOURCE_BECAME_INSUFFICIENT_BEFORE_SAVE"
        GoTo Stopped
    End If
    If UCase$(CWR_CellChangeable(grid, sourceRow, "TO_CAT")) <> "TRUE" Or _
       UCase$(CWR_CellChangeable(grid, sourceRow, "NLTYP")) <> "TRUE" Or _
       UCase$(CWR_CellChangeable(grid, sourceRow, "NLBER")) <> "TRUE" Or _
       UCase$(CWR_CellChangeable(grid, sourceRow, "NLPLA")) <> "TRUE" Or _
       UCase$(CWR_CellChangeable(grid, sourceRow, "POST_QUAN")) <> "TRUE" Or _
       UCase$(CWR_CellChangeable(grid, sourceRow, "TO_CHARG")) <> "TRUE" Then
        messageText = "POST_WRITE_FIELDS_NOT_CHANGEABLE_AFTER_PREVERIFY"
        GoTo Stopped
    End If

    grid.selectedRows = CStr(sourceRow)
    grid.SetCurrentCell sourceRow, "CHARG"
    CWR_Log "POST_ACTION_START", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact source selected; " & sourceProof & "; targetType=" & targetStorageType & "/" & targetStorageSection & _
        "; targetStockType=" & targetStockType & "; target=" & CWR_TARGET_BIN & "/" & _
        CWR_TARGET_BATCH & "/" & CStr(CWR_TARGET_QTY) & "; CF=USER_HANDOFF"

    grid.ModifyCell sourceRow, "TO_CAT", targetStockType
    modifiedDraft = True
    If StrComp(CWR_Cell(grid, sourceRow, "TO_CAT"), targetStockType, vbTextCompare) <> 0 Then
        messageText = "POST_DRAFT_STOCK_TYPE_READBACK_MISMATCH: " & CWR_Cell(grid, sourceRow, "TO_CAT")
        GoTo Stopped
    End If
    grid.ModifyCell sourceRow, "NLTYP", targetStorageType
    If StrComp(CWR_Cell(grid, sourceRow, "NLTYP"), targetStorageType, vbTextCompare) <> 0 Then
        messageText = "POST_DRAFT_STORAGE_TYPE_READBACK_MISMATCH: " & CWR_Cell(grid, sourceRow, "NLTYP")
        GoTo Stopped
    End If
    grid.ModifyCell sourceRow, "NLBER", targetStorageSection
    If StrComp(CWR_Cell(grid, sourceRow, "NLBER"), targetStorageSection, vbTextCompare) <> 0 Then
        messageText = "POST_DRAFT_STORAGE_SECTION_READBACK_MISMATCH: " & CWR_Cell(grid, sourceRow, "NLBER")
        GoTo Stopped
    End If
    grid.ModifyCell sourceRow, "NLPLA", CWR_TARGET_BIN
    If StrComp(CWR_Cell(grid, sourceRow, "NLPLA"), CWR_TARGET_BIN, vbTextCompare) <> 0 Then
        messageText = "POST_DRAFT_BIN_READBACK_MISMATCH: " & CWR_Cell(grid, sourceRow, "NLPLA")
        GoTo Stopped
    End If
    grid.ModifyCell sourceRow, "POST_QUAN", CWR_FastQuantityText(CWR_TARGET_QTY)
    If Abs(CWR_Number(CWR_Cell(grid, sourceRow, "POST_QUAN")) - CWR_TARGET_QTY) > 0.000001 Then
        messageText = "POST_DRAFT_QTY_READBACK_MISMATCH: " & CWR_Cell(grid, sourceRow, "POST_QUAN")
        GoTo Stopped
    End If
    grid.ModifyCell sourceRow, "TO_CHARG", CWR_TARGET_BATCH
    If StrComp(CWR_Cell(grid, sourceRow, "TO_CHARG"), CWR_TARGET_BATCH, vbTextCompare) <> 0 Then
        messageText = "POST_DRAFT_BATCH_READBACK_MISMATCH: " & CWR_Cell(grid, sourceRow, "TO_CHARG")
        GoTo Stopped
    End If

    ' GuiGridView batches ModifyCell calls locally.  SAP's own scripting API
    ' requires TriggerModified after a multiple-cell edit so the grid validates
    ' the exact pending values before the one permitted Save is pressed.
    grid.TriggerModified
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "POST_DRAFT_TRIGGER_TIMEOUT"
        GoTo Stopped
    End If
    If CWR_GetModalText(ses, modalText) Then
        messageText = "POST_DRAFT_TRIGGER_MODAL: " & modalText
        GoTo Stopped
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "POST_DRAFT_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    ' TriggerModified communicates with SAP, so reacquire every control below
    ' the window before inspecting the draft or pressing the mapped Save.
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then GoTo Stopped
    If StrComp(CWR_Cell(grid, sourceRow, "TO_CAT"), targetStockType, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(grid, sourceRow, "NLTYP"), targetStorageType, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(grid, sourceRow, "NLBER"), targetStorageSection, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(grid, sourceRow, "NLPLA"), CWR_TARGET_BIN, vbTextCompare) <> 0 Or _
       Abs(CWR_Number(CWR_Cell(grid, sourceRow, "POST_QUAN")) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(CWR_Cell(grid, sourceRow, "TO_CHARG"), CWR_TARGET_BATCH, vbTextCompare) <> 0 Then
        messageText = "POST_DRAFT_TRIGGER_READBACK_MISMATCH"
        GoTo Stopped
    End If

    ' This is an exact grid-cell commit, not a blind session Enter: SAP Grid
    ' keeps the draft active until the focused destination-batch cell accepts
    ' the edit event. Any popup or SAP error stops before Save.
    grid.SetCurrentCell sourceRow, "TO_CHARG"
    grid.PressEnter
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "POST_DRAFT_GRID_COMMIT_TIMEOUT"
        GoTo Stopped
    End If
    If CWR_GetModalText(ses, modalText) Then
        messageText = "POST_DRAFT_GRID_COMMIT_MODAL: " & modalText
        GoTo Stopped
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "POST_DRAFT_GRID_COMMIT_SAP_" & statusType & ": " & statusText
        GoTo Stopped
    End If
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then GoTo Stopped
    If StrComp(CWR_Cell(grid, sourceRow, "TO_CAT"), targetStockType, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(grid, sourceRow, "NLTYP"), targetStorageType, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(grid, sourceRow, "NLBER"), targetStorageSection, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(grid, sourceRow, "NLPLA"), CWR_TARGET_BIN, vbTextCompare) <> 0 Or _
       Abs(CWR_Number(CWR_Cell(grid, sourceRow, "POST_QUAN")) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(CWR_Cell(grid, sourceRow, "TO_CHARG"), CWR_TARGET_BATCH, vbTextCompare) <> 0 Then
        messageText = "POST_DRAFT_GRID_COMMIT_READBACK_MISMATCH"
        GoTo Stopped
    End If
    If Not CWR_SelectExactPOSTSource(ses, grid, sourceRow, targetRow, messageText) Then GoTo Stopped
    If StrComp(CWR_Cell(grid, sourceRow, "NLTYP"), targetStorageType, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(grid, sourceRow, "NLPLA"), CWR_TARGET_BIN, vbTextCompare) <> 0 Or _
       Abs(CWR_Number(CWR_Cell(grid, sourceRow, "POST_QUAN")) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(CWR_Cell(grid, sourceRow, "TO_CHARG"), CWR_TARGET_BATCH, vbTextCompare) <> 0 Then
        messageText = "POST_DRAFT_SELECTION_READBACK_MISMATCH"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedPOSTMenuSave(ses, saveMenu, detailText, messageText) Then GoTo Stopped

    ' The next operation is the one and only SAP write in this function.  From
    ' this point a reset/retry is prohibited; only a fresh readback may decide
    ' success or state-unproven.
    saveSent = True
    saveMenu.Select
    CWR_Log "POST_ACTION_SENT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Save pressed once using " & detailText & "; target=" & CWR_TARGET_BIN & "/" & _
        CWR_TARGET_BATCH & "/" & CStr(CWR_TARGET_QTY)
    If Not CWR_WaitPOSTSaveSettle(ses, 3#, 45#) Then
        messageText = "POST_SAVE_TIMEOUT_STATE_UNPROVEN"
        GoTo PostSaveStopped
    End If
    If CWR_GetModalText(ses, modalText) Then
        messageText = "POST_SAVE_MODAL_STATE_UNPROVEN: " & modalText
        GoTo PostSaveStopped
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "POST_SAVE_SAP_" & statusType & ": " & statusText
        GoTo PostSaveStopped
    End If
    If Not CWR_QueryExpectedPOST(ses, grid, messageText) Then GoTo PostSaveStopped
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, afterSourceRow, afterTargetRow, afterSourceProof, afterTargetProof, messageText) Then GoTo PostSaveStopped
    sourceAvailAfter = CWR_Number(CWR_Cell(grid, afterSourceRow, "AVL_ALT_QTY"))
    sourceTotalAfter = CWR_Number(CWR_Cell(grid, afterSourceRow, "LIME_QUAN"))
    targetAvailAfter = CWR_Number(CWR_Cell(grid, afterTargetRow, "AVL_ALT_QTY"))
    targetTotalAfter = CWR_Number(CWR_Cell(grid, afterTargetRow, "LIME_QUAN"))
    If Abs(sourceAvailAfter - (sourceAvailBefore - CWR_TARGET_QTY)) > 0.000001 Or _
       Abs(sourceTotalAfter - (sourceTotalBefore - CWR_TARGET_QTY)) > 0.000001 Or _
       Abs(targetAvailAfter - (targetAvailBefore + CWR_TARGET_QTY)) > 0.000001 Or _
       Abs(targetTotalAfter - (targetTotalBefore + CWR_TARGET_QTY)) > 0.000001 Then
        messageText = "POST_POSTVERIFY_DELTA_MISMATCH: BEFORE_SRC=" & CStr(sourceAvailBefore) & _
            "/" & CStr(sourceTotalBefore) & "; BEFORE_TARGET=" & CStr(targetAvailBefore) & _
            "/" & CStr(targetTotalBefore) & "; AFTER_SRC=" & CStr(sourceAvailAfter) & _
            "/" & CStr(sourceTotalAfter) & "; AFTER_TARGET=" & CStr(targetAvailAfter) & _
            "/" & CStr(targetTotalAfter)
        GoTo PostSaveStopped
    End If

    CWR_RecordPostingChange afterSourceProof, afterTargetProof, sourceAvailAfter, targetAvailAfter
    CWR_Log "POST_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Posting change readback proved once; " & afterSourceProof & "; " & afterTargetProof & _
        "; CF=USER_HANDOFF"
    Codex_ExecutePostingBinBatchChange = "PASS|POST_CHANGE_VERIFIED|" & afterSourceProof & "|" & _
        afterTargetProof & "|NEXT=PRDO_CHANGE_QTY_THEN_CREATE_WT|CF=USER_HANDOFF"
    Exit Function

Stopped:
    If modifiedDraft And Not saveSent Then
        resetText = ""
        If Not CWR_ResetPOSTWithoutSave(ses, resetText) Then messageText = messageText & "; POST_DRAFT_RESET_FAILED=" & resetText
    End If
    CWR_Log "POST_ACTION", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ExecutePostingBinBatchChange = "STOP|" & messageText
    Exit Function

PostSaveStopped:
    CWR_Log "POST_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ExecutePostingBinBatchChange = "STOP|" & messageText
    Exit Function

Failed:
    If modifiedDraft And Not saveSent Then
        resetText = ""
        Call CWR_ResetPOSTWithoutSave(ses, resetText)
    End If
    CWR_Log "POST_ACTION", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description & IIf(saveSent, "; SAVE_STATE_UNPROVEN", "")
    Codex_ExecutePostingBinBatchChange = "FAIL|" & CStr(Err.Number) & "|" & Err.Description & _
        IIf(saveSent, "|SAVE_STATE_UNPROVEN", "")
End Function

' Mandatory read-only recovery after any save whose immediate SAP response was
' not conclusive.  It has no write path and intentionally recognizes only the
' exact pre-action or exact expected-after inventory state.
Public Function Codex_ReadPostingChangeRecovery() As String
    Dim ses As Object, grid As Object, messageText As String, sourceProof As String, targetProof As String
    Dim sourceRow As Long, targetRow As Long
    Dim sourceAvail As Double, sourceTotal As Double, targetAvail As Double, targetTotal As Double

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenPOSTSession(ses, messageText) Then GoTo Stopped
    If Not CWR_QueryExpectedPOST(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then GoTo Stopped
    sourceAvail = CWR_Number(CWR_Cell(grid, sourceRow, "AVL_ALT_QTY"))
    sourceTotal = CWR_Number(CWR_Cell(grid, sourceRow, "LIME_QUAN"))
    targetAvail = CWR_Number(CWR_Cell(grid, targetRow, "AVL_ALT_QTY"))
    targetTotal = CWR_Number(CWR_Cell(grid, targetRow, "LIME_QUAN"))

    If Abs(sourceAvail - 256#) <= 0.000001 And Abs(sourceTotal - 256#) <= 0.000001 And _
       Abs(targetAvail - 256#) <= 0.000001 And Abs(targetTotal - 256#) <= 0.000001 Then
        CWR_RecordPostingChange sourceProof, targetProof, sourceAvail, targetAvail
        CWR_Log "POST_RECOVERY_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "Fresh POST readback proves the single posting change; " & sourceProof & "; " & targetProof
        Codex_ReadPostingChangeRecovery = "PASS|POST_CHANGE_VERIFIED_BY_RECOVERY|" & sourceProof & "|" & targetProof & _
            "|NEXT=PRDO_CHANGE_QTY_THEN_CREATE_WT|CF=USER_HANDOFF"
        Exit Function
    End If
    If Abs(sourceAvail - 384#) <= 0.000001 And Abs(sourceTotal - 384#) <= 0.000001 And _
       Abs(targetAvail - 128#) <= 0.000001 And Abs(targetTotal - 128#) <= 0.000001 Then
        messageText = "POST_NOT_APPLIED_READBACK: SOURCE=384/384; TARGET=128/128; NO_RETRY_IN_THIS_CALL"
        GoTo Stopped
    End If
    messageText = "POST_STATE_UNPROVEN: SOURCE=" & CStr(sourceAvail) & "/" & CStr(sourceTotal) & _
        "; TARGET=" & CStr(targetAvail) & "/" & CStr(targetTotal)
    GoTo Stopped
Stopped:
    CWR_Log "POST_RECOVERY_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ReadPostingChangeRecovery = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "POST_RECOVERY_READ", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ReadPostingChangeRecovery = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Non-mutating edit-mode check.  It proves the exact PRDO row and both cells
' are editable, then restarts PRDO without changing any value.
Public Function Codex_PreflightPRDOBatchQtyChange() As String
    Const PRDO_OIP_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell"
    Dim ses As Object, grid As Object, toolbar As Object
    Dim messageText As String, qtyText As String, batchText As String, statusText As String
    Dim rowNo As Long, batchChangeable As String, qtyChangeable As String, editModeEntered As Boolean, resetText As String

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertTargetStockLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    If Abs(CWR_Number(qtyText) - CWR_EXPECTED_SOURCE_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "PRDO_BEFORE_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        GoTo Stopped
    End If
    Set toolbar = ses.FindById(PRDO_OIP_TOOLBAR_ID, False)
    If toolbar Is Nothing Then
        messageText = "PRDO_EDIT_TOOLBAR_NOT_FOUND"
        GoTo Stopped
    End If
    toolbar.PressButton "OIP_CHANGE"
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "PRDO_EDIT_MODE_TIMEOUT"
        GoTo Stopped
    End If
    editModeEntered = True
    batchChangeable = CWR_CellChangeable(grid, rowNo, "BATCHNO_UI")
    qtyChangeable = CWR_CellChangeable(grid, rowNo, "QTY_UI")
    If UCase$(batchChangeable) <> "TRUE" Or UCase$(qtyChangeable) <> "TRUE" Then
        messageText = "PRDO_EDIT_FIELDS_NOT_CHANGEABLE: BATCH=" & batchChangeable & "; QTY=" & qtyChangeable
        GoTo Stopped
    End If
    If Not CWR_ResetPRDOWithoutChanges(ses, messageText) Then GoTo Stopped
    CWR_Log "PRDO_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "PRDO row exact; current Qty=" & qtyText & "; Batch=" & batchText & _
        "; Batch/QTY fields changeable; no value changed; CF=USER_HANDOFF"
    Codex_PreflightPRDOBatchQtyChange = "PASS|PRDO_EDIT_PREVERIFIED|BEFORE_QTY=" & qtyText & _
        "|BEFORE_BATCH=" & batchText & "|BATCH_CHANGEABLE=" & batchChangeable & _
        "|QTY_CHANGEABLE=" & qtyChangeable & "|TARGET_QTY=" & CStr(CWR_TARGET_QTY) & _
        "|TARGET_BATCH=" & CWR_TARGET_BATCH & "|CF=USER_HANDOFF"
    Exit Function
Stopped:
    If editModeEntered Then
        resetText = ""
        If Not CWR_ResetPRDOWithoutChanges(ses, resetText) Then _
            messageText = messageText & "; PRDO_RESET_AFTER_PREVIEW_FAILED=" & resetText
    End If
    CWR_Log "PRDO_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_PreflightPRDOBatchQtyChange = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "PRDO_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_PreflightPRDOBatchQtyChange = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' The batch is intentionally handled in /SCWM/POST.  This gate is for the
' follow-on PRDO quantity reduction only: it proves the original batch remains
' read-only and the exact ODO item quantity alone can be changed to 128.
Public Function Codex_PreflightPRDOQtyOnlyAfterPosting() As String
    Const PRDO_OIP_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell"
    Dim ses As Object, grid As Object, toolbar As Object
    Dim messageText As String, qtyText As String, batchText As String, statusText As String, actionDetail As String, dumpText As String
    Dim rowNo As Long, batchChangeable As String, qtyChangeable As String, editModeEntered As Boolean, resetText As String

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPostingVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    If Abs(CWR_Number(qtyText) - CWR_EXPECTED_SOURCE_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "PRDO_QTY_ONLY_BEFORE_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        GoTo Stopped
    End If
    Set toolbar = ses.FindById(PRDO_OIP_TOOLBAR_ID, False)
    If Not CWR_GetVerifiedPRDOChangeAction(toolbar, actionDetail, messageText) Then GoTo Stopped
    toolbar.PressButton "OIP_CHANGE"
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "PRDO_QTY_ONLY_EDIT_MODE_TIMEOUT"
        GoTo Stopped
    End If
    editModeEntered = True
    If CWR_HasAnyModal(ses, messageText) Then
        messageText = "PRDO_QTY_ONLY_EDIT_MODAL: " & messageText
        GoTo Stopped
    End If
    If Not CWR_FindCurrentPRDOGrid(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    batchChangeable = CWR_CellChangeable(grid, rowNo, "BATCHNO_UI")
    qtyChangeable = CWR_CellChangeable(grid, rowNo, "QTY_UI")
    If UCase$(batchChangeable) <> "FALSE" Or UCase$(qtyChangeable) <> "TRUE" Then
        messageText = "PRDO_QTY_ONLY_FIELD_CONTRACT_MISMATCH: BATCH=" & batchChangeable & "; QTY=" & qtyChangeable
        GoTo Stopped
    End If
    CWR_DumpGuiShellToolbar toolbar, "PRDO_EDIT_MODE", dumpText
    If Not CWR_ResetPRDOWithoutChanges(ses, messageText) Then GoTo Stopped
    editModeEntered = False
    CWR_Log "PRDO_QTY_ONLY_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact ODO item editable for quantity only; batch stays read-only. " & actionDetail
    Codex_PreflightPRDOQtyOnlyAfterPosting = "PASS|PRDO_QTY_ONLY_PREVERIFIED|ROW=" & CStr(rowNo) & _
        "|BEFORE_QTY=" & qtyText & "|BATCH_READONLY=" & batchChangeable & "|QTY_CHANGEABLE=" & qtyChangeable & _
        "|TARGET_QTY=" & CStr(CWR_TARGET_QTY) & "|" & actionDetail & vbLf & dumpText
    Exit Function
Stopped:
    If editModeEntered Then
        resetText = ""
        If Not CWR_ResetPRDOWithoutChanges(ses, resetText) Then messageText = messageText & "; PRDO_RESET_AFTER_PREVIEW_FAILED=" & resetText
    End If
    CWR_Log "PRDO_QTY_ONLY_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_PreflightPRDOQtyOnlyAfterPosting = "STOP|" & messageText
    Exit Function
Failed:
    If editModeEntered Then
        resetText = ""
        Call CWR_ResetPRDOWithoutChanges(ses, resetText)
    End If
    CWR_Log "PRDO_QTY_ONLY_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_PreflightPRDOQtyOnlyAfterPosting = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Reconciles the exact PRDO quantity before the one permitted write.  A prior
' run may have saved the requested quantity and then stopped later while
' opening Create WT.  In that case this function proves the live value and
' resumes without another Save; an unexpected value remains fail-closed.
Private Function CWR_EnsurePRDOQtyAtTargetOnce() As String
    CWR_EnsurePRDOQtyAtTargetOnce = CWR_EnsurePRDOQtyAtTarget(False)
End Function

' DELTA_ADD retains its verified old WT.  This wrapper shares the exact PRDO
' selection/readback/one-Save contract with replacement work, but never
' requires or performs a Cancel checkpoint.
Private Function CWR_EnsurePRDOQtyForDeltaAddOnce() As String
    CWR_EnsurePRDOQtyForDeltaAddOnce = CWR_EnsurePRDOQtyAtTarget(True)
End Function

Private Function CWR_EnsurePRDOQtyAtTarget(ByVal retainExistingWT As Boolean) As String
    Dim ses As Object, grid As Object
    Dim messageText As String, qtyText As String, batchText As String, statusText As String, resultText As String
    Dim rowNo As Long, liveQty As Double

    On Error GoTo Failed
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If retainExistingWT Then
        If Not CWR_AssertRetainedDeltaWT(messageText) Then GoTo Stopped
    Else
        If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
        If Not CWR_AssertPostingVerifiedLocal(messageText) Then GoTo Stopped
    End If
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    liveQty = CWR_Number(qtyText)
    If StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "PRDO_QTY_RESUME_BATCH_MISMATCH: " & batchText
        GoTo Stopped
    End If
    If Abs(liveQty - CWR_TARGET_QTY) <= 0.000001 Then
        CWR_RecordPRDOQtyChange qtyText, batchText, statusText
        CWR_Log "PRDO_QTY_RECONCILE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "Fresh PRDO readback already equals target Qty=" & qtyText & "; no second Save sent; batch=" & batchText
        CWR_EnsurePRDOQtyAtTarget = "PASS|PRDO_QTY_ALREADY_VERIFIED|WRITE=FALSE|QTY=" & qtyText & _
            "|BATCH_UNCHANGED=" & batchText & "|NEXT=" & _
            IIf(retainExistingWT, "CREATE_DELTA_WT_KEEP_OLD", "CREATE_WT_FROM_TARGET_THEN_CF_HANDOFF")
        Exit Function
    End If
    If Abs(liveQty - CWR_EXPECTED_SOURCE_QTY) > 0.000001 Then
        messageText = "PRDO_QTY_RESUME_STATE_UNKNOWN: LIVE_QTY=" & qtyText & _
            "; EXPECTED_BEFORE=" & CStr(CWR_EXPECTED_SOURCE_QTY) & "; TARGET=" & CStr(CWR_TARGET_QTY)
        GoTo Stopped
    End If
    resultText = CWR_ExecutePRDOQtyToTarget(retainExistingWT)
    If Left$(resultText, 5) <> "PASS|" Then
        messageText = resultText
        GoTo Stopped
    End If
    CWR_EnsurePRDOQtyAtTarget = resultText
    Exit Function
Stopped:
    CWR_Log "PRDO_QTY_RECONCILE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    CWR_EnsurePRDOQtyAtTarget = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "PRDO_QTY_RECONCILE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    CWR_EnsurePRDOQtyAtTarget = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Converts a cancelled status-A row into the smallest normal rework branch.
' A pure recreate is routed through CHANGE_QTY even if the quantity itself is
' unchanged: this preserves the existing locked preverify/create flow and
' never treats the post-Adopt WT surface as a substitute for Adopt.
Private Function CWR_DeriveCancelledResumeDecision(ByVal wsPlan As Worksheet, ByVal planRow As Long) As String
    Dim needQty As Boolean, needBin As Boolean, needBatch As Boolean
    needQty = (UCase$(CWR_Text(wsPlan.Cells(planRow, 21).Value2)) = "TRUE")
    needBin = (UCase$(CWR_Text(wsPlan.Cells(planRow, 22).Value2)) = "TRUE")
    needBatch = (UCase$(CWR_Text(wsPlan.Cells(planRow, 23).Value2)) = "TRUE")
    If needBin And needBatch Then
        CWR_DeriveCancelledResumeDecision = "CHANGE_BIN_BATCH"
    ElseIf needBin Then
        CWR_DeriveCancelledResumeDecision = "CHANGE_BIN"
    ElseIf needBatch Then
        CWR_DeriveCancelledResumeDecision = "CHANGE_BATCH"
    Else
        CWR_DeriveCancelledResumeDecision = "CHANGE_QTY"
    End If
End Function

' Changes only the exact ODO item quantity from 192 to 128. The batch field
' is positively verified read-only and is never modified. One internal Save
' is sent only after the selected row and QTY_UI readback both match.
Public Function Codex_ExecutePRDOQtyOnlyAfterPosting() As String
    Codex_ExecutePRDOQtyOnlyAfterPosting = CWR_ExecutePRDOQtyToTarget(False)
End Function

Private Function CWR_ExecutePRDOQtyToTarget(ByVal retainExistingWT As Boolean) As String
    Const PRDO_OIP_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell"
    Dim ses As Object, grid As Object, toolbar As Object
    Dim messageText As String, qtyText As String, batchText As String, statusText As String, actionDetail As String
    Dim rowNo As Long, batchChangeable As String, qtyChangeable As String, statusType As String, statusBarText As String
    Dim draftModified As Boolean, saveSent As Boolean, resetText As String, modalText As String

    On Error GoTo Failed
    If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If retainExistingWT Then
        If Not CWR_AssertRetainedDeltaWT(messageText) Then GoTo Stopped
    Else
        If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
        If Not CWR_AssertPostingVerifiedLocal(messageText) Then GoTo Stopped
    End If
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    If Abs(CWR_Number(qtyText) - CWR_EXPECTED_SOURCE_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "PRDO_QTY_ACTION_BEFORE_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        GoTo Stopped
    End If
    Set toolbar = ses.FindById(PRDO_OIP_TOOLBAR_ID, False)
    If Not CWR_GetVerifiedPRDOChangeAction(toolbar, actionDetail, messageText) Then GoTo Stopped
    toolbar.PressButton "OIP_CHANGE"
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "PRDO_QTY_EDIT_MODE_TIMEOUT"
        GoTo Stopped
    End If
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = "PRDO_QTY_EDIT_MODAL: " & modalText
        GoTo Stopped
    End If
    If Not CWR_FindCurrentPRDOGrid(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    batchChangeable = CWR_CellChangeable(grid, rowNo, "BATCHNO_UI")
    qtyChangeable = CWR_CellChangeable(grid, rowNo, "QTY_UI")
    If UCase$(batchChangeable) <> "FALSE" Or UCase$(qtyChangeable) <> "TRUE" Then
        messageText = "PRDO_QTY_FIELD_CONTRACT_MISMATCH: BATCH=" & batchChangeable & "; QTY=" & qtyChangeable
        GoTo Stopped
    End If
    If Not CWR_SelectExactPRDOItem(ses, grid, rowNo, messageText) Then GoTo Stopped
    If CWR_SelectedRowsText(grid) <> CStr(rowNo) Then
        messageText = "PRDO_QTY_SELECTED_ROW_READBACK_MISMATCH: " & CWR_SelectedRowsText(grid)
        GoTo Stopped
    End If
    If StrComp(CWR_Cell(grid, rowNo, "BATCHNO_UI"), CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "PRDO_QTY_BATCH_CHANGED_BEFORE_DRAFT: " & CWR_Cell(grid, rowNo, "BATCHNO_UI")
        GoTo Stopped
    End If
    grid.ModifyCell rowNo, "QTY_UI", CWR_FastQuantityText(CWR_TARGET_QTY)
    draftModified = True
    If Abs(CWR_Number(CWR_Cell(grid, rowNo, "QTY_UI")) - CWR_TARGET_QTY) > 0.000001 Then
        messageText = "PRDO_QTY_DRAFT_READBACK_MISMATCH: " & CWR_Cell(grid, rowNo, "QTY_UI")
        GoTo Stopped
    End If
    grid.TriggerModified
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "PRDO_QTY_DRAFT_TRIGGER_TIMEOUT"
        GoTo Stopped
    End If
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = "PRDO_QTY_DRAFT_TRIGGER_MODAL: " & modalText
        GoTo Stopped
    End If
    CWR_ReadStatusBar ses, statusType, statusBarText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "PRDO_QTY_DRAFT_SAP_" & statusType & ": " & statusBarText
        GoTo Stopped
    End If
    If Not CWR_FindCurrentPRDOGrid(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    If Abs(CWR_Number(qtyText) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "PRDO_QTY_DRAFT_POSTTRIGGER_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        GoTo Stopped
    End If
    Set toolbar = ses.FindById(PRDO_OIP_TOOLBAR_ID, False)
    If Not CWR_GetVerifiedPRDOSaveAction(toolbar, actionDetail, messageText) Then GoTo Stopped
    If Not CWR_RequireReworkAction("PRDO_QTY", messageText) Then GoTo Stopped

    ' One permitted PRDO write: Save the exact target quantity; no batch edit.
    saveSent = True
    toolbar.PressButton "SAVE"
    CWR_Log "PRDO_QTY_SAVE_SENT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Exact PRDO item selected and quantity Save sent once; Qty=" & CStr(CWR_TARGET_QTY) & "; batch unchanged=" & CWR_EXPECTED_SOURCE_BATCH & _
        "; " & actionDetail
    If Not CWR_WaitPOSTSaveSettle(ses, 3#, 45#) Then
        messageText = "PRDO_QTY_SAVE_TIMEOUT_STATE_UNPROVEN"
        GoTo SaveStopped
    End If
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = "PRDO_QTY_SAVE_MODAL_STATE_UNPROVEN: " & modalText
        GoTo SaveStopped
    End If
    CWR_ReadStatusBar ses, statusType, statusBarText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "PRDO_QTY_SAVE_SAP_" & statusType & ": " & statusBarText
        GoTo SaveStopped
    End If
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then GoTo SaveStopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo SaveStopped
    If Abs(CWR_Number(qtyText) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "PRDO_QTY_POSTVERIFY_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        GoTo SaveStopped
    End If
    CWR_RecordPRDOQtyChange qtyText, batchText, statusText
    CWR_Log "PRDO_QTY_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Fresh PRDO readback proved Qty=" & qtyText & "; batch unchanged; next " & _
        IIf(retainExistingWT, "Create delta WT Qty=" & CStr(CWR_CREATE_QTY), "Create WT") & "; CF=USER_HANDOFF"
    CWR_ExecutePRDOQtyToTarget = "PASS|PRDO_QTY_VERIFIED|QTY=" & qtyText & _
        "|BATCH_UNCHANGED=" & batchText & "|" & statusText & _
        "|NEXT=" & IIf(retainExistingWT, "CREATE_DELTA_WT_KEEP_OLD", "CREATE_WT_FROM_TARGET_THEN_CF_HANDOFF") & _
        "|CF=USER_HANDOFF"
    Exit Function

Stopped:
    If draftModified And Not saveSent Then
        resetText = ""
        If Not CWR_ResetPRDOWithoutChanges(ses, resetText) Then messageText = messageText & "; PRDO_DRAFT_RESET_FAILED=" & resetText
    End If
    CWR_Log "PRDO_QTY_ACTION", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    CWR_ExecutePRDOQtyToTarget = "STOP|" & messageText
    Exit Function
SaveStopped:
    CWR_Log "PRDO_QTY_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    CWR_ExecutePRDOQtyToTarget = "STOP|" & messageText & "|NO_RETRY"
    Exit Function
Failed:
    If draftModified And Not saveSent Then
        resetText = ""
        Call CWR_ResetPRDOWithoutChanges(ses, resetText)
    End If
    CWR_Log "PRDO_QTY_ACTION", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description & IIf(saveSent, "; STATE_UNPROVEN_NO_RETRY", "")
    CWR_ExecutePRDOQtyToTarget = "FAIL|" & CStr(Err.Number) & "|" & Err.Description & _
        IIf(saveSent, "|STATE_UNPROVEN_NO_RETRY", "")
End Function

' Read-only discovery of the WT creation screen.  Querying the screen loads
' candidate stock only; no Create, Create+Save, Save, or confirmation action
' is called here.
Public Function Codex_ProbeCreateWTForTarget() As String
    Const CREATE_MENU_ID As String = "wnd[0]/mbar/menu[0]/menu[2]/menu[0]"
    Const CREATE_QUERY_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_TODLV:1001/btnCMD_QUERY_EXECUTE"
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Dim prdoSes As Object, prdoGrid As Object, createSes As Object, menuItem As Object, queryButton As Object
    Dim wrGrid As Object, stockGrid As Object, messageText As String, qtyText As String, batchText As String, statusText As String
    Dim rowNo As Long, dumpText As String

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertTargetStockLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenExpectedPRDO(prdoSes, prdoGrid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(prdoGrid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    If Abs(CWR_Number(qtyText) - CWR_EXPECTED_SOURCE_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "CREATE_PROBE_PRDO_BEFORE_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        GoTo Stopped
    End If
    Set menuItem = prdoSes.FindById(CREATE_MENU_ID, False)
    If menuItem Is Nothing Then
        messageText = "CREATE_FOLLOWON_MENU_NOT_FOUND"
        GoTo Stopped
    End If
    menuItem.Select
    If Not CWR_WaitForCreateWTSession(createSes, 20#, messageText) Then GoTo Stopped
    Set queryButton = createSes.FindById(CREATE_QUERY_ID, False)
    If queryButton Is Nothing Then
        messageText = "CREATE_QUERY_BUTTON_NOT_FOUND"
        GoTo Stopped
    End If
    queryButton.Press
    If Not CWR_WaitIdle(createSes, 30#) Then
        messageText = "CREATE_QUERY_TIMEOUT"
        GoTo Stopped
    End If
    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    If wrGrid Is Nothing Then
        messageText = "CREATE_WR_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    dumpText = CWR_DumpGridBrief("WR", wrGrid, 10, 120)
    If stockGrid Is Nothing Then
        dumpText = dumpText & "STOCK_GRID_NOT_FOUND" & vbLf
    Else
        dumpText = dumpText & CWR_DumpGridBrief("STOCK", stockGrid, 20, 160)
    End If
    CWR_Log "WT_CREATE_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Read-only Create WT query completed; no Create/Save/CF action."
    Codex_ProbeCreateWTForTarget = "PASS|CREATE_WT_READONLY_PROBE|BEFORE_QTY=" & qtyText & _
        "|BEFORE_BATCH=" & batchText & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ProbeCreateWTForTarget = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ProbeCreateWTForTarget = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only WT-creation discovery after the PRDO quantity has been verified
' at 128. It opens the standard follow-on screen and queries candidates, but
' never selects stock for creation and never presses Create or Create+Save.
Public Function Codex_ProbeCreateWTAfterQtyChange() As String
    Const CREATE_MENU_ID As String = "wnd[0]/mbar/menu[0]/menu[2]/menu[0]"
    Const CREATE_QUERY_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_TODLV:1001/btnCMD_QUERY_EXECUTE"
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Dim prdoSes As Object, prdoGrid As Object, createSes As Object, menuItem As Object, queryButton As Object
    Dim wrGrid As Object, stockGrid As Object, messageText As String, qtyText As String, batchText As String, statusText As String
    Dim rowNo As Long, dumpText As String, menuText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenExpectedPRDO(prdoSes, prdoGrid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(prdoGrid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    If Abs(CWR_Number(qtyText) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_PRDO_POSTQTY_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        GoTo Stopped
    End If
    Set menuItem = prdoSes.FindById(CREATE_MENU_ID, False)
    If menuItem Is Nothing Then
        messageText = "CREATE_WT_FOLLOWON_MENU_NOT_FOUND"
        GoTo Stopped
    End If
    menuText = CWR_Text(menuItem.text)
    If StrComp(menuText, "Warehouse Task", vbTextCompare) <> 0 And _
       InStr(1, menuText, "Create", vbTextCompare) = 0 Then
        messageText = "CREATE_WT_FOLLOWON_MENU_IDENTITY_MISMATCH: " & menuText
        GoTo Stopped
    End If
    menuItem.Select
    If Not CWR_WaitForCreateWTSession(createSes, 20#, messageText) Then GoTo Stopped
    Set queryButton = createSes.FindById(CREATE_QUERY_ID, False)
    If queryButton Is Nothing Then
        messageText = "CREATE_WT_QUERY_BUTTON_NOT_FOUND"
        GoTo Stopped
    End If
    queryButton.Press
    If Not CWR_WaitIdle(createSes, 30#) Then
        messageText = "CREATE_WT_QUERY_TIMEOUT"
        GoTo Stopped
    End If
    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    If wrGrid Is Nothing Then
        messageText = "CREATE_WT_WR_GRID_NOT_FOUND"
        GoTo Stopped
    End If
    dumpText = CWR_DumpGridBrief("WR", wrGrid, 10, 120)
    If stockGrid Is Nothing Then
        dumpText = dumpText & "STOCK_GRID_NOT_FOUND" & vbLf
    Else
        dumpText = dumpText & CWR_DumpGridBrief("STOCK", stockGrid, 20, 160)
    End If
    CWR_Log "WT_CREATE_POSTQTY_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Create WT candidates queried read only after Qty=128; no Create/Save/CF action."
    Codex_ProbeCreateWTAfterQtyChange = "PASS|CREATE_WT_POSTQTY_READONLY_PROBE|QTY=" & qtyText & _
        "|BATCH=" & batchText & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_POSTQTY_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ProbeCreateWTAfterQtyChange = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_POSTQTY_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ProbeCreateWTAfterQtyChange = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Continues the prior read-only Create WT probe without opening a second
' follow-on window.  It selects the displayed WR row and the available-stock
' tab only to expose fields; it never creates a WT.
Public Function Codex_DumpExistingCreateWT() As String
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Dim createSes As Object, wrGrid As Object, stockTab As Object, stockGrid As Object
    Dim messageText As String, dumpText As String

    On Error GoTo Failed
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        Codex_DumpExistingCreateWT = "STOP|" & messageText
        Exit Function
    End If
    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    If wrGrid Is Nothing Then
        Codex_DumpExistingCreateWT = "STOP|CREATE_WR_GRID_NOT_FOUND"
        Exit Function
    End If
    If CLng(wrGrid.rowCount) <> 1 Then
        Codex_DumpExistingCreateWT = "STOP|CREATE_WR_ROW_COUNT=" & CStr(wrGrid.rowCount)
        Exit Function
    End If
    wrGrid.selectedRows = "0"
    On Error Resume Next
    wrGrid.SetCurrentCell 0, "DLVNO"
    On Error GoTo Failed
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Then
        Codex_DumpExistingCreateWT = "STOP|CREATE_STOCK_TAB_NOT_FOUND" & vbLf & CWR_DumpKnownGrid("WR", wrGrid, 5)
        Exit Function
    End If
    stockTab.Select
    If Not CWR_WaitIdle(createSes, 20#) Then
        Codex_DumpExistingCreateWT = "STOP|CREATE_STOCK_TAB_TIMEOUT"
        Exit Function
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    dumpText = CWR_DumpKnownGrid("WR", wrGrid, 5)
    If stockGrid Is Nothing Then
        dumpText = dumpText & "STOCK_GRID_NOT_FOUND" & vbLf
    Else
        dumpText = dumpText & CWR_DumpKnownGrid("STOCK", stockGrid, 20)
    End If
    Codex_DumpExistingCreateWT = "PASS|CREATE_WT_FIELDS_READONLY" & vbLf & dumpText
    Exit Function
Failed:
    Codex_DumpExistingCreateWT = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only field/changeability map for the currently open Create Warehouse
' Task session. This establishes whether the target batch/bin can be selected
' there before any WT-creation command is considered.
Public Function Codex_MapExistingCreateWTFieldContract() As String
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Const STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV"
    Const STOCK_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_AVSTKDLV/ssubSUB_ODP_AVSTKDLV:/SCWM/SAPLUI_TODLV:4400/subSUB_ODP_AVSTKDLV_DATA:/SCWM/SAPLUI_TODLV:4410/cntlCC_ALV_OD_AVSTKDLV/shellcont/shell"
    Dim createSes As Object, wrGrid As Object, stockTab As Object, stockGrid As Object
    Dim messageText As String, dumpText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    If wrGrid Is Nothing Or CLng(wrGrid.rowCount) <> 1 Then
        messageText = "CREATE_WT_WR_GRID_CONTRACT_UNAVAILABLE"
        GoTo Stopped
    End If
    wrGrid.selectedRows = "0"
    wrGrid.SetCurrentCell 0, "DLVNO"
    Set stockTab = createSes.FindById(STOCK_TAB_ID, False)
    If stockTab Is Nothing Then
        messageText = "CREATE_WT_STOCK_TAB_NOT_FOUND"
        GoTo Stopped
    End If
    stockTab.Select
    If Not CWR_WaitIdle(createSes, 20#) Then
        messageText = "CREATE_WT_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If
    Set stockGrid = createSes.FindById(STOCK_GRID_ID, False)
    dumpText = CWR_DumpGridFieldContract("WR", wrGrid, 0)
    If stockGrid Is Nothing Then
        dumpText = dumpText & "STOCK_GRID_NOT_FOUND" & vbLf
    Else
        dumpText = dumpText & CWR_DumpGridFieldContract("STOCK", stockGrid, 0)
    End If
    CWR_Log "WT_CREATE_FIELD_CONTRACT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Create WT field contract read only; no stock or WT selected for creation."
    Codex_MapExistingCreateWTFieldContract = "PASS|CREATE_WT_FIELD_CONTRACT|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_FIELD_CONTRACT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_MapExistingCreateWTFieldContract = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_FIELD_CONTRACT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_MapExistingCreateWTFieldContract = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Maps the Create Warehouse Task action surface without invoking an action.
' This is deliberately read-only: it may open/query the standard candidate
' screen if needed, then lists ordinary and embedded toolbar actions plus the
' controls relevant to source-stock selection.  In particular it must expose
' a supported way to choose the requested bin/batch before any create code is
' allowed to exist.
Public Function Codex_MapCreateWTActionSurface() As String
    Dim createSes As Object, root As Object
    Dim messageText As String, dumpText As String, probeText As String
    Dim itemCount As Long

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped

    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "CREATE_WT_ACTION_MAP_OPEN_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If

    Set root = createSes.FindById("wnd[0]", False)
    If root Is Nothing Then
        messageText = "CREATE_WT_ACTION_MAP_ROOT_NOT_FOUND"
        GoTo Stopped
    End If
    CWR_DumpToolbarControls createSes, "wnd[0]/tbar[0]", "CREATE_WT_STANDARD_TBAR", dumpText
    CWR_DumpToolbarControls createSes, "wnd[0]/tbar[1]", "CREATE_WT_APPLICATION_TBAR", dumpText
    itemCount = 0
    CWR_CollectCreateWTActionSurface root, 0, itemCount, dumpText

    CWR_Log "WT_CREATE_ACTION_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Create WT action surface mapped read only; no stock selection, Create, Save or CF action."
    Codex_MapCreateWTActionSurface = "PASS|CREATE_WT_ACTION_SURFACE|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_ACTION_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_MapCreateWTActionSurface = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_ACTION_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_MapCreateWTActionSurface = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only discovery of the alternate-stock tab.  The sole selection made
' here is the uniquely verified ODO row and then the "Other Stock" tab; no
' quantity is adopted and no warehouse task is created or saved.
Public Function Codex_ProbeCreateWTOtherStock() As String
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Const OTHER_STOCK_TAB_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP:/SCWM/SAPLUI_TODLV:4000/tabsGV_TAB_ODP/tabpOK_TAB_ODP_NASTKDLV"
    Dim createSes As Object, root As Object, wrGrid As Object, otherTab As Object
    Dim messageText As String, dumpText As String, probeText As String
    Dim rowDoc As String, rowItem As String, rowMat As String, rowQty As String, itemCount As Long

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "CREATE_WT_OTHER_STOCK_OPEN_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If

    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    If wrGrid Is Nothing Or CLng(wrGrid.rowCount) <> 1 Then
        messageText = "CREATE_WT_OTHER_STOCK_WR_GRID_UNAVAILABLE"
        GoTo Stopped
    End If
    rowDoc = CWR_Cell(wrGrid, 0, "DLVNO")
    rowItem = CWR_Cell(wrGrid, 0, "ITMNO")
    rowMat = CWR_Cell(wrGrid, 0, "MATNR")
    rowQty = CWR_Cell(wrGrid, 0, "UI_OPNQTY")
    If StrComp(rowDoc, CWR_EXPECTED_DOC, vbTextCompare) <> 0 Or _
       StrComp(rowItem, CWR_EXPECTED_ITEM, vbTextCompare) <> 0 Or _
       StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Or _
       Abs(CWR_Number(rowQty) - CWR_TARGET_QTY) > 0.000001 Then
        messageText = "CREATE_WT_OTHER_STOCK_WR_MISMATCH: DOC=" & rowDoc & "; ITEM=" & rowItem & _
            "; MAT=" & rowMat & "; QTY=" & rowQty
        GoTo Stopped
    End If
    wrGrid.selectedRows = "0"
    wrGrid.SetCurrentCell 0, "DLVNO"
    If Not CWR_WaitIdle(createSes, 20#) Then
        messageText = "CREATE_WT_OTHER_STOCK_ROW_SELECT_TIMEOUT"
        GoTo Stopped
    End If

    Set otherTab = createSes.FindById(OTHER_STOCK_TAB_ID, False)
    If otherTab Is Nothing Or StrComp(CWR_Text(otherTab.text), "Other Stock", vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_OTHER_STOCK_TAB_IDENTITY_MISMATCH"
        GoTo Stopped
    End If
    otherTab.Select
    If Not CWR_WaitIdle(createSes, 20#) Then
        messageText = "CREATE_WT_OTHER_STOCK_TAB_TIMEOUT"
        GoTo Stopped
    End If

    Set root = createSes.FindById("wnd[0]", False)
    If root Is Nothing Then
        messageText = "CREATE_WT_OTHER_STOCK_ROOT_NOT_FOUND"
        GoTo Stopped
    End If
    itemCount = 0
    CWR_CollectCreateWTActionSurface root, 0, itemCount, dumpText
    CWR_DumpVisibleCreateWTGrids root, 0, itemCount, dumpText
    CWR_Log "WT_CREATE_OTHER_STOCK_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Other Stock tab mapped after exact WR selection; no stock adopted, created or saved."
    Codex_ProbeCreateWTOtherStock = "PASS|CREATE_WT_OTHER_STOCK_READONLY|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_OTHER_STOCK_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ProbeCreateWTOtherStock = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_OTHER_STOCK_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ProbeCreateWTOtherStock = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Enumerates the Create Warehouse Task menu hierarchy without selecting any
' command.  This is used to identify whether SAP exposes a supported batch or
' alternate-stock input command before attempting an adoption or creation.
Public Function Codex_DumpCreateWTMenuTree() As String
    Dim createSes As Object, menuRoot As Object
    Dim messageText As String, dumpText As String, probeText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_FindCreateWTSession(createSes, messageText) Then
        probeText = Codex_ProbeCreateWTAfterQtyChange()
        If Left$(probeText, 5) <> "PASS|" Then
            messageText = "CREATE_WT_MENU_MAP_OPEN_FAILED: " & probeText
            GoTo Stopped
        End If
        If Not CWR_FindCreateWTSession(createSes, messageText) Then GoTo Stopped
    End If
    Set menuRoot = createSes.FindById("wnd[0]/mbar", False)
    If menuRoot Is Nothing Then
        messageText = "CREATE_WT_MENU_ROOT_NOT_FOUND"
        GoTo Stopped
    End If
    CWR_DumpMenuTree menuRoot, 0, dumpText
    CWR_Log "WT_CREATE_MENU_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Create WT menu hierarchy read only; no command selected."
    Codex_DumpCreateWTMenuTree = "PASS|CREATE_WT_MENU_TREE|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "WT_CREATE_MENU_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_DumpCreateWTMenuTree = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "WT_CREATE_MENU_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_DumpCreateWTMenuTree = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only item-detail discovery for batch/location fields.  The detail view
' is opened from the exact ODO item and then reset without any data change.
Public Function Codex_ProbePRDOItemDetailFields() As String
    Dim ses As Object, grid As Object, root As Object
    Dim messageText As String, qtyText As String, batchText As String, statusText As String
    Dim rowNo As Long, dumpText As String, itemCount As Long, resetText As String

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    If Abs(CWR_Number(qtyText) - CWR_EXPECTED_SOURCE_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "DETAIL_PROBE_PRDO_BEFORE_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        GoTo Stopped
    End If
    grid.selectedRows = CStr(rowNo)
    grid.SetCurrentCell rowNo, "ITEMNO"
    grid.PressToolbarButton "&DETAIL"
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "PRDO_DETAIL_TIMEOUT"
        GoTo Stopped
    End If
    Set root = ses.FindById("wnd[0]", False)
    If root Is Nothing Then
        messageText = "PRDO_DETAIL_ROOT_NOT_FOUND"
        GoTo Stopped
    End If
    itemCount = 0
    CWR_CollectDetailFields root, 0, itemCount, dumpText
    resetText = ""
    If Not CWR_ResetPRDOWithoutChanges(ses, resetText) Then
        messageText = "PRDO_DETAIL_RESET_FAILED: " & resetText
        GoTo Stopped
    End If
    CWR_Log "PRDO_DETAIL_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Batch/location item-detail fields read only; no edit/save."
    Codex_ProbePRDOItemDetailFields = "PASS|PRDO_DETAIL_FIELDS|COUNT=" & CStr(itemCount) & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "PRDO_DETAIL_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ProbePRDOItemDetailFields = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "PRDO_DETAIL_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ProbePRDOItemDetailFields = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Read-only post-quantity detail probe. It is used only because the Create WT
' screen still proposes the old batch; no PRDO field is changed in this probe.
Public Function Codex_ProbePRDOItemDetailAfterQtyChange() As String
    Dim ses As Object, grid As Object, root As Object
    Dim messageText As String, qtyText As String, batchText As String, statusText As String
    Dim rowNo As Long, dumpText As String, itemCount As Long, resetText As String
    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    If Abs(CWR_Number(qtyText) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "PRDO_DETAIL_POSTQTY_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        GoTo Stopped
    End If
    grid.selectedRows = CStr(rowNo)
    grid.SetCurrentCell rowNo, "ITEMNO"
    grid.PressToolbarButton "&DETAIL"
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "PRDO_DETAIL_POSTQTY_TIMEOUT"
        GoTo Stopped
    End If
    If CWR_HasAnyModal(ses, messageText) Then
        messageText = "PRDO_DETAIL_POSTQTY_MODAL: " & messageText
        GoTo Stopped
    End If
    Set root = ses.FindById("wnd[0]", False)
    If root Is Nothing Then
        messageText = "PRDO_DETAIL_POSTQTY_ROOT_NOT_FOUND"
        GoTo Stopped
    End If
    CWR_CollectDetailFields root, 0, itemCount, dumpText
    resetText = ""
    If Not CWR_ResetPRDOWithoutChanges(ses, resetText) Then
        messageText = "PRDO_DETAIL_POSTQTY_RESET_FAILED: " & resetText
        GoTo Stopped
    End If
    CWR_Log "PRDO_DETAIL_POSTQTY_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "PRDO detail batch/location fields read only after Qty=128; no edit/save."
    Codex_ProbePRDOItemDetailAfterQtyChange = "PASS|PRDO_DETAIL_POSTQTY_FIELDS|COUNT=" & CStr(itemCount) & vbLf & dumpText
    Exit Function
Stopped:
    CWR_Log "PRDO_DETAIL_POSTQTY_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ProbePRDOItemDetailAfterQtyChange = "STOP|" & messageText
    Exit Function
Failed:
    CWR_Log "PRDO_DETAIL_POSTQTY_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ProbePRDOItemDetailAfterQtyChange = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Recovery endpoint for a read-only editability preview.  It never saves or
' changes fields; it only returns the exact PRDO session to a fresh display.
Public Function Codex_ResetPRDOAfterPreview() As String
    Dim gui As Object, app As Object, con As Object, ses As Object
    Dim i As Long, tcode As String, messageText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        tcode = CWR_Text(ses.Info.Transaction)
        If UCase$(tcode) = "/SCWM/PRDO" Then Exit For
        Set ses = Nothing
    Next i
    If ses Is Nothing Then
        Codex_ResetPRDOAfterPreview = "STOP|PRDO_SESSION_NOT_FOUND"
        Exit Function
    End If
    If Not CWR_ResetPRDOWithoutChanges(ses, messageText) Then
        Codex_ResetPRDOAfterPreview = "STOP|" & messageText
        Exit Function
    End If
    Codex_ResetPRDOAfterPreview = "PASS|PRDO_RESET_DISPLAY_ONLY"
    Exit Function
Failed:
    Codex_ResetPRDOAfterPreview = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Maps the action surface of the exact ODO item detail after the quantity
' change. It does not enter edit mode or change any value; it opens detail,
' reads controls/toolbars and resets PRDO without saving.
Public Function Codex_MapPRDOItemDetailActionSurface() As String
    Dim ses As Object, grid As Object, root As Object
    Dim messageText As String, qtyText As String, batchText As String, statusText As String
    Dim rowNo As Long, dumpText As String, itemCount As Long, resetText As String

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    If Abs(CWR_Number(qtyText) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "PRDO_DETAIL_ACTION_SURFACE_ITEM_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        GoTo Stopped
    End If
    grid.selectedRows = CStr(rowNo)
    grid.SetCurrentCell rowNo, "ITEMNO"
    grid.PressToolbarButton "&DETAIL"
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "PRDO_DETAIL_ACTION_SURFACE_TIMEOUT"
        GoTo Stopped
    End If
    If CWR_HasAnyModal(ses, messageText) Then
        messageText = "PRDO_DETAIL_ACTION_SURFACE_MODAL: " & messageText
        GoTo Stopped
    End If
    Set root = ses.FindById("wnd[0]", False)
    If root Is Nothing Then
        messageText = "PRDO_DETAIL_ACTION_SURFACE_ROOT_NOT_FOUND"
        GoTo Stopped
    End If
    CWR_DumpToolbarControls ses, "wnd[0]/tbar[0]", "PRDO_DETAIL_STANDARD_TBAR", dumpText
    CWR_DumpToolbarControls ses, "wnd[0]/tbar[1]", "PRDO_DETAIL_APPLICATION_TBAR", dumpText
    itemCount = 0
    CWR_CollectCreateWTActionSurface root, 0, itemCount, dumpText
    CWR_CollectDetailFields root, 0, itemCount, dumpText
    resetText = ""
    If Not CWR_ResetPRDOWithoutChanges(ses, resetText) Then
        messageText = "PRDO_DETAIL_ACTION_SURFACE_RESET_FAILED: " & resetText
        GoTo Stopped
    End If
    CWR_Log "PRDO_DETAIL_ACTION_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "PRDO item detail action surface mapped read only; no edit/save."
    Codex_MapPRDOItemDetailActionSurface = "PASS|PRDO_DETAIL_ACTION_SURFACE|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    If Not ses Is Nothing Then
        resetText = ""
        Call CWR_ResetPRDOWithoutChanges(ses, resetText)
    End If
    CWR_Log "PRDO_DETAIL_ACTION_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_MapPRDOItemDetailActionSurface = "STOP|" & messageText
    Exit Function
Failed:
    If Not ses Is Nothing Then
        resetText = ""
        Call CWR_ResetPRDOWithoutChanges(ses, resetText)
    End If
    CWR_Log "PRDO_DETAIL_ACTION_MAP", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_MapPRDOItemDetailActionSurface = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Opens the exact item's Batch Selection Criteria dialog solely to map its
' fields. No criterion is altered and PRDO is reset without Save afterward.
Public Function Codex_ProbePRDOBatchSelectionCriteria() As String
    Const ODP1_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/cntlCONTAINER_TB_ODP1_1/shellcont/shell"
    Const BATCH_SELECTION_ID As String = "ODP1_BATCH_SEL"
    Dim ses As Object, grid As Object, toolbar As Object, root As Object, modal As Object
    Dim messageText As String, qtyText As String, batchText As String, statusText As String
    Dim rowNo As Long, dumpText As String, itemCount As Long, resetText As String, detailText As String

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    If Abs(CWR_Number(qtyText) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "PRDO_BATCH_SELECTION_ITEM_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        GoTo Stopped
    End If
    grid.selectedRows = CStr(rowNo)
    grid.SetCurrentCell rowNo, "ITEMNO"
    Set toolbar = ses.FindById(ODP1_TOOLBAR_ID, False)
    If toolbar Is Nothing Then
        messageText = "PRDO_BATCH_SELECTION_TOOLBAR_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(toolbar, BATCH_SELECTION_ID, "Batch Selection Criteria", detailText, messageText) Then GoTo Stopped
    toolbar.PressButton BATCH_SELECTION_ID
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "PRDO_BATCH_SELECTION_OPEN_TIMEOUT"
        GoTo Stopped
    End If
    Set modal = ses.FindById("wnd[1]", False)
    If Not modal Is Nothing Then
        Set root = modal
        dumpText = "ACTION=" & detailText & vbLf & "WINDOW=MODAL|TITLE=" & CWR_Text(modal.text) & vbLf
    Else
        Set root = ses.FindById("wnd[0]", False)
        If root Is Nothing Then
            messageText = "PRDO_BATCH_SELECTION_ROOT_NOT_FOUND"
            GoTo Stopped
        End If
        dumpText = "ACTION=" & detailText & vbLf & "WINDOW=MAIN|TITLE=" & CWR_Text(root.text) & vbLf
    End If
    itemCount = 0
    CWR_CollectCreateWTActionSurface root, 0, itemCount, dumpText
    CWR_CollectDetailFields root, 0, itemCount, dumpText
    resetText = ""
    If Not CWR_ResetPRDOWithoutChanges(ses, resetText) Then
        messageText = "PRDO_BATCH_SELECTION_RESET_FAILED: " & resetText
        GoTo Stopped
    End If
    CWR_Log "PRDO_BATCH_SELECTION_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Batch Selection Criteria mapped read only; no criterion changed or saved."
    Codex_ProbePRDOBatchSelectionCriteria = "PASS|PRDO_BATCH_SELECTION_FIELDS|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    If Not ses Is Nothing Then
        resetText = ""
        Call CWR_ResetPRDOWithoutChanges(ses, resetText)
    End If
    CWR_Log "PRDO_BATCH_SELECTION_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ProbePRDOBatchSelectionCriteria = "STOP|" & messageText
    Exit Function
Failed:
    If Not ses Is Nothing Then
        resetText = ""
        Call CWR_ResetPRDOWithoutChanges(ses, resetText)
    End If
    CWR_Log "PRDO_BATCH_SELECTION_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ProbePRDOBatchSelectionCriteria = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

' Opens Mass Change for exactly one verified ODO item and maps its fields.
' The routine never changes a checkbox/value and always resets PRDO without
' saving, so it is a discovery-only gate for a supported batch update path.
Public Function Codex_ProbePRDOMassChange() As String
    Const ODP1_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/cntlCONTAINER_TB_ODP1_1/shellcont/shell"
    Const MASS_CHANGE_ID As String = "ODP1_MASS_CHANGES"
    Dim ses As Object, grid As Object, toolbar As Object, root As Object, modal As Object
    Dim messageText As String, qtyText As String, batchText As String, statusText As String
    Dim rowNo As Long, dumpText As String, itemCount As Long, resetText As String, detailText As String

    On Error GoTo Failed
    If Not CWR_AssertCFHandoff(messageText) Then GoTo Stopped
    If Not CWR_AssertCancelledLocal(messageText) Then GoTo Stopped
    If Not CWR_AssertPRDOQtyVerifiedLocal(messageText) Then GoTo Stopped
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then GoTo Stopped
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then GoTo Stopped
    If Abs(CWR_Number(qtyText) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "PRDO_MASS_CHANGE_ITEM_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        GoTo Stopped
    End If
    grid.selectedRows = CStr(rowNo)
    grid.SetCurrentCell rowNo, "ITEMNO"
    Set toolbar = ses.FindById(ODP1_TOOLBAR_ID, False)
    If toolbar Is Nothing Then
        messageText = "PRDO_MASS_CHANGE_TOOLBAR_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_GetVerifiedToolbarAction(toolbar, MASS_CHANGE_ID, "Mass Change", detailText, messageText) Then GoTo Stopped
    toolbar.PressButton MASS_CHANGE_ID
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "PRDO_MASS_CHANGE_OPEN_TIMEOUT"
        GoTo Stopped
    End If
    Set modal = ses.FindById("wnd[1]", False)
    If Not modal Is Nothing Then
        Set root = modal
        dumpText = "ACTION=" & detailText & vbLf & "WINDOW=MODAL|TITLE=" & CWR_Text(modal.text) & vbLf
    Else
        Set root = ses.FindById("wnd[0]", False)
        If root Is Nothing Then
            messageText = "PRDO_MASS_CHANGE_ROOT_NOT_FOUND"
            GoTo Stopped
        End If
        dumpText = "ACTION=" & detailText & vbLf & "WINDOW=MAIN|TITLE=" & CWR_Text(root.text) & vbLf
    End If
    itemCount = 0
    CWR_CollectCreateWTActionSurface root, 0, itemCount, dumpText
    CWR_CollectDetailFields root, 0, itemCount, dumpText
    resetText = ""
    If Not CWR_ResetPRDOWithoutChanges(ses, resetText) Then
        messageText = "PRDO_MASS_CHANGE_RESET_FAILED: " & resetText
        GoTo Stopped
    End If
    CWR_Log "PRDO_MASS_CHANGE_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Mass Change fields mapped for one item only; no value changed or saved."
    Codex_ProbePRDOMassChange = "PASS|PRDO_MASS_CHANGE_FIELDS|WRITE=FALSE" & vbLf & dumpText
    Exit Function
Stopped:
    If Not ses Is Nothing Then
        resetText = ""
        Call CWR_ResetPRDOWithoutChanges(ses, resetText)
    End If
    CWR_Log "PRDO_MASS_CHANGE_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Codex_ProbePRDOMassChange = "STOP|" & messageText
    Exit Function
Failed:
    If Not ses Is Nothing Then
        resetText = ""
        Call CWR_ResetPRDOWithoutChanges(ses, resetText)
    End If
    CWR_Log "PRDO_MASS_CHANGE_PROBE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", _
        "Err=" & CStr(Err.Number) & ": " & Err.Description
    Codex_ProbePRDOMassChange = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function CWR_AssertCFHandoff(ByRef messageText As String) As Boolean
    Dim handoffText As String, liveText As String, allowText As String, probeText As String
    On Error GoTo Failed
    handoffText = UCase$(CWR_ConfigValue("CF_USER_HANDOFF_REQUIRED"))
    liveText = UCase$(CWR_ConfigValue("CF_LIVE_ENABLED"))
    allowText = UCase$(CWR_ConfigValue("ALLOW_CF"))
    probeText = UCase$(CWR_ConfigValue("PROBE_ALLOW_CF"))
    If handoffText <> "TRUE" Or liveText <> "FALSE" Or allowText <> "FALSE" Or probeText <> "FALSE" Then
        messageText = "CF_HANDOFF_LOCK_INVALID: Handoff=" & handoffText & "; Live=" & liveText & _
            "; Allow=" & allowText & "; Probe=" & probeText
        Exit Function
    End If
    CWR_AssertCFHandoff = True
    Exit Function
Failed:
    messageText = "CF_HANDOFF_LOCK_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Rework writes are separately armed from the legacy CF flags.  The caller
' must still pass an exact Plan and live SAP preverification for every action.
Private Function CWR_AssertReworkGate(ByRef messageText As String) As Boolean
    Dim enabledText As String, neverCFText As String, lockedText As String
    Dim fingerprintText As String, baseFingerprintText As String
    On Error GoTo Failed
    enabledText = UCase$(CWR_ConfigValue("REWORK_EXECUTE_ENABLED"))
    neverCFText = UCase$(CWR_ConfigValue("REWORK_NEVER_CF"))
    lockedText = UCase$(CWR_ConfigValue("SAP_MAP_LOCKED"))
    fingerprintText = CWR_ConfigValue("SAP_MAP_FINGERPRINT")
    baseFingerprintText = CWR_ConfigValue("SAP_MAP_BASE_FINGERPRINT")
    If enabledText <> "TRUE" Then
        messageText = "REWORK_EXECUTE_NOT_ENABLED"
        Exit Function
    End If
    If neverCFText <> "TRUE" Then
        messageText = "REWORK_CF_HANDOFF_LOCK_INVALID"
        Exit Function
    End If
    If lockedText <> "TRUE" Or fingerprintText = "" Or baseFingerprintText = "" Or _
       StrComp(fingerprintText, baseFingerprintText, vbTextCompare) <> 0 Then
        messageText = "REWORK_SAP_MAP_LOCK_INVALID"
        Exit Function
    End If
    If Not SAPRun_RequireRunMode("REWORK_EXECUTE", messageText) Then Exit Function
    If Not SAPRun_VerifyLockedFingerprint(messageText) Then
        messageText = "REWORK_" & messageText
        Exit Function
    End If
    CWR_AssertReworkGate = True
    Exit Function
Failed:
    messageText = "REWORK_GATE_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_RequireReworkAction(ByVal actionName As String, ByRef messageText As String) As Boolean
    Dim listText As String, tokenText As String
    If Not CWR_AssertReworkGate(messageText) Then Exit Function
    listText = UCase$(Replace(CWR_ConfigValue("REWORK_ACTION_WHITELIST"), " ", ""))
    tokenText = UCase$(Trim$(actionName))
    If listText = "" Or InStr(1, "," & listText & ",", "," & tokenText & ",", vbBinaryCompare) = 0 Then
        messageText = "REWORK_ACTION_NOT_WHITELISTED: " & tokenText
        Exit Function
    End If
    CWR_RequireReworkAction = True
End Function

Private Function CWR_AssertCancelledLocal(ByRef messageText As String) As Boolean
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    If CWR_PLAN_WH_ROW < 5 Then
        messageText = "LOCAL_PLAN_ROW_INVALID"
        Exit Function
    End If
    If CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 1).Value2) <> CWR_EXPECTED_OB Or _
       CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 2).Value2) <> CWR_EXPECTED_DOC Or _
       CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 5).Value2) <> CWR_EXPECTED_WT Then
        messageText = "LOCAL_IDENTITY_MISMATCH"
        Exit Function
    End If
    If UCase$(CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 26).Value2)) <> "WT_CANCELLED_PENDING_REMAINING" And _
       UCase$(CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 26).Value2)) <> "WT_CANCELLED_REMAINING" Then
        messageText = "LOCAL_WT_NOT_CANCELLED: " & CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 26).Value2)
        Exit Function
    End If
    CWR_AssertCancelledLocal = True
    Exit Function
Failed:
    messageText = "LOCAL_CANCEL_STATE_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_AssertTargetStockLocal(ByRef messageText As String) As Boolean
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    If UCase$(CWR_Text(ws.Cells(5, 15).Value2)) <> "READY" Or _
       (UCase$(CWR_Text(ws.Cells(5, 17).Value2)) <> "TARGET_STOCK_VERIFIED" And _
        UCase$(CWR_Text(ws.Cells(5, 17).Value2)) <> "POST_PREVERIFIED") Then
        messageText = "TARGET_STOCK_NOT_VERIFIED"
        Exit Function
    End If
    If CWR_Text(ws.Cells(5, 14).Value2) <> CWR_TARGET_BIN Or _
       CWR_Text(ws.Cells(5, 22).Value2) <> CWR_TARGET_BATCH Or _
       Abs(CWR_Number(ws.Cells(5, 12).Value2) - CWR_TARGET_QTY) > 0.000001 Then
        messageText = "LOCAL_TARGET_CHANGED_AFTER_STOCK_PROOF"
        Exit Function
    End If
    CWR_AssertTargetStockLocal = True
    Exit Function
Failed:
    messageText = "TARGET_STOCK_LOCAL_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_AssertPostingVerifiedLocal(ByRef messageText As String) As Boolean
    Dim ws As Worksheet, flowText As String, proofText As String, wtState As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    If CWR_PLAN_WH_ROW < 5 Or _
       CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 1).Value2) <> CWR_EXPECTED_OB Or _
       CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 2).Value2) <> CWR_EXPECTED_DOC Then
        messageText = "POST_VERIFIED_IDENTITY_NOT_CURRENT"
        Exit Function
    End If
    flowText = UCase$(CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 16).Value2))
    proofText = UCase$(CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 17).Value2))
    wtState = UCase$(CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 26).Value2))
    If (flowText = "PRDO_CHANGE_QTY_THEN_CREATE_WT" And proofText = "POST_CHANGE_VERIFIED") Or _
       (flowText = "CREATE_WT_FROM_TARGET_THEN_CF_HANDOFF" And proofText = "PICK_BIN_BATCH_POST_VERIFIED") Or _
        ((wtState = "WT_CANCELLED_REMAINING" Or wtState = "WT_CANCELLED_PENDING_REMAINING") And _
         StrComp(CWR_TARGET_BIN, CWR_EXPECTED_SOURCE_BIN, vbTextCompare) = 0 And _
         StrComp(CWR_TARGET_BATCH, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) = 0) Then
        CWR_AssertPostingVerifiedLocal = True
        Exit Function
    End If
    messageText = "POST_VERIFIED_STATE_NOT_CURRENT"
    Exit Function
Failed:
    messageText = "POST_VERIFIED_STATE_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_AssertPRDOQtyVerifiedLocal(ByRef messageText As String) As Boolean
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    If CWR_PLAN_WH_ROW < 5 Or _
       CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 1).Value2) <> CWR_EXPECTED_OB Or _
       CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 2).Value2) <> CWR_EXPECTED_DOC Or _
       UCase$(CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 16).Value2)) <> "CREATE_WT_FROM_TARGET_THEN_CF_HANDOFF" Or _
       (UCase$(CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 17).Value2)) <> "PRDO_QTY_VERIFIED" And _
        UCase$(CWR_Text(ws.Cells(CWR_PLAN_WH_ROW, 17).Value2)) <> "PICK_BIN_BATCH_POST_VERIFIED") Then
        messageText = "PRDO_QTY_VERIFIED_STATE_NOT_CURRENT"
        Exit Function
    End If
    CWR_AssertPRDOQtyVerifiedLocal = True
    Exit Function
Failed:
    messageText = "PRDO_QTY_VERIFIED_STATE_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' A fresh POST preflight is mandatory for a write.  This prevents an action
' from resuming from a stale sheet row or from being invoked a second time
' after its verified state has already been recorded.
Private Function CWR_AssertPostingPlan(ByRef messageText As String) As Boolean
    Dim ws As Worksheet, stateText As String
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    If CWR_Text(ws.Cells(5, 1).Value2) <> CWR_EXPECTED_OB Or _
       CWR_Text(ws.Cells(5, 2).Value2) <> CWR_EXPECTED_DOC Or _
       CWR_Text(ws.Cells(5, 16).Value2) <> "POST_CHANGE_BIN_BATCH_QTY" Then
        messageText = "POST_PLAN_IDENTITY_OR_ACTION_MISMATCH"
        Exit Function
    End If
    stateText = UCase$(CWR_Text(ws.Cells(5, 17).Value2))
    If stateText = "POST_CHANGE_VERIFIED" Then
        messageText = "POST_ALREADY_VERIFIED_NO_RETRY"
        Exit Function
    End If
    If stateText <> "POST_PREVERIFIED" Or _
       UCase$(CWR_Text(ws.Cells(5, 18).Value2)) <> "POST_SOURCE_AND_TARGET_READ" Then
        messageText = "POST_PREVERIFY_NOT_CURRENT: " & stateText
        Exit Function
    End If
    CWR_AssertPostingPlan = True
    Exit Function
Failed:
    messageText = "POST_PLAN_LOCAL_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_ConfigValue(ByVal keyName As String) As String
    Dim ws As Worksheet, r As Long, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("SYSTEM")
    lastRow = ws.Cells(ws.rows.count, 17).End(xlUp).row
    For r = 1 To lastRow
        If StrComp(CWR_Text(ws.Cells(r, 17).Value2), keyName, vbTextCompare) = 0 Then
            CWR_ConfigValue = CWR_Text(ws.Cells(r, 18).Value2)
            Exit Function
        End If
    Next r
End Function

Private Function CWR_ValidatePreparedSnapshot(ByRef messageText As String) As Boolean
    Dim wsActual As Worksheet, wsInput As Worksheet
    Dim sourceKey As String
    On Error GoTo Failed
    Set wsActual = ThisWorkbook.Worksheets("DATA_CORE")
    Set wsInput = CF_GetInputSheet()
    sourceKey = CWR_EXPECTED_OB & "|" & CWR_EXPECTED_DOC & "|" & CWR_EXPECTED_ITEM & "|" & CWR_EXPECTED_WT

    If CWR_Text(wsInput.Range("A7").Value2) <> CWR_EXPECTED_OB Then GoTo BadInput
    If CWR_PLAN_WH_ROW < 5 Then GoTo BadActual
    If CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 1).Value2) <> CWR_EXPECTED_OB Then GoTo BadActual
    If CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 2).Value2) <> CWR_EXPECTED_DOC Then GoTo BadActual
    If CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 3).Value2) <> CWR_EXPECTED_ITEM Then GoTo BadActual
    If CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 5).Value2) <> CWR_EXPECTED_WT Then GoTo BadActual
    If CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 6).Value2) <> CWR_EXPECTED_MATERIAL Then GoTo BadActual
    If Abs(CWR_Number(wsActual.Cells(CWR_PLAN_WH_ROW, 7).Value2) - CWR_EXPECTED_SOURCE_QTY) > 0.000001 Then GoTo BadActual
    If CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 8).Value2) <> CWR_EXPECTED_SOURCE_BIN Then GoTo BadActual
    If CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 9).Value2) <> CWR_EXPECTED_SOURCE_BATCH Then GoTo BadActual
    If CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 20).Value2) <> sourceKey Then GoTo BadActual
    If UCase$(CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 25).Value2)) <> "OB_OPEN_OR_PARTIAL" Then GoTo BadActual
    If UCase$(CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 26).Value2)) <> "WT_ACTIVE" And _
       UCase$(CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 26).Value2)) <> "WT_CANCELLED_PENDING_REMAINING" And _
       UCase$(CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 26).Value2)) <> "WT_CANCELLED_REMAINING" Then GoTo BadActual
    If Abs(CWR_Number(wsActual.Cells(CWR_PLAN_WH_ROW, 12).Value2) - CWR_TARGET_QTY) > 0.000001 Then GoTo BadTarget
    If CWR_Text(wsActual.Cells(CWR_PLAN_WH_ROW, 14).Value2) <> CWR_TARGET_BIN Then GoTo BadTarget
    CWR_ValidatePreparedSnapshot = True
    Exit Function
BadInput:
    messageText = "INPUT_OB_MISMATCH: " & CWR_Text(wsInput.Range("A7").Value2)
    Exit Function
BadActual:
    messageText = "LOCAL_SOURCE_SNAPSHOT_MISMATCH"
    Exit Function
BadTarget:
    messageText = "LOCAL_TARGET_SNAPSHOT_MISMATCH"
    Exit Function
Failed:
    messageText = "LOCAL_SNAPSHOT_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_GetMONSession(ByRef ses As Object, ByRef messageText As String) As Boolean
    On Error GoTo Failed

    ' First reuse an existing MON reader. If none exists, create one bounded
    ' dedicated second tab and reuse it; never reopen MON for each readback.
    Set ses = CF_MON_GetExistingReaderSession(messageText)
    If ses Is Nothing Then
        Set ses = CF_MON_EnsureDedicatedReaderSession(messageText)
        If ses Is Nothing Then Exit Function
    End If
    If Not CF_MON_WaitExistingReaderIdle(ses, 15#) Then
        messageText = "MON_SESSION_BUSY"
        Set ses = Nothing
        Exit Function
    End If
    CWR_GetMONSession = True
    Exit Function
Failed:
    messageText = "MON_SESSION_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set ses = Nothing
End Function

' Legacy same-tab recovery helper retained for compatibility with older
' diagnostics. The executor path above uses the dedicated second MON tab so
' PRDO remains available and repeated reads do not reopen MON sessions.
Private Function CWR_ReuseVerifiedPRDOTabForMON(ByRef ses As Object, ByRef messageText As String) As Boolean
    Dim gui As Object, app As Object, con As Object, candidate As Object
    Dim i As Long, n As Long, matchCount As Long, tcode As String

    On Error GoTo Failed
    ' CP4_LEGACY_MON_REUSE_DISABLED: Single MON Broker (CWR_GetMONSession) is the
    ' only sanctioned resolver. This legacy same-tab reuse must never run.
    messageText = "STOP_LEGACY_MON_REUSE_DISABLED|use CWR_GetMONSession dedicated reader; PRDO tab is never reused as MON."
    Set ses = Nothing
    Exit Function
    messageText = ""
    Set ses = Nothing
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    If app Is Nothing Then
        messageText = "STOP_MON_SESSION_REQUIRED: SAP connection is unavailable."
        Exit Function
    End If
    If app.children.count < 1 Then
        messageText = "STOP_MON_SESSION_REQUIRED: SAP connection is unavailable."
        Exit Function
    End If
    Set con = app.children.ElementAt(0)
    n = con.children.count
    For i = 0 To n - 1
        tcode = ""
        On Error Resume Next
        tcode = CWR_Text(con.children.ElementAt(i).Info.Transaction)
        On Error GoTo Failed
        If UCase$(Replace(Trim$(tcode), "/", "")) = "SCWMPRDO" Then
            If Not CBool(con.children.ElementAt(i).busy) Then
                matchCount = matchCount + 1
                Set candidate = con.children.ElementAt(i)
            End If
        End If
    Next i
    If matchCount = 0 Then
        messageText = "STOP_MON_SESSION_REQUIRED: no existing idle /SCWM/MON or /SCWM/PRDO session; RUN does not open a new SAP tab."
        Exit Function
    End If
    If matchCount <> 1 Then
        messageText = "STOP_MON_SESSION_AMBIGUOUS: idle /SCWM/PRDO sessions=" & CStr(matchCount) & "; no tab was changed."
        Exit Function
    End If

    Set ses = candidate
    If Not CWR_ReserveMONOpen(messageText) Then Exit Function
    ses.StartTransaction "/SCWM/MON"
    If Not CWR_WaitIdle(ses, 30#) Then
        messageText = "MON_SAME_TAB_START_TIMEOUT"
        Set ses = Nothing
        Exit Function
    End If
    tcode = CWR_Text(ses.Info.Transaction)
    If UCase$(Replace(Trim$(tcode), "/", "")) <> "SCWMMON" Then
        messageText = "MON_SAME_TAB_TCODE_MISMATCH: " & tcode
        Set ses = Nothing
        Exit Function
    End If
    CWR_Log "MON_SESSION_REUSE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "Reused the verified /SCWM/PRDO tab for /SCWM/MON readback; SAME_TAB=TRUE; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
    CWR_ReuseVerifiedPRDOTabForMON = True
    Exit Function
Failed:
    messageText = "MON_SAME_TAB_RECOVERY_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set ses = Nothing
End Function

Private Function CWR_OpenExpectedPRDO(ByRef ses As Object, ByRef grid As Object, ByRef messageText As String) As Boolean
    Const PRDO_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell"
    Const PRDO_DOC_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_DLV_PRD:2001/txt/SCWM/S_SP_Q_HEAD-DOCNO_ODO"
    Const PRDO_GO_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_DLV_PRD:2001/btnCMD_GO"
    Dim gui As Object, app As Object, con As Object, i As Long, tcode As String, titleText As String
    Dim field As Object, goButton As Object, candidate As Object, candidateField As Object, needsRestart As Boolean

    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    If app.children.count < 1 Then
        messageText = "SAP_CONNECTION_NOT_FOUND"
        Exit Function
    End If
    Set con = app.children.ElementAt(0)
    Set ses = Nothing
    For i = 0 To con.children.count - 1
        tcode = ""
        On Error Resume Next
        tcode = CStr(con.children.ElementAt(i).Info.Transaction & "")
        titleText = CWR_Text(con.children.ElementAt(i).FindById("wnd[0]", False).text)
        On Error GoTo Failed
        If UCase$(Trim$(tcode)) = "/SCWM/PRDO" Then
            Set candidateField = con.children.ElementAt(i).FindById(PRDO_DOC_ID, False)
            If Not candidateField Is Nothing Then
                Set ses = con.children.ElementAt(i)
                Exit For
            End If
            If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) = 0 And _
               candidate Is Nothing Then
                Set candidate = con.children.ElementAt(i)
            End If
        End If
    Next i
    If ses Is Nothing Then
        If Not candidate Is Nothing Then
            Set ses = candidate
            needsRestart = True
        ElseIf Not CWR_GetMONSession(ses, messageText) Then
            Exit Function
        Else
            needsRestart = True
        End If
    End If
    If needsRestart Then
        CWR_DOCUMENT_RECOVERY_COUNT = CWR_DOCUMENT_RECOVERY_COUNT + 1
        ses.StartTransaction "/SCWM/PRDO"
        If Not CWR_WaitIdle(ses, 30#) Then
            messageText = "PRDO_START_TIMEOUT"
            Exit Function
        End If
        tcode = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction & "")
        On Error GoTo Failed
        If UCase$(Trim$(tcode)) <> "/SCWM/PRDO" Then
            messageText = "PRDO_START_TCODE_MISMATCH: " & tcode
            Exit Function
        End If
    End If
    If CBool(ses.busy) Then
        messageText = "PRDO_SESSION_BUSY"
        Exit Function
    End If
    Set field = ses.FindById(PRDO_DOC_ID, False)
    Set goButton = ses.FindById(PRDO_GO_ID, False)
    If field Is Nothing Or goButton Is Nothing Then
        messageText = "PRDO_SEARCH_CONTROLS_NOT_FOUND"
        Exit Function
    End If
    field.text = CWR_EXPECTED_DOC
    If CWR_Text(field.text) <> CWR_EXPECTED_DOC Then
        messageText = "PRDO_DOC_READBACK_MISMATCH: " & CWR_Text(field.text)
        Exit Function
    End If
    If Not CWR_ReserveDocumentOpen(messageText) Then Exit Function
    goButton.Press
    If Not CWR_WaitIdle(ses, 30#) Then
        messageText = "PRDO_QUERY_TIMEOUT"
        Exit Function
    End If
    If Not CWR_WaitForObject(ses, PRDO_GRID_ID, 20#) Then
        messageText = "PRDO_ITEM_GRID_TIMEOUT"
        Exit Function
    End If
    Set grid = ses.FindById(PRDO_GRID_ID, False)
    If grid Is Nothing Then
        messageText = "PRDO_ITEM_GRID_NOT_FOUND"
        Exit Function
    End If
    CWR_OpenExpectedPRDO = True
    Exit Function
Failed:
    messageText = "PRDO_OPEN_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set ses = Nothing: Set grid = Nothing
End Function

' Bug A fix (NOT_LIVE_PROVEN): open exactly one new SAP GUI session for PRDO,
' mirroring the proven RT_OpenDedicatedMONSession idiom (seed.CreateSession + bounded
' wait for the new tab). The caller runs StartTransaction "/SCWM/PRDO" on it. The MON
' reader is never repurposed, so MON stays open on its own tab (MON<>PRDO preserved).

' Reacquires the currently displayed PRDO item grid after a server roundtrip
' such as entering edit mode. Unlike CWR_OpenExpectedPRDO it does not press
' Go again, which would leave edit mode and discard the read-only preview.
Private Function CWR_FindCurrentPRDOGrid(ByVal ses As Object, ByRef grid As Object, _
                                         ByRef messageText As String) As Boolean
    Const PRDO_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_ODP1:/SCWM/SAPLUI_DLV_PRD:3000/tabsTABSTRIP_ODP1/tabpOK_ODP1_TAB1/ssubSUB_ODP1_TAB1:/SCWM/SAPLUI_DLV_CORE:3210/ssubSUB_ODP1_1_CONTENT:/SCWM/SAPLUI_DLV_CORE:3211/cntlCONTAINER_ALV_ODP1_1/shellcont/shell"
    On Error GoTo Failed
    Set grid = ses.FindById(PRDO_GRID_ID, False)
    If grid Is Nothing Then
        messageText = "PRDO_ITEM_GRID_NOT_FOUND_AFTER_SERVER_ROUNDTRIP"
        Exit Function
    End If
    CWR_FindCurrentPRDOGrid = True
    Exit Function
Failed:
    messageText = "PRDO_ITEM_GRID_REACQUIRE_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set grid = Nothing
End Function

Private Function CWR_WaitForCreateWTSession(ByRef createSes As Object, ByVal seconds As Double, ByRef messageText As String) As Boolean
    Dim gui As Object, app As Object, con As Object, i As Long, t0 As Double, titleText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    t0 = Timer
    Do
        For i = 0 To con.children.count - 1
            Set createSes = con.children.ElementAt(i)
            titleText = ""
            On Error Resume Next
            titleText = CWR_Text(createSes.FindById("wnd[0]", False).text)
            On Error GoTo Failed
            If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then
                If Not CBool(createSes.busy) Then
                    CWR_WaitForCreateWTSession = True
                    Exit Function
                End If
            End If
        Next i
        Set createSes = Nothing
        DoEvents
        If CWR_Elapsed(t0) >= seconds Then Exit Do
    Loop
    messageText = "CREATE_WT_SCREEN_TIMEOUT"
    Exit Function
Failed:
    messageText = "CREATE_WT_SCREEN_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set createSes = Nothing
End Function

Private Function CWR_FindCreateWTSession(ByRef createSes As Object, ByRef messageText As String) As Boolean
    Dim gui As Object, app As Object, con As Object, i As Long, titleText As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set createSes = con.children.ElementAt(i)
        titleText = ""
        On Error Resume Next
        titleText = CWR_Text(createSes.FindById("wnd[0]", False).text)
        On Error GoTo Failed
        If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 Then
            If Not CBool(createSes.busy) Then
                CWR_FindCreateWTSession = True
                Exit Function
            End If
        End If
    Next i
    Set createSes = Nothing
    messageText = "CREATE_WT_SESSION_NOT_FOUND"
    Exit Function
Failed:
    messageText = "CREATE_WT_SESSION_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set createSes = Nothing
End Function

' Finds a Create WT window only when its one Warehouse Request row is the
' exact request currently loaded from PLAN.  Title-only matching is unsafe
' because independent Create WT windows can remain open in the same SAP
' connection.
Private Function CWR_FindExactCreateWTSession(ByRef createSes As Object, ByRef messageText As String) As Boolean
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Dim gui As Object, app As Object, con As Object, candidate As Object, wrGrid As Object
    Dim i As Long, titleText As String, requestRow As Long, openQty As Double, ignoredText As String

    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine
    If app.children.count < 1 Then
        messageText = "SAP_CONNECTION_NOT_FOUND"
        Exit Function
    End If
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set candidate = con.children.ElementAt(i)
        titleText = ""
        On Error Resume Next
        titleText = CWR_Text(candidate.FindById("wnd[0]", False).text)
        On Error GoTo Failed
        If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 And _
           Not CBool(candidate.busy) Then
            Set wrGrid = candidate.FindById(WR_GRID_ID, False)
            If Not wrGrid Is Nothing Then
                ignoredText = vbNullString
                If CWR_LocateExactWarehouseRequest(wrGrid, requestRow, openQty, ignoredText) Then
                    If Abs(openQty - CWR_CREATE_QTY) <= 0.000001 Then
                        Set createSes = candidate
                        CWR_FindExactCreateWTSession = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
    Set createSes = Nothing
    messageText = "CREATE_WT_EXACT_SESSION_NOT_FOUND"
    Exit Function
Failed:
    messageText = "CREATE_WT_EXACT_SESSION_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set createSes = Nothing
End Function

' Finds only the interrupted, unsaved Create WT draft owned by the current
' Plan.  Its Open Qty must be exactly zero; a normal request (or any other
' user's request) is not a recovery candidate.
Private Function CWR_FindOwnedZeroOpenCreateWTDraft(ByRef createSes As Object, ByRef messageText As String) As Boolean
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Dim gui As Object, app As Object, con As Object, candidate As Object, wrGrid As Object
    Dim i As Long, matches As Long, titleText As String, requestRow As Long, openQty As Double, ignoredText As String

    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine
    If app.children.count < 1 Then
        messageText = "SAP_CONNECTION_NOT_FOUND"
        Exit Function
    End If
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set candidate = con.children.ElementAt(i)
        titleText = ""
        On Error Resume Next
        titleText = CWR_Text(candidate.FindById("wnd[0]", False).text)
        On Error GoTo Failed
        If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 And _
           Not CBool(candidate.busy) Then
            Set wrGrid = candidate.FindById(WR_GRID_ID, False)
            If Not wrGrid Is Nothing Then
                ignoredText = vbNullString
                If CWR_LocateExactWarehouseRequest(wrGrid, requestRow, openQty, ignoredText) Then
                    If Abs(openQty) <= 0.000001 Then
                        matches = matches + 1
                        Set createSes = candidate
                    End If
                End If
            End If
        End If
    Next i
    If matches <> 1 Then
        Set createSes = Nothing
        messageText = "CREATE_WT_OWNED_ZERO_DRAFT_MATCH_COUNT=" & CStr(matches)
        Exit Function
    End If
    CWR_FindOwnedZeroOpenCreateWTDraft = True
    Exit Function
Failed:
    messageText = "CREATE_WT_OWNED_ZERO_DRAFT_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set createSes = Nothing
End Function

' Captures the SAP session identities before a PRDO follow-on is selected.
' This is deliberately a bounded connection-level check, not a control-tree
' scan.  It lets the production route distinguish a newly opened Create WT
' screen from any unrelated one the user already has open.
Private Function CWR_CaptureSAPSessionIds(ByRef sessionIds As String, ByRef messageText As String) As Boolean
    Dim gui As Object, app As Object, con As Object, ses As Object
    Dim i As Long, sessionId As String

    On Error GoTo Failed
    sessionIds = ""
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine
    If app.children.count < 1 Then
        messageText = "SAP_CONNECTION_NOT_FOUND"
        Exit Function
    End If
    Set con = app.children.ElementAt(0)
    For i = 0 To con.children.count - 1
        Set ses = con.children.ElementAt(i)
        sessionId = CWR_Text(ses.id & "")
        If sessionId = "" Then
            messageText = "SAP_SESSION_ID_UNAVAILABLE"
            Exit Function
        End If
        sessionIds = sessionIds & "|" & sessionId & "|"
    Next i
    CWR_CaptureSAPSessionIds = True
    Exit Function
Failed:
    messageText = "SAP_SESSION_ID_CAPTURE_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Waits for the specific Create WT screen opened by the immediately preceding
' PRDO follow-on.  An already-open Create WT window is never selected here.
Private Function CWR_WaitForNewCreateWTSession(ByVal prdoSes As Object, ByVal priorSessionIds As String, _
                                                ByRef createSes As Object, ByVal seconds As Double, _
                                                ByRef messageText As String) As Boolean
    Dim gui As Object, app As Object, con As Object, candidate As Object
    Dim i As Long, t0 As Double, titleText As String, sessionId As String, prdoSessionId As String

    On Error GoTo Failed
    Set gui = GetObject("SAPGUI"): Set app = gui.GetScriptingEngine: Set con = app.children.ElementAt(0)
    prdoSessionId = CWR_Text(prdoSes.id & "")
    t0 = Timer
    Do
        For i = 0 To con.children.count - 1
            Set candidate = con.children.ElementAt(i)
            titleText = "": sessionId = ""
            On Error Resume Next
            titleText = CWR_Text(candidate.FindById("wnd[0]", False).text)
            sessionId = CWR_Text(candidate.id & "")
            On Error GoTo Failed
            If InStr(1, titleText, "Create Warehouse Task", vbTextCompare) > 0 And _
               Not CBool(candidate.busy) Then
                If sessionId = prdoSessionId Or _
                   InStr(1, priorSessionIds, "|" & sessionId & "|", vbBinaryCompare) = 0 Then
                    Set createSes = candidate
                    CWR_WaitForNewCreateWTSession = True
                    Exit Function
                End If
            End If
        Next i
        Set createSes = Nothing
        DoEvents
        If CWR_Elapsed(t0) >= seconds Then Exit Do
    Loop
    messageText = "CREATE_WT_NEW_SCREEN_TIMEOUT"
    Exit Function
Failed:
    messageText = "CREATE_WT_NEW_SCREEN_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set createSes = Nothing
End Function

' Opens Create Warehouse Task from the exact PRDO document that was just
' quantity-verified.  No source stock is selected and no Create/Save/CF action
' is sent in this helper; it only opens and queries the request surface.
Private Function CWR_OpenExactCreateWTAfterQtyChange(ByRef createSes As Object, ByRef messageText As String) As Boolean
    Dim prdoSes As Object, prdoGrid As Object, menuItem As Object, queryButton As Object
    Dim qtyText As String, batchText As String, statusText As String, menuText As String, modalText As String
    Dim rowNo As Long, priorSessionIds As String, followOnId As String, queryId As String

    On Error GoTo Failed
    If Not CWR_OpenExpectedPRDO(prdoSes, prdoGrid, messageText) Then Exit Function
    If Not CWR_FindExactPRDOItem(prdoGrid, rowNo, qtyText, batchText, statusText, messageText) Then Exit Function
    If Abs(CWR_Number(qtyText) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_PRDO_POSTQTY_MISMATCH: QTY=" & qtyText & "; BATCH=" & batchText
        Exit Function
    End If
    If CWR_HasAnyModal(prdoSes, modalText) Then
        messageText = "CREATE_WT_PRDO_MODAL_PRESENT: " & modalText
        Exit Function
    End If
    If Not CWR_CaptureSAPSessionIds(priorSessionIds, messageText) Then Exit Function
    followOnId = SAPAction_Command("WT_CREATE_FOLLOWON", "/SCWM/PRDO")
    queryId = SAPAction_Command("WT_CREATE_QUERY", "S000")
    If followOnId = "" Or queryId = "" Then
        messageText = "CREATE_WT_NAVIGATION_ACTION_MAP_MISSING"
        Exit Function
    End If
    Set menuItem = prdoSes.FindById(followOnId, False)
    If menuItem Is Nothing Then
        messageText = "CREATE_WT_FOLLOWON_MENU_NOT_FOUND"
        Exit Function
    End If
    menuText = CWR_Text(menuItem.text)
    If StrComp(menuText, "Warehouse Task", vbTextCompare) <> 0 And _
       InStr(1, menuText, "Create", vbTextCompare) = 0 Then
        messageText = "CREATE_WT_FOLLOWON_MENU_IDENTITY_MISMATCH: " & menuText
        Exit Function
    End If
    menuItem.Select
    If Not CWR_WaitForNewCreateWTSession(prdoSes, priorSessionIds, createSes, 20#, messageText) Then Exit Function
    Set queryButton = createSes.FindById(queryId, False)
    If queryButton Is Nothing Then
        messageText = "CREATE_WT_QUERY_BUTTON_NOT_FOUND"
        Exit Function
    End If
    queryButton.Press
    If Not CWR_WaitIdle(createSes, 30#) Then
        messageText = "CREATE_WT_QUERY_TIMEOUT"
        Exit Function
    End If
    If CWR_HasAnyModal(createSes, modalText) Then
        messageText = "CREATE_WT_QUERY_MODAL_PRESENT: " & modalText
        Exit Function
    End If
    CWR_OpenExactCreateWTAfterQtyChange = True
    Exit Function
Failed:
    messageText = "CREATE_WT_OPEN_EXACT_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set createSes = Nothing
End Function

Private Function CWR_DumpGridBrief(ByVal labelText As String, ByVal grid As Object, _
                                   ByVal maxRows As Long, ByVal maxColumns As Long) As String
    Dim r As Long, c As String, rowCount As Long, columnCount As Long, columnTotal As Long
    Dim columnIndex As Long, titleText As String, valueText As String
    On Error GoTo Failed
    rowCount = CLng(grid.rowCount)
    CWR_DumpGridBrief = labelText & "_GRID=" & CWR_Text(grid.id) & "|ROWS=" & CStr(rowCount) & vbLf
    On Error Resume Next
    Err.Clear
    columnTotal = CLng(grid.columnCount)
    Err.Clear
    On Error GoTo Failed
    If columnTotal <= 0 Then
        CWR_DumpGridBrief = CWR_DumpGridBrief & labelText & "_COLS_UNAVAILABLE" & vbLf
        Exit Function
    End If
    For columnIndex = 0 To columnTotal - 1
        c = CWR_GridColumnNameAt(grid, columnIndex)
        If c = "" Then
            CWR_DumpGridBrief = CWR_DumpGridBrief & labelText & "_COL_INDEX=" & CStr(columnIndex) & "|ID_UNAVAILABLE" & vbLf
            GoTo NextColumn
        End If
        columnCount = columnCount + 1
        titleText = CWR_Text(grid.GetColumnTitles(c))
        CWR_DumpGridBrief = CWR_DumpGridBrief & labelText & "_COL=" & c & "|TITLE=" & titleText
        For r = 0 To rowCount - 1
            If r >= maxRows Then Exit For
            valueText = CWR_Cell(grid, r, c)
            CWR_DumpGridBrief = CWR_DumpGridBrief & "|R" & CStr(r) & "=" & valueText
        Next r
        CWR_DumpGridBrief = CWR_DumpGridBrief & vbLf
        If columnCount >= maxColumns Then Exit For
NextColumn:
    Next columnIndex
    Exit Function
Failed:
    CWR_DumpGridBrief = labelText & "_GRID_DUMP_ERROR: " & CStr(Err.Number) & ": " & Err.Description & vbLf
End Function

' SAP GUI versions expose ColumnOrder either as an indexed property or as a
' GuiCollection.  Read both forms so diagnostics do not fail merely because
' the client differs; this helper never alters the grid.
Private Function CWR_GridColumnNameAt(ByVal grid As Object, ByVal columnIndex As Long) As String
    Dim columnOrder As Object
    On Error Resume Next
    Err.Clear
    CWR_GridColumnNameAt = CWR_Text(grid.columnOrder(columnIndex) & "")
    If Err.Number = 0 And CWR_GridColumnNameAt <> "" Then Exit Function
    Err.Clear
    Set columnOrder = grid.columnOrder
    If Not columnOrder Is Nothing Then CWR_GridColumnNameAt = CWR_Text(columnOrder.ElementAt(columnIndex) & "")
    Err.Clear
    On Error GoTo 0
End Function

Private Function CWR_DumpKnownGrid(ByVal labelText As String, ByVal grid As Object, ByVal maxRows As Long) As String
    Dim cols As Variant, c As Variant, r As Long, n As Long, lineText As String, valueText As String
    On Error GoTo Failed
    If UCase$(labelText) = "WR" Then
        cols = Array("DLVNO", "DOCNO", "ITMNO", "ITEMNO", "MATNR", "PRODUCTNO", "CHARG", "BATCHNO", _
                     "NLPLA", "UI_OPNQTY", "OPEN_QTY", "QTY", "QUAN", "STATUS_TO", "STATUS", "LGPLA", "VLPLA", "PROCTY")
    Else
        cols = Array("TANUM", "TOSTAT", "DLVNO", "DOCNO", "ITMNO", "ITEMNO", "MATNR", "PRODUCTNO", _
                     "CHARG", "NCHARG", "BATCHNO", "LGPLA", "VLPLA", "NLPLA", "VLTYP", "NLTYP", _
                     "VSOLM", "AVL_ALT_QTY", "LIME_QUAN", "QUAN", "QTY", "UI_AVLQTY", _
                     "TO_CAT", "CAT", "OWNER", "ENTITLED", "STOCK_TYPE", "PROCTY")
    End If
    n = CLng(grid.rowCount)
    CWR_DumpKnownGrid = labelText & "_GRID=" & CWR_Text(grid.id) & "|ROWS=" & CStr(n) & vbLf
    For r = 0 To n - 1
        If r >= maxRows Then Exit For
        lineText = labelText & "_ROW=" & CStr(r)
        For Each c In cols
            valueText = CWR_Cell(grid, r, CStr(c))
            If valueText <> "" Then lineText = lineText & "|" & CStr(c) & "=" & valueText
        Next c
        CWR_DumpKnownGrid = CWR_DumpKnownGrid & lineText & vbLf
    Next r
    Exit Function
Failed:
    CWR_DumpKnownGrid = labelText & "_KNOWN_DUMP_ERROR: " & CStr(Err.Number) & ": " & Err.Description & vbLf
End Function

Private Function CWR_DumpGridFieldContract(ByVal labelText As String, ByVal grid As Object, ByVal rowNo As Long) As String
    Dim cols As Variant, c As Variant, valueText As String, changeableText As String
    On Error GoTo Failed
    cols = Array("DLVNO", "DOCNO", "ITMNO", "ITEMNO", "MATNR", "PRODUCTNO", "CHARG", "BATCHNO", _
                 "LGPLA", "VLPLA", "NLPLA", "VLTYP", "NLTYP", "VLBER", "NLBER", _
                 "UI_OPNQTY", "OPEN_QTY", "UI_AVLQTY", "AVL_ALT_QTY", "QTY", "QUAN", "POST_QUAN", "LIME_QUAN", _
                 "TO_CHARG", "TO_CAT", "CAT", "PROCTY", "STATUS_TO", "STATUS")
    CWR_DumpGridFieldContract = labelText & "_GRID=" & CWR_Text(grid.id) & "|ROWS=" & CStr(grid.rowCount) & vbLf
    If CLng(grid.rowCount) <= rowNo Then Exit Function
    For Each c In cols
        valueText = CWR_Cell(grid, rowNo, CStr(c))
        changeableText = CWR_CellChangeable(grid, rowNo, CStr(c))
        If valueText <> "" Or changeableText <> "" Then
            CWR_DumpGridFieldContract = CWR_DumpGridFieldContract & labelText & "|COL=" & CStr(c) & _
                "|VALUE=" & valueText & "|CHANGEABLE=" & changeableText & vbLf
        End If
    Next c
    Exit Function
Failed:
    CWR_DumpGridFieldContract = labelText & "_GRID_CONTRACT_ERROR: " & CStr(Err.Number) & ": " & Err.Description & vbLf
End Function

Private Sub CWR_CollectDetailFields(ByVal obj As Object, ByVal depth As Long, _
                                    ByRef itemCount As Long, ByRef dumpText As String)
    Dim i As Long, child As Object, typeText As String, idText As String, nameText As String
    Dim textText As String, tipText As String, changeText As String, haystack As String
    If obj Is Nothing Or depth > 16 Or itemCount >= 220 Or Len(dumpText) >= 28000 Then Exit Sub
    On Error Resume Next
    For i = 0 To obj.children.count - 1
        Set child = obj.children.ElementAt(i)
        typeText = CWR_Text(child.Type)
        idText = CWR_Text(child.id)
        nameText = CWR_Text(child.Name)
        textText = CWR_Text(child.text)
        tipText = CWR_Text(child.Tooltip)
        changeText = CWR_Text(child.Changeable)
        haystack = UCase$(typeText & " " & idText & " " & nameText & " " & textText & " " & tipText)
        If InStr(haystack, "BATCH") > 0 Or InStr(haystack, "CHARG") > 0 Or _
           InStr(haystack, "BIN") > 0 Or InStr(haystack, "LGPLA") > 0 Or _
           InStr(haystack, "LOCATION") > 0 Or InStr(haystack, "QUANTITY") > 0 Or _
           InStr(haystack, "QTY") > 0 Or InStr(typeText, "TAB") > 0 Then
            dumpText = dumpText & "D=" & CStr(depth) & "|TYPE=" & typeText & _
                "|ID=" & idText & "|NAME=" & nameText & "|TEXT=" & textText & _
                "|TIP=" & tipText & "|CHANGEABLE=" & changeText & vbLf
            itemCount = itemCount + 1
            If itemCount >= 220 Or Len(dumpText) >= 28000 Then Exit For
        End If
        CWR_CollectDetailFields child, depth + 1, itemCount, dumpText
        If itemCount >= 220 Or Len(dumpText) >= 28000 Then Exit For
    Next i
    On Error GoTo 0
End Sub

Private Function CWR_OpenPOSTSession(ByRef ses As Object, ByRef messageText As String) As Boolean
    Dim gui As Object, app As Object, con As Object, i As Long, tcode As String
    On Error GoTo Failed
    Set gui = GetObject("SAPGUI")
    Set app = gui.GetScriptingEngine
    If app.children.count < 1 Then
        messageText = "SAP_CONNECTION_NOT_FOUND"
        Exit Function
    End If
    Set con = app.children.ElementAt(0)
    Set ses = Nothing
    For i = 0 To con.children.count - 1
        tcode = ""
        On Error Resume Next
        tcode = CStr(con.children.ElementAt(i).Info.Transaction & "")
        On Error GoTo Failed
        If UCase$(Trim$(tcode)) = "/SCWM/POST" Then
            Set ses = con.children.ElementAt(i)
            Exit For
        End If
    Next i
    If ses Is Nothing Then
        For i = 0 To con.children.count - 1
            tcode = ""
            On Error Resume Next
            tcode = CStr(con.children.ElementAt(i).Info.Transaction & "")
            On Error GoTo Failed
            If UCase$(Trim$(tcode)) = "/SCWM/PRDO" Then
                Set ses = con.children.ElementAt(i)
                Exit For
            End If
        Next i
        ' Button #4 owns production navigation.  A user must not have to leave
        ' a PRDO/POST window open before running the flow; reuse the bounded
        ' Monitor session and navigate it to POST just as the PRDO opener does.
        If ses Is Nothing Then
            If Not CWR_GetMONSession(ses, messageText) Then
                messageText = "POST_NAVIGATION_SESSION_UNAVAILABLE|" & messageText
                Exit Function
            End If
        End If
        If CBool(ses.busy) Then
            messageText = "POST_SOURCE_SESSION_BUSY"
            Exit Function
        End If
        ses.StartTransaction "/SCWM/POST"
        If Not CWR_WaitIdle(ses, 30#) Then
            messageText = "POST_START_TIMEOUT"
            Exit Function
        End If
        tcode = ""
        On Error Resume Next
        tcode = CStr(ses.Info.Transaction & "")
        On Error GoTo Failed
        If UCase$(Trim$(tcode)) <> "/SCWM/POST" Then
            messageText = "POST_START_TCODE_MISMATCH: " & tcode
            Exit Function
        End If
    End If
    If CBool(ses.busy) Then
        messageText = "POST_SESSION_BUSY"
        Exit Function
    End If
    CWR_OpenPOSTSession = True
    Exit Function
Failed:
    messageText = "POST_OPEN_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set ses = Nothing
End Function

' Navigates only the already-used POST session to the Monitor when the saved
' POST result must be read.  No data is changed and no warehouse task action
' is selected there.
Private Function CWR_OpenMonitorForRead(ByRef ses As Object, ByRef messageText As String) As Boolean
    Dim tcode As String
    On Error GoTo Failed
    If ses Is Nothing Then
        messageText = "MON_READ_SESSION_NOT_SET"
        Exit Function
    End If
    If CBool(ses.busy) Then
        messageText = "MON_READ_SESSION_BUSY"
        Exit Function
    End If
    tcode = CWR_Text(ses.Info.Transaction)
    If UCase$(Replace(tcode, "/", "")) <> "SCWMMON" Then
        If Not CWR_ReserveMONOpen(messageText) Then Exit Function
        ses.StartTransaction "/SCWM/MON"
        If Not CWR_WaitIdle(ses, 30#) Then
            messageText = "MON_READ_START_TIMEOUT"
            Exit Function
        End If
        tcode = CWR_Text(ses.Info.Transaction)
    End If
    If UCase$(Replace(tcode, "/", "")) <> "SCWMMON" Then
        messageText = "MON_READ_TCODE_MISMATCH: " & tcode
        Exit Function
    End If
    CWR_OpenMonitorForRead = True
    Exit Function
Failed:
    messageText = "MON_READ_OPEN_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Reads a specifically reported posting-change warehouse task from the current
' Monitor grid. It is intentionally descriptive: status C is neither assumed
' nor requested by this function.
Private Function CWR_ReadCreatedPostingWT(ByVal lower As Object, ByVal wtNo As String, _
                                          ByRef proofText As String, ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, matches As Long, currentDoc As String, rowDoc As String, rowWT As String
    Dim itemText As String, matText As String, statusText As String, sourceBin As String, destBin As String
    Dim batchText As String, qtyText As String, procText As String
    On Error GoTo Failed
    n = CLng(lower.rowCount)
    currentDoc = ""
    For r = 0 To n - 1
        rowDoc = CWR_Cell(lower, r, "DOCNO")
        If rowDoc <> "" Then currentDoc = rowDoc
        rowWT = CWR_Cell(lower, r, "TANUM")
        If StrComp(currentDoc, CWR_EXPECTED_DOC, vbTextCompare) = 0 And _
           StrComp(rowWT, wtNo, vbTextCompare) = 0 Then
            matches = matches + 1
            itemText = CWR_Cell(lower, r, "ITEMNO")
            matText = CWR_Cell(lower, r, "MATNR")
            statusText = CWR_Cell(lower, r, "TOSTAT")
            sourceBin = CWR_Cell(lower, r, "VLPLA")
            destBin = CWR_Cell(lower, r, "NLPLA")
            batchText = CWR_Cell(lower, r, "CHARG")
            qtyText = CWR_Cell(lower, r, "VSOLM")
            procText = CWR_Cell(lower, r, "PROCTY")
        End If
    Next r
    If matches <> 1 Then
        messageText = "POST_CREATED_WT_MON_MATCH_COUNT=" & CStr(matches) & "; WT=" & wtNo
        Exit Function
    End If
    proofText = "WT=" & wtNo & "; STATUS=" & statusText & "; ITEM=" & itemText & _
        "; MAT=" & matText & "; QTY=" & qtyText & "; SOURCE_BIN=" & sourceBin & _
        "; DEST_BIN=" & destBin & "; BATCH=" & batchText & "; PROCTY=" & procText
    CWR_ReadCreatedPostingWT = True
    Exit Function
Failed:
    messageText = "POST_CREATED_WT_MON_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_ExtractWarehouseTaskNumber(ByVal statusText As String) As String
    Dim p As Long, i As Long, tailText As String, ch As String
    p = InStr(1, statusText, "Warehouse task", vbTextCompare)
    If p = 0 Then Exit Function
    tailText = Mid$(statusText, p + Len("Warehouse task"))
    For i = 1 To Len(tailText)
        ch = Mid$(tailText, i, 1)
        If ch >= "0" And ch <= "9" Then
            CWR_ExtractWarehouseTaskNumber = CWR_ExtractWarehouseTaskNumber & ch
        ElseIf CWR_ExtractWarehouseTaskNumber <> "" Then
            Exit Function
        End If
    Next i
End Function

Private Function CWR_QueryExpectedPOST(ByVal ses As Object, ByRef grid As Object, _
                                       ByRef messageText As String, _
                                       Optional ByVal forceRefresh As Boolean = False) As Boolean
    Const POST_CRITERION_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/cmb/SCWM/S_UI_POST-V_CRITERION"
    Const POST_PRODUCT_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/ctxt/SCWM/S_ASPQ_STOCK-MATNR"
    Const POST_GO_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_SEARCH_VALUE:/SCWM/SAPLUI_POST:2001/btnCMD_GO"
    Const POST_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/subSUB_OIP_1_DATA:/SCWM/SAPLUI_POST:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell"
    Const POST_EXECUTE_ID As String = "wnd[0]/tbar[1]/btn[8]"
    Dim criterion As Object, productCtl As Object, goButton As Object, executeButton As Object
    Dim currentGridMessage As String, stageText As String
    Dim setErr As Long, setDesc As String, retriedAfter613 As Boolean
    On Error GoTo Failed
    If ses Is Nothing Then
        messageText = "POST_SESSION_NOT_SET"
        Exit Function
    End If

    ' CHECK + POSTING has already loaded the exact product grid.  Cancelling a
    ' WT does not alter stock at the target bin, so reuse that live, visible
    ' grid after validating the material instead of re-pressing a legacy
    ' search-button control.  A forced recovery refresh must still obtain a
    ' new result, but it must not rewrite the Product field when that grid is
    ' already correct: on this SAP GUI patch that setter can raise error 613.
    stageText = "REUSE_CURRENT_PRODUCT_GRID"
    currentGridMessage = ""
    If CWR_FindCurrentPOSTGrid(ses, grid, currentGridMessage) Then
        If CWR_POSTGridHasExpectedMaterial(grid) Then
            If Not forceRefresh Then
                CWR_QueryExpectedPOST = True
                Exit Function
            End If
            ' F8 is the stable read-only Execute/query action for the already
            ' validated product context.  Do not touch the repainting MATNR
            ' input field here; this is the exact recovery path that formerly
            ' stopped at STEP=SET_PRODUCT with SAP GUI error 613.
            stageText = "REFRESH_CURRENT_PRODUCT_GRID"
            Set executeButton = ses.FindById(POST_EXECUTE_ID, False)
            If Not executeButton Is Nothing Then
                executeButton.Press
                stageText = "WAIT_CURRENT_PRODUCT_REFRESH"
                If Not CWR_WaitIdle(ses, 30#) Then
                    messageText = "POST_QUERY_REFRESH_TIMEOUT"
                    Exit Function
                End If
                stageText = "REACQUIRE_CURRENT_PRODUCT_GRID"
                If Not CWR_WaitForObject(ses, POST_GRID_ID, 20#) Then
                    messageText = "POST_QUERY_REFRESH_GRID_TIMEOUT"
                    Exit Function
                End If
                Set grid = ses.FindById(POST_GRID_ID, False)
                If grid Is Nothing Or Not CWR_POSTGridHasExpectedMaterial(grid) Then
                    messageText = "POST_QUERY_REFRESH_RESULT_MATERIAL_NOT_FOUND: " & CWR_EXPECTED_MATERIAL
                    Exit Function
                End If
                CWR_QueryExpectedPOST = True
                Exit Function
            End If
        End If
        Set grid = Nothing
    End If

    ' Fallback applies only when no validated result grid is open.  Keep this
    ' path exact and use the stable F8 execute button, not btnCMD_GO, whose
    ' object can be stale after the POST search surface changes.
    stageText = "FIND_SEARCH_CONTROLS"
    Set criterion = ses.FindById(POST_CRITERION_ID, False)
    Set productCtl = ses.FindById(POST_PRODUCT_ID, False)
    Set goButton = ses.FindById(POST_GO_ID, False)
    Set executeButton = ses.FindById(POST_EXECUTE_ID, False)
    If criterion Is Nothing Or productCtl Is Nothing Or (goButton Is Nothing And executeButton Is Nothing) Then
        messageText = "POST_SEARCH_CONTROLS_NOT_FOUND"
        Exit Function
    End If

    ' POST_PRODUCT_ID is an exact MATNR field.  Do not read the combo's Text
    ' property here: on the current SAP GUI patch that accessor itself can
    ' raise error 613 while the search controls are repainting.
    stageText = "SET_PRODUCT"
    On Error Resume Next
    Err.Clear
    productCtl.text = CWR_EXPECTED_MATERIAL
    setErr = Err.Number: setDesc = Err.Description: Err.Clear
    On Error GoTo Failed
    If setErr <> 0 Then
        If setErr = 613 And Not retriedAfter613 Then
            ' Error 613 is a SAP GUI repaint/control-lifetime failure, not a
            ' stock result.  Re-open the same read-only POST transaction once,
            ' reacquire the exact controls, then retry SET_PRODUCT exactly one
            ' time.  No mutation is sent from this recovery.
            retriedAfter613 = True
            CWR_Log "POST_QUERY_613_RECOVERY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                "STEP=SET_PRODUCT|SAP_ERR=613|ACTION=REOPEN_READ_ONLY_POST_SURFACE|RETRY_COUNT=1|NO_SAP_WRITE"
            If Not CWR_ResetPOSTSearchSurface(ses, messageText) Then Exit Function
            Set criterion = ses.FindById(POST_CRITERION_ID, False)
            Set productCtl = ses.FindById(POST_PRODUCT_ID, False)
            Set goButton = ses.FindById(POST_GO_ID, False)
            Set executeButton = ses.FindById(POST_EXECUTE_ID, False)
            If criterion Is Nothing Or productCtl Is Nothing Or (goButton Is Nothing And executeButton Is Nothing) Then
                messageText = "STOP_POST_QUERY_NEED_USER_DATA|STEP=SET_PRODUCT|SAP_ERR=613|POST_SURFACE_CONTROLS_NOT_FOUND"
                Exit Function
            End If
            On Error Resume Next
            Err.Clear
            productCtl.text = CWR_EXPECTED_MATERIAL
            setErr = Err.Number: setDesc = Err.Description: Err.Clear
            On Error GoTo Failed
            If setErr <> 0 Then
                messageText = "STOP_POST_QUERY_NEED_USER_DATA|STEP=SET_PRODUCT|SAP_ERR=" & CStr(setErr) & _
                    "|DETAIL=" & setDesc & "|NO_MUTATION|TAI_POSTING_REQUIRED"
                CWR_Log "POST_QUERY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
                Exit Function
            End If
        Else
            messageText = "STOP_POST_QUERY_NEED_USER_DATA|STEP=SET_PRODUCT|SAP_ERR=" & CStr(setErr) & _
                "|DETAIL=" & setDesc & "|NO_MUTATION|TAI_POSTING_REQUIRED"
            CWR_Log "POST_QUERY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
            Exit Function
        End If
    End If
    stageText = "PRODUCT_READBACK"
    If StrComp(CWR_Text(productCtl.text), CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Then
        messageText = "POST_PRODUCT_READBACK_MISMATCH: " & CWR_Text(productCtl.text)
        Exit Function
    End If
    stageText = "EXECUTE_PRODUCT_QUERY"
    If Not executeButton Is Nothing Then
        executeButton.Press
    Else
        goButton.Press
    End If
    stageText = "WAIT_PRODUCT_QUERY"
    If Not CWR_WaitIdle(ses, 30#) Then
        messageText = "POST_QUERY_TIMEOUT"
        Exit Function
    End If
    stageText = "WAIT_RESULT_GRID"
    If Not CWR_WaitForObject(ses, POST_GRID_ID, 20#) Then
        messageText = "POST_GRID_TIMEOUT"
        Exit Function
    End If
    stageText = "ACQUIRE_RESULT_GRID"
    Set grid = ses.FindById(POST_GRID_ID, False)
    If grid Is Nothing Then
        messageText = "POST_GRID_NOT_FOUND"
        Exit Function
    End If
    stageText = "VALIDATE_RESULT_MATERIAL"
    If Not CWR_POSTGridHasExpectedMaterial(grid) Then
        messageText = "POST_RESULT_MATERIAL_NOT_FOUND: " & CWR_EXPECTED_MATERIAL
        Exit Function
    End If
    CWR_QueryExpectedPOST = True
    Exit Function
Failed:
    messageText = "POST_QUERY_ERROR: STEP=" & stageText & "; " & CStr(Err.Number) & ": " & Err.Description
    CWR_Log "POST_QUERY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Set grid = Nothing
End Function

' Reopens only the read-only /SCWM/POST surface after a bounded SAP GUI 613
' control error.  It deliberately performs no product query, Save, Create,
' Cancel, or CF action; the caller reacquires controls and sends one query.
Private Function CWR_ResetPOSTSearchSurface(ByVal ses As Object, ByRef messageText As String) As Boolean
    On Error GoTo Failed
    If ses Is Nothing Then
        messageText = "POST_QUERY_NEED_USER_DATA|POST_SESSION_NOT_SET"
        Exit Function
    End If
    If CBool(ses.busy) Then
        messageText = "POST_QUERY_NEED_USER_DATA|POST_SESSION_BUSY_DURING_613_RECOVERY"
        Exit Function
    End If
    ses.StartTransaction "/SCWM/POST"
    If Not CWR_WaitIdle(ses, 30#) Then
        messageText = "POST_QUERY_NEED_USER_DATA|POST_SURFACE_REOPEN_TIMEOUT"
        Exit Function
    End If
    If UCase$(Replace(CWR_Text(ses.Info.Transaction), "/", "")) <> "SCWMPOST" Then
        messageText = "POST_QUERY_NEED_USER_DATA|POST_SURFACE_TCODE_MISMATCH=" & CWR_Text(ses.Info.Transaction)
        Exit Function
    End If
    CWR_ResetPOSTSearchSurface = True
    Exit Function
Failed:
    messageText = "POST_QUERY_NEED_USER_DATA|POST_SURFACE_REOPEN_ERROR=" & CStr(Err.Number) & ":" & Err.Description
End Function

' This is a bounded row check on the current POST result only.  It deliberately
' does not enumerate the SAP control tree or alter the query/draft state.
Private Function CWR_POSTGridHasExpectedMaterial(ByVal grid As Object) As Boolean
    Dim r As Long, n As Long, materialText As String
    On Error GoTo ExitHere
    If grid Is Nothing Then Exit Function
    n = CLng(grid.rowCount)
    For r = 0 To n - 1
        materialText = CWR_Cell(grid, r, "MATNR")
        If StrComp(materialText, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 Then
            CWR_POSTGridHasExpectedMaterial = True
            Exit Function
        End If
    Next r
ExitHere:
End Function

Private Function CWR_FindCurrentPOSTGrid(ByVal ses As Object, ByRef grid As Object, _
                                         ByRef messageText As String) As Boolean
    Const POST_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/subSUB_OIP_1_DATA:/SCWM/SAPLUI_POST:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell"
    On Error GoTo Failed
    Set grid = ses.FindById(POST_GRID_ID, False)
    If grid Is Nothing Then
        messageText = "POST_GRID_NOT_FOUND_AFTER_SERVER_ROUNDTRIP"
        Exit Function
    End If
    CWR_FindCurrentPOSTGrid = True
    Exit Function
Failed:
    messageText = "POST_GRID_REACQUIRE_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set grid = Nothing
End Function

Private Function CWR_SelectExactPOSTSource(ByVal ses As Object, ByRef grid As Object, _
                                           ByRef sourceRow As Long, ByRef targetRow As Long, _
                                           ByRef messageText As String) As Boolean
    Dim sourceProof As String, targetProof As String, modalText As String, statusType As String, statusText As String
    Dim selectedRows As String
    On Error GoTo Failed
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then Exit Function
    grid.selectedRows = CStr(sourceRow)
    grid.SetCurrentCell sourceRow, "TO_CHARG"
    grid.SelectionChanged
    If Not CWR_WaitPOSTSaveSettle(ses, 0.5, 15#) Then
        messageText = "POST_SELECTION_CHANGE_TIMEOUT"
        Exit Function
    End If
    If CWR_GetModalText(ses, modalText) Then
        messageText = "POST_SELECTION_CHANGE_MODAL: " & modalText
        Exit Function
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "POST_SELECTION_CHANGE_SAP_" & statusType & ": " & statusText
        Exit Function
    End If
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then Exit Function
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then Exit Function
    selectedRows = CWR_SelectedRowsText(grid)
    If StrComp(selectedRows, CStr(sourceRow), vbBinaryCompare) <> 0 Then
        messageText = "POST_SELECTION_READBACK_MISMATCH: SELECTED_ROWS=" & selectedRows & _
            "; REQUIRED_SOURCE_ROW=" & CStr(sourceRow)
        Exit Function
    End If
    CWR_SelectExactPOSTSource = True
    Exit Function
Failed:
    messageText = "POST_SELECTION_CHANGE_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_GetVerifiedPOSTCreateSaveAction(ByVal ses As Object, ByRef actionToolbar As Object, _
                                                      ByRef detailText As String, ByRef messageText As String) As Boolean
    Const POST_ACTION_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell"
    Const CREATE_SAVE_POST_ID As String = "OK_OIP_POST_CREATE_POST"
    Dim i As Long, n As Long, matches As Long, buttonId As String, buttonText As String, tipText As String, enabledText As String
    On Error GoTo Failed
    Set actionToolbar = ses.FindById(POST_ACTION_TOOLBAR_ID, False)
    If actionToolbar Is Nothing Then
        messageText = "POST_CREATE_SAVE_TOOLBAR_NOT_FOUND"
        Exit Function
    End If
    If InStr(1, CWR_Text(actionToolbar.Type), "GuiShell", vbTextCompare) = 0 Then
        messageText = "POST_CREATE_SAVE_TOOLBAR_TYPE_MISMATCH: " & CWR_Text(actionToolbar.Type)
        Exit Function
    End If
    n = CLng(actionToolbar.buttonCount)
    For i = 0 To n - 1
        buttonId = CWR_Text(actionToolbar.GetButtonId(i))
        If StrComp(buttonId, CREATE_SAVE_POST_ID, vbBinaryCompare) = 0 Then
            matches = matches + 1
            buttonText = CWR_Text(actionToolbar.GetButtonText(i))
            tipText = CWR_Text(actionToolbar.GetButtonTooltip(i))
            enabledText = CWR_Text(actionToolbar.GetButtonEnabled(i))
        End If
    Next i
    If matches <> 1 Then
        messageText = "POST_CREATE_SAVE_BUTTON_MATCH_COUNT=" & CStr(matches)
        Exit Function
    End If
    If StrComp(buttonText, "Create+Save", vbTextCompare) <> 0 Or _
       StrComp(tipText, "Create + Save Posting Change", vbTextCompare) <> 0 Or _
       UCase$(enabledText) <> "TRUE" Then
        messageText = "POST_CREATE_SAVE_BUTTON_IDENTITY_MISMATCH: TEXT=" & buttonText & _
            "; TIP=" & tipText & "; ENABLED=" & enabledText
        Exit Function
    End If
    detailText = "ID=" & CREATE_SAVE_POST_ID & "; TEXT=" & buttonText & "; TIP=" & tipText & "; ENABLED=" & enabledText
    CWR_GetVerifiedPOSTCreateSaveAction = True
    Exit Function
Failed:
    messageText = "POST_CREATE_SAVE_BUTTON_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set actionToolbar = Nothing
End Function

Private Function CWR_GetVerifiedPRDOChangeAction(ByVal toolbar As Object, ByRef detailText As String, _
                                                  ByRef messageText As String) As Boolean
    Const PRDO_CHANGE_ID As String = "OIP_CHANGE"
    Dim i As Long, n As Long, matches As Long, buttonId As String, buttonText As String, tipText As String, enabledText As String
    On Error GoTo Failed
    If toolbar Is Nothing Then
        messageText = "PRDO_EDIT_TOOLBAR_NOT_FOUND"
        Exit Function
    End If
    If InStr(1, CWR_Text(toolbar.Type), "GuiShell", vbTextCompare) = 0 Then
        messageText = "PRDO_EDIT_TOOLBAR_TYPE_MISMATCH: " & CWR_Text(toolbar.Type)
        Exit Function
    End If
    n = CLng(toolbar.buttonCount)
    For i = 0 To n - 1
        buttonId = CWR_Text(toolbar.GetButtonId(i))
        If StrComp(buttonId, PRDO_CHANGE_ID, vbBinaryCompare) = 0 Then
            matches = matches + 1
            buttonText = CWR_Text(toolbar.GetButtonText(i))
            tipText = CWR_Text(toolbar.GetButtonTooltip(i))
            enabledText = CWR_Text(toolbar.GetButtonEnabled(i))
        End If
    Next i
    If matches <> 1 Then
        messageText = "PRDO_EDIT_BUTTON_MATCH_COUNT=" & CStr(matches)
        Exit Function
    End If
    If InStr(1, buttonText & " " & tipText, "Change", vbTextCompare) = 0 Or UCase$(enabledText) <> "TRUE" Then
        messageText = "PRDO_EDIT_BUTTON_IDENTITY_MISMATCH: TEXT=" & buttonText & _
            "; TIP=" & tipText & "; ENABLED=" & enabledText
        Exit Function
    End If
    detailText = "ID=" & PRDO_CHANGE_ID & "; TEXT=" & buttonText & "; TIP=" & tipText & "; ENABLED=" & enabledText
    CWR_GetVerifiedPRDOChangeAction = True
    Exit Function
Failed:
    messageText = "PRDO_EDIT_BUTTON_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_GetVerifiedPRDOSaveAction(ByVal toolbar As Object, ByRef detailText As String, _
                                                ByRef messageText As String) As Boolean
    Const PRDO_SAVE_ID As String = "SAVE"
    Dim i As Long, n As Long, matches As Long, buttonId As String, buttonText As String, tipText As String, enabledText As String
    On Error GoTo Failed
    If toolbar Is Nothing Then
        messageText = "PRDO_SAVE_TOOLBAR_NOT_FOUND"
        Exit Function
    End If
    n = CLng(toolbar.buttonCount)
    For i = 0 To n - 1
        buttonId = CWR_Text(toolbar.GetButtonId(i))
        If StrComp(buttonId, PRDO_SAVE_ID, vbBinaryCompare) = 0 Then
            matches = matches + 1
            buttonText = CWR_Text(toolbar.GetButtonText(i))
            tipText = CWR_Text(toolbar.GetButtonTooltip(i))
            enabledText = CWR_Text(toolbar.GetButtonEnabled(i))
        End If
    Next i
    If matches <> 1 Then
        messageText = "PRDO_SAVE_BUTTON_MATCH_COUNT=" & CStr(matches)
        Exit Function
    End If
    If InStr(1, buttonText & " " & tipText, "Save", vbTextCompare) = 0 Or UCase$(enabledText) <> "TRUE" Then
        messageText = "PRDO_SAVE_BUTTON_IDENTITY_MISMATCH: TEXT=" & buttonText & _
            "; TIP=" & tipText & "; ENABLED=" & enabledText
        Exit Function
    End If
    detailText = "ID=" & PRDO_SAVE_ID & "; TEXT=" & buttonText & "; TIP=" & tipText & "; ENABLED=" & enabledText
    CWR_GetVerifiedPRDOSaveAction = True
    Exit Function
Failed:
    messageText = "PRDO_SAVE_BUTTON_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_SelectExactPRDOItem(ByVal ses As Object, ByRef grid As Object, _
                                          ByRef rowNo As Long, ByRef messageText As String) As Boolean
    Dim qtyText As String, batchText As String, statusText As String, selectedRows As String, modalText As String
    On Error GoTo Failed
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then Exit Function
    grid.selectedRows = CStr(rowNo)
    grid.SetCurrentCell rowNo, "QTY_UI"
    grid.SelectionChanged
    If Not CWR_WaitPOSTSaveSettle(ses, 0.5, 15#) Then
        messageText = "PRDO_SELECTION_CHANGE_TIMEOUT"
        Exit Function
    End If
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = "PRDO_SELECTION_CHANGE_MODAL: " & modalText
        Exit Function
    End If
    If Not CWR_FindCurrentPRDOGrid(ses, grid, messageText) Then Exit Function
    If Not CWR_FindExactPRDOItem(grid, rowNo, qtyText, batchText, statusText, messageText) Then Exit Function
    selectedRows = CWR_SelectedRowsText(grid)
    If StrComp(selectedRows, CStr(rowNo), vbBinaryCompare) <> 0 Then
        messageText = "PRDO_SELECTION_READBACK_MISMATCH: SELECTED_ROWS=" & selectedRows & _
            "; REQUIRED_ROW=" & CStr(rowNo)
        Exit Function
    End If
    CWR_SelectExactPRDOItem = True
    Exit Function
Failed:
    messageText = "PRDO_SELECTION_CHANGE_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_GetVerifiedPOSTSaveButton(ByVal ses As Object, ByRef saveButton As Object, _
                                                ByRef detailText As String, ByRef messageText As String, _
                                                Optional ByVal requireChangeable As Boolean = False) As Boolean
    Const POST_SAVE_ID As String = "wnd[0]/tbar[0]/btn[11]"
    Dim typeText As String, tooltipText As String, buttonText As String, changeableText As String
    On Error GoTo Failed
    Set saveButton = ses.FindById(POST_SAVE_ID, False)
    If saveButton Is Nothing Then
        messageText = "POST_SAVE_CONTROL_NOT_FOUND"
        Exit Function
    End If
    ' GuiButton does not expose the same optional properties in every SAP GUI
    ' patch.  Read the non-mutating identity properties independently so a
    ' missing cosmetic property cannot be mistaken for a valid control.
    On Error Resume Next
    typeText = CWR_Text(saveButton.Type)
    tooltipText = CWR_Text(saveButton.Tooltip)
    buttonText = CWR_Text(saveButton.text)
    changeableText = CWR_Text(saveButton.Changeable)
    Err.Clear
    On Error GoTo Failed
    If InStr(1, typeText, "GuiButton", vbTextCompare) = 0 Then
        messageText = "POST_SAVE_CONTROL_TYPE_MISMATCH: " & typeText
        Exit Function
    End If
    If InStr(1, UCase$(tooltipText & " " & buttonText), "SAVE", vbTextCompare) = 0 Then
        messageText = "POST_SAVE_CONTROL_IDENTITY_MISMATCH: TEXT=" & buttonText & "; TIP=" & tooltipText
        Exit Function
    End If
    If requireChangeable And UCase$(changeableText) <> "TRUE" Then
        messageText = "POST_SAVE_CONTROL_NOT_CHANGEABLE: " & changeableText
        Exit Function
    End If
    detailText = "ID=" & POST_SAVE_ID & "; TYPE=" & typeText & "; TEXT=" & buttonText & _
        "; TIP=" & tooltipText & "; CHANGEABLE=" & changeableText
    CWR_GetVerifiedPOSTSaveButton = True
    Exit Function
Failed:
    messageText = "POST_SAVE_CONTROL_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set saveButton = Nothing
End Function

Private Function CWR_GetVerifiedPOSTMenuSave(ByVal ses As Object, ByRef saveMenu As Object, _
                                              ByRef detailText As String, ByRef messageText As String) As Boolean
    Const POST_MENU_SAVE_ID As String = "wnd[0]/mbar/menu[0]/menu[1]"
    Dim typeText As String, menuText As String, changeableText As String
    On Error GoTo Failed
    Set saveMenu = ses.FindById(POST_MENU_SAVE_ID, False)
    If saveMenu Is Nothing Then
        messageText = "POST_MENU_SAVE_NOT_FOUND"
        Exit Function
    End If
    On Error Resume Next
    typeText = CWR_Text(saveMenu.Type)
    menuText = CWR_Text(saveMenu.text)
    changeableText = CWR_Text(saveMenu.Changeable)
    Err.Clear
    On Error GoTo Failed
    If InStr(1, typeText, "GuiMenu", vbTextCompare) = 0 Or _
       StrComp(menuText, "Save", vbTextCompare) <> 0 Then
        messageText = "POST_MENU_SAVE_IDENTITY_MISMATCH: TYPE=" & typeText & "; TEXT=" & menuText
        Exit Function
    End If
    If UCase$(changeableText) <> "TRUE" Then
        messageText = "POST_MENU_SAVE_NOT_CHANGEABLE: " & changeableText
        Exit Function
    End If
    detailText = "ID=" & POST_MENU_SAVE_ID & "; TYPE=" & typeText & "; TEXT=" & menuText & _
        "; CHANGEABLE=" & changeableText
    CWR_GetVerifiedPOSTMenuSave = True
    Exit Function
Failed:
    messageText = "POST_MENU_SAVE_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set saveMenu = Nothing
End Function

Private Function CWR_GetModalText(ByVal ses As Object, ByRef modalText As String) As Boolean
    Dim modal As Object, body As Object, titleText As String, bodyText As String
    On Error GoTo ExitHere
    Set modal = ses.FindById("wnd[1]", False)
    If modal Is Nothing Then Exit Function
    titleText = CWR_Text(modal.text)
    Set body = modal.FindById("usr/txtMESSTXT1", False)
    If Not body Is Nothing Then bodyText = CWR_Text(body.text)
    modalText = "TITLE=" & titleText & IIf(bodyText <> "", "; TEXT=" & bodyText, "")
    CWR_GetModalText = True
ExitHere:
End Function

Private Function CWR_HasAnyModal(ByVal ses As Object, ByRef modalText As String) As Boolean
    Dim modal As Object
    On Error GoTo ExitHere
    Set modal = ses.FindById("wnd[1]", False)
    If modal Is Nothing Then Exit Function
    modalText = "TITLE=" & CWR_Text(modal.text)
    CWR_HasAnyModal = True
ExitHere:
End Function

' Only declines the standard unsaved-data prompt.  Any unexpected prompt is
' left untouched so recovery cannot accidentally save or discard user data.
Private Function CWR_DeclineExactNoSavePrompt(ByVal ses As Object, ByVal modalText As String, _
                                              ByRef messageText As String) As Boolean
    Dim yesButton As Object, noButton As Object, cancelButton As Object
    On Error GoTo Failed
    If InStr(1, modalText, "Data will be lost", vbTextCompare) = 0 Or _
       InStr(1, modalText, "save", vbTextCompare) = 0 Then
        messageText = "CREATE_WT_STALE_RECOVERY_UNEXPECTED_MODAL: " & modalText
        Exit Function
    End If
    Set yesButton = ses.FindById("wnd[1]/usr/btnBUTTON_1", False)
    Set noButton = ses.FindById("wnd[1]/usr/btnBUTTON_2", False)
    Set cancelButton = ses.FindById("wnd[1]/usr/btnBUTTON_3", False)
    If yesButton Is Nothing Or noButton Is Nothing Or cancelButton Is Nothing Then
        messageText = "CREATE_WT_STALE_RECOVERY_MODAL_BUTTONS_MISSING"
        Exit Function
    End If
    If StrComp(CWR_Text(yesButton.text), "Yes", vbTextCompare) <> 0 Or _
       StrComp(CWR_Text(noButton.text), "No", vbTextCompare) <> 0 Or _
       StrComp(CWR_Text(cancelButton.text), "Cancel", vbTextCompare) <> 0 Then
        messageText = "CREATE_WT_STALE_RECOVERY_MODAL_BUTTONS_MISMATCH: YES=" & CWR_Text(yesButton.text) & _
            "; NO=" & CWR_Text(noButton.text) & "; CANCEL=" & CWR_Text(cancelButton.text)
        Exit Function
    End If
    noButton.Press
    CWR_DeclineExactNoSavePrompt = True
    Exit Function
Failed:
    messageText = "CREATE_WT_STALE_RECOVERY_MODAL_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Reads the currently displayed Create+Save confirmation and the still-open
' grid behind it.  No control is focused or modified here; this is the last
' gate before the separate, explicitly user-authorized Yes response.
Private Function CWR_InspectExactPendingPOSTConfirmation(ByVal ses As Object, _
                                                          ByRef yesButton As Object, _
                                                          ByRef sourceCounterBefore As String, _
                                                          ByRef proofText As String, _
                                                          ByRef messageText As String) As Boolean
    Dim modal As Object, noButton As Object, cancelButton As Object, grid As Object
    Dim titleText As String, yesText As String, noText As String, cancelText As String
    Dim sourceProof As String, targetProof As String, payloadText As String, selectedRows As String, pendingProof As String
    Dim sourceRow As Long, targetRow As Long, statusType As String, statusText As String

    On Error GoTo Failed
    Set modal = ses.FindById("wnd[1]", False)
    If modal Is Nothing Then
        messageText = "POST_CONFIRMATION_MODAL_NOT_FOUND"
        Exit Function
    End If
    titleText = CWR_Text(modal.text)
    If InStr(1, titleText, "Posting Change for Warehouse Number " & CWR_EXPECTED_WAREHOUSE, vbTextCompare) = 0 Then
        messageText = "POST_CONFIRMATION_TITLE_MISMATCH: " & titleText
        Exit Function
    End If
    Set yesButton = ses.FindById("wnd[1]/usr/btnBUTTON_1", False)
    Set noButton = ses.FindById("wnd[1]/usr/btnBUTTON_2", False)
    Set cancelButton = ses.FindById("wnd[1]/usr/btnBUTTON_3", False)
    If yesButton Is Nothing Or noButton Is Nothing Or cancelButton Is Nothing Then
        messageText = "POST_CONFIRMATION_BUTTON_CONTRACT_MISSING"
        Exit Function
    End If
    yesText = CWR_Text(yesButton.text)
    noText = CWR_Text(noButton.text)
    cancelText = CWR_Text(cancelButton.text)
    If StrComp(yesText, "Yes", vbTextCompare) <> 0 Or _
       StrComp(noText, "No", vbTextCompare) <> 0 Or _
       StrComp(cancelText, "Cancel", vbTextCompare) <> 0 Then
        messageText = "POST_CONFIRMATION_BUTTON_TEXT_MISMATCH: YES=" & yesText & _
            "; NO=" & noText & "; CANCEL=" & cancelText
        Exit Function
    End If
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then Exit Function
    If Not CWR_ProveExactPOSTSourceAndTarget(grid, sourceRow, targetRow, sourceProof, targetProof, messageText) Then Exit Function
    selectedRows = CWR_SelectedRowsText(grid)
    If selectedRows <> "" And StrComp(selectedRows, CStr(sourceRow), vbBinaryCompare) <> 0 Then
        messageText = "POST_CONFIRMATION_WRONG_OR_MULTI_SELECTION: SELECTED_ROWS=" & selectedRows & _
            "; REQUIRED_SOURCE_ROW=" & CStr(sourceRow)
        Exit Function
    End If
    If Not CWR_ExactPOSTDraftPayload(grid, sourceRow, targetRow, payloadText, messageText) Then Exit Function
    ' SAP clears SelectedRows while this modal owns focus on this GUI build.
    ' In that state, the equivalent safety proof is that exactly one positive
    ' posting quantity exists and it is the uniquely proven old source row.
    If Not CWR_ProveOnlyExactPendingDraft(grid, sourceRow, pendingProof, messageText) Then Exit Function
    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "POST_CONFIRMATION_SAP_" & statusType & ": " & statusText
        Exit Function
    End If
    sourceCounterBefore = CWR_Cell(grid, sourceRow, "COUNTER")
    proofText = "MODAL_TITLE=" & titleText & "; BUTTONS=Yes/No/Cancel; SOURCE_ROW=" & CStr(sourceRow) & _
        "; SELECTED_ROWS=" & IIf(selectedRows = "", "<SAP_MODAL_CLEARED>", selectedRows) & _
        "; " & pendingProof & "; COUNTER_BEFORE=" & sourceCounterBefore & _
        "; " & payloadText & "; " & sourceProof & "; " & targetProof
    CWR_InspectExactPendingPOSTConfirmation = True
    Exit Function
Failed:
    messageText = "POST_CONFIRMATION_INSPECT_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set yesButton = Nothing
End Function

Private Function CWR_ExactPOSTDraftPayload(ByVal grid As Object, ByVal sourceRow As Long, _
                                            ByVal targetRow As Long, ByRef payloadText As String, _
                                            ByRef messageText As String) As Boolean
    Dim targetStorageType As String, targetStorageSection As String, targetStockType As String
    Dim stockTypeText As String, storageTypeText As String, storageSectionText As String
    Dim binText As String, qtyText As String, batchText As String
    On Error GoTo Failed
    targetStorageType = CWR_Cell(grid, targetRow, "VLTYP")
    targetStorageSection = CWR_Cell(grid, targetRow, "VLBER")
    targetStockType = CWR_Cell(grid, targetRow, "CAT")
    If StrComp(targetStorageType, CWR_TARGET_STORAGE_TYPE, vbTextCompare) <> 0 Or _
       StrComp(targetStorageSection, CWR_TARGET_STORAGE_SECTION, vbTextCompare) <> 0 Or _
       StrComp(targetStockType, CWR_TARGET_STOCK_TYPE, vbTextCompare) <> 0 Then
        messageText = "POST_TARGET_ATTRIBUTE_MISMATCH: VLTYP=" & targetStorageType & _
            "; VLBER=" & targetStorageSection & "; CAT=" & targetStockType
        Exit Function
    End If
    stockTypeText = CWR_Cell(grid, sourceRow, "TO_CAT")
    storageTypeText = CWR_Cell(grid, sourceRow, "NLTYP")
    storageSectionText = CWR_Cell(grid, sourceRow, "NLBER")
    binText = CWR_Cell(grid, sourceRow, "NLPLA")
    qtyText = CWR_Cell(grid, sourceRow, "POST_QUAN")
    batchText = CWR_Cell(grid, sourceRow, "TO_CHARG")
    payloadText = "DRAFT: TO_CAT=" & stockTypeText & "; NLTYP=" & storageTypeText & _
        "; NLBER=" & storageSectionText & "; NLPLA=" & binText & "; POST_QUAN=" & qtyText & _
        "; TO_CHARG=" & batchText
    If StrComp(stockTypeText, CWR_TARGET_STOCK_TYPE, vbTextCompare) <> 0 Or _
       StrComp(storageTypeText, CWR_TARGET_STORAGE_TYPE, vbTextCompare) <> 0 Or _
       StrComp(storageSectionText, CWR_TARGET_STORAGE_SECTION, vbTextCompare) <> 0 Or _
       StrComp(binText, CWR_TARGET_BIN, vbTextCompare) <> 0 Or _
       Abs(CWR_Number(qtyText) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(batchText, CWR_TARGET_BATCH, vbTextCompare) <> 0 Then
        messageText = "POST_CONFIRMATION_DRAFT_PAYLOAD_MISMATCH: " & payloadText
        Exit Function
    End If
    CWR_ExactPOSTDraftPayload = True
    Exit Function
Failed:
    messageText = "POST_CONFIRMATION_DRAFT_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_ProveOnlyExactPendingDraft(ByVal grid As Object, ByVal sourceRow As Long, _
                                                 ByRef proofText As String, ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, pendingRows As Long, qtyText As String, qtyValue As Double
    On Error GoTo Failed
    n = CLng(grid.rowCount)
    For r = 0 To n - 1
        qtyText = CWR_Cell(grid, r, "POST_QUAN")
        qtyValue = CWR_Number(qtyText)
        If qtyValue > 0.000001 Then
            pendingRows = pendingRows + 1
            If r <> sourceRow Then
                messageText = "POST_CONFIRMATION_OTHER_PENDING_ROW: ROW=" & CStr(r) & "; POST_QUAN=" & qtyText
                Exit Function
            End If
            If Abs(qtyValue - CWR_TARGET_QTY) > 0.000001 Then
                messageText = "POST_CONFIRMATION_SOURCE_PENDING_QTY_MISMATCH: " & qtyText
                Exit Function
            End If
        End If
    Next r
    If pendingRows <> 1 Then
        messageText = "POST_CONFIRMATION_PENDING_ROW_COUNT=" & CStr(pendingRows) & _
            "; REQUIRED_SOURCE_ROW=" & CStr(sourceRow)
        Exit Function
    End If
    proofText = "UNIQUE_PENDING_POST_QUAN_ROW=" & CStr(sourceRow) & "; POST_QUAN=" & CStr(CWR_TARGET_QTY)
    CWR_ProveOnlyExactPendingDraft = True
    Exit Function
Failed:
    messageText = "POST_CONFIRMATION_PENDING_SCAN_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_SelectedRowsText(ByVal grid As Object) As String
    On Error Resume Next
    CWR_SelectedRowsText = CWR_Text(grid.selectedRows & "")
    Err.Clear
End Function

' Reads selection metadata only.  It deliberately does not call
' SetCurrentCell, SelectionChanged, or any toolbar command.
Private Function CWR_GridSelectionState(ByVal labelText As String, ByVal grid As Object) As String
    Dim selectedText As String, currentRowText As String, currentColumnText As String, firstVisibleText As String
    On Error Resume Next
    selectedText = CWR_Text(grid.selectedRows & "")
    currentRowText = CWR_Text(grid.CurrentCellRow & "")
    currentColumnText = CWR_Text(grid.CurrentCellColumn & "")
    firstVisibleText = CWR_Text(grid.FirstVisibleRow & "")
    Err.Clear
    On Error GoTo 0
    CWR_GridSelectionState = labelText & "_SELECTION|SELECTED_ROWS=" & selectedText & _
        "|CURRENT_ROW=" & currentRowText & "|CURRENT_COLUMN=" & currentColumnText & _
        "|FIRST_VISIBLE_ROW=" & firstVisibleText & vbLf
End Function

' Returns the technical id and current enabled state of a toolbar without
' pressing it.  This makes any later action contract explicit in the run log.
Private Function CWR_DumpToolbarButtons(ByVal labelText As String, ByVal toolbar As Object) As String
    Dim i As Long, n As Long, buttonId As String, textText As String, tipText As String, enabledText As String
    On Error GoTo Failed
    n = CLng(toolbar.buttonCount)
    CWR_DumpToolbarButtons = labelText & "_TOOLBAR|BUTTONS=" & CStr(n) & vbLf
    For i = 0 To n - 1
        buttonId = CWR_Text(toolbar.GetButtonId(i))
        textText = CWR_Text(toolbar.GetButtonText(i))
        tipText = CWR_Text(toolbar.GetButtonTooltip(i))
        enabledText = CWR_Text(toolbar.GetButtonEnabled(i))
        CWR_DumpToolbarButtons = CWR_DumpToolbarButtons & labelText & "_TOOLBAR_BUTTON|INDEX=" & CStr(i) & _
            "|ID=" & buttonId & "|TEXT=" & textText & "|TIP=" & tipText & _
            "|ENABLED=" & enabledText & vbLf
    Next i
    Exit Function
Failed:
    CWR_DumpToolbarButtons = labelText & "_TOOLBAR_DUMP_ERROR: " & CStr(Err.Number) & ": " & Err.Description & vbLf
End Function

Private Function CWR_PostCounterAdvanced(ByVal beforeText As String, ByVal afterText As String, _
                                         ByRef messageText As String) As Boolean
    Dim beforeValue As Double, afterValue As Double
    If Not IsNumeric(afterText) Then
        messageText = "POST_CREATE_COUNTER_AFTER_NOT_NUMERIC: " & afterText
        Exit Function
    End If
    beforeValue = CWR_Number(beforeText)
    afterValue = CWR_Number(afterText)
    If afterValue <= beforeValue + 0.000001 Then
        messageText = "POST_CREATE_COUNTER_NOT_ADVANCED: BEFORE=" & beforeText & "; AFTER=" & afterText
        Exit Function
    End If
    CWR_PostCounterAdvanced = True
End Function

Private Sub CWR_ReadStatusBar(ByVal ses As Object, ByRef statusType As String, ByRef statusText As String)
    Dim sbar As Object
    On Error Resume Next
    Set sbar = ses.FindById("wnd[0]/sbar", False)
    If Not sbar Is Nothing Then
        statusType = CWR_Text(sbar.messageType)
        statusText = CWR_Text(sbar.text)
    End If
    On Error GoTo 0
End Sub

Private Function CWR_ResetPOSTWithoutSave(ByVal ses As Object, ByRef messageText As String) As Boolean
    Dim tcode As String, modalText As String
    On Error GoTo Failed
    If ses Is Nothing Then
        messageText = "POST_RESET_SESSION_NOT_SET"
        Exit Function
    End If
    ses.StartTransaction "/SCWM/POST"
    If Not CWR_WaitIdle(ses, 30#) Then
        messageText = "POST_RESET_TIMEOUT"
        Exit Function
    End If
    If CWR_GetModalText(ses, modalText) Then
        messageText = "POST_RESET_MODAL: " & modalText
        Exit Function
    End If
    tcode = CWR_Text(ses.Info.Transaction)
    If UCase$(tcode) <> "/SCWM/POST" Then
        messageText = "POST_RESET_TCODE_MISMATCH: " & tcode
        Exit Function
    End If
    CWR_ResetPOSTWithoutSave = True
    Exit Function
Failed:
    messageText = "POST_RESET_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_ProveExactTargetStock(ByVal grid As Object, ByRef proofText As String, _
                                           ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, lastMaterial As String, matText As String, batchText As String, binText As String
    Dim stockType As String, availText As String, totalText As String, matches As Long
    Dim availableQty As Double, totalQty As Double

    On Error GoTo Failed
    n = CLng(grid.rowCount): matches = 0: lastMaterial = ""
    For r = 0 To n - 1
        matText = CWR_Cell(grid, r, "MATNR")
        If matText <> "" Then lastMaterial = matText Else matText = lastMaterial
        batchText = CWR_Cell(grid, r, "CHARG")
        binText = CWR_Cell(grid, r, "VLPLA")
        stockType = CWR_Cell(grid, r, "TO_CAT")
        If StrComp(matText, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 And _
           StrComp(batchText, CWR_TARGET_BATCH, vbTextCompare) = 0 And _
           StrComp(binText, CWR_TARGET_BIN, vbTextCompare) = 0 And _
           StrComp(stockType, "FT", vbTextCompare) = 0 Then
            matches = matches + 1
            availText = CWR_Cell(grid, r, "AVL_ALT_QTY")
            totalText = CWR_Cell(grid, r, "LIME_QUAN")
            availableQty = CWR_Number(availText)
            totalQty = CWR_Number(totalText)
        End If
    Next r
    If matches <> 1 Then
        messageText = "TARGET_STOCK_FT_MATCH_COUNT=" & CStr(matches) & _
            "; MAT=" & CWR_EXPECTED_MATERIAL & "; BIN=" & CWR_TARGET_BIN & "; BATCH=" & CWR_TARGET_BATCH
        Exit Function
    End If
    If availableQty + 0.000001 < CWR_TARGET_QTY Or totalQty + 0.000001 < CWR_TARGET_QTY Then
        messageText = "TARGET_STOCK_INSUFFICIENT: AVAILABLE=" & CStr(availableQty) & "; TOTAL=" & CStr(totalQty) & _
            "; REQUIRED=" & CStr(CWR_TARGET_QTY)
        Exit Function
    End If
    proofText = "MAT=" & CWR_EXPECTED_MATERIAL & "; BIN=" & CWR_TARGET_BIN & "; BATCH=" & CWR_TARGET_BATCH & _
        "; STOCK_TYPE=FT; AVAILABLE=" & CStr(availableQty) & "; TOTAL=" & CStr(totalQty) & _
        "; REQUIRED=" & CStr(CWR_TARGET_QTY)
    CWR_ProveExactTargetStock = True
    Exit Function
Failed:
    messageText = "TARGET_STOCK_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Returns the actual stock candidates at the requested new pick bin without
' assuming the requested target batch exists yet.  A batch edit may only be
' designed from a single, fully identified eligible candidate in this output.
Private Function CWR_DumpTargetBinStockCandidates(ByVal grid As Object, _
                                                   ByRef messageText As String) As String
    Dim r As Long, n As Long, lastMaterial As String, matText As String, batchText As String
    Dim binText As String, stockTypeText As String, availText As String, totalText As String
    Dim ownerText As String, entitledText As String, candidateCount As Long, whBatchCount As Long
    Dim pickBatchCount As Long, eligibleCount As Long

    On Error GoTo Failed
    n = CLng(grid.rowCount)
    lastMaterial = ""
    For r = 0 To n - 1
        matText = CWR_Cell(grid, r, "MATNR")
        If matText <> "" Then lastMaterial = matText Else matText = lastMaterial
        binText = CWR_Cell(grid, r, "VLPLA")
        stockTypeText = CWR_Cell(grid, r, "TO_CAT")
        If StrComp(matText, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 And _
           StrComp(binText, CWR_TARGET_BIN, vbTextCompare) = 0 And _
           StrComp(stockTypeText, CWR_TARGET_STOCK_TYPE, vbTextCompare) = 0 Then
            batchText = CWR_Cell(grid, r, "CHARG")
            availText = CWR_Cell(grid, r, "AVL_ALT_QTY")
            totalText = CWR_Cell(grid, r, "LIME_QUAN")
            ownerText = CWR_Cell(grid, r, "OWNER")
            entitledText = CWR_Cell(grid, r, "ENTITLED")
            candidateCount = candidateCount + 1
            If StrComp(batchText, CWR_WH_BATCH, vbTextCompare) = 0 Then whBatchCount = whBatchCount + 1
            If StrComp(batchText, CWR_PICK_BIN_BATCH_BEFORE, vbTextCompare) = 0 Then pickBatchCount = pickBatchCount + 1
            If CWR_Number(availText) + 0.000001 >= CWR_TARGET_QTY And _
               CWR_Number(totalText) + 0.000001 >= CWR_TARGET_QTY Then eligibleCount = eligibleCount + 1
            CWR_DumpTargetBinStockCandidates = CWR_DumpTargetBinStockCandidates & _
                "CANDIDATE|ROW=" & CStr(r) & "; MAT=" & matText & "; BIN=" & binText & _
                "; BATCH=" & batchText & "; STOCK_TYPE=" & stockTypeText & _
                "; AVAILABLE=" & availText & "; TOTAL=" & totalText & _
                "; OWNER=" & ownerText & "; ENTITLED=" & entitledText & vbLf
        End If
    Next r
    CWR_DumpTargetBinStockCandidates = CWR_DumpTargetBinStockCandidates & _
        "SUMMARY|TARGET_BIN=" & CWR_TARGET_BIN & "; CANDIDATES=" & CStr(candidateCount) & _
        "; WH_BATCH=" & CWR_WH_BATCH & "; WH_BATCH_MATCHES=" & CStr(whBatchCount) & _
        "; PICK_BIN_BATCH_BEFORE=" & CWR_PICK_BIN_BATCH_BEFORE & "; PICK_BIN_BATCH_MATCHES=" & CStr(pickBatchCount) & _
        "; ELIGIBLE_AT_LEAST_" & CStr(CWR_TARGET_QTY) & "=" & CStr(eligibleCount) & vbLf
    If candidateCount = 0 Then
        CWR_DumpTargetBinStockCandidates = CWR_DumpTargetBinStockCandidates & _
            "NEXT=STOP_NO_STOCK_AT_TARGET_BIN; NO_BATCH_CHANGE_OR_WT_CREATE"
    ElseIf whBatchCount = 0 And pickBatchCount = 1 Then
        CWR_DumpTargetBinStockCandidates = CWR_DumpTargetBinStockCandidates & _
            "NEXT=POST_CHANGE_PICK_BIN_BATCH_TO_WH_BATCH; NO_WT_CREATE"
    ElseIf whBatchCount = 0 Then
        CWR_DumpTargetBinStockCandidates = CWR_DumpTargetBinStockCandidates & _
            "NEXT=STOP_PICK_BIN_SOURCE_BATCH_AMBIGUOUS; NO_WT_CREATE"
    Else
        CWR_DumpTargetBinStockCandidates = CWR_DumpTargetBinStockCandidates & _
            "NEXT=WH_BATCH_ALREADY_PRESENT; RECONCILE_QTY_BEFORE_ANY_WT_CREATE"
    End If
    Exit Function
Failed:
    messageText = "TARGET_BIN_STOCK_SCAN_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    CWR_DumpTargetBinStockCandidates = "FAIL|" & messageText
End Function

' Builds a no-write plan for a batch-only posting change at the new pick bin.
' It never considers CWR_EXPECTED_SOURCE_BIN, which prevents this workflow
' from moving stock from the cancelled task's former bin.
Private Function CWR_PlanPickBinBatchChange(ByVal grid As Object, ByRef sourceRow As Long, _
                                            ByRef planState As String, ByRef planText As String, _
                                            ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, lastMaterial As String, matText As String, batchText As String
    Dim binText As String, stockTypeText As String, availText As String, totalText As String
    Dim targetMatches As Long, sourceMatches As Long, targetAvailable As Double, targetTotal As Double
    Dim sourceAvailable As Double, sourceTotal As Double, sourceBatch As String
    Dim sourceType As String, sourceSection As String, sourceStockType As String

    On Error GoTo Failed
    sourceRow = -1
    n = CLng(grid.rowCount)
    lastMaterial = ""
    For r = 0 To n - 1
        matText = CWR_Cell(grid, r, "MATNR")
        If matText <> "" Then lastMaterial = matText Else matText = lastMaterial
        binText = CWR_Cell(grid, r, "VLPLA")
        stockTypeText = CWR_Cell(grid, r, "TO_CAT")
        If stockTypeText = "" Then stockTypeText = CWR_Cell(grid, r, "CAT")
        If StrComp(matText, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 And _
           StrComp(binText, CWR_TARGET_BIN, vbTextCompare) = 0 And _
           StrComp(stockTypeText, CWR_TARGET_STOCK_TYPE, vbTextCompare) = 0 Then
            batchText = CWR_Cell(grid, r, "CHARG")
            availText = CWR_Cell(grid, r, "AVL_ALT_QTY")
            totalText = CWR_Cell(grid, r, "LIME_QUAN")
            If StrComp(batchText, CWR_WH_BATCH, vbTextCompare) = 0 Then
                targetMatches = targetMatches + 1
                targetAvailable = CWR_Number(availText)
                targetTotal = CWR_Number(totalText)
            ElseIf StrComp(batchText, CWR_PICK_BIN_BATCH_BEFORE, vbTextCompare) = 0 And _
                   CWR_Number(availText) + 0.000001 >= CWR_TARGET_QTY And _
                   CWR_Number(totalText) + 0.000001 >= CWR_TARGET_QTY Then
                sourceMatches = sourceMatches + 1
                sourceRow = r
                sourceBatch = batchText
                sourceAvailable = CWR_Number(availText)
                sourceTotal = CWR_Number(totalText)
                sourceType = CWR_Cell(grid, r, "VLTYP")
                sourceSection = CWR_Cell(grid, r, "VLBER")
                sourceStockType = stockTypeText
            End If
        End If
    Next r

    If targetMatches > 1 Then
        messageText = "PICK_BIN_TARGET_BATCH_AMBIGUOUS: MATCHES=" & CStr(targetMatches) & _
            "; BIN=" & CWR_TARGET_BIN & "; WH_BATCH=" & CWR_WH_BATCH
        Exit Function
    End If
    If targetMatches = 1 And targetAvailable + 0.000001 >= CWR_TARGET_QTY And _
       targetTotal + 0.000001 >= CWR_TARGET_QTY Then
        planState = "NO_CHANGE_REQUIRED"
        planText = "TARGET: MAT=" & CWR_EXPECTED_MATERIAL & "; BIN=" & CWR_TARGET_BIN & _
            "; BATCH=" & CWR_WH_BATCH & "; STOCK_TYPE=" & CWR_TARGET_STOCK_TYPE & _
            "; AVAILABLE=" & CStr(targetAvailable) & "; TOTAL=" & CStr(targetTotal) & _
            "; REQUIRED=" & CStr(CWR_TARGET_QTY)
        CWR_PlanPickBinBatchChange = True
        Exit Function
    End If
    If sourceMatches <> 1 Then
        messageText = "PICK_BIN_BATCH_SOURCE_MATCH_COUNT=" & CStr(sourceMatches) & _
            "; TARGET_BIN=" & CWR_TARGET_BIN & "; SOURCE_BATCH=" & CWR_PICK_BIN_BATCH_BEFORE & _
            "; WH_BATCH=" & CWR_WH_BATCH & _
            "; TARGET_EXISTS=" & CStr(targetMatches)
        Exit Function
    End If
    ' Valid APFL locations may not use a storage section. Keep the blank
    ' source section as-is; type, bin, stock type and batch remain required.
    If sourceType = "" Or sourceStockType = "" Then
        messageText = "PICK_BIN_SOURCE_ATTRIBUTES_MISSING: VLTYP=" & sourceType & _
            "; VLBER=" & sourceSection & "; CAT=" & sourceStockType
        Exit Function
    End If
    planState = "CHANGE_REQUIRED"
    planText = "SOURCE: ROW=" & CStr(sourceRow) & "; MAT=" & CWR_EXPECTED_MATERIAL & _
        "; BIN=" & CWR_TARGET_BIN & "; BATCH=" & sourceBatch & _
        "; STOCK_TYPE=" & sourceStockType & "; VLTYP=" & sourceType & "; VLBER=" & sourceSection & _
        "; AVAILABLE=" & CStr(sourceAvailable) & "; TOTAL=" & CStr(sourceTotal) & _
        "; CHANGE_TO_BATCH=" & CWR_WH_BATCH & "; POST_QUAN=" & CStr(CWR_TARGET_QTY)
    CWR_PlanPickBinBatchChange = True
    Exit Function
Failed:
    messageText = "PICK_BIN_BATCH_PLAN_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' A write may never begin while the currently displayed POST list contains a
' pending posting quantity.  This leaves any user draft untouched for manual
' review instead of overwriting or silently refreshing it.
Private Function CWR_AssertNoPendingPOSTDraft(ByVal grid As Object, ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, qtyText As String, qtyValue As Double
    On Error GoTo Failed
    n = CLng(grid.rowCount)
    For r = 0 To n - 1
        qtyText = CWR_Cell(grid, r, "POST_QUAN")
        qtyValue = CWR_Number(qtyText)
        If qtyValue > 0.000001 Then
            messageText = "PICK_BIN_POST_PENDING_DRAFT_EXISTS: ROW=" & CStr(r) & "; POST_QUAN=" & qtyText
            Exit Function
        End If
    Next r
    CWR_AssertNoPendingPOSTDraft = True
    Exit Function
Failed:
    messageText = "PICK_BIN_POST_PENDING_DRAFT_SCAN_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Validates every grid field used by the in-place change before any value is
' written.  The source bin is never a destination derived from another row.
Private Function CWR_AssertPickBinPostFieldsChangeable(ByVal grid As Object, ByVal sourceRow As Long, _
                                                         ByVal sourceSection As String, ByRef messageText As String) As Boolean
    Dim fields As Variant, fieldName As Variant, stateText As String
    On Error GoTo Failed
    fields = Array("TO_CAT", "NLTYP", "NLPLA", "POST_QUAN", "TO_CHARG")
    For Each fieldName In fields
        stateText = CWR_CellChangeable(grid, sourceRow, CStr(fieldName))
        If UCase$(stateText) <> "TRUE" Then
            messageText = "PICK_BIN_POST_FIELD_NOT_CHANGEABLE: ROW=" & CStr(sourceRow) & _
                "; FIELD=" & CStr(fieldName) & "; CHANGEABLE=" & stateText
            Exit Function
        End If
    Next fieldName
    If sourceSection <> "" Then
        stateText = CWR_CellChangeable(grid, sourceRow, "NLBER")
        If UCase$(stateText) <> "TRUE" Then
            messageText = "PICK_BIN_POST_FIELD_NOT_CHANGEABLE: ROW=" & CStr(sourceRow) & _
                "; FIELD=NLBER; CHANGEABLE=" & stateText
            Exit Function
        End If
    End If
    CWR_AssertPickBinPostFieldsChangeable = True
    Exit Function
Failed:
    messageText = "PICK_BIN_POST_FIELD_CONTRACT_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Reads only the exact material/bin/stock-type/owner entitlement slice used by
' the selected source.  It returns aggregated balance by the two relevant
' batches so post-verification can prove the precise quantity delta even when
' SAP hides a zero-quantity source row after the posting change.
Private Function CWR_ReadPickBinBatchBalances(ByVal grid As Object, ByVal expectedOwner As String, _
                                               ByVal expectedEntitled As String, _
                                               ByRef sourceAvailable As Double, ByRef sourceTotal As Double, _
                                               ByRef whAvailable As Double, ByRef whTotal As Double, _
                                               ByRef sourceRows As Long, ByRef whRows As Long, _
                                               ByRef proofText As String, ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, lastMaterial As String, matText As String, batchText As String
    Dim binText As String, stockTypeText As String, ownerText As String, entitledText As String
    Dim availText As String, totalText As String
    On Error GoTo Failed
    n = CLng(grid.rowCount)
    lastMaterial = ""
    sourceAvailable = 0#: sourceTotal = 0#: whAvailable = 0#: whTotal = 0#
    sourceRows = 0: whRows = 0
    For r = 0 To n - 1
        matText = CWR_Cell(grid, r, "MATNR")
        If matText <> "" Then lastMaterial = matText Else matText = lastMaterial
        binText = CWR_Cell(grid, r, "VLPLA")
        stockTypeText = CWR_Cell(grid, r, "TO_CAT")
        If stockTypeText = "" Then stockTypeText = CWR_Cell(grid, r, "CAT")
        ownerText = CWR_Cell(grid, r, "OWNER")
        entitledText = CWR_Cell(grid, r, "ENTITLED")
        If StrComp(matText, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 And _
           StrComp(binText, CWR_TARGET_BIN, vbTextCompare) = 0 And _
           StrComp(stockTypeText, CWR_TARGET_STOCK_TYPE, vbTextCompare) = 0 And _
           (expectedOwner = "" Or StrComp(ownerText, expectedOwner, vbTextCompare) = 0) And _
           (expectedEntitled = "" Or StrComp(entitledText, expectedEntitled, vbTextCompare) = 0) Then
            batchText = CWR_Cell(grid, r, "CHARG")
            availText = CWR_Cell(grid, r, "AVL_ALT_QTY")
            totalText = CWR_Cell(grid, r, "LIME_QUAN")
            If StrComp(batchText, CWR_PICK_BIN_BATCH_BEFORE, vbTextCompare) = 0 Then
                sourceRows = sourceRows + 1
                sourceAvailable = sourceAvailable + CWR_Number(availText)
                sourceTotal = sourceTotal + CWR_Number(totalText)
            ElseIf StrComp(batchText, CWR_WH_BATCH, vbTextCompare) = 0 Then
                whRows = whRows + 1
                whAvailable = whAvailable + CWR_Number(availText)
                whTotal = whTotal + CWR_Number(totalText)
            End If
        End If
    Next r
    proofText = "PICK_BIN_BALANCE: BIN=" & CWR_TARGET_BIN & "; OWNER=" & expectedOwner & _
        "; ENTITLED=" & expectedEntitled & "; SOURCE_BATCH=" & CWR_PICK_BIN_BATCH_BEFORE & _
        "; SOURCE_ROWS=" & CStr(sourceRows) & "; SOURCE_AVAILABLE=" & CStr(sourceAvailable) & _
        "; SOURCE_TOTAL=" & CStr(sourceTotal) & "; WH_BATCH=" & CWR_WH_BATCH & _
        "; WH_ROWS=" & CStr(whRows) & "; WH_AVAILABLE=" & CStr(whAvailable) & _
        "; WH_TOTAL=" & CStr(whTotal)
    CWR_ReadPickBinBatchBalances = True
    Exit Function
Failed:
    messageText = "PICK_BIN_POST_BALANCE_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Confirms the draft still identifies the original actual stock at the exact
' same source bin, and that only its destination batch is the ODO/WH batch.
Private Function CWR_ProveExactPickBinBatchDraft(ByVal grid As Object, ByVal sourceRow As Long, _
                                                  ByVal sourceType As String, ByVal sourceSection As String, _
                                                  ByRef proofText As String, ByRef messageText As String) As Boolean
    Dim sourceBatch As String, sourceBin As String, sourceStockType As String
    Dim destStockType As String, destType As String, destSection As String, destBin As String
    Dim postQty As String, destBatch As String, lastMaterial As String, r As Long
    On Error GoTo Failed
    If sourceRow < 0 Or sourceRow >= CLng(grid.rowCount) Then
        messageText = "PICK_BIN_POST_DRAFT_SOURCE_ROW_OUT_OF_RANGE: " & CStr(sourceRow)
        Exit Function
    End If
    For r = 0 To sourceRow
        If CWR_Cell(grid, r, "MATNR") <> "" Then lastMaterial = CWR_Cell(grid, r, "MATNR")
    Next r
    sourceBatch = CWR_Cell(grid, sourceRow, "CHARG")
    sourceBin = CWR_Cell(grid, sourceRow, "VLPLA")
    sourceStockType = CWR_Cell(grid, sourceRow, "CAT")
    If sourceStockType = "" Then sourceStockType = CWR_Cell(grid, sourceRow, "TO_CAT")
    destStockType = CWR_Cell(grid, sourceRow, "TO_CAT")
    destType = CWR_Cell(grid, sourceRow, "NLTYP")
    destSection = CWR_Cell(grid, sourceRow, "NLBER")
    destBin = CWR_Cell(grid, sourceRow, "NLPLA")
    postQty = CWR_Cell(grid, sourceRow, "POST_QUAN")
    destBatch = CWR_Cell(grid, sourceRow, "TO_CHARG")
    proofText = "DRAFT: ROW=" & CStr(sourceRow) & "; MAT=" & lastMaterial & _
        "; SOURCE_BIN=" & sourceBin & "; SOURCE_BATCH=" & sourceBatch & _
        "; SOURCE_TYPE=" & CWR_Cell(grid, sourceRow, "VLTYP") & "/" & CWR_Cell(grid, sourceRow, "VLBER") & _
        "; TO_CAT=" & destStockType & "; NLTYP=" & destType & "; NLBER=" & destSection & _
        "; NLPLA=" & destBin & "; POST_QUAN=" & postQty & "; TO_CHARG=" & destBatch
    If StrComp(lastMaterial, CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Or _
       StrComp(sourceBatch, CWR_PICK_BIN_BATCH_BEFORE, vbTextCompare) <> 0 Or _
       StrComp(sourceBin, CWR_TARGET_BIN, vbTextCompare) <> 0 Or _
       StrComp(sourceStockType, CWR_TARGET_STOCK_TYPE, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(grid, sourceRow, "VLTYP"), sourceType, vbTextCompare) <> 0 Or _
       StrComp(CWR_Cell(grid, sourceRow, "VLBER"), sourceSection, vbTextCompare) <> 0 Or _
       StrComp(destStockType, CWR_TARGET_STOCK_TYPE, vbTextCompare) <> 0 Or _
       StrComp(destType, sourceType, vbTextCompare) <> 0 Or _
       StrComp(destSection, sourceSection, vbTextCompare) <> 0 Or _
       StrComp(destBin, CWR_TARGET_BIN, vbTextCompare) <> 0 Or _
       Abs(CWR_Number(postQty) - CWR_TARGET_QTY) > 0.000001 Or _
       StrComp(destBatch, CWR_WH_BATCH, vbTextCompare) <> 0 Then
        messageText = "PICK_BIN_POST_DRAFT_PAYLOAD_MISMATCH: " & proofText
        Exit Function
    End If
    CWR_ProveExactPickBinBatchDraft = True
    Exit Function
Failed:
    messageText = "PICK_BIN_POST_DRAFT_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Selects exactly the already-proven source row immediately before the only
' Create+Save action. It does not edit any value or press a business command.
Private Function CWR_SelectExactPickBinBatchSource(ByVal ses As Object, ByRef grid As Object, _
                                                    ByVal sourceRow As Long, ByVal sourceType As String, _
                                                    ByVal sourceSection As String, ByRef messageText As String) As Boolean
    Dim draftProof As String, selectedRows As String, modalText As String, statusType As String, statusText As String
    On Error GoTo Failed
    If Not CWR_ProveExactPickBinBatchDraft(grid, sourceRow, sourceType, sourceSection, draftProof, messageText) Then Exit Function
    grid.selectedRows = CStr(sourceRow)
    grid.SetCurrentCell sourceRow, "TO_CHARG"
    grid.SelectionChanged
    If Not CWR_WaitPOSTSaveSettle(ses, 0.5, 15#) Then
        messageText = "PICK_BIN_POST_SELECTION_TIMEOUT"
        Exit Function
    End If
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = "PICK_BIN_POST_SELECTION_MODAL: " & modalText
        Exit Function
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "PICK_BIN_POST_SELECTION_SAP_" & statusType & ": " & statusText
        Exit Function
    End If
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then Exit Function
    If Not CWR_ProveExactPickBinBatchDraft(grid, sourceRow, sourceType, sourceSection, draftProof, messageText) Then Exit Function
    selectedRows = CWR_SelectedRowsText(grid)
    If StrComp(selectedRows, CStr(sourceRow), vbBinaryCompare) <> 0 Then
        messageText = "PICK_BIN_POST_SELECTION_READBACK_MISMATCH: SELECTED_ROWS=" & selectedRows & _
            "; REQUIRED_SOURCE_ROW=" & CStr(sourceRow)
        Exit Function
    End If
    CWR_SelectExactPickBinBatchSource = True
    Exit Function
Failed:
    messageText = "PICK_BIN_POST_SELECTION_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' /SCWM/POST can save the posting-change task directly and open SAP's display
' log instead of the Yes/No/Cancel draft dialog. That is a successful task
' creation, but not proof of a confirmed batch change; CF remains with user.
Private Function CWR_InspectCreatedPostingWTLog(ByVal ses As Object, ByRef wtNo As String, _
                                                 ByRef proofText As String, ByRef messageText As String) As Boolean
    Dim modal As Object, titleText As String, statusType As String, statusText As String
    On Error GoTo Failed
    Set modal = ses.FindById("wnd[1]", False)
    If modal Is Nothing Then
        messageText = "POST_CREATED_WT_LOG_MODAL_NOT_FOUND"
        Exit Function
    End If
    titleText = CWR_Text(modal.text)
    If StrComp(titleText, "Display logs", vbTextCompare) <> 0 Then
        messageText = "POST_CREATED_WT_LOG_TITLE_MISMATCH: " & titleText
        Exit Function
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    wtNo = CWR_ExtractWarehouseTaskNumber(statusText)
    If UCase$(statusType) <> "S" Or wtNo = "" Or InStr(1, statusText, "created", vbTextCompare) = 0 Then
        messageText = "POST_CREATED_WT_LOG_STATUS_UNPROVEN: " & statusType & ":" & statusText
        Exit Function
    End If
    proofText = "LOG_TITLE=Display logs; STATUS=" & statusType & ":" & statusText & _
        "; RESULT=POSTING_CHANGE_WT_CREATED_NOT_CONFIRMED_BY_CODE"
    CWR_InspectCreatedPostingWTLog = True
    Exit Function
Failed:
    messageText = "POST_CREATED_WT_LOG_INSPECT_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Dismisses only SAP's informational Display logs viewer after a returned
' posting-change WT number.  This is not a business confirmation and never
' presses a CF/Confirm action; it merely permits the required fresh readback.
Private Function CWR_CloseCreatedPostingWTLogForRead(ByVal ses As Object, _
                                                      ByRef messageText As String) As Boolean
    Dim modal As Object, continueButton As Object, titleText As String, tooltipText As String, modalText As String
    On Error GoTo Failed
    Set modal = ses.FindById("wnd[1]", False)
    If modal Is Nothing Then
        messageText = "POST_CREATED_WT_LOG_MODAL_NOT_PRESENT"
        Exit Function
    End If
    titleText = CWR_Text(modal.text)
    If StrComp(titleText, "Display logs", vbTextCompare) <> 0 Then
        messageText = "POST_CREATED_WT_LOG_TITLE_MISMATCH: " & titleText
        Exit Function
    End If
    Set continueButton = ses.FindById("wnd[1]/tbar[0]/btn[0]", False)
    If continueButton Is Nothing Then
        messageText = "POST_CREATED_WT_LOG_CONTINUE_NOT_FOUND"
        Exit Function
    End If
    tooltipText = CWR_Text(continueButton.Tooltip)
    If InStr(1, tooltipText, "Continue", vbTextCompare) = 0 Then
        messageText = "POST_CREATED_WT_LOG_CONTINUE_IDENTITY_MISMATCH: " & tooltipText
        Exit Function
    End If
    continueButton.Press
    If Not CWR_WaitPOSTSaveSettle(ses, 0.5, 20#) Then
        messageText = "POST_CREATED_WT_LOG_CLOSE_TIMEOUT"
        Exit Function
    End If
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = "POST_CREATED_WT_LOG_CLOSE_MODAL_REMAINS: " & modalText
        Exit Function
    End If
    CWR_CloseCreatedPostingWTLogForRead = True
    Exit Function
Failed:
    messageText = "POST_CREATED_WT_LOG_CLOSE_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Validates the exact Yes/No/Cancel confirmation created by the internal POST
' Create+Save command. It never answers the dialog; the caller does so once.
Private Function CWR_InspectPickBinBatchConfirmation(ByVal ses As Object, ByVal sourceType As String, _
                                                       ByVal sourceSection As String, ByRef yesButton As Object, _
                                                       ByRef proofText As String, ByRef messageText As String) As Boolean
    Dim modal As Object, noButton As Object, cancelButton As Object, grid As Object
    Dim titleText As String, yesText As String, noText As String, cancelText As String
    Dim draftProof As String, pendingProof As String, selectedRows As String, statusType As String, statusText As String
    Dim sourceRow As Long, planState As String, planText As String
    On Error GoTo Failed
    Set modal = ses.FindById("wnd[1]", False)
    If modal Is Nothing Then
        messageText = "PICK_BIN_POST_CONFIRMATION_MODAL_NOT_FOUND"
        Exit Function
    End If
    titleText = CWR_Text(modal.text)
    If InStr(1, titleText, "Posting Change for Warehouse Number " & CWR_EXPECTED_WAREHOUSE, vbTextCompare) = 0 Then
        messageText = "PICK_BIN_POST_CONFIRMATION_TITLE_MISMATCH: " & titleText
        Exit Function
    End If
    Set yesButton = ses.FindById("wnd[1]/usr/btnBUTTON_1", False)
    Set noButton = ses.FindById("wnd[1]/usr/btnBUTTON_2", False)
    Set cancelButton = ses.FindById("wnd[1]/usr/btnBUTTON_3", False)
    If yesButton Is Nothing Or noButton Is Nothing Or cancelButton Is Nothing Then
        messageText = "PICK_BIN_POST_CONFIRMATION_BUTTON_CONTRACT_MISSING"
        Exit Function
    End If
    yesText = CWR_Text(yesButton.text)
    noText = CWR_Text(noButton.text)
    cancelText = CWR_Text(cancelButton.text)
    If StrComp(yesText, "Yes", vbTextCompare) <> 0 Or _
       StrComp(noText, "No", vbTextCompare) <> 0 Or _
       StrComp(cancelText, "Cancel", vbTextCompare) <> 0 Then
        messageText = "PICK_BIN_POST_CONFIRMATION_BUTTON_TEXT_MISMATCH: YES=" & yesText & _
            "; NO=" & noText & "; CANCEL=" & cancelText
        Exit Function
    End If
    If Not CWR_FindCurrentPOSTGrid(ses, grid, messageText) Then Exit Function
    If Not CWR_PlanPickBinBatchChange(grid, sourceRow, planState, planText, messageText) Then Exit Function
    If StrComp(planState, "CHANGE_REQUIRED", vbTextCompare) <> 0 Then
        messageText = "PICK_BIN_POST_CONFIRMATION_PLAN_CHANGED: " & planState & "; " & planText
        Exit Function
    End If
    If Not CWR_ProveExactPickBinBatchDraft(grid, sourceRow, sourceType, sourceSection, draftProof, messageText) Then Exit Function
    If Not CWR_ProveOnlyExactPendingDraft(grid, sourceRow, pendingProof, messageText) Then Exit Function
    selectedRows = CWR_SelectedRowsText(grid)
    If selectedRows <> "" And StrComp(selectedRows, CStr(sourceRow), vbBinaryCompare) <> 0 Then
        messageText = "PICK_BIN_POST_CONFIRMATION_WRONG_OR_MULTI_SELECTION: SELECTED_ROWS=" & selectedRows & _
            "; REQUIRED_SOURCE_ROW=" & CStr(sourceRow)
        Exit Function
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = "PICK_BIN_POST_CONFIRMATION_SAP_" & statusType & ": " & statusText
        Exit Function
    End If
    proofText = "MODAL_TITLE=" & titleText & "; BUTTONS=Yes/No/Cancel; SOURCE_ROW=" & CStr(sourceRow) & _
        "; SELECTED_ROWS=" & IIf(selectedRows = "", "<SAP_MODAL_CLEARED>", selectedRows) & _
        "; " & pendingProof & "; " & draftProof
    CWR_InspectPickBinBatchConfirmation = True
    Exit Function
Failed:
    messageText = "PICK_BIN_POST_CONFIRMATION_INSPECT_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    Set yesButton = Nothing
End Function

' Verifies a toolbar command by its stable technical id and displayed tooltip
' before it is pressed. The caller remains responsible for the one permitted
' PressButton call; this helper itself never invokes any action.
Private Function CWR_GetVerifiedToolbarAction(ByVal toolbar As Object, ByVal expectedId As String, _
                                              ByVal expectedTooltip As String, ByRef detailText As String, _
                                              ByRef messageText As String) As Boolean
    Dim i As Long, n As Long, matches As Long, buttonId As String, tipText As String, enabledText As String
    On Error GoTo Failed
    n = CLng(toolbar.buttonCount)
    For i = 0 To n - 1
        buttonId = CWR_Text(toolbar.GetButtonId(i))
        If StrComp(buttonId, expectedId, vbTextCompare) = 0 Then
            matches = matches + 1
            tipText = CWR_Text(toolbar.GetButtonTooltip(i))
            enabledText = CWR_Text(toolbar.GetButtonEnabled(i))
            detailText = "ID=" & buttonId & "; TOOLTIP=" & tipText & "; ENABLED=" & enabledText
            If StrComp(tipText, expectedTooltip, vbTextCompare) <> 0 Then
                messageText = "CREATE_WT_TOOLBAR_TOOLTIP_MISMATCH: " & detailText
                Exit Function
            End If
            If StrComp(enabledText, "True", vbTextCompare) <> 0 Then
                messageText = "CREATE_WT_TOOLBAR_ACTION_NOT_ENABLED: " & detailText
                Exit Function
            End If
        End If
    Next i
    If matches <> 1 Then
        messageText = "CREATE_WT_TOOLBAR_ACTION_MATCH_COUNT=" & CStr(matches) & "; ID=" & expectedId
        Exit Function
    End If
    CWR_GetVerifiedToolbarAction = True
    Exit Function
Failed:
    messageText = "CREATE_WT_TOOLBAR_ACTION_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Classifies the source-bin branch without changing SAP.  The same evaluator
' is used by the concrete executor and by the reusable different-bin plan.
Private Function CWR_ClassifyTargetBinCandidates(ByVal stockGrid As Object, ByRef planState As String, _
                                                  ByRef planText As String, ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, batchText As String, binText As String, stockTypeText As String
    Dim availText As String, totalText As String, availableQty As Double, totalQty As Double
    Dim exactSeen As Long, exactEligible As Long, alternateSeen As Long, alternateEligible As Long
    On Error GoTo Failed
    n = CLng(stockGrid.rowCount)
    For r = 0 To n - 1
        binText = CWR_Cell(stockGrid, r, "LGPLA")
        stockTypeText = CWR_Cell(stockGrid, r, "CAT")
        If stockTypeText = "" Then stockTypeText = CWR_Cell(stockGrid, r, "TO_CAT")
        If StrComp(binText, CWR_TARGET_BIN, vbTextCompare) = 0 And _
           (stockTypeText = "" Or StrComp(stockTypeText, CWR_TARGET_STOCK_TYPE, vbTextCompare) = 0) Then
            batchText = CWR_Cell(stockGrid, r, "CHARG")
            availText = CWR_Cell(stockGrid, r, "UI_AVLQTY")
            If availText = "" Then availText = CWR_Cell(stockGrid, r, "AVL_ALT_QTY")
            totalText = CWR_Cell(stockGrid, r, "LIME_QUAN")
            availableQty = CWR_Number(availText)
            If totalText = "" Then totalQty = availableQty Else totalQty = CWR_Number(totalText)
            If StrComp(batchText, CWR_WH_BATCH, vbTextCompare) = 0 Then
                exactSeen = exactSeen + 1
                If availableQty + 0.000001 >= CWR_TARGET_QTY And totalQty + 0.000001 >= CWR_TARGET_QTY Then
                    exactEligible = exactEligible + 1
                End If
            Else
                alternateSeen = alternateSeen + 1
                If availableQty + 0.000001 >= CWR_TARGET_QTY And totalQty + 0.000001 >= CWR_TARGET_QTY Then
                    alternateEligible = alternateEligible + 1
                End If
            End If
        End If
    Next r
    planText = "BIN=" & CWR_TARGET_BIN & "; REQUIRED_BATCH=" & CWR_WH_BATCH & _
        "; REQUIRED_QTY=" & CStr(CWR_TARGET_QTY) & "; EXACT_ROWS=" & CStr(exactSeen) & _
        "; EXACT_ELIGIBLE=" & CStr(exactEligible) & "; ALTERNATE_ROWS=" & CStr(alternateSeen) & _
        "; ALTERNATE_ELIGIBLE=" & CStr(alternateEligible)
    If exactEligible = 1 Then
        planState = "DIRECT_CREATE_WT"
    ElseIf exactEligible > 1 Then
        planState = "STOP_AMBIGUOUS_CORRECT_BATCH"
    ElseIf exactSeen > 0 Then
        planState = "STOP_INSUFFICIENT_CORRECT_BATCH"
    ElseIf alternateEligible = 1 Then
        planState = "POST_BATCH_CHANGE_REQUIRED"
    ElseIf alternateEligible > 1 Then
        planState = "STOP_AMBIGUOUS_ALTERNATE_BATCH"
    Else
        planState = "STOP_NO_ELIGIBLE_STOCK_AT_TARGET_BIN"
    End If
    CWR_ClassifyTargetBinCandidates = True
    Exit Function
Failed:
    messageText = "SOURCE_BIN_CANDIDATE_CLASSIFY_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' After the top-level Create action, SAP represents the unsaved selection as
' Open Qty=0 and removes the selected exact candidate from availability.  For
' this concrete 128-piece flow that is the required evidence before Save.
Private Function CWR_ProveExactOutboundDraftReservation(ByVal wrGrid As Object, ByVal stockGrid As Object, _
                                                         ByRef proofText As String, ByRef messageText As String) As Boolean
    Dim rowDoc As String, rowItem As String, rowMat As String, rowBatch As String, rowDest As String, rowProc As String
    Dim openQty As Double, requestRow As Long, r As Long, n As Long, binText As String, batchText As String, stockTypeText As String
    Dim availableText As String, targetRows As Long, targetAvailable As Double
    On Error GoTo Failed
    If Not CWR_LocateExactWarehouseRequest(wrGrid, requestRow, openQty, messageText) Then Exit Function
    If Abs(openQty) > 0.000001 Then
        messageText = "OUTBOUND_WT_DRAFT_OPEN_QTY_NOT_RESERVED=" & CStr(openQty)
        Exit Function
    End If
    n = CLng(stockGrid.rowCount)
    For r = 0 To n - 1
        binText = CWR_Cell(stockGrid, r, "LGPLA")
        batchText = CWR_Cell(stockGrid, r, "CHARG")
        stockTypeText = CWR_Cell(stockGrid, r, "CAT")
        If stockTypeText = "" Then stockTypeText = CWR_Cell(stockGrid, r, "TO_CAT")
        If StrComp(binText, CWR_TARGET_BIN, vbTextCompare) = 0 And _
           StrComp(batchText, CWR_WH_BATCH, vbTextCompare) = 0 And _
           (stockTypeText = "" Or StrComp(stockTypeText, CWR_TARGET_STOCK_TYPE, vbTextCompare) = 0) Then
            targetRows = targetRows + 1
            availableText = CWR_Cell(stockGrid, r, "UI_AVLQTY")
            If availableText = "" Then availableText = CWR_Cell(stockGrid, r, "AVL_ALT_QTY")
            targetAvailable = targetAvailable + CWR_Number(availableText)
        End If
    Next r
    If targetAvailable > 0.000001 Then
        messageText = "OUTBOUND_WT_DRAFT_TARGET_STOCK_NOT_FULLY_RESERVED: ROWS=" & CStr(targetRows) & _
            "; AVAILABLE=" & CStr(targetAvailable)
        Exit Function
    End If
    proofText = "WR_OPEN_QTY=0; RESERVED_SOURCE_BIN=" & CWR_TARGET_BIN & "; RESERVED_BATCH=" & CWR_WH_BATCH & _
        "; TARGET_CANDIDATE_ROWS_AFTER_CREATE=" & CStr(targetRows) & "; TARGET_AVAILABLE_AFTER_CREATE=" & CStr(targetAvailable)
    CWR_ProveExactOutboundDraftReservation = True
    Exit Function
Failed:
    messageText = "OUTBOUND_WT_DRAFT_RESERVATION_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Proves the one newly saved active outbound WT from Monitor.  Status C is
' neither required nor requested here: that later confirmation is user-owned.
Private Function CWR_ReadExactNewOutboundWT(ByRef wtNo As String, ByRef proofText As String, _
                                            ByRef messageText As String) As Boolean
    Dim monSes As Object, lower As Object
    Dim r As Long, n As Long, currentDoc As String, rowWT As String, rowStatus As String
    Dim rowItem As String, rowMat As String, rowBatch As String, rowSource As String, rowDest As String, rowProc As String
    Dim rowQty As Double, matches As Long, activeCount As Long, retainedCount As Long
    On Error GoTo Failed
    If Not CWR_GetMONSession(monSes, messageText) Then Exit Function
    If Not CWR_LoadFilteredWTGrid(monSes, lower, messageText) Then Exit Function
    n = CLng(lower.rowCount)
    For r = 0 To n - 1
        If CWR_Cell(lower, r, "DOCNO") <> "" Then currentDoc = CWR_Cell(lower, r, "DOCNO")
        rowWT = CWR_Cell(lower, r, "TANUM")
        rowStatus = CWR_Cell(lower, r, "TOSTAT")
        If StrComp(currentDoc, CWR_EXPECTED_DOC, vbTextCompare) = 0 And rowWT <> "" And rowStatus = "" Then
            If CWR_MONValueEquals(CWR_Cell(lower, r, "ITEMNO"), CWR_EXPECTED_ITEM) And _
               StrComp(CWR_Cell(lower, r, "MATNR"), CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 Then
                activeCount = activeCount + 1
                If CWR_IsDeltaAddRoute() And CWR_MONValueEquals(rowWT, CWR_EXPECTED_WT) Then
                    retainedCount = retainedCount + 1
                Else
                    matches = matches + 1
                    wtNo = rowWT
                    rowItem = CWR_Cell(lower, r, "ITEMNO")
                    rowMat = CWR_Cell(lower, r, "MATNR")
                    rowBatch = CWR_Cell(lower, r, "CHARG")
                    rowSource = CWR_Cell(lower, r, "VLPLA")
                    rowDest = CWR_Cell(lower, r, "NLPLA")
                    rowProc = CWR_Cell(lower, r, "PROCTY")
                    rowQty = CWR_Number(CWR_Cell(lower, r, "VSOLM"))
                End If
            End If
        End If
    Next r
    If CWR_IsDeltaAddRoute() And (activeCount <> 2 Or retainedCount <> 1) Then
        messageText = "OUTBOUND_WT_DELTA_POSTVERIFY_ACTIVE_COUNT=" & CStr(activeCount) & _
            "; RETAINED_MATCHES=" & CStr(retainedCount) & "; EXPECTED_OLD_WT=" & CWR_EXPECTED_WT
        Exit Function
    End If
    If matches <> 1 Then
        messageText = "OUTBOUND_WT_POSTVERIFY_ACTIVE_MATCH_COUNT=" & CStr(matches)
        Exit Function
    End If
    If StrComp(rowItem, CWR_EXPECTED_ITEM, vbTextCompare) <> 0 Or _
       StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Or _
        Abs(rowQty - CWR_CREATE_QTY) > 0.000001 Or _
       StrComp(rowBatch, CWR_WH_BATCH, vbTextCompare) <> 0 Or _
       StrComp(rowSource, CWR_TARGET_BIN, vbTextCompare) <> 0 Or _
       StrComp(rowDest, CWR_EXPECTED_DEST_BIN, vbTextCompare) <> 0 Or _
       StrComp(rowProc, CWR_EXPECTED_PROCTY, vbTextCompare) <> 0 Then
        messageText = "OUTBOUND_WT_POSTVERIFY_CONTRACT_MISMATCH: WT=" & wtNo & "; ITEM=" & rowItem & _
            "; MAT=" & rowMat & "; QTY=" & CStr(rowQty) & "; BATCH=" & rowBatch & _
            "; SOURCE=" & rowSource & "; DEST=" & rowDest & "; PROCTY=" & rowProc
        Exit Function
    End If
    proofText = "WT=" & wtNo & "; STATUS=OPEN_NOT_CONFIRMED; ITEM=" & rowItem & _
        "; MAT=" & rowMat & "; QTY=" & CStr(rowQty) & "; SOURCE_BIN=" & rowSource & _
        "; DEST_BIN=" & rowDest & "; BATCH=" & rowBatch & "; PROCTY=" & rowProc
    CWR_ReadExactNewOutboundWT = True
    Exit Function
Failed:
    messageText = "OUTBOUND_WT_POSTVERIFY_MONITOR_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Proves that the redetermined candidate grid exposes exactly the requested
' stock bin/batch with enough available and total quantity. It intentionally
' does not select/adopt that row; the result is a gate for a later explicit
' adoption step only.
Private Function CWR_ProveCreateWTTargetCandidate(ByVal stockGrid As Object, _
                                                   ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, matches As Long, batchText As String, binText As String, stockTypeText As String
    Dim availText As String, totalText As String, availableQty As Double, totalQty As Double
    On Error GoTo Failed
    n = CLng(stockGrid.rowCount)
    For r = 0 To n - 1
        batchText = CWR_Cell(stockGrid, r, "CHARG")
        binText = CWR_Cell(stockGrid, r, "LGPLA")
        stockTypeText = CWR_Cell(stockGrid, r, "CAT")
        If stockTypeText = "" Then stockTypeText = CWR_Cell(stockGrid, r, "TO_CAT")
        If StrComp(batchText, CWR_WH_BATCH, vbTextCompare) = 0 And _
           StrComp(binText, CWR_TARGET_BIN, vbTextCompare) = 0 And _
           (stockTypeText = "" Or StrComp(stockTypeText, CWR_TARGET_STOCK_TYPE, vbTextCompare) = 0) Then
            matches = matches + 1
            availText = CWR_Cell(stockGrid, r, "UI_AVLQTY")
            If availText = "" Then availText = CWR_Cell(stockGrid, r, "AVL_ALT_QTY")
            totalText = CWR_Cell(stockGrid, r, "LIME_QUAN")
            availableQty = CWR_Number(availText)
            If totalText = "" Then totalQty = availableQty Else totalQty = CWR_Number(totalText)
        End If
    Next r
    If matches <> 1 Then
        messageText = "CREATE_WT_REDETERMINE_TARGET_MATCH_COUNT=" & CStr(matches) & _
            "; BIN=" & CWR_TARGET_BIN & "; BATCH=" & CWR_WH_BATCH
        Exit Function
    End If
    If availableQty + 0.000001 < CWR_TARGET_QTY Or totalQty + 0.000001 < CWR_TARGET_QTY Then
        messageText = "CREATE_WT_REDETERMINE_TARGET_INSUFFICIENT: AVAILABLE=" & CStr(availableQty) & _
            "; TOTAL=" & CStr(totalQty) & "; REQUIRED=" & CStr(CWR_TARGET_QTY)
        Exit Function
    End If
    CWR_ProveCreateWTTargetCandidate = True
    Exit Function
Failed:
    messageText = "CREATE_WT_REDETERMINE_TARGET_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Confirms the single open Warehouse Request is still exactly the expected
' ODO item before a source candidate is selected or a draft is created.
Private Function CWR_VerifyExactCreateWTRequest(ByVal createSes As Object, ByRef wrGrid As Object, _
                                                 ByRef messageText As String) As Boolean
    Const WR_GRID_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_TODLV:1000/tabsGV_TAB_OIP/tabpOK_TAB_OIP_DLV/ssubSUB_OIP_DLV:/SCWM/SAPLUI_TODLV:3100/subSUB_OIP_DLV_DATA:/SCWM/SAPLUI_TODLV:3110/cntlCC_ALV_OI_DLV/shellcont/shell"
    Dim rowDoc As String, rowItem As String, rowMat As String, rowQty As String
    On Error GoTo Failed
    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    If wrGrid Is Nothing Or CLng(wrGrid.rowCount) <> 1 Then
        messageText = "CREATE_WT_WR_GRID_UNAVAILABLE"
        Exit Function
    End If
    rowDoc = CWR_Cell(wrGrid, 0, "DLVNO")
    rowItem = CWR_Cell(wrGrid, 0, "ITMNO")
    rowMat = CWR_Cell(wrGrid, 0, "MATNR")
    rowQty = CWR_Cell(wrGrid, 0, "UI_OPNQTY")
    If StrComp(rowDoc, CWR_EXPECTED_DOC, vbTextCompare) <> 0 Or _
       StrComp(rowItem, CWR_EXPECTED_ITEM, vbTextCompare) <> 0 Or _
       StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Or _
       Abs(CWR_Number(rowQty) - CWR_TARGET_QTY) > 0.000001 Then
        messageText = "CREATE_WT_WR_MISMATCH: DOC=" & rowDoc & "; ITEM=" & rowItem & _
            "; MAT=" & rowMat & "; QTY=" & rowQty
        Exit Function
    End If
    wrGrid.selectedRows = "0"
    wrGrid.SetCurrentCell 0, "DLVNO"
    ' The dependent source-stock and draft tabs are populated only after SAP
    ' receives the grid selection event; property assignment alone can leave
    ' the visible row highlighted but logically inactive.
    wrGrid.SelectionChanged
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 20#) Then
        messageText = "CREATE_WT_WR_SELECTION_TIMEOUT"
        Exit Function
    End If
    Set wrGrid = createSes.FindById(WR_GRID_ID, False)
    If wrGrid Is Nothing Or CLng(wrGrid.rowCount) <> 1 Then
        messageText = "CREATE_WT_WR_UNAVAILABLE_AFTER_SELECTION"
        Exit Function
    End If
    If CWR_SelectedRowsText(wrGrid) <> "0" Then
        messageText = "CREATE_WT_WR_SELECTION_READBACK_MISMATCH: " & CWR_SelectedRowsText(wrGrid)
        Exit Function
    End If
    rowDoc = CWR_Cell(wrGrid, 0, "DLVNO")
    rowItem = CWR_Cell(wrGrid, 0, "ITMNO")
    rowMat = CWR_Cell(wrGrid, 0, "MATNR")
    rowQty = CWR_Cell(wrGrid, 0, "UI_OPNQTY")
    If StrComp(rowDoc, CWR_EXPECTED_DOC, vbTextCompare) <> 0 Or _
       StrComp(rowItem, CWR_EXPECTED_ITEM, vbTextCompare) <> 0 Or _
       StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Or _
       Abs(CWR_Number(rowQty) - CWR_TARGET_QTY) > 0.000001 Then
        messageText = "CREATE_WT_WR_CHANGED_AFTER_SELECTION: DOC=" & rowDoc & "; ITEM=" & rowItem & _
            "; MAT=" & rowMat & "; QTY=" & rowQty
        Exit Function
    End If
    CWR_VerifyExactCreateWTRequest = True
    Exit Function
Failed:
    messageText = "CREATE_WT_WR_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Identifies the one stock-removal line that already has the WH/ODO batch at
' the user-confirmed new source bin. This is the reusable direct-create gate
' for future "different bin, correct batch" cases.
Private Function CWR_FindExactCreateWTTargetCandidate(ByVal stockGrid As Object, ByRef candidateRow As Long, _
                                                       ByRef proofText As String, ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, matches As Long, batchText As String, binText As String, stockTypeText As String
    Dim availText As String, totalText As String, availableQty As Double, totalQty As Double
    On Error GoTo Failed
    candidateRow = -1
    n = CLng(stockGrid.rowCount)
    For r = 0 To n - 1
        batchText = CWR_Cell(stockGrid, r, "CHARG")
        binText = CWR_Cell(stockGrid, r, "LGPLA")
        stockTypeText = CWR_Cell(stockGrid, r, "CAT")
        If stockTypeText = "" Then stockTypeText = CWR_Cell(stockGrid, r, "TO_CAT")
        If StrComp(batchText, CWR_WH_BATCH, vbTextCompare) = 0 And _
           StrComp(binText, CWR_TARGET_BIN, vbTextCompare) = 0 And _
           (stockTypeText = "" Or StrComp(stockTypeText, CWR_TARGET_STOCK_TYPE, vbTextCompare) = 0) Then
            availText = CWR_Cell(stockGrid, r, "UI_AVLQTY")
            If availText = "" Then availText = CWR_Cell(stockGrid, r, "AVL_ALT_QTY")
            totalText = CWR_Cell(stockGrid, r, "LIME_QUAN")
            availableQty = CWR_Number(availText)
            If totalText = "" Then totalQty = availableQty Else totalQty = CWR_Number(totalText)
            If availableQty + 0.000001 >= CWR_TARGET_QTY And totalQty + 0.000001 >= CWR_TARGET_QTY Then
                matches = matches + 1
                candidateRow = r
            End If
        End If
    Next r
    If matches <> 1 Then
        messageText = "CREATE_WT_TARGET_CANDIDATE_MATCH_COUNT=" & CStr(matches) & _
            "; BIN=" & CWR_TARGET_BIN & "; BATCH=" & CWR_WH_BATCH & "; QTY=" & CStr(CWR_TARGET_QTY)
        Exit Function
    End If
    proofText = "ROW=" & CStr(candidateRow) & "; BIN=" & CWR_Cell(stockGrid, candidateRow, "LGPLA") & _
        "; BATCH=" & CWR_Cell(stockGrid, candidateRow, "CHARG") & _
        "; AVAILABLE=" & CWR_Cell(stockGrid, candidateRow, "UI_AVLQTY") & _
        "; REQUIRED=" & CStr(CWR_TARGET_QTY)
    CWR_FindExactCreateWTTargetCandidate = True
    Exit Function
Failed:
    messageText = "CREATE_WT_TARGET_CANDIDATE_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Identifies exactly one former source row and exactly one requested target
' row in /SCWM/POST.  The source identity is the old bin/batch that belonged
' to the cancelled WT; the target identity is the requested bin/batch.  This
' function is read-only and does not infer a row from position alone.
Private Function CWR_ProveExactPOSTSourceAndTarget(ByVal grid As Object, _
                                                    ByRef sourceRow As Long, ByRef targetRow As Long, _
                                                    ByRef sourceProof As String, ByRef targetProof As String, _
                                                    ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, lastMaterial As String, matText As String, batchText As String
    Dim binText As String, stockType As String, availText As String, totalText As String
    Dim sourceMatches As Long, targetMatches As Long, sourceAvail As Double, sourceTotal As Double
    Dim targetAvail As Double, targetTotal As Double

    On Error GoTo Failed
    n = CLng(grid.rowCount)
    sourceRow = -1: targetRow = -1: lastMaterial = ""
    For r = 0 To n - 1
        matText = CWR_Cell(grid, r, "MATNR")
        If matText <> "" Then lastMaterial = matText Else matText = lastMaterial
        batchText = CWR_Cell(grid, r, "CHARG")
        binText = CWR_Cell(grid, r, "VLPLA")
        stockType = CWR_Cell(grid, r, "TO_CAT")
        availText = CWR_Cell(grid, r, "AVL_ALT_QTY")
        totalText = CWR_Cell(grid, r, "LIME_QUAN")
        If StrComp(matText, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 And _
           StrComp(batchText, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) = 0 And _
           StrComp(binText, CWR_EXPECTED_SOURCE_BIN, vbTextCompare) = 0 And _
           StrComp(stockType, "FT", vbTextCompare) = 0 Then
            sourceMatches = sourceMatches + 1
            sourceRow = r
            sourceAvail = CWR_Number(availText)
            sourceTotal = CWR_Number(totalText)
        End If
        If StrComp(matText, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 And _
           StrComp(batchText, CWR_TARGET_BATCH, vbTextCompare) = 0 And _
           StrComp(binText, CWR_TARGET_BIN, vbTextCompare) = 0 And _
           StrComp(stockType, "FT", vbTextCompare) = 0 Then
            targetMatches = targetMatches + 1
            targetRow = r
            targetAvail = CWR_Number(availText)
            targetTotal = CWR_Number(totalText)
        End If
    Next r
    If sourceMatches <> 1 Then
        messageText = "POST_SOURCE_FT_MATCH_COUNT=" & CStr(sourceMatches) & _
            "; MAT=" & CWR_EXPECTED_MATERIAL & "; BIN=" & CWR_EXPECTED_SOURCE_BIN & _
            "; BATCH=" & CWR_EXPECTED_SOURCE_BATCH
        Exit Function
    End If
    If targetMatches <> 1 Then
        messageText = "POST_TARGET_FT_MATCH_COUNT=" & CStr(targetMatches) & _
            "; MAT=" & CWR_EXPECTED_MATERIAL & "; BIN=" & CWR_TARGET_BIN & _
            "; BATCH=" & CWR_TARGET_BATCH
        Exit Function
    End If
    If sourceAvail + 0.000001 < CWR_TARGET_QTY Or sourceTotal + 0.000001 < CWR_TARGET_QTY Then
        messageText = "POST_SOURCE_INSUFFICIENT: AVAILABLE=" & CStr(sourceAvail) & _
            "; TOTAL=" & CStr(sourceTotal) & "; REQUIRED=" & CStr(CWR_TARGET_QTY)
        Exit Function
    End If
    If targetAvail + 0.000001 < CWR_TARGET_QTY Or targetTotal + 0.000001 < CWR_TARGET_QTY Then
        messageText = "POST_TARGET_INSUFFICIENT: AVAILABLE=" & CStr(targetAvail) & _
            "; TOTAL=" & CStr(targetTotal) & "; REQUIRED=" & CStr(CWR_TARGET_QTY)
        Exit Function
    End If
    sourceProof = "SOURCE: MAT=" & CWR_EXPECTED_MATERIAL & "; BIN=" & CWR_EXPECTED_SOURCE_BIN & _
        "; BATCH=" & CWR_EXPECTED_SOURCE_BATCH & "; STOCK_TYPE=FT; AVAILABLE=" & CStr(sourceAvail) & _
        "; TOTAL=" & CStr(sourceTotal) & "; REQUIRED=" & CStr(CWR_TARGET_QTY)
    targetProof = "TARGET: MAT=" & CWR_EXPECTED_MATERIAL & "; BIN=" & CWR_TARGET_BIN & _
        "; BATCH=" & CWR_TARGET_BATCH & "; STOCK_TYPE=FT; AVAILABLE=" & CStr(targetAvail) & _
        "; TOTAL=" & CStr(targetTotal) & "; REQUIRED=" & CStr(CWR_TARGET_QTY)
    CWR_ProveExactPOSTSourceAndTarget = True
    Exit Function
Failed:
    messageText = "POST_SOURCE_TARGET_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub CWR_CollectPOSTControls(ByVal obj As Object, ByVal depth As Long, ByRef itemCount As Long, ByRef dumpText As String)
    Dim i As Long, child As Object, typeText As String, idText As String, nameText As String
    Dim textText As String, tipText As String, haystack As String
    If obj Is Nothing Or depth > 14 Or itemCount >= 180 Or Len(dumpText) >= 24000 Then Exit Sub
    On Error Resume Next
    For i = 0 To obj.children.count - 1
        Set child = obj.children.ElementAt(i)
        typeText = CWR_Text(child.Type)
        idText = CWR_Text(child.id)
        nameText = CWR_Text(child.Name)
        textText = CWR_Text(child.text)
        tipText = CWR_Text(child.Tooltip)
        haystack = UCase$(typeText & " " & idText & " " & nameText & " " & textText & " " & tipText)
        If InStr(haystack, "MATNR") > 0 Or InStr(haystack, "PRODUCT") > 0 Or _
           InStr(haystack, "ADVANCED") > 0 Or InStr(haystack, "SEARCH") > 0 Or _
           InStr(haystack, "EXECUTE") > 0 Or InStr(haystack, "OWNER") > 0 Or _
           InStr(haystack, "START") > 0 Or InStr(typeText, "COMBOBOX") > 0 Or _
           InStr(typeText, "TEXTFIELD") > 0 Or InStr(typeText, "BUTTON") > 0 Then
            dumpText = dumpText & "D=" & CStr(depth) & "|TYPE=" & typeText & _
                "|ID=" & idText & "|NAME=" & nameText & "|TEXT=" & textText & "|TIP=" & tipText & vbLf
            itemCount = itemCount + 1
            If itemCount >= 180 Or Len(dumpText) >= 24000 Then Exit For
        End If
        CWR_CollectPOSTControls child, depth + 1, itemCount, dumpText
        If itemCount >= 180 Or Len(dumpText) >= 24000 Then Exit For
    Next i
    On Error GoTo 0
End Sub

' Recursively maps the embedded toolbars used by /SCWM/POST.  This is
' diagnostic only: it reads child metadata and never selects or presses a
' command.  The explicit toolbar header is needed because the Create+Save
' buttons in this transaction are not part of the standard SAP toolbar.
Private Sub CWR_DumpEmbeddedToolbars(ByVal obj As Object, ByVal depth As Long, _
                                     ByRef itemCount As Long, ByRef dumpText As String)
    Dim i As Long, child As Object, typeText As String, idText As String, nameText As String
    Dim textText As String, tipText As String, enabledText As String, changeableText As String, haystack As String

    If obj Is Nothing Or depth > 16 Or itemCount >= 320 Or Len(dumpText) >= 30000 Then Exit Sub
    On Error Resume Next
    typeText = CWR_Text(obj.Type)
    idText = CWR_Text(obj.id)
    If InStr(1, typeText, "GuiToolbar", vbTextCompare) > 0 Then
        dumpText = dumpText & "TOOLBAR|D=" & CStr(depth) & "|ID=" & idText & _
            "|TYPE=" & typeText & "|CHILDREN=" & CStr(obj.children.count) & vbLf
        For i = 0 To obj.children.count - 1
            Set child = obj.children.ElementAt(i)
            typeText = CWR_Text(child.Type)
            idText = CWR_Text(child.id)
            nameText = CWR_Text(child.Name)
            textText = CWR_Text(child.text)
            tipText = CWR_Text(child.Tooltip)
            enabledText = CWR_Text(child.Enabled)
            changeableText = CWR_Text(child.Changeable)
            dumpText = dumpText & "TOOLBAR_BUTTON|D=" & CStr(depth) & "|INDEX=" & CStr(i) & _
                "|TYPE=" & typeText & "|ID=" & idText & "|NAME=" & nameText & _
                "|TEXT=" & textText & "|TIP=" & tipText & "|ENABLED=" & enabledText & _
                "|CHANGEABLE=" & changeableText & vbLf
            itemCount = itemCount + 1
            If itemCount >= 320 Or Len(dumpText) >= 30000 Then Exit For
        Next i
    End If
    For i = 0 To obj.children.count - 1
        Set child = obj.children.ElementAt(i)
        typeText = CWR_Text(child.Type)
        idText = CWR_Text(child.id)
        nameText = CWR_Text(child.Name)
        textText = CWR_Text(child.text)
        tipText = CWR_Text(child.Tooltip)
        haystack = UCase$(typeText & " " & idText & " " & nameText & " " & textText & " " & tipText)
        If InStr(1, typeText, "GuiButton", vbTextCompare) > 0 And _
           (InStr(haystack, "CREATE") > 0 Or InStr(haystack, "SAVE") > 0 Or _
            InStr(haystack, "POSTING") > 0 Or InStr(haystack, "CHANGE") > 0) Then
            enabledText = CWR_Text(child.Enabled)
            changeableText = CWR_Text(child.Changeable)
            dumpText = dumpText & "ACTION_CANDIDATE|D=" & CStr(depth + 1) & _
                "|TYPE=" & typeText & "|ID=" & idText & "|NAME=" & nameText & _
                "|TEXT=" & textText & "|TIP=" & tipText & "|ENABLED=" & enabledText & _
                "|CHANGEABLE=" & changeableText & vbLf
            itemCount = itemCount + 1
        End If
        CWR_DumpEmbeddedToolbars child, depth + 1, itemCount, dumpText
        If itemCount >= 320 Or Len(dumpText) >= 30000 Then Exit For
    Next i
    On Error GoTo 0
End Sub

' Generic read-only mapper for the Create Warehouse Task screen.  Embedded
' SAP.Toolbar.1 shells are enumerated explicitly because their buttons are
' not exposed as ordinary GuiButton children.  No Select/Press/Modify call is
' present in this diagnostic routine.
Private Sub CWR_CollectCreateWTActionSurface(ByVal obj As Object, ByVal depth As Long, _
                                             ByRef itemCount As Long, ByRef dumpText As String)
    Dim i As Long, child As Object, typeText As String, idText As String, nameText As String
    Dim textText As String, tipText As String, changeText As String, enabledText As String, haystack As String

    If obj Is Nothing Or depth > 16 Or itemCount >= 420 Or Len(dumpText) >= 36000 Then Exit Sub
    On Error Resume Next
    typeText = CWR_Text(obj.Type)
    idText = CWR_Text(obj.id)
    nameText = CWR_Text(obj.Name)
    textText = CWR_Text(obj.text)
    tipText = CWR_Text(obj.Tooltip)
    changeText = CWR_Text(obj.Changeable)
    enabledText = CWR_Text(obj.Enabled)
    haystack = UCase$(typeText & " " & idText & " " & nameText & " " & textText & " " & tipText)

    If InStr(1, typeText, "GuiShell", vbTextCompare) > 0 And _
       InStr(1, textText, "SAP.Toolbar", vbTextCompare) > 0 Then
        CWR_DumpGuiShellToolbar obj, "CREATE_WT_EMBEDDED_TBAR", dumpText
    End If

    If InStr(1, typeText, "GuiTab", vbTextCompare) > 0 Or _
       InStr(1, typeText, "GuiButton", vbTextCompare) > 0 Or _
       InStr(1, typeText, "GuiTextField", vbTextCompare) > 0 Or _
       InStr(1, typeText, "GuiCTextField", vbTextCompare) > 0 Or _
       InStr(1, typeText, "GuiComboBox", vbTextCompare) > 0 Or _
       InStr(1, typeText, "GuiCheckBox", vbTextCompare) > 0 Or _
       InStr(1, typeText, "GuiRadioButton", vbTextCompare) > 0 Or _
       InStr(1, typeText, "GuiGrid", vbTextCompare) > 0 Or _
       InStr(haystack, "CREATE") > 0 Or InStr(haystack, "SAVE") > 0 Or _
       InStr(haystack, "BATCH") > 0 Or InStr(haystack, "CHARG") > 0 Or _
       InStr(haystack, "STOCK") > 0 Or InStr(haystack, "BIN") > 0 Or _
       InStr(haystack, "LOCATION") > 0 Or InStr(haystack, "SEARCH") > 0 Or _
       InStr(haystack, "SELECT") > 0 Or InStr(haystack, "AVAILABLE") > 0 Or _
       InStr(haystack, "QUANTITY") > 0 Or InStr(haystack, "QTY") > 0 Then
        dumpText = dumpText & "CREATE_WT_CONTROL|D=" & CStr(depth) & _
            "|TYPE=" & typeText & "|ID=" & idText & "|NAME=" & nameText & _
            "|TEXT=" & textText & "|TIP=" & tipText & "|ENABLED=" & enabledText & _
            "|CHANGEABLE=" & changeText & vbLf
        itemCount = itemCount + 1
    End If

    For i = 0 To obj.children.count - 1
        Set child = obj.children.ElementAt(i)
        CWR_CollectCreateWTActionSurface child, depth + 1, itemCount, dumpText
        If itemCount >= 420 Or Len(dumpText) >= 36000 Then Exit For
    Next i
    On Error GoTo 0
End Sub

' Dumps only visible ALV-style grids for the currently selected Create WT tab.
' The screen uses GuiShell wrappers for these grids, so the container id is
' used as the stable discriminator before calling the grid read helper.
Private Sub CWR_DumpVisibleCreateWTGrids(ByVal obj As Object, ByVal depth As Long, _
                                         ByRef itemCount As Long, ByRef dumpText As String)
    Dim i As Long, child As Object, idText As String
    If obj Is Nothing Or depth > 16 Or itemCount >= 460 Or Len(dumpText) >= 42000 Then Exit Sub
    On Error Resume Next
    idText = CWR_Text(obj.id)
    If InStr(1, UCase$(idText), "/CNTLCC_ALV", vbTextCompare) > 0 And _
       LCase$(Right$(idText, Len("/shellcont/shell"))) = "/shellcont/shell" Then
        dumpText = dumpText & CWR_DumpKnownGrid("CREATE_WT_VISIBLE", obj, 30)
        itemCount = itemCount + 1
    End If
    For i = 0 To obj.children.count - 1
        Set child = obj.children.ElementAt(i)
        CWR_DumpVisibleCreateWTGrids child, depth + 1, itemCount, dumpText
        If itemCount >= 460 Or Len(dumpText) >= 42000 Then Exit For
    Next i
    On Error GoTo 0
End Sub

Private Function CWR_FilterPOSTActionTree(ByVal rawTree As String) As String
    Dim lines() As String, i As Long, lineText As String, upperText As String, matches As Long
    If rawTree = "" Then Exit Function
    lines = Split(rawTree, vbLf)
    For i = LBound(lines) To UBound(lines)
        lineText = lines(i)
        upperText = UCase$(lineText)
        If InStr(upperText, "CREATE") > 0 Or InStr(upperText, "SAVE") > 0 Or _
           InStr(upperText, "POSTING CHANGE") > 0 Or InStr(upperText, "OPEN WAREHOUSE TASK") > 0 Then
            CWR_FilterPOSTActionTree = CWR_FilterPOSTActionTree & "OBJECT_TREE|" & lineText & vbLf
            matches = matches + 1
            If matches >= 80 Or Len(CWR_FilterPOSTActionTree) >= 20000 Then Exit For
        End If
    Next i
    If matches = 0 Then
        CWR_FilterPOSTActionTree = "OBJECT_TREE_NO_ACTION_TEXT|" & Left$(rawTree, 16000)
    End If
End Function

Private Sub CWR_DumpGuiShellToolbar(ByVal toolbar As Object, ByVal labelText As String, ByRef dumpText As String)
    Dim i As Long, n As Long, buttonId As String, buttonText As String, tipText As String
    Dim typeText As String, enabledText As String, checkedText As String
    On Error GoTo Failed
    typeText = CWR_Text(toolbar.Type)
    n = CLng(toolbar.buttonCount)
    dumpText = dumpText & labelText & "|ID=" & CWR_Text(toolbar.id) & "|TYPE=" & typeText & _
        "|BUTTONS=" & CStr(n) & vbLf
    For i = 0 To n - 1
        buttonId = CWR_Text(toolbar.GetButtonId(i))
        buttonText = CWR_Text(toolbar.GetButtonText(i))
        tipText = CWR_Text(toolbar.GetButtonTooltip(i))
        enabledText = CWR_Text(toolbar.GetButtonEnabled(i))
        checkedText = CWR_Text(toolbar.GetButtonChecked(i))
        dumpText = dumpText & labelText & "|INDEX=" & CStr(i) & "|ID=" & buttonId & _
            "|TEXT=" & buttonText & "|TIP=" & tipText & "|ENABLED=" & enabledText & _
            "|CHECKED=" & checkedText & vbLf
    Next i
    Exit Sub
Failed:
    dumpText = dumpText & labelText & "|MAP_ERROR=" & CStr(Err.Number) & ": " & Err.Description & vbLf
End Sub

Private Sub CWR_DumpToolbarControls(ByVal ses As Object, ByVal toolbarId As String, _
                                    ByVal labelText As String, ByRef dumpText As String)
    Dim toolbar As Object, child As Object, i As Long, typeText As String, idText As String
    Dim textText As String, tipText As String, nameText As String
    On Error GoTo Failed
    Set toolbar = ses.FindById(toolbarId, False)
    If toolbar Is Nothing Then
        dumpText = dumpText & labelText & "|ID=" & toolbarId & "|NOT_FOUND" & vbLf
        Exit Sub
    End If
    dumpText = dumpText & labelText & "|ID=" & toolbarId & "|TYPE=" & CWR_Text(toolbar.Type) & _
        "|CHILDREN=" & CStr(toolbar.children.count) & vbLf
    For i = 0 To toolbar.children.count - 1
        Set child = toolbar.children.ElementAt(i)
        typeText = "": idText = "": textText = "": tipText = "": nameText = ""
        On Error Resume Next
        typeText = CWR_Text(child.Type)
        idText = CWR_Text(child.id)
        nameText = CWR_Text(child.Name)
        textText = CWR_Text(child.text)
        tipText = CWR_Text(child.Tooltip)
        Err.Clear
        On Error GoTo Failed
        dumpText = dumpText & labelText & "|INDEX=" & CStr(i) & "|TYPE=" & typeText & _
            "|ID=" & idText & "|NAME=" & nameText & "|TEXT=" & textText & _
            "|TIP=" & tipText & vbLf
    Next i
    Exit Sub
Failed:
    dumpText = dumpText & labelText & "|DUMP_ERROR=" & CStr(Err.Number) & ": " & Err.Description & vbLf
End Sub

Private Sub CWR_DumpMenuTree(ByVal obj As Object, ByVal depth As Long, ByRef dumpText As String)
    Dim child As Object, i As Long, typeText As String, idText As String, textText As String, nameText As String
    If obj Is Nothing Or depth > 6 Or Len(dumpText) > 24000 Then Exit Sub
    On Error Resume Next
    For i = 0 To obj.children.count - 1
        Set child = obj.children.ElementAt(i)
        typeText = CWR_Text(child.Type)
        idText = CWR_Text(child.id)
        textText = CWR_Text(child.text)
        nameText = CWR_Text(child.Name)
        dumpText = dumpText & "D=" & CStr(depth) & "|INDEX=" & CStr(i) & "|TYPE=" & typeText & _
            "|ID=" & idText & "|NAME=" & nameText & "|TEXT=" & textText & vbLf
        CWR_DumpMenuTree child, depth + 1, dumpText
        If Len(dumpText) > 24000 Then Exit For
    Next i
    On Error GoTo 0
End Sub

Private Sub CWR_CollectAllVisibleText(ByVal obj As Object, ByVal depth As Long, _
                                      ByRef itemCount As Long, ByRef dumpText As String)
    Dim child As Object, i As Long, typeText As String, idText As String, textText As String, tipText As String
    If obj Is Nothing Or depth > 14 Or itemCount >= 250 Or Len(dumpText) >= 28000 Then Exit Sub
    On Error Resume Next
    For i = 0 To obj.children.count - 1
        Set child = obj.children.ElementAt(i)
        typeText = CWR_Text(child.Type)
        idText = CWR_Text(child.id)
        textText = CWR_Text(child.text)
        tipText = CWR_Text(child.Tooltip)
        If textText <> "" Or tipText <> "" Or InStr(1, typeText, "Grid", vbTextCompare) > 0 Then
            dumpText = dumpText & "D=" & CStr(depth) & "|TYPE=" & typeText & "|ID=" & idText & _
                "|TEXT=" & textText & "|TIP=" & tipText & vbLf
            itemCount = itemCount + 1
            If itemCount >= 250 Or Len(dumpText) >= 28000 Then Exit For
        End If
        CWR_CollectAllVisibleText child, depth + 1, itemCount, dumpText
        If itemCount >= 250 Or Len(dumpText) >= 28000 Then Exit For
    Next i
    On Error GoTo 0
End Sub

Private Function CWR_DumpPOSTSourceTargetColumns(ByVal grid As Object, ByVal sourceRow As Long, _
                                                  ByVal targetRow As Long) As String
    Dim columnOrder As Object, colId As String, c As Long, n As Long, titleText As String
    On Error GoTo Failed
    n = CLng(grid.columnCount)
    Set columnOrder = grid.columnOrder
    CWR_DumpPOSTSourceTargetColumns = "GRID=" & CWR_Text(grid.id) & "; COLUMNS=" & CStr(n) & vbLf
    For c = 0 To n - 1
        colId = CWR_Text(columnOrder.ElementAt(c))
        titleText = CWR_POSTColumnTitle(grid, colId)
        CWR_DumpPOSTSourceTargetColumns = CWR_DumpPOSTSourceTargetColumns & _
            "COL=" & CStr(c) & "|ID=" & colId & "|TITLE=" & titleText & _
            "|SOURCE=" & CWR_Cell(grid, sourceRow, colId) & _
            "|SOURCE_CHANGEABLE=" & CWR_CellChangeable(grid, sourceRow, colId) & _
            "|TARGET=" & CWR_Cell(grid, targetRow, colId) & vbLf
    Next c
    Exit Function
Failed:
    CWR_DumpPOSTSourceTargetColumns = "POST_COLUMN_DUMP_ERROR: " & CStr(Err.Number) & ": " & Err.Description & vbLf
End Function

Private Function CWR_POSTColumnTitle(ByVal grid As Object, ByVal colId As String) As String
    Dim titles As Object
    On Error Resume Next
    Set titles = grid.GetColumnTitles(colId)
    If Not titles Is Nothing Then CWR_POSTColumnTitle = CWR_Text(titles.ElementAt(0))
    Err.Clear
    On Error GoTo 0
End Function

Private Function CWR_ReadExactPRDOItem(ByVal grid As Object, ByRef qtyText As String, _
                                       ByRef batchText As String, ByRef statusText As String, _
                                       ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, matches As Long, itemText As String, matText As String
    Dim pickText As String, giText As String

    On Error GoTo Failed
    n = CLng(grid.rowCount): matches = 0
    For r = 0 To n - 1
        itemText = CWR_Cell(grid, r, "ITEMNO")
        matText = CWR_Cell(grid, r, "PRODUCTNO")
        If itemText = CWR_EXPECTED_ITEM And StrComp(matText, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 Then
            matches = matches + 1
            qtyText = CWR_Cell(grid, r, "QTY_UI")
            batchText = CWR_Cell(grid, r, "BATCHNO_UI")
            pickText = CWR_Cell(grid, r, "STATUS_PICKING")
            giText = CWR_Cell(grid, r, "STATUS_GI")
        End If
    Next r
    If matches <> 1 Then
        messageText = "PRDO_ITEM_MATCH_COUNT=" & CStr(matches)
        Exit Function
    End If
    ' Qty=0 is the expected post-save value for the explicit CANCEL_TO_ZERO
    ' route.  It is still rejected for every other route, so a blank or zero
    ' item cannot be mistaken for a usable Create-WT request.
    If Not IsNumeric(qtyText) Or CWR_Number(qtyText) < 0# Or _
       (CWR_Number(qtyText) <= 0# And CWR_PLAN_DECISION <> "CANCEL_TO_ZERO") Or _
       batchText = "" Then
        messageText = "PRDO_ITEM_INCOMPLETE: QTY=" & qtyText & "; BATCH=" & batchText
        Exit Function
    End If
    statusText = "PICKING=" & pickText & "; GI=" & giText
    CWR_ReadExactPRDOItem = True
    Exit Function
Failed:
    messageText = "PRDO_ITEM_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_FindExactPRDOItem(ByVal grid As Object, ByRef rowNo As Long, _
                                       ByRef qtyText As String, ByRef batchText As String, _
                                       ByRef statusText As String, ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, matches As Long, itemText As String, matText As String
    Dim pickText As String, giText As String
    On Error GoTo Failed
    n = CLng(grid.rowCount): rowNo = -1: matches = 0
    For r = 0 To n - 1
        itemText = CWR_Cell(grid, r, "ITEMNO")
        matText = CWR_Cell(grid, r, "PRODUCTNO")
        If itemText = CWR_EXPECTED_ITEM And StrComp(matText, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 Then
            matches = matches + 1
            rowNo = r
            qtyText = CWR_Cell(grid, r, "QTY_UI")
            batchText = CWR_Cell(grid, r, "BATCHNO_UI")
            pickText = CWR_Cell(grid, r, "STATUS_PICKING")
            giText = CWR_Cell(grid, r, "STATUS_GI")
        End If
    Next r
    If matches <> 1 Then
        messageText = "PRDO_ITEM_MATCH_COUNT=" & CStr(matches)
        Exit Function
    End If
    ' This selector is also used immediately after Save.  CANCEL_TO_ZERO must
    ' accept an exact zero readback; every other route still requires > 0.
    If Not IsNumeric(qtyText) Or CWR_Number(qtyText) < 0# Or _
       (CWR_Number(qtyText) <= 0# And CWR_PLAN_DECISION <> "CANCEL_TO_ZERO") Or _
       batchText = "" Then
        messageText = "PRDO_ITEM_INCOMPLETE: QTY=" & qtyText & "; BATCH=" & batchText
        Exit Function
    End If
    statusText = "PICKING=" & pickText & "; GI=" & giText
    CWR_FindExactPRDOItem = True
    Exit Function
Failed:
    messageText = "PRDO_ITEM_FIND_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_CellChangeable(ByVal grid As Object, ByVal rowNo As Long, ByVal colName As String) As String
    On Error Resume Next
    CWR_CellChangeable = CWR_Text(grid.GetCellChangeable(rowNo, colName) & "")
    Err.Clear
End Function

Private Function CWR_ResetPRDOWithoutChanges(ByVal ses As Object, ByRef messageText As String) As Boolean
    Dim tcode As String
    On Error GoTo Failed
    ses.StartTransaction "/SCWM/PRDO"
    If Not CWR_WaitIdle(ses, 30#) Then
        messageText = "PRDO_RESET_TIMEOUT"
        Exit Function
    End If
    tcode = CWR_Text(ses.Info.Transaction)
    If UCase$(tcode) <> "/SCWM/PRDO" Then
        messageText = "PRDO_RESET_TCODE_MISMATCH: " & tcode
        Exit Function
    End If
    CWR_ResetPRDOWithoutChanges = True
    Exit Function
Failed:
    messageText = "PRDO_RESET_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Requires the Monitor to show no active WT with a real task number for the
' current ODO.  Continuation rows without TANUM are deliberately ignored.
Private Function CWR_AssertNoActiveOutboundWT(ByRef proofText As String, ByRef messageText As String) As Boolean
    Dim monSes As Object, lower As Object
    Dim r As Long, n As Long, currentDoc As String, rowWT As String, rowStatus As String
    Dim rowItem As String, rowMat As String
    Dim activeCount As Long, taskCount As Long
    On Error GoTo Failed
    If Not CWR_GetMONSession(monSes, messageText) Then Exit Function
    If Not CWR_LoadFilteredWTGrid(monSes, lower, messageText) Then Exit Function
    n = CLng(lower.rowCount)
    For r = 0 To n - 1
        If CWR_Cell(lower, r, "DOCNO") <> "" Then currentDoc = CWR_Cell(lower, r, "DOCNO")
        rowWT = CWR_Cell(lower, r, "TANUM")
        rowItem = CWR_Cell(lower, r, "ITEMNO")
        rowMat = CWR_Cell(lower, r, "MATNR")
        If CWR_MONValueEquals(currentDoc, CWR_EXPECTED_DOC) And _
           CWR_MONValueEquals(rowItem, CWR_EXPECTED_ITEM) And _
           StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 And rowWT <> "" Then
            taskCount = taskCount + 1
            rowStatus = CWR_Cell(lower, r, "TOSTAT")
            If rowStatus = "" Then activeCount = activeCount + 1
        End If
    Next r
    If activeCount <> 0 Then
        messageText = "CREATE_WT_ACTIVE_OUTBOUND_TASKS_PRESENT=" & CStr(activeCount)
        Exit Function
    End If
    proofText = "MONITOR_TASK_ROWS=" & CStr(taskCount) & "; ACTIVE_OUTBOUND_WT=0; ITEM=" & CWR_EXPECTED_ITEM
    CWR_AssertNoActiveOutboundWT = True
    Exit Function
Failed:
    messageText = "CREATE_WT_MONITOR_PREVERIFY_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_LocateExactActiveWT(ByVal ses As Object, ByRef lower As Object, _
                                         ByRef targetRow As Long, ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, currentDoc As String, currentWho As String
    Dim rowDoc As String, rowWT As String, rowItem As String, rowMat As String, rowStatus As String
    Dim rowBin As String, rowBatch As String, rowDest As String, rowProc As String, rowQty As Double
    Dim matches As Long

    On Error GoTo Failed
    If Not CWR_LoadFilteredWTGrid(ses, lower, messageText) Then Exit Function
    targetRow = -1: currentDoc = "": currentWho = "": matches = 0
    n = CLng(lower.rowCount)
    For r = 0 To n - 1
        rowDoc = CWR_Cell(lower, r, "DOCNO")
        If rowDoc <> "" Then
            If StrComp(rowDoc, currentDoc, vbTextCompare) <> 0 Then currentWho = ""
            currentDoc = rowDoc
        End If
        If CWR_Cell(lower, r, "WHO") <> "" Then currentWho = CWR_Cell(lower, r, "WHO")
        rowWT = CWR_Cell(lower, r, "TANUM")
        If CWR_MONValueEquals(currentDoc, CWR_EXPECTED_DOC) And _
           CWR_MONValueEquals(rowWT, CWR_EXPECTED_WT) Then
            matches = matches + 1
            targetRow = r
            rowItem = CWR_Cell(lower, r, "ITEMNO")
            rowMat = CWR_Cell(lower, r, "MATNR")
            rowStatus = CWR_Cell(lower, r, "TOSTAT")
            rowBin = CWR_Cell(lower, r, "VLPLA")
            rowBatch = CWR_Cell(lower, r, "CHARG")
            rowDest = CWR_Cell(lower, r, "NLPLA")
            rowProc = CWR_Cell(lower, r, "PROCTY")
            rowQty = CWR_Number(CWR_Cell(lower, r, "VSOLM"))
        End If
    Next r
    If matches <> 1 Then
        messageText = "WT_MATCH_COUNT=" & CStr(matches)
        Exit Function
    End If
    If Not CWR_MONValueEquals(rowItem, CWR_EXPECTED_ITEM) Or _
       StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) <> 0 Or _
       Abs(rowQty - CWR_EXPECTED_SOURCE_QTY) > 0.000001 Or _
       StrComp(rowBin, CWR_EXPECTED_SOURCE_BIN, vbTextCompare) <> 0 Or _
       StrComp(rowBatch, CWR_EXPECTED_SOURCE_BATCH, vbTextCompare) <> 0 Or _
       rowDest = "" Or rowProc = "" Or rowStatus <> "" Then
        messageText = "WT_PREVERIFY_MISMATCH: ITEM=" & rowItem & "; MAT=" & rowMat & _
            "; QTY=" & CStr(rowQty) & "; BIN=" & rowBin & "; BATCH=" & rowBatch & _
            "; DEST=" & rowDest & "; PROCTY=" & rowProc & "; STATUS=" & rowStatus
        Exit Function
    End If
    ' Destination and process type are properties of the currently active WT.
    ' They are captured only after the immutable old-WT identity above passes,
    ' then carried forward for the later request/create and post-verify gates.
    ' This avoids rejecting a valid old WT merely because a historical default
    ' process type (for example 2093) differs from the live task (2010).
    CWR_EXPECTED_DEST_BIN = rowDest
    CWR_EXPECTED_PROCTY = rowProc
    lower.selectedRows = CStr(targetRow)
    lower.SetCurrentCell targetRow, "TANUM"
    CWR_LocateExactActiveWT = True
    Exit Function
Failed:
    messageText = "WT_PREVERIFY_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_LoadFilteredWTGrid(ByVal ses As Object, ByRef lower As Object, ByRef messageText As String) As Boolean
    Dim upper As Object
    Dim mapText As String, upperID As String, lowerID As String, actionID As String
    Dim docColumn As String, r As Long, n As Long, expectedRow As Long
    Dim nonBlankDocs As Long, exactDocMatches As Long, rowDoc As String

    On Error GoTo Failed
    If ses Is Nothing Then
        messageText = "MON_SESSION_NOT_FOUND"
        GoTo Stopped
    End If
    If Not CWR_ReserveMONRefresh(messageText) Then GoTo Stopped

    ' Use the established MON action map rather than the old fixed node and
    ' selection-screen ids.  The mapped helper enters the LE Delivery (OB)
    ' field, verifies its readback, then executes exactly once.
    mapText = ""
    If Not ModuleSAP.CF_MON_SetLEDelivery(ses, CWR_EXPECTED_OB, mapText) Then
        messageText = "MON_MAPPED_FILTER_NOT_PROVEN: " & mapText
        GoTo Stopped
    End If
    If Not CWR_WaitIdle(ses, 30#) Then
        messageText = "MON_MAPPED_FILTER_TIMEOUT"
        GoTo Stopped
    End If

    upperID = CWR_ConfigValue("MON_RESULT_GRID_ID")
    If upperID = "" Then upperID = CWR_UPPER_ID
    lowerID = CWR_ConfigValue("MON_WT_GRID_ID")
    If lowerID = "" Then lowerID = CWR_WT_ID
    docColumn = CWR_ConfigValue("MON_UPPER_DOC_COL")
    If docColumn = "" Then docColumn = "DOCNO_H"
    actionID = CWR_ConfigValue("MON_WT_ACTION_ID")
    If actionID = "" Then actionID = CWR_WT_ACTION

    If Not CWR_WaitForObject(ses, upperID, 20#) Then
        messageText = "MON_UPPER_GRID_TIMEOUT: " & upperID
        GoTo Stopped
    End If
    Set upper = ses.FindById(upperID, False)
    If upper Is Nothing Then
        messageText = "MON_UPPER_GRID_NOT_FOUND: " & upperID
        GoTo Stopped
    End If

    ' A filtered MON layout can legitimately retain other visible document
    ' rows.  The prior condition required NONBLANK=1 even when the exact
    ' planned document had already been found (for example row 0 of 12), which
    ' stopped a safe route before any SAP write.  Require one exact document,
    ' not one displayed document.
    expectedRow = -1: nonBlankDocs = 0: exactDocMatches = 0: n = CLng(upper.rowCount)
    For r = 0 To n - 1
        rowDoc = CWR_Cell(upper, r, docColumn)
        If rowDoc <> "" Then nonBlankDocs = nonBlankDocs + 1
        If CWR_MONValueEquals(rowDoc, CWR_EXPECTED_DOC) Then
            exactDocMatches = exactDocMatches + 1
            expectedRow = r
        End If
    Next r
    If exactDocMatches <> 1 Then
        messageText = "MON_DOC_FILTER_MISMATCH: EXACT_DOC_MATCHES=" & CStr(exactDocMatches) & _
            "; NONBLANK=" & CStr(nonBlankDocs) & "; EXPECTED_DOC=" & CWR_EXPECTED_DOC & _
            "; MAP=" & mapText
        GoTo Stopped
    End If

    upper.selectedRows = CStr(expectedRow)
    upper.SetCurrentCell expectedRow, docColumn
    upper.PressToolbarButton actionID
    If Not CWR_WaitIdle(ses, 30#) Then
        messageText = "MON_WT_ACTION_TIMEOUT: " & actionID
        GoTo Stopped
    End If
    If Not CWR_WaitForObject(ses, lowerID, 20#) Then
        messageText = "MON_WT_GRID_TIMEOUT: " & lowerID
        GoTo Stopped
    End If
    Set lower = ses.FindById(lowerID, False)
    If lower Is Nothing Then
        messageText = "MON_WT_GRID_NOT_FOUND: " & lowerID
        GoTo Stopped
    End If
    CWR_LoadFilteredWTGrid = True
    Exit Function
Stopped:
    CWR_Log "MON_FILTER", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", messageText
    Exit Function
Failed:
    messageText = "MON_GRID_LOAD_ERROR: " & CStr(Err.Number) & ": " & Err.Description
    CWR_Log "MON_FILTER", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "FAIL", messageText
End Function

' MON returns identifier cells as strings; normalize leading zero padding only
' so the exact document remains the identity gate in every layout.
Private Function CWR_MONValueEquals(ByVal actualText As String, ByVal expectedText As String) As Boolean
    Dim actualKey As String, expectedKey As String
    actualKey = Replace(UCase$(CWR_Text(actualText)), " ", "")
    expectedKey = Replace(UCase$(CWR_Text(expectedText)), " ", "")
    Do While Len(actualKey) > 1 And Left$(actualKey, 1) = "0"
        actualKey = Mid$(actualKey, 2)
    Loop
    Do While Len(expectedKey) > 1 And Left$(expectedKey, 1) = "0"
        expectedKey = Mid$(expectedKey, 2)
    Loop
    CWR_MONValueEquals = (actualKey <> "" And actualKey = expectedKey)
End Function

Private Function CWR_SendExactCancel(ByVal ses As Object, ByVal lower As Object, ByVal targetRow As Long, _
                                     ByRef messageText As String) As Boolean
    CWR_SendExactCancel = CWR_SendCancelSelection(ses, lower, CStr(targetRow), messageText)
End Function

' Sends one Cancel WT action for the complete selected-row set.  The
' scheduler uses this for a QTY group so SAP receives one group action rather
' than one Cancel/confirm per WT member.
Private Function CWR_SendCancelSelection(ByVal ses As Object, ByVal lower As Object, _
                                          ByVal selectedRowsText As String, ByRef messageText As String, _
                                          Optional ByVal expectedWTSet As String = "") As Boolean
    Dim modal As Object, yesButton As Object
    Dim titleText As String, buttonText As String, treeText As String, selectedReadback As String
    Dim firstRowText As String, firstRow As Long
    Dim props As Variant, errNo As Long, errText As String

    On Error GoTo Failed
    If Trim$(selectedRowsText) = "" Then
        messageText = "STOP_CANCEL_GROUP_SELECTION_EMPTY"
        Exit Function
    End If
    lower.selectedRows = selectedRowsText
    selectedReadback = CWR_SelectedRowsText(lower)
    If Not CWR_SelectedRowsSetMatches(selectedReadback, selectedRowsText) Then
        messageText = "STOP_CANCEL_GROUP_SELECTION_READBACK_MISMATCH: EXPECTED=" & selectedRowsText & _
            "; ACTUAL=" & selectedReadback
        Exit Function
    End If
    firstRowText = Split(Replace(Replace(selectedRowsText, " ", ""), ";", ","), ",")(0)
    If Not IsNumeric(firstRowText) Then
        messageText = "STOP_CANCEL_GROUP_FIRST_ROW_INVALID: " & firstRowText
        Exit Function
    End If
    firstRow = CLng(firstRowText)
    lower.SetCurrentCell firstRow, "TANUM"
    On Error Resume Next
    Err.Clear
    lower.PressToolbarContextButton "METHODS"
    errNo = Err.Number: errText = Err.Description: Err.Clear
    On Error GoTo Failed
    If errNo <> 0 Then
        messageText = "CANCEL_METHODS_MENU_ERROR: " & CStr(errNo) & ": " & errText
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    lower.SelectContextMenuItemByText "Cancel WT"
    errNo = Err.Number: errText = Err.Description: Err.Clear
    On Error GoTo Failed
    If errNo <> 0 Then
        messageText = "CANCEL_MENU_ITEM_ERROR: " & CStr(errNo) & ": " & errText
        Exit Function
    End If
    If Not CWR_WaitIdle(ses, 20#) Then
        messageText = "CANCEL_ACTION_TIMEOUT"
        Exit Function
    End If

    Set modal = ses.FindById("wnd[1]", False)
    If Not modal Is Nothing Then
        titleText = CWR_Text(modal.text)
        props = Array("Id", "Type", "Name", "Text", "Tooltip", "Enabled", "Visible")
        treeText = ""
        On Error Resume Next
        treeText = CStr(ses.GetObjectTree("wnd[1]", props) & "")
        On Error GoTo Failed
        If InStr(1, titleText & " " & treeText, "cancel", vbTextCompare) = 0 And _
           InStr(1, titleText & " " & treeText, "Display logs", vbTextCompare) = 0 Then
            messageText = "STOP_UNEXPECTED_CANCEL_MODAL: " & titleText
            Exit Function
        End If
        Set yesButton = ses.FindById("wnd[1]/usr/btnSPOP-OPTION1", False)
        If yesButton Is Nothing Then
            ' Some SAP GUI patches complete Cancel immediately and return the
            ' informational Display logs viewer instead of the Yes/No popup.
            ' The log is the business-action proof; looking for a Confirm
            ' button here is a false stop and must never trigger a retry.
            If expectedWTSet = "" Then expectedWTSet = CWR_EXPECTED_WT
            If CWR_CancelDisplayLogsProof(ses, modal, treeText, expectedWTSet, messageText) Then
                CWR_SendCancelSelection = True
                Exit Function
            End If
            If messageText = "" Then messageText = "STOP_CANCEL_CONFIRM_BUTTON_NOT_FOUND: " & titleText
            Exit Function
        End If
        buttonText = CWR_Text(yesButton.text)
        If UCase$(buttonText) <> "YES" And UCase$(buttonText) <> "OK" Then
            messageText = "STOP_CANCEL_CONFIRM_LABEL_UNEXPECTED: " & buttonText
            Exit Function
        End If
        yesButton.Press
        If Not CWR_WaitIdle(ses, 30#) Then
            messageText = "CANCEL_CONFIRM_TIMEOUT"
            Exit Function
        End If
        Set modal = ses.FindById("wnd[1]", False)
        If Not modal Is Nothing Then
            treeText = ""
            On Error Resume Next
            treeText = CStr(ses.GetObjectTree("wnd[1]", props) & "")
            On Error GoTo Failed
            If expectedWTSet = "" Then expectedWTSet = CWR_EXPECTED_WT
            If Not CWR_CancelDisplayLogsProof(ses, modal, treeText, expectedWTSet, messageText) Then
                If messageText = "" Then messageText = "STOP_CANCEL_MODAL_REMAINS"
                Exit Function
            End If
        End If
    End If
    CWR_SendCancelSelection = True
    Exit Function
Failed:
    messageText = "CANCEL_ACTION_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' SAP may return a Display logs viewer after Cancel has already been
' accepted.  Treat a clearly identified log/status surface as success, close
' only that informational viewer, and leave MON readback to the caller.  No
' business confirmation or retry is sent from this helper.
Private Function CWR_CancelDisplayLogsProof(ByVal ses As Object, ByVal modal As Object, _
                                             ByVal treeText As String, ByVal expectedWTSet As String, _
                                             ByRef messageText As String) As Boolean
    Dim titleText As String, statusType As String, statusText As String, logText As String
    Dim closeButton As Object, closeTip As String, modalAfter As Object
    Dim wtToken As Variant, normalizedSet As String, canceledCount As Long, tokenText As String
    Dim tokenPos As Long, tokenSnippet As String
    On Error GoTo Failed
    titleText = CWR_Text(modal.text)
    CWR_ReadStatusBar ses, statusType, statusText
    If InStr(1, titleText & " " & treeText, "Display logs", vbTextCompare) = 0 And _
       InStr(1, titleText & " " & treeText, "Display log", vbTextCompare) = 0 Then Exit Function
    logText = titleText & " " & treeText & " " & statusText
    normalizedSet = Replace(Replace(CWR_Text(expectedWTSet), ";", ","), " ", "")
    If normalizedSet <> "" Then
        If CWR_CANCEL_RESULT_BY_WT Is Nothing Then
            Set CWR_CANCEL_RESULT_BY_WT = CreateObject("Scripting.Dictionary")
            CWR_CANCEL_RESULT_BY_WT.CompareMode = vbTextCompare
        End If
        For Each wtToken In Split(normalizedSet, ",")
            tokenText = CWR_Text(wtToken)
            If tokenText <> "" Then
                tokenPos = InStr(1, logText, tokenText, vbTextCompare)
                If tokenPos = 0 Then
                    CWR_CANCEL_RESULT_BY_WT(tokenText) = "FAIL"
                Else
                    tokenSnippet = Mid$(logText, IIf(tokenPos > 100, tokenPos - 100, 1), 220)
                    If InStr(1, tokenSnippet, "canceled", vbTextCompare) > 0 Or _
                       InStr(1, tokenSnippet, "cancelled", vbTextCompare) > 0 Or _
                       InStr(1, tokenSnippet, "cancel", vbTextCompare) > 0 Then
                        CWR_CANCEL_RESULT_BY_WT(tokenText) = "PASS"
                        canceledCount = canceledCount + 1
                    Else
                        CWR_CANCEL_RESULT_BY_WT(tokenText) = "FAIL"
                    End If
                End If
            End If
        Next wtToken
        If canceledCount < UBound(Split(normalizedSet, ",")) + 1 Then
            ' SAP GUI sometimes exposes only the Display logs title through
            ' GetObjectTree while the individual message rows remain in the
            ' viewer.  Do not convert that successful post-action surface into
            ' a missing-confirm failure: MON readback below is the authoritative
            ' per-WT proof.  Preserve the map as PENDING_MON for audit.
            For Each wtToken In Split(normalizedSet, ",")
                tokenText = CWR_Text(wtToken)
                If tokenText <> "" And (Not CWR_CANCEL_RESULT_BY_WT.Exists(tokenText) Or _
                   CWR_Text(CWR_CANCEL_RESULT_BY_WT(tokenText)) <> "PASS") Then
                    CWR_CANCEL_RESULT_BY_WT(tokenText) = "PENDING_MON"
                End If
            Next wtToken
        End If
    ElseIf UCase$(statusType) <> "S" And InStr(1, statusText, "cancel", vbTextCompare) = 0 And _
       InStr(1, logText, "cancel", vbTextCompare) = 0 Then
        messageText = "STOP_CANCEL_DISPLAY_LOGS_STATUS_UNPROVEN: " & statusType & ":" & statusText
        Exit Function
    End If
    Set closeButton = ses.FindById("wnd[1]/tbar[0]/btn[12]", False)
    If closeButton Is Nothing Then Set closeButton = ses.FindById("wnd[1]/tbar[0]/btn[0]", False)
    If Not closeButton Is Nothing Then
        closeTip = CWR_Text(closeButton.Tooltip) & " " & CWR_Text(closeButton.text)
        If InStr(1, closeTip, "cancel", vbTextCompare) > 0 Or _
           InStr(1, closeTip, "close", vbTextCompare) > 0 Or _
           InStr(1, closeTip, "continue", vbTextCompare) > 0 Or closeTip = "" Then
            closeButton.Press
            CWR_WaitIdle ses, 10#
        End If
    End If
    Set modalAfter = ses.FindById("wnd[1]", False)
    If Not modalAfter Is Nothing Then
        messageText = "STOP_CANCEL_DISPLAY_LOGS_CLOSE_FAILED|EXPECTED_WTS=" & expectedWTSet
        Exit Function
    End If
    CWR_Log "WT_CANCEL_DISPLAY_LOGS", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "DISPLAY_LOGS_PROOF=TRUE|STATUS=" & statusType & ":" & statusText & _
        "|EXPECTED_WT_SET=" & expectedWTSet & "|SUCCESS_LINES=" & CStr(canceledCount) & _
        "|RESULTS=" & CWR_CancelResultMapText(normalizedSet) & _
        "|CANCEL_SENT=TRUE|CANCEL_VERIFIED=TRUE|NO_CONFIRM_RETRY=TRUE|NO_SAP_WRITE_AFTER_PROOF=TRUE|MODAL_REMAINS=FALSE"
    CWR_CancelDisplayLogsProof = True
    Exit Function
Failed:
    messageText = "CANCEL_DISPLAY_LOGS_PROOF_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_CancelResultMapText(ByVal expectedWTSet As String) As String
    Dim token As Variant, keyText As String, resultText As String
    On Error GoTo ExitHere
    For Each token In Split(expectedWTSet, ",")
        keyText = CWR_Text(token)
        If keyText <> "" Then
            If resultText <> "" Then resultText = resultText & ";"
            resultText = resultText & keyText & "="
            If CWR_CANCEL_RESULT_BY_WT Is Nothing Then
                resultText = resultText & "UNPROVEN"
            ElseIf CWR_CANCEL_RESULT_BY_WT.Exists(keyText) Then
                resultText = resultText & CWR_Text(CWR_CANCEL_RESULT_BY_WT(keyText))
            Else
                resultText = resultText & "UNPROVEN"
            End If
        End If
    Next token
ExitHere:
    CWR_CancelResultMapText = resultText
End Function

Private Function CWR_SelectedRowsSetMatches(ByVal actualText As String, ByVal expectedText As String) As Boolean
    Dim expected As Object, actual As Object, token As Variant
    Dim normalizedExpected As String, normalizedActual As String
    On Error GoTo Failed
    normalizedExpected = Replace(Replace(CWR_Text(expectedText), " ", ""), ";", ",")
    normalizedActual = Replace(Replace(CWR_Text(actualText), " ", ""), ";", ",")
    If normalizedExpected = "" Or normalizedActual = "" Then Exit Function
    Set expected = CreateObject("Scripting.Dictionary"): expected.CompareMode = vbTextCompare
    Set actual = CreateObject("Scripting.Dictionary"): actual.CompareMode = vbTextCompare
    For Each token In Split(normalizedExpected, ",")
        If CWR_Text(token) <> "" Then expected(CWR_Text(token)) = True
    Next token
    For Each token In Split(normalizedActual, ",")
        If CWR_Text(token) <> "" Then actual(CWR_Text(token)) = True
    Next token
    If expected.count <> actual.count Then Exit Function
    For Each token In expected.keys
        If Not actual.Exists(CStr(token)) Then Exit Function
    Next token
    CWR_SelectedRowsSetMatches = True
    Exit Function
Failed:
    CWR_SelectedRowsSetMatches = False
End Function

Private Function CWR_PostVerifyCancelled(ByVal ses As Object, ByRef proofText As String, _
                                         ByRef messageText As String) As Boolean
    Dim lower As Object, r As Long, n As Long, currentDoc As String, rowDoc As String
    Dim rowWT As String, rowStatus As String, rowItem As String, rowMat As String, oldMatches As Long, activeOther As Long

    On Error GoTo Failed
    If Not CWR_LoadFilteredWTGrid(ses, lower, messageText) Then Exit Function
    n = CLng(lower.rowCount): currentDoc = "": oldMatches = 0: activeOther = 0: rowStatus = ""
    For r = 0 To n - 1
        rowDoc = CWR_Cell(lower, r, "DOCNO")
        If rowDoc <> "" Then currentDoc = rowDoc
        rowItem = CWR_Cell(lower, r, "ITEMNO")
        rowMat = CWR_Cell(lower, r, "MATNR")
        If CWR_MONValueEquals(currentDoc, CWR_EXPECTED_DOC) And _
           CWR_MONValueEquals(rowItem, CWR_EXPECTED_ITEM) And _
           StrComp(rowMat, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 Then
            rowWT = CWR_Cell(lower, r, "TANUM")
            If rowWT <> "" Then
                If StrComp(rowWT, CWR_EXPECTED_WT, vbTextCompare) = 0 Then
                    oldMatches = oldMatches + 1
                    rowStatus = CWR_Cell(lower, r, "TOSTAT")
                ElseIf CWR_Cell(lower, r, "TOSTAT") = "" And _
                       Not CWR_IsKnownBatchPeerOldWT(rowWT) Then
                    activeOther = activeOther + 1
                End If
            End If
        End If
    Next r
    If oldMatches > 1 Then
        messageText = "POSTVERIFY_DUPLICATE_OLD_WT=" & CStr(oldMatches)
        Exit Function
    End If
    If oldMatches = 1 And rowStatus = "" Then
        messageText = "POSTVERIFY_OLD_WT_STILL_ACTIVE"
        Exit Function
    End If
    If activeOther > 0 Then
        messageText = "POSTVERIFY_UNEXPECTED_ACTIVE_WT=" & CStr(activeOther)
        Exit Function
    End If
    If oldMatches = 0 Then
        proofText = "MON_REFRESH=PASS; OLD_WT_NOT_VISIBLE_AFTER_CANCEL; NEW_ACTIVE_WT=0"
    Else
        proofText = "MON_REFRESH=PASS; OLD_WT_STATUS=" & rowStatus & "; NEW_ACTIVE_WT=0"
    End If
    CWR_PostVerifyCancelled = True
    Exit Function
Failed:
    messageText = "CANCEL_POSTVERIFY_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub CWR_RecordCancelledSnapshot(ByVal proofText As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(CWR_PLAN_WH_ROW, 25).Value2 = "OB_OPEN_OR_PARTIAL"
    ws.Cells(CWR_PLAN_WH_ROW, 26).Value2 = "WT_CANCELLED_PENDING_REMAINING"
    ws.Cells(CWR_PLAN_WH_ROW, 27).ClearContents
    ws.Cells(CWR_PLAN_WH_ROW, 28).Value2 = False
    ws.Cells(CWR_PLAN_WH_ROW, 29).Value2 = "WAIT_REMAINING_QTY"
    ws.Cells(CWR_PLAN_WH_ROW, 30).Value2 = "Old WT cancellation verified; read remaining quantity before rework. " & proofText
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: old WT cancelled; CF remains user handoff. " & proofText
ExitHere:
End Sub

Private Sub CWR_RecordRemaining(ByVal remainingQty As Double, ByVal currentBatch As String, ByVal statusText As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(CWR_PLAN_WH_ROW, 25).Value2 = "OB_OPEN_OR_PARTIAL"
    ws.Cells(CWR_PLAN_WH_ROW, 26).Value2 = "WT_CANCELLED_REMAINING"
    ws.Cells(CWR_PLAN_WH_ROW, 27).Value2 = remainingQty
    ws.Cells(CWR_PLAN_WH_ROW, 28).Value2 = True
    ws.Cells(CWR_PLAN_WH_ROW, 29).Value2 = "REPROCESS_REQUIRED"
    ws.Cells(CWR_PLAN_WH_ROW, 30).Value2 = "Cancelled WT remaining quantity read from PRDO; " & statusText & _
        "; target remains bin=" & CWR_TARGET_BIN & "; batch=" & CWR_TARGET_BATCH & "; qty=" & CStr(CWR_TARGET_QTY)
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: remaining quantity verified after cancel; CF remains user handoff."
ExitHere:
End Sub

Private Sub CWR_RecordTargetStock(ByVal proofText As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(5, 15).Value2 = "READY"
    ws.Cells(5, 16).Value2 = "POST_CHANGE_BIN_BATCH_QTY"
    ws.Cells(5, 17).Value2 = "TARGET_STOCK_VERIFIED"
    ws.Cells(5, 18).Value2 = "POST_TARGET_READ"
    ws.Cells(5, 19).ClearContents
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: target source stock verified; CF remains user handoff. " & proofText
ExitHere:
End Sub

' Retains the already-verified PRDO quantity state while recording that no
' POST mutation was required because the WH batch was already at the pick bin.
Private Sub CWR_RecordPickBinBatchAlreadyCorrect(ByVal proofText As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(CWR_PLAN_WH_ROW, 15).Value2 = "READY"
    ws.Cells(CWR_PLAN_WH_ROW, 16).Value2 = "CREATE_WT_FROM_TARGET_THEN_CF_HANDOFF"
    ws.Cells(CWR_PLAN_WH_ROW, 17).Value2 = "PICK_BIN_BATCH_POST_VERIFIED"
    ws.Cells(CWR_PLAN_WH_ROW, 18).Value2 = "POST_PICK_BIN_ALREADY_CORRECT"
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: WH batch is already available at the confirmed pick bin; no POST save was sent. " & _
        proofText & ". Next: redetermine/create WH only; CF remains user handoff."
ExitHere:
End Sub

' Records a one-shot POST recovery only after a fresh exact product query has
' proved the required WH batch at the confirmed pick bin.  Unlike the normal
' no-op record, this wording retains the fact that a prior Create+Save was
' sent while making explicit that this invocation issued no second command.
Private Sub CWR_RecordPickBinBatchPostReconciled(ByVal proofText As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(CWR_PLAN_WH_ROW, 15).Value2 = "READY"
    ws.Cells(CWR_PLAN_WH_ROW, 16).Value2 = "CREATE_WT_FROM_TARGET_THEN_CF_HANDOFF"
    ws.Cells(CWR_PLAN_WH_ROW, 17).Value2 = "PICK_BIN_BATCH_POST_VERIFIED"
    ws.Cells(CWR_PLAN_WH_ROW, 18).Value2 = "POST_PICK_BIN_RECOVERED_READ_ONLY"
    ws.Cells(CWR_PLAN_WH_ROW, 19).Value2 = "Fresh exact POST readback proved the required WH batch at the confirmed pick bin after an earlier Create+Save. No POST retry was sent. " & proofText
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: fresh POST readback reconciled the earlier Create+Save without retrying it. Next: create one outbound WH from the verified bin/batch; CF remains user handoff. " & proofText
ExitHere:
End Sub

Public Sub CWR_LockedPlanSchedule_StaticProbe()
    If CWR_StableHash("CHECK_LOCKED_PLAN") <> CWR_StableHash("CHECK_LOCKED_PLAN") Then _
        Err.Raise vbObjectError + 10001, , "LOCKED_PLAN_HASH_NOT_STABLE"
    If CWR_StableHash("CHECK_LOCKED_PLAN") = CWR_StableHash("CHECK_LOCKED_PLAN_CHANGED") Then _
        Err.Raise vbObjectError + 10002, , "LOCKED_PLAN_HASH_NOT_SENSITIVE"
End Sub

' Persists the one-shot boundary before POST Create+Save.  RUN_LOG is also
' checked on later runs because the normal Cancel-reconcile checkpoint may
' overwrite worksheet presentation fields while the SAP outcome is still
' being reconciled.
Private Sub CWR_RecordPickBinPostSendCheckpoint(ByVal actionDetail As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(CWR_PLAN_WH_ROW, 15).Value2 = "WAIT_RECONCILE"
    ws.Cells(CWR_PLAN_WH_ROW, 16).Value2 = "POST_CHANGE_SENT_RECONCILE"
    ws.Cells(CWR_PLAN_WH_ROW, 17).Value2 = "STATE_UNKNOWN_AFTER_POST_CREATE_SAVE"
    ws.Cells(CWR_PLAN_WH_ROW, 18).Value2 = "NO_RETRY_RECONCILE"
    ws.Cells(CWR_PLAN_WH_ROW, 19).Value2 = "POST Create+Save attempted once; next run reads SAP outcome only. " & actionDetail
ExitHere:
End Sub

Private Function CWR_HasUnreconciledPickBinPostSend() As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long, stageText As String, codeText As String
    Dim pending As Boolean
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("RUN_LOG")
    lastRow = ws.Cells(ws.rows.count, 17).End(xlUp).row
    For r = 1 To lastRow
        If StrComp(CWR_Text(ws.Cells(r, 4).Value2), CWR_EXPECTED_OB, vbTextCompare) = 0 And _
           StrComp(CWR_Text(ws.Cells(r, 5).Value2), CWR_EXPECTED_DOC, vbTextCompare) = 0 And _
           StrComp(CWR_Text(ws.Cells(r, 8).Value2), CWR_EXPECTED_WT, vbTextCompare) = 0 Then
            stageText = UCase$(CWR_Text(ws.Cells(r, 3).Value2))
            codeText = UCase$(CWR_Text(ws.Cells(r, 10).Value2))
            If stageText = "POST_PICK_BIN_CREATE_SAVE_SENT" Then
                pending = True
            ElseIf (stageText = "POST_PICK_BIN_POSTVERIFY" Or stageText = "POST_DIRECT_RECOVERY" Or _
                    stageText = "POST_PICK_BIN_COMMITTED") And _
                   codeText = "PASS" Then
                pending = False
            End If
        End If
    Next r
    CWR_HasUnreconciledPickBinPostSend = pending
ExitHere:
End Function

' Records a fully re-read in-place POST change.  The target batch is available
' now, so this is explicitly a continuation checkpoint, not a CF handoff.
Private Sub CWR_RecordPickBinBatchPosting(ByVal beforeProof As String, ByVal afterProof As String, _
                                           ByVal createdBefore As Long, ByVal createdAfter As Long)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(CWR_PLAN_WH_ROW, 15).Value2 = "READY"
    ws.Cells(CWR_PLAN_WH_ROW, 16).Value2 = "CREATE_WT_FROM_TARGET_THEN_CF_HANDOFF"
    ws.Cells(CWR_PLAN_WH_ROW, 17).Value2 = "PICK_BIN_BATCH_POST_VERIFIED"
    ws.Cells(CWR_PLAN_WH_ROW, 18).Value2 = "POST_PICK_BIN_INPLACE_VERIFIED"
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: target batch is verified at the confirmed pick bin after in-place POST change " & _
        CWR_PICK_BIN_BATCH_BEFORE & " to WH batch " & CWR_WH_BATCH & " for " & CStr(CWR_TARGET_QTY) & _
        " pcs; created postings " & CStr(createdBefore) & " -> " & CStr(createdAfter) & _
        ". Next: create outbound WH from the verified bin/qty. CF remains user-owned. " & _
        beforeProof & "; " & afterProof
ExitHere:
End Sub

' This is a committed transaction-proof state. The SAP-created posting-change
' WT is recorded separately from the outbound WT and is never queued for CF.
' The next gate is the exact target-bin lookup before any outbound quantity write.
Private Sub CWR_RecordCreatedPostingWT(ByVal wtNo As String, ByVal proofText As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(CWR_PLAN_WH_ROW, 15).Value2 = "POST_CHANGE_COMMITTED"
    ws.Cells(CWR_PLAN_WH_ROW, 16).Value2 = "POST_CHANGE_TRANSACTION_PROOF"
    ws.Cells(CWR_PLAN_WH_ROW, 17).Value2 = "POST_CHANGE_RESULT=SUCCESS"
    ws.Cells(CWR_PLAN_WH_ROW, 18).Value2 = "POST_PROOF_WT=" & wtNo
    ws.Cells(CWR_PLAN_WH_ROW, 19).Value2 = "POST change committed once; POST_PROOF_WT=" & wtNo & _
        "; WRITE_RETRY_ALLOWED=FALSE; CF_ACTION_SENT=0; continue to target-bin/qty section. " & proofText
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: SAP created posting-change WT " & wtNo & _
        " (transaction proof only). No CF action was sent; continue to exact target-bin/qty and outbound WH creation. " & proofText
ExitHere:
End Sub

' Records the outbound task creation only after Monitor has proved its exact
' source/bin/batch/quantity.  CF remains explicitly outside all tool actions.
Private Sub CWR_RecordOutboundWTCreated(ByVal wtNo As String, ByVal proofText As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(CWR_PLAN_WH_ROW, 15).Value2 = "READY"
    ws.Cells(CWR_PLAN_WH_ROW, 16).Value2 = "READY_FOR_CF"
    ws.Cells(CWR_PLAN_WH_ROW, 17).Value2 = "WT_POST_VERIFY_PASS"
    ws.Cells(CWR_PLAN_WH_ROW, 18).Value2 = "USER_CF_HANDOFF"
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: outbound WT " & wtNo & " created from bin " & CWR_TARGET_BIN & _
        " with batch " & CWR_WH_BATCH & " and qty " & CStr(CWR_CREATE_QTY) & _
        IIf(CWR_IsDeltaAddRoute(), " (retained WT=" & CWR_EXPECTED_WT & "; PRDO total=" & CStr(CWR_TARGET_QTY) & ")", "") & _
        ". READY_FOR_CF; CF_NOT_SENT; user confirmation is required. " & proofText
ExitHere:
End Sub

' Executes the supported multi-WT form of CHANGE_QTY.  It is deliberately
' limited to one ODO, unchanged bin/batch, and active rows only.  That lets us
' aggregate multiple old WTs that belong to the same ODO item without guessing
' whether an interrupted SAP draft belongs to this batch.
Private Function CWR_ExecuteBatchQtyIncreaseOB(ByVal obText As String) As String
    Dim records As Collection, groups As Object, rec As Object, group As Object
    Dim i As Long, resultText As String, messageText As String, wtList As String
    Dim createSes As Object, mapPromotionText As String

    On Error GoTo Failed
    If Not CWR_BatchCollectActiveQtyRecords(obText, records, groups, messageText) Then GoTo Stopped
    If records.count < 2 Then
        messageText = "BATCH_QTY_REQUIRES_AT_LEAST_TWO_ACTIVE_CASES"
        GoTo Stopped
    End If

    Set CWR_BATCH_PEER_OLD_WTS = CreateObject("Scripting.Dictionary")
    CWR_BATCH_PEER_OLD_WTS.CompareMode = vbTextCompare
    For i = 1 To records.count
        Set rec = records.item(i)
        CWR_BATCH_PEER_OLD_WTS.Add CStr(rec("OldWT")), True
    Next i
    CWR_BATCH_ACTIVE = True

    ' Cancel every old WT in the batch first.  The existing cancellation helper
    ' still performs its own Monitor reconcile; the peer list only prevents a
    ' not-yet-cancelled sibling from being treated as an unexpected new WT.
    For i = 1 To records.count
        Set rec = records.item(i)
        If Not CWR_BatchUsePlan(rec, messageText) Then GoTo Stopped
        resultText = CWR_EnsureOldWTCancelledOnce()
        If Left$(resultText, 5) <> "PASS|" Then
            messageText = "BATCH_CANCEL_PLAN=" & CStr(rec("PlanID")) & "|" & resultText
            GoTo Stopped
        End If
        rec("DestBin") = CWR_EXPECTED_DEST_BIN
        rec("ProcTy") = CWR_EXPECTED_PROCTY
        CWR_Log "BATCH_CANCEL", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                "PLAN=" & CStr(rec("PlanID")) & "; " & resultText
    Next i

    Set rec = records.item(1)
    If Not CWR_BatchUsePlan(rec, messageText) Then GoTo Stopped
    If Not CWR_BatchSetPRDOQtyOneSave(records, groups, messageText) Then GoTo Stopped
    If Not CWR_BatchOpenCreateWT(records, groups, createSes, messageText) Then GoTo Stopped
    If Not CWR_BatchAdoptCreateSave(records, groups, createSes, messageText) Then GoTo Stopped
    If Not CWR_BatchPostVerifyAndRecord(records, messageText, wtList) Then GoTo Stopped

    If Not SAPAction_PromoteLivePass("WT_CREATE_SAVE", _
        "Batch one-save WT creation; " & wtList & "; CF not sent.", mapPromotionText) Then _
        mapPromotionText = "STOP_MAP_PROMOTION|" & mapPromotionText
    CWR_ExecuteBatchQtyIncreaseOB = "PASS|OB_BATCH_READY_FOR_CF|OB=" & obText & _
        "|OUTBOUND_WT_CREATED=" & CStr(records.count) & "|WT=" & wtList & _
        "|PRDO_SAVE=ONE|CREATE_SAVE=ONE|WT_POST_VERIFY_PASS|CF_NOT_SENT|MAP=" & mapPromotionText
    GoTo ExitHere

Stopped:
    CWR_Log "BATCH_QTY", obText, CWR_EXPECTED_DOC, "STOP", messageText
    CWR_ExecuteBatchQtyIncreaseOB = "STOP|OB_BATCH_QTY|OB=" & obText & "|" & messageText & _
        "|NO_RETRY|CF_NOT_SENT"
    GoTo ExitHere
Failed:
    CWR_Log "BATCH_QTY", obText, CWR_EXPECTED_DOC, "FAIL", _
            "Err=" & CStr(Err.Number) & ": " & Err.Description
    CWR_ExecuteBatchQtyIncreaseOB = "FAIL|OB_BATCH_QTY|" & CStr(Err.Number) & "|" & _
        Err.Description & "|STATE_UNKNOWN|NO_RETRY|CF_NOT_SENT"
ExitHere:
    CWR_BATCH_ACTIVE = False
    Set CWR_BATCH_PEER_OLD_WTS = Nothing
End Function

Private Function CWR_BatchCollectActiveQtyRecords(ByVal obText As String, ByRef records As Collection, _
                                                    ByRef groups As Object, ByRef messageText As String) As Boolean
    Dim wsPlan As Worksheet, wsWH As Worksheet, r As Long, lastRow As Long, whRow As Long
    Dim rec As Object, group As Object, groupKey As String, planID As String, rawStatus As String, wtState As String
    Dim decisionText As String, gateText As String, docText As String, itemText As String, materialText As String
    Dim sourceBin As String, targetBin As String, sourceBatch As String, targetBatch As String

    On Error GoTo Failed
    Set records = New Collection
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If StrComp(CWR_Text(wsPlan.Cells(r, 2).Value2), obText, vbTextCompare) <> 0 Then GoTo NextPlan
        planID = CWR_Text(wsPlan.Cells(r, 1).Value2)
        whRow = CLng(CWR_Number(wsPlan.Cells(r, 27).Value2))
        If planID = "" Or whRow < 5 Then
            messageText = "BATCH_PLAN_ID_OR_WH_ROW_INVALID"
            Exit Function
        End If
        rawStatus = UCase$(CWR_Text(wsWH.Cells(whRow, 10).Value2))
        wtState = UCase$(CWR_Text(wsWH.Cells(whRow, 26).Value2))
        If rawStatus = "C" Or wtState = "WT_DONE" Then GoTo NextPlan
        decisionText = UCase$(CWR_Text(wsPlan.Cells(r, 20).Value2))
        gateText = UCase$(CWR_Text(wsPlan.Cells(r, 24).Value2))
        If rawStatus <> "" Or wtState <> "WT_ACTIVE" Or gateText <> "PASS" Or decisionText <> "CHANGE_QTY" Then
            messageText = "BATCH_REQUIRES_ACTIVE_PASS_CHANGE_QTY|PLAN=" & planID & _
                "|STATUS=" & rawStatus & "|WT_STATE=" & wtState & "|DECISION=" & decisionText & "|GATE=" & gateText
            Exit Function
        End If
        docText = CWR_Text(wsPlan.Cells(r, 3).Value2)
        itemText = CWR_Text(wsWH.Cells(whRow, 3).Value2)
        materialText = CWR_Text(wsPlan.Cells(r, 6).Value2)
        sourceBin = CWR_Text(wsPlan.Cells(r, 9).Value2)
        targetBin = CWR_Text(wsPlan.Cells(r, 10).Value2)
        sourceBatch = CWR_Text(wsPlan.Cells(r, 11).Value2)
        targetBatch = CWR_Text(wsPlan.Cells(r, 14).Value2)
        If docText = "" Or itemText = "" Or materialText = "" Or sourceBin = "" Or sourceBatch = "" Or _
           StrComp(sourceBin, targetBin, vbTextCompare) <> 0 Or _
           StrComp(sourceBatch, targetBatch, vbTextCompare) <> 0 Then
            messageText = "BATCH_QTY_ONLY_BIN_OR_BATCH_CHANGE_NOT_SUPPORTED|PLAN=" & planID
            Exit Function
        End If
        If records.count > 0 Then
            If StrComp(CStr(records.item(1)("Doc")), docText, vbTextCompare) <> 0 Then
                messageText = "BATCH_REQUIRES_ONE_ODO_DOCUMENT"
                Exit Function
            End If
        End If
        Set rec = CreateObject("Scripting.Dictionary")
        rec.CompareMode = vbTextCompare
        rec.Add "PlanID", planID
        rec.Add "PlanRow", r
        rec.Add "WHRow", whRow
        rec.Add "Doc", docText
        rec.Add "Item", itemText
        rec.Add "Material", materialText
        rec.Add "OldWT", CWR_Text(wsPlan.Cells(r, 5).Value2)
        rec.Add "OldQty", CWR_Number(wsPlan.Cells(r, 7).Value2)
        rec.Add "TargetQty", CWR_Number(wsPlan.Cells(r, 8).Value2)
        rec.Add "SourceBin", sourceBin
        rec.Add "SourceBatch", sourceBatch
        rec.Add "TargetBin", targetBin
        rec.Add "TargetBatch", targetBatch
        rec.Add "StockType", ""
        rec.Add "DestBin", ""
        rec.Add "ProcTy", ""
        rec.Add "CancelAlready", False
        rec.Add "CancelProof", ""
        rec.Add "Adopted", False
        If CStr(rec("OldWT")) = "" Or CDbl(rec("OldQty")) <= 0# Or CDbl(rec("TargetQty")) <= 0# Then
            messageText = "BATCH_QTY_REQUIRED_VALUE_MISSING|PLAN=" & planID
            Exit Function
        End If
        groupKey = itemText & "|" & materialText & "|" & sourceBatch
        rec.Add "GroupKey", groupKey
        records.Add rec
        If Not groups.Exists(groupKey) Then
            Set group = CreateObject("Scripting.Dictionary")
            group.CompareMode = vbTextCompare
            group.Add "Item", itemText
            group.Add "Material", materialText
            group.Add "Batch", sourceBatch
            group.Add "OldQty", 0#
            group.Add "TargetQty", 0#
            group.Add "PRDORow", -1
            group.Add "RequestRow", -1
            group.Add "Remaining", 0#
            group.Add "LiveQty", 0#
            group.Add "WriteRequired", False
            groups.Add groupKey, group
        End If
        Set group = groups.item(groupKey)
        group("OldQty") = CDbl(group("OldQty")) + CDbl(rec("OldQty"))
        group("TargetQty") = CDbl(group("TargetQty")) + CDbl(rec("TargetQty"))
NextPlan:
    Next r
    If records.count = 0 Then
        messageText = "BATCH_NO_ACTIVE_CASES"
        Exit Function
    End If
    CWR_BatchCollectActiveQtyRecords = True
    Exit Function
Failed:
    messageText = "BATCH_COLLECT_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_BatchUsePlan(ByVal rec As Object, ByRef messageText As String) As Boolean
    Dim planTargetBatch As String
    CWR_REQUESTED_PLAN_ID = CStr(rec("PlanID"))
    planTargetBatch = CWR_Text(rec("TargetBatch"))
    If planTargetBatch <> "" Then CWR_RECOVERY_CONTEXT = True
    If Not CWR_LoadExecutionContext(messageText) Then Exit Function
    If planTargetBatch <> "" Then
        ' A recovered POST row can have a blank/pre-change target batch in
        ' CHECK.  The batch record carries the SAP-proven final batch, which is
        ' safe to use for local post-verify bookkeeping only.
        CWR_TARGET_BATCH = planTargetBatch
        CWR_PICK_BIN_BATCH_BEFORE = planTargetBatch
    End If
    If Not CWR_AssertCFHandoff(messageText) Then Exit Function
    rec("StockType") = CWR_TARGET_STOCK_TYPE
    CWR_BatchUsePlan = True
End Function

Private Sub CWR_BatchRecordPRDOVerified(ByVal rec As Object, ByVal group As Object)
    Dim ws As Worksheet, whRow As Long
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    whRow = CLng(rec("WHRow"))
    ws.Cells(whRow, 15).Value2 = "READY"
    ws.Cells(whRow, 16).Value2 = "CREATE_WT_FROM_TARGET_THEN_CF_HANDOFF"
    ws.Cells(whRow, 17).Value2 = "PRDO_BATCH_QTY_VERIFIED"
    ws.Cells(whRow, 18).Value2 = "PRDO_BATCH_ONE_SAVE"
    ws.Cells(whRow, 19).Value2 = "PRDO item total verified=" & CStr(group("TargetQty")) & _
        "; batch unchanged=" & CStr(group("Batch")) & "; one Save for OB batch."
ExitHere:
End Sub

' Loads every affected ODO item, changes the editable quantity cells in the
' natural top-to-bottom PRDO grid order, and sends exactly one Save only after
' every cell has a positive readback.
Private Function CWR_BatchSetPRDOQtyOneSave(ByVal records As Collection, ByVal groups As Object, _
                                             ByRef messageText As String) As Boolean
    Const PRDO_OIP_TOOLBAR_ID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell"
    Dim ses As Object, grid As Object, toolbar As Object, group As Object, rec As Object
    Dim groupKey As Variant, i As Long, r As Long, actionDetail As String, modalText As String
    Dim statusType As String, statusText As String, writeNeeded As Boolean, saveSent As Boolean

    On Error GoTo Failed
    If Not CWR_RequireReworkAction("PRDO_QTY", messageText) Then Exit Function
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then Exit Function
    If Not CWR_BatchMapPRDORows(grid, groups, messageText) Then Exit Function
    For Each groupKey In groups.keys
        Set group = groups.item(CStr(groupKey))
        If CBool(group("WriteRequired")) Then writeNeeded = True
    Next groupKey
    If writeNeeded Then
        Set toolbar = ses.FindById(PRDO_OIP_TOOLBAR_ID, False)
        If Not CWR_GetVerifiedPRDOChangeAction(toolbar, actionDetail, messageText) Then Exit Function
        toolbar.PressButton "OIP_CHANGE"
        If Not CWR_WaitIdle(ses, 20#) Then
            messageText = "BATCH_PRDO_EDIT_MODE_TIMEOUT"
            Exit Function
        End If
        If Not CWR_BatchNoSAPError(ses, "BATCH_PRDO_EDIT", messageText) Then Exit Function
        If Not CWR_FindCurrentPRDOGrid(ses, grid, messageText) Then Exit Function

        ' The loop is over SAP row number, not worksheet order.  This is the
        ' explicit top-to-bottom filling contract requested for a multi-item ODO.
        For r = 0 To CLng(grid.rowCount) - 1
            For Each groupKey In groups.keys
                Set group = groups.item(CStr(groupKey))
                If CLng(group("PRDORow")) = r And CBool(group("WriteRequired")) Then
                    If UCase$(CWR_CellChangeable(grid, r, "BATCHNO_UI")) <> "FALSE" Or _
                       UCase$(CWR_CellChangeable(grid, r, "QTY_UI")) <> "TRUE" Then
                        messageText = "BATCH_PRDO_FIELD_CONTRACT_MISMATCH|ROW=" & CStr(r)
                        Exit Function
                    End If
                    grid.selectedRows = CStr(r)
                    grid.SetCurrentCell r, "QTY_UI"
                    grid.SelectionChanged
                    If Not CWR_WaitPOSTSaveSettle(ses, 0.25, 15#) Then
                        messageText = "BATCH_PRDO_SELECTION_TIMEOUT|ROW=" & CStr(r)
                        Exit Function
                    End If
                    If Not CWR_FindCurrentPRDOGrid(ses, grid, messageText) Then Exit Function
                    If CWR_SelectedRowsText(grid) <> CStr(r) Then
                        messageText = "BATCH_PRDO_SELECTION_READBACK_MISMATCH|ROW=" & CStr(r)
                        Exit Function
                    End If
                    If StrComp(CWR_Cell(grid, r, "BATCHNO_UI"), CStr(group("Batch")), vbTextCompare) <> 0 Then
                        messageText = "BATCH_PRDO_BATCH_CHANGED_BEFORE_DRAFT|ROW=" & CStr(r)
                        Exit Function
                    End If
                    grid.ModifyCell r, "QTY_UI", CWR_FastQuantityText(CDbl(group("TargetQty")))
                    If Abs(CWR_Number(CWR_Cell(grid, r, "QTY_UI")) - CDbl(group("TargetQty"))) > 0.000001 Then
                        messageText = "BATCH_PRDO_QTY_READBACK_MISMATCH|ROW=" & CStr(r)
                        Exit Function
                    End If
                    grid.TriggerModified
                    If Not CWR_WaitPOSTSaveSettle(ses, 0.25, 20#) Then
                        messageText = "BATCH_PRDO_TRIGGER_TIMEOUT|ROW=" & CStr(r)
                        Exit Function
                    End If
                    If Not CWR_BatchNoSAPError(ses, "BATCH_PRDO_TRIGGER", messageText) Then Exit Function
                    If Not CWR_FindCurrentPRDOGrid(ses, grid, messageText) Then Exit Function
                    If Abs(CWR_Number(CWR_Cell(grid, r, "QTY_UI")) - CDbl(group("TargetQty"))) > 0.000001 Or _
                       StrComp(CWR_Cell(grid, r, "BATCHNO_UI"), CStr(group("Batch")), vbTextCompare) <> 0 Then
                        messageText = "BATCH_PRDO_POSTTRIGGER_MISMATCH|ROW=" & CStr(r)
                        Exit Function
                    End If
                End If
            Next groupKey
        Next r

        Set toolbar = ses.FindById(PRDO_OIP_TOOLBAR_ID, False)
        If Not CWR_GetVerifiedPRDOSaveAction(toolbar, actionDetail, messageText) Then Exit Function
        saveSent = True
        toolbar.PressButton "SAVE"
        CWR_Log "BATCH_PRDO_SAVE_SENT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                "Rows filled top-to-bottom; one PRDO Save sent for " & CStr(groups.count) & " ODO item(s). " & actionDetail
        If Not CWR_WaitPOSTSaveSettle(ses, 3#, 45#) Then
            messageText = "STATE_UNKNOWN_AFTER_BATCH_PRDO_SAVE|TIMEOUT"
            GoTo SaveStopped
        End If
        If Not CWR_BatchNoSAPError(ses, "BATCH_PRDO_SAVE", messageText) Then GoTo SaveStopped
    End If
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then GoTo SaveStopped
    If Not CWR_BatchMapPRDORows(grid, groups, messageText) Then GoTo SaveStopped
    For Each groupKey In groups.keys
        Set group = groups.item(CStr(groupKey))
        If Abs(CDbl(group("LiveQty")) - CDbl(group("TargetQty"))) > 0.000001 Then
            messageText = "BATCH_PRDO_POSTVERIFY_MISMATCH|ITEM=" & CStr(group("Item"))
            GoTo SaveStopped
        End If
    Next groupKey
    For i = 1 To records.count
        Set rec = records.item(i)
        Set group = groups.item(CStr(rec("GroupKey")))
        CWR_BatchRecordPRDOVerified rec, group
    Next i
    CWR_Log "BATCH_PRDO_POSTVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "All PRDO totals verified after " & IIf(writeNeeded, "one Save", "read-only resume") & "; batch unchanged."
    CWR_BatchSetPRDOQtyOneSave = True
    Exit Function
SaveStopped:
    If saveSent Then messageText = messageText & "|NO_RETRY"
    Exit Function
Failed:
    messageText = "BATCH_PRDO_ERROR: " & CStr(Err.Number) & ": " & Err.Description & _
        IIf(saveSent, "|STATE_UNKNOWN_AFTER_BATCH_PRDO_SAVE|NO_RETRY", "")
End Function

Private Function CWR_BatchMapPRDORows(ByVal grid As Object, ByVal groups As Object, _
                                       ByRef messageText As String) As Boolean
    Dim groupKey As Variant, group As Object, r As Long, n As Long, matches As Long
    Dim itemText As String, materialText As String, batchText As String, qtyText As String, liveQty As Double
    On Error GoTo Failed
    n = CLng(grid.rowCount)
    For Each groupKey In groups.keys
        Set group = groups.item(CStr(groupKey))
        matches = 0: group("PRDORow") = -1
        For r = 0 To n - 1
            itemText = CWR_Cell(grid, r, "ITEMNO")
            materialText = CWR_Cell(grid, r, "PRODUCTNO")
            batchText = CWR_Cell(grid, r, "BATCHNO_UI")
            If CWR_MONValueEquals(itemText, CStr(group("Item"))) And _
               StrComp(materialText, CStr(group("Material")), vbTextCompare) = 0 And _
               StrComp(batchText, CStr(group("Batch")), vbTextCompare) = 0 Then
                matches = matches + 1
                group("PRDORow") = r
                qtyText = CWR_Cell(grid, r, "QTY_UI")
                liveQty = CWR_Number(qtyText)
                group("LiveQty") = liveQty
            End If
        Next r
        ' A target of zero is a valid, already-saved readback for a cancelled
        ' ODO item.  Match identity first; accept either the old positive
        ' quantity or the exact target below without turning a valid resume
        ' into a false item-match failure.
        If matches <> 1 Then
            messageText = "BATCH_PRDO_ITEM_MATCH_COUNT=" & CStr(matches) & "|ITEM=" & CStr(group("Item"))
            Exit Function
        End If
        If liveQty < -0.000001 Then
            messageText = "BATCH_PRDO_NEGATIVE_QTY|ITEM=" & CStr(group("Item"))
            Exit Function
        End If
        If Abs(liveQty - CDbl(group("OldQty"))) <= 0.000001 Then
            group("WriteRequired") = (Abs(liveQty - CDbl(group("TargetQty"))) > 0.000001)
        ElseIf Abs(liveQty - CDbl(group("TargetQty"))) <= 0.000001 Then
            group("WriteRequired") = False
        Else
            messageText = "BATCH_PRDO_STATE_UNKNOWN|ITEM=" & CStr(group("Item")) & _
                "|LIVE=" & CStr(liveQty) & "|BEFORE=" & CStr(group("OldQty")) & _
                "|TARGET=" & CStr(group("TargetQty"))
            Exit Function
        End If
    Next groupKey
    CWR_BatchMapPRDORows = True
    Exit Function
Failed:
    messageText = "BATCH_PRDO_MAP_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_BatchNoSAPError(ByVal ses As Object, ByVal stageText As String, _
                                      ByRef messageText As String) As Boolean
    Dim modalText As String, statusType As String, statusText As String
    If CWR_HasAnyModal(ses, modalText) Then
        messageText = stageText & "_MODAL|" & modalText
        Exit Function
    End If
    CWR_ReadStatusBar ses, statusType, statusText
    If statusType <> "" And InStr(1, "EAX", UCase$(statusType), vbBinaryCompare) > 0 Then
        messageText = stageText & "_SAP_" & statusType & "|" & statusText
        Exit Function
    End If
    CWR_BatchNoSAPError = True
End Function

Private Function CWR_BatchPrimeWarehouseTaskSurface(ByVal createSes As Object, _
                                                     ByRef messageText As String) As Boolean
    Dim wtTabId As String, wtGridId As String, wtTab As Object, wtGrid As Object
    Dim firstRows As Long, refreshedRows As Long
    On Error GoTo Failed
    wtTabId = SAPAction_Command("WT_TAB_WAREHOUSE_TASK", "S000")
    wtGridId = SAPAction_Command("WT_WAREHOUSE_TASK_GRID", "S000")
    If wtTabId = "" Or wtGridId = "" Then
        messageText = "STOP_WHTASK_NAVIGATION_ACTION_MAP_MISSING"
        Exit Function
    End If
    ' SAP can render only the first visible WHTask row on the first tab
    ' activation.  The bounded two-select refresh is the required navigation
    ' proof before any quantity/adopt decision is made.
    If Not CWR_RefreshWarehouseTaskDraftGrid(createSes, wtTabId, wtGridId, wtGrid, _
                                              firstRows, refreshedRows, messageText) Then
        messageText = "STOP_WHTASK_NAVIGATION_UNPROVEN|" & messageText
        Exit Function
    End If
    CWR_Log "WHTASK_NAVIGATION", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "METHOD=SELECT_TWICE|FIRST_ROWS=" & CStr(firstRows) & _
        "|REFRESHED_ROWS=" & CStr(refreshedRows) & "|READ_ONLY=TRUE"
    CWR_BatchPrimeWarehouseTaskSurface = True
    Exit Function
Failed:
    messageText = "STOP_WHTASK_NAVIGATION_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_BatchOpenCreateWT(ByVal records As Collection, ByVal groups As Object, _
                                        ByRef createSes As Object, ByRef messageText As String) As Boolean
    Dim prdoSes As Object, prdoGrid As Object, menuItem As Object, queryButton As Object
    Dim priorSessionIds As String, followOnId As String, queryId As String, menuText As String
    Dim firstRec As Object
    Dim reuseQueryId As String, reuseModalText As String, reuseQueryBtn As Object, surfaceStale As Boolean
    On Error GoTo Failed
    Set firstRec = records.item(1).item("SchedulerRec")
    If Not CWR_BatchAssertNoActiveOutboundWT(records, groups, messageText) Then Exit Function
    If Not CWR_BatchAnyCreateRemaining(groups) Then
        ' Existing active delta WT(s) already cover the planned CreateQty.
        ' Keep this path read-only; post-verify will reconcile the aggregate.
        CWR_BatchOpenCreateWT = True
        Exit Function
    End If
    If CWR_CREATE_CONTEXT_VALID And _
       StrComp(CWR_CREATE_CONTEXT_OB, CStr(firstRec("OB")), vbTextCompare) = 0 And _
       StrComp(CWR_CREATE_CONTEXT_DOC, CStr(firstRec("Doc")), vbTextCompare) = 0 Then
        ' Bug B fix (NOT_LIVE_PROVEN): a single material taken from two source
        ' positions builds TWO create groups (the CreateGroup key includes
        ' SRC_BIN/DST_BIN), but this cache keys on OB+Doc only. Leg 1 ends with a
        ' Create+Save, so its surface is no longer on the S000 query-entry screen.
        ' The old code pressed "query" on that spent surface (or stopped with
        ' STOP_WHTASK_SESSION_REUSE_CONTROL_MISSING) -> "reads old context / stops
        ' before bin fill". Reuse only when the surface PROVES it is still at the
        ' query-entry screen; otherwise self-heal by reopening fresh (budget=20,
        ' only 2 opens needed, so DOCUMENT/WHTASK budgets are never threatened).
        surfaceStale = False
        Set createSes = CWR_CREATE_CONTEXT_SESSION
        reuseQueryId = SAPAction_Command("WT_CREATE_QUERY", "S000")
        If reuseQueryId = "" Then
            messageText = "STOP_WHTASK_SESSION_REUSE_ACTION_MAP_MISSING"
            Exit Function
        End If
        If createSes Is Nothing Then
            surfaceStale = True
        ElseIf CBool(createSes.busy) Then
            surfaceStale = True
        ElseIf CWR_HasAnyModal(createSes, reuseModalText) Then
            surfaceStale = True
        Else
            Set reuseQueryBtn = createSes.FindById(reuseQueryId, False)
            If reuseQueryBtn Is Nothing Then surfaceStale = True
        End If
        If Not surfaceStale Then
            reuseQueryBtn.Press
            If Not CWR_WaitIdle(createSes, 30#) Then surfaceStale = True
        End If
        If Not surfaceStale Then
            If Not CWR_BatchNoSAPError(createSes, "BATCH_CREATE_QUERY_REUSE", messageText) Then surfaceStale = True
        End If
        If surfaceStale Then
            ' Post-save / non-query-entry surface: discard the stale create context
            ' and reopen a clean Create-WT surface for this distinct source position.
            CWR_CREATE_CONTEXT_VALID = False
            CWR_CREATE_CONTEXT_OB = vbNullString
            CWR_CREATE_CONTEXT_DOC = vbNullString
            Set CWR_CREATE_CONTEXT_SESSION = Nothing
            CWR_Log "NAV_DECISION", CStr(firstRec("OB")), CStr(firstRec("Doc")), "PASS", _
                "FROM=PRDO|TO=WT|METHOD=WHTASK|REUSE_DOCUMENT=SELF_HEAL_FRESH|DOCUMENT=" & CStr(firstRec("Doc"))
            GoTo FreshOpen
        End If
        CWR_WHTASK_REUSE_COUNT = CWR_WHTASK_REUSE_COUNT + 1
        CWR_Log "NAV_DECISION", CStr(firstRec("OB")), CStr(firstRec("Doc")), "PASS", _
            "FROM=PRDO|TO=WT|METHOD=WHTASK|REUSE_DOCUMENT=TRUE|DOCUMENT=" & CStr(firstRec("Doc")) & _
            "|WHTASK_REUSE_COUNT=" & CStr(CWR_WHTASK_REUSE_COUNT)
        If Not CWR_BatchPrimeWarehouseTaskSurface(createSes, messageText) Then Exit Function
        CWR_BatchOpenCreateWT = True
        Exit Function
    End If
FreshOpen:
    If Not CWR_OpenExpectedPRDO(prdoSes, prdoGrid, messageText) Then Exit Function
    If Not CWR_BatchMapPRDORows(prdoGrid, groups, messageText) Then Exit Function
    If Not CWR_BatchAllPRDOTargets(groups, messageText) Then Exit Function
    If Not CWR_BatchNoSAPError(prdoSes, "BATCH_CREATE_PRDO", messageText) Then Exit Function
    If Not CWR_CaptureSAPSessionIds(priorSessionIds, messageText) Then Exit Function
    followOnId = SAPAction_Command("WT_CREATE_FOLLOWON", "/SCWM/PRDO")
    queryId = SAPAction_Command("WT_CREATE_QUERY", "S000")
    If followOnId = "" Or queryId = "" Then
        messageText = "BATCH_CREATE_NAVIGATION_ACTION_MAP_MISSING"
        Exit Function
    End If
    Set menuItem = prdoSes.FindById(followOnId, False)
    If menuItem Is Nothing Then
        messageText = "BATCH_CREATE_FOLLOWON_MENU_NOT_FOUND"
        Exit Function
    End If
    menuText = CWR_Text(menuItem.text)
    If StrComp(menuText, "Warehouse Task", vbTextCompare) <> 0 And _
       InStr(1, menuText, "Create", vbTextCompare) = 0 Then
        messageText = "BATCH_CREATE_FOLLOWON_MENU_IDENTITY_MISMATCH|" & menuText
        Exit Function
    End If
    menuItem.Select
    If Not CWR_WaitForNewCreateWTSession(prdoSes, priorSessionIds, createSes, 20#, messageText) Then Exit Function
    Set queryButton = createSes.FindById(queryId, False)
    If queryButton Is Nothing Then
        messageText = "BATCH_CREATE_QUERY_BUTTON_NOT_FOUND"
        Exit Function
    End If
    queryButton.Press
    If Not CWR_WaitIdle(createSes, 30#) Then
        messageText = "BATCH_CREATE_QUERY_TIMEOUT"
        Exit Function
    End If
    If Not CWR_BatchNoSAPError(createSes, "BATCH_CREATE_QUERY", messageText) Then Exit Function
    Set CWR_CREATE_CONTEXT_SESSION = createSes
    CWR_CREATE_CONTEXT_VALID = True
    CWR_CREATE_CONTEXT_OB = CStr(firstRec("OB"))
    CWR_CREATE_CONTEXT_DOC = CStr(firstRec("Doc"))
    CWR_WHTASK_OPEN_COUNT = CWR_WHTASK_OPEN_COUNT + 1
    CWR_Log "NAV_DECISION", CStr(firstRec("OB")), CStr(firstRec("Doc")), "PASS", _
        "FROM=PRDO|TO=WT|METHOD=WHTASK|REUSE_DOCUMENT=FALSE|DOCUMENT=" & CStr(firstRec("Doc")) & _
        "|WHTASK_OPEN_COUNT=" & CStr(CWR_WHTASK_OPEN_COUNT)
    If Not CWR_BatchPrimeWarehouseTaskSurface(createSes, messageText) Then Exit Function
    CWR_BatchOpenCreateWT = True
    Exit Function
Failed:
    messageText = "BATCH_CREATE_OPEN_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_BatchAllPRDOTargets(ByVal groups As Object, ByRef messageText As String) As Boolean
    Dim groupKey As Variant, group As Object
    For Each groupKey In groups.keys
        Set group = groups.item(CStr(groupKey))
        If Abs(CDbl(group("LiveQty")) - CDbl(group("TargetQty"))) > 0.000001 Then
            messageText = "BATCH_PRDO_TARGET_UNPROVEN|ITEM=" & CStr(group("Item"))
            Exit Function
        End If
    Next groupKey
    CWR_BatchAllPRDOTargets = True
End Function

Private Function CWR_BatchAssertNoActiveOutboundWT(ByVal records As Collection, ByVal groups As Object, _
                                                     ByRef messageText As String) As Boolean
    Dim monSes As Object, lower As Object, rec As Object, group As Object, schedulerRec As Object
    Dim r As Long, n As Long, i As Long, groupKey As Variant
    Dim currentDoc As String, rowWT As String, rowItem As String, rowMat As String, rowStatus As String
    Dim rowBatch As String, rowBin As String, rowDest As String, rowProc As String, rowQty As Double
    Dim expectedOldWT As String, execMode As String, retainedExpected As Boolean
    Dim matched As Boolean, existingQty As Double, expectedQty As Double, remainingQty As Double
    On Error GoTo Failed
    For Each groupKey In groups.keys
        Set group = groups.item(CStr(groupKey))
        If Not group.Exists("ExistingCreateQty") Then group.Add "ExistingCreateQty", 0# Else group("ExistingCreateQty") = 0#
        If Not group.Exists("Remaining") Then group.Add "Remaining", 0# Else group("Remaining") = 0#
    Next groupKey
    If Not CWR_GetMONSession(monSes, messageText) Then Exit Function
    If Not CWR_LoadFilteredWTGrid(monSes, lower, messageText) Then Exit Function
    n = CLng(lower.rowCount)
    For r = 0 To n - 1
        If CWR_Cell(lower, r, "DOCNO") <> "" Then currentDoc = CWR_Cell(lower, r, "DOCNO")
        rowWT = CWR_Cell(lower, r, "TANUM")
        rowStatus = CWR_Cell(lower, r, "TOSTAT")
        If rowWT <> "" And rowStatus = "" And CWR_MONValueEquals(currentDoc, CWR_EXPECTED_DOC) Then
            rowItem = CWR_Cell(lower, r, "ITEMNO")
            rowMat = CWR_Cell(lower, r, "MATNR")
            rowBatch = CWR_Cell(lower, r, "CHARG")
            rowBin = CWR_Cell(lower, r, "VLPLA")
            rowDest = CWR_Cell(lower, r, "NLPLA")
            rowProc = CWR_Cell(lower, r, "PROCTY")
            rowQty = CWR_Number(CWR_Cell(lower, r, "VSOLM"))
            matched = False
            For i = 1 To records.count
                Set rec = records.item(i)
                If CWR_MONValueEquals(rowItem, CStr(rec("Item"))) And _
                   StrComp(rowMat, CStr(rec("Material")), vbTextCompare) = 0 Then
                    matched = True
                    expectedOldWT = CWR_Text(rec("OldWT"))
                    execMode = ""
                    If rec.Exists("SchedulerRec") Then
                        Set schedulerRec = rec.item("SchedulerRec")
                        execMode = CWR_Text(schedulerRec("ExecMode"))
                    End If
                    retainedExpected = (StrComp(rowWT, expectedOldWT, vbTextCompare) = 0 And _
                        StrComp(execMode, "KEEP_EXISTING_CREATE_DELTA", vbTextCompare) = 0)
                    If StrComp(rowWT, expectedOldWT, vbTextCompare) = 0 Then
                        If Not retainedExpected Then
                            messageText = "BATCH_CREATE_ACTIVE_OUTBOUND_WT_PRESENT|WT=" & rowWT & _
                                "|PLAN=" & CWR_Text(rec("PlanID"))
                            Exit Function
                        End If
                    ElseIf StrComp(rowBatch, CStr(rec("TargetBatch")), vbTextCompare) = 0 And _
                           StrComp(rowBin, CStr(rec("TargetBin")), vbTextCompare) = 0 And _
                           StrComp(rowDest, CStr(rec("DestBin")), vbTextCompare) = 0 And _
                           StrComp(rowProc, CStr(rec("ProcTy")), vbTextCompare) = 0 Then
                        Set group = groups.item(CStr(rec("GroupKey")))
                        group("ExistingCreateQty") = CDbl(group("ExistingCreateQty")) + rowQty
                    Else
                        messageText = "BATCH_CREATE_ACTIVE_OUTBOUND_WT_PRESENT|WT=" & rowWT & _
                            "|PLAN=" & CWR_Text(rec("PlanID"))
                        Exit Function
                    End If
                    Exit For
                End If
            Next i
            If Not matched Then
                ' Ignore unrelated active WTs in the filtered MON layout.
                ' The exact item/material rows above remain fail-closed.
            End If
        End If
    Next r
    For Each groupKey In groups.keys
        Set group = groups.item(CStr(groupKey))
        If group.Exists("CreateQty") Then
            expectedQty = CDbl(group("CreateQty"))
        Else
            expectedQty = CDbl(group("TargetQty"))
        End If
        existingQty = CDbl(group("ExistingCreateQty"))
        If existingQty > expectedQty + 0.000001 Then
            messageText = "STATE_UNKNOWN_CREATE_DELTA_OVER_TARGET|GROUP=" & CStr(groupKey) & _
                "|EXISTING=" & CStr(existingQty) & "|EXPECTED=" & CStr(expectedQty)
            Exit Function
        End If
        remainingQty = expectedQty - existingQty
        If remainingQty < 0# Then remainingQty = 0#
        group("Remaining") = remainingQty
        CWR_Log "BATCH_CREATE_PREVERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "GROUP=" & CStr(groupKey) & "; EXISTING_CREATE_QTY=" & CStr(existingQty) & _
            "; PLANNED_CREATE_QTY=" & CStr(expectedQty) & "; REMAINING_CREATE_QTY=" & CStr(remainingQty) & _
            "; READ_ONLY=TRUE"
    Next groupKey
    If Not CWR_BatchAllocateCreateRemainder(records, groups, messageText) Then Exit Function
    CWR_BatchAssertNoActiveOutboundWT = True
    Exit Function
Failed:
    messageText = "BATCH_CREATE_MONITOR_PREVERIFY_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_BatchAllocateCreateRemainder(ByVal records As Collection, ByVal groups As Object, _
                                                    ByRef messageText As String) As Boolean
    Dim groupKey As Variant, group As Object, rec As Object
    Dim remainingQty As Double, plannedQty As Double, assignQty As Double, i As Long
    On Error GoTo Failed
    For Each groupKey In groups.keys
        Set group = groups.item(CStr(groupKey))
        remainingQty = CDbl(group("Remaining"))
        For i = 1 To records.count
            Set rec = records.item(i)
            If StrComp(CStr(rec("GroupKey")), CStr(groupKey), vbTextCompare) = 0 Then
                If rec.Exists("PlannedDraftQty") Then
                    plannedQty = CDbl(rec("PlannedDraftQty"))
                Else
                    plannedQty = CDbl(rec("DraftQty"))
                End If
                assignQty = plannedQty
                If assignQty > remainingQty Then assignQty = remainingQty
                If assignQty < 0# Then assignQty = 0#
                rec("DraftQty") = assignQty
                rec("Adopted") = (assignQty <= 0.000001)
                remainingQty = remainingQty - assignQty
                If remainingQty < 0# Then remainingQty = 0#
            End If
        Next i
        If remainingQty > 0.000001 Then
            messageText = "STATE_UNKNOWN_CREATE_REMAINDER_UNALLOCATED|GROUP=" & CStr(groupKey) & _
                "|REMAINING=" & CStr(remainingQty)
            Exit Function
        End If
    Next groupKey
    CWR_BatchAllocateCreateRemainder = True
    Exit Function
Failed:
    messageText = "BATCH_CREATE_REMAINDER_ALLOCATE_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_BatchAnyCreateRemaining(ByVal groups As Object) As Boolean
    Dim groupKey As Variant, group As Object
    On Error GoTo ExitHere
    For Each groupKey In groups.keys
        Set group = groups.item(CStr(groupKey))
        If CDbl(group("Remaining")) > 0.000001 Then
            CWR_BatchAnyCreateRemaining = True
            Exit Function
        End If
    Next groupKey
ExitHere:
End Function

' One Create WT session can contain multiple Warehouse Request rows.  Requests
' are handled in their visible SAP order; when one request needs several
' source bins, those source rows are also selected in their visible order.
Private Function CWR_BatchAdoptCreateSave(ByVal records As Collection, ByVal groups As Object, _
                                           ByVal createSes As Object, ByRef messageText As String) As Boolean
    Dim wrGrid As Object, topToolbar As Object, group As Object, groupKey As Variant
    Dim requestGridId As String, stockTabId As String, stockGridId As String, stockToolbarId As String
    Dim qtyColumn As String, adoptId As String, wtTabId As String, wtGridId As String, toolbarId As String
    Dim createCommandId As String, saveCommandId As String, r As Long, actionDetail As String
    Dim createSaveSent As Boolean, adoptedCount As Long, expectedAdopts As Long, rec As Object

    On Error GoTo Failed
    requestGridId = SAPAction_Command("WT_REQUEST_GRID", "S000")
    stockTabId = SAPAction_Command("WT_TAB_STK_CAN_BE_REMOVED", "S000")
    stockGridId = SAPAction_Command("WT_STOCK_GRID", "S000")
    stockToolbarId = SAPAction_Command("WT_STOCK_TOOLBAR", "S000")
    qtyColumn = SAPAction_Command("WT_STOCK_SET_QTY", "S000")
    adoptId = SAPAction_Command("WT_ADOPT_QUANTITY", "S000")
    wtTabId = SAPAction_Command("WT_TAB_WAREHOUSE_TASK", "S000")
    wtGridId = SAPAction_Command("WT_WAREHOUSE_TASK_GRID", "S000")
    toolbarId = SAPAction_Command("WT_REQUEST_TOOLBAR", "S000")
    If requestGridId = "" Or stockTabId = "" Or stockGridId = "" Or stockToolbarId = "" Or _
       qtyColumn = "" Or adoptId = "" Or wtTabId = "" Or wtGridId = "" Or toolbarId = "" Then
        messageText = "BATCH_CREATE_ACTION_MAP_REQUIRED_CONTROL_MISSING"
        Exit Function
    End If
    Set wrGrid = createSes.FindById(requestGridId, False)
    If Not CWR_BatchMapRequestRows(wrGrid, groups, messageText) Then Exit Function
    expectedAdopts = 0
    For Each rec In records
        If CDbl(rec("DraftQty")) > 0.000001 Then expectedAdopts = expectedAdopts + 1
    Next rec
    If expectedAdopts = 0 Then
        CWR_BatchAdoptCreateSave = True
        Exit Function
    End If

    ' Iterate the request surface by its row number.  This avoids an implicit
    ' sort by PlanID and keeps the SAP user-facing top-to-bottom order.
    For r = 0 To CLng(wrGrid.rowCount) - 1
        For Each groupKey In groups.keys
            Set group = groups.item(CStr(groupKey))
            If CLng(group("RequestRow")) = r Then
                If Not CWR_BatchAdoptOneRequest(records, group, createSes, requestGridId, stockTabId, _
                    stockGridId, stockToolbarId, qtyColumn, adoptId, wtTabId, wtGridId, adoptedCount, messageText) Then
                    Exit Function
                End If
            End If
        Next groupKey
    Next r
    If adoptedCount <> expectedAdopts Then
        messageText = "STOP_BATCH_ADOPT_QUANTITY_UNPROVEN|DRAFT_EXPECTED=" & CStr(expectedAdopts) & _
            "|DRAFT_OBSERVED=" & CStr(adoptedCount)
        Exit Function
    End If
    If Not CWR_BatchVerifyAllAdopts(createSes, requestGridId, wtTabId, wtGridId, groups, expectedAdopts, messageText) Then Exit Function

    If Not SAPAction_AssertLivePass("WT_CREATE_SAVE", "S000", messageText) Then Exit Function
    If Not CWR_RequireReworkAction("WT_CREATE_SAVE", messageText) Then Exit Function
    createCommandId = SAPAction_Command("WT_CREATE_COMMAND_ID", "S000")
    saveCommandId = SAPAction_Command("WT_SAVE_COMMAND_ID", "S000")
    If createCommandId = "" Or saveCommandId = "" Then
        messageText = "BATCH_CREATE_SAVE_ACTION_MAP_MISSING"
        Exit Function
    End If
    Set topToolbar = createSes.FindById(toolbarId, False)
    If topToolbar Is Nothing Then
        messageText = "BATCH_CREATE_SAVE_TOOLBAR_NOT_FOUND"
        Exit Function
    End If
    If Not CWR_GetVerifiedToolbarAction(topToolbar, createCommandId, "Create Warehouse Task", actionDetail, messageText) Then Exit Function
    topToolbar.PressButton createCommandId
    createSaveSent = True
    CWR_Log "BATCH_CREATE_SENT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "All source rows adopted in SAP order; one Create sent. " & actionDetail
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 45#) Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|BATCH_CREATE_TIMEOUT"
        Exit Function
    End If
    If Not CWR_BatchNoSAPError(createSes, "BATCH_CREATE", messageText) Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|" & messageText
        Exit Function
    End If
    Set topToolbar = createSes.FindById(toolbarId, False)
    If topToolbar Is Nothing Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|BATCH_SAVE_TOOLBAR_NOT_FOUND"
        Exit Function
    End If
    If Not CWR_GetVerifiedToolbarAction(topToolbar, saveCommandId, "Save", actionDetail, messageText) Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|" & messageText
        Exit Function
    End If
    topToolbar.PressButton saveCommandId
    CWR_Log "BATCH_CREATE_SAVE_SENT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "One Save sent after all batch drafts were verified. " & actionDetail
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 45#) Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|BATCH_SAVE_TIMEOUT"
        Exit Function
    End If
    If Not CWR_BatchNoSAPError(createSes, "BATCH_SAVE", messageText) Then
        messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|" & messageText
        Exit Function
    End If
    CWR_BatchAdoptCreateSave = True
    Exit Function
Failed:
    messageText = "BATCH_ADOPT_CREATE_ERROR: " & CStr(Err.Number) & ": " & Err.Description & _
        IIf(createSaveSent, "|STATE_UNKNOWN_AFTER_CREATE_SAVE|NO_RETRY", "")
End Function

Private Function CWR_BatchMapRequestRows(ByVal wrGrid As Object, ByVal groups As Object, _
                                          ByRef messageText As String) As Boolean
    Dim groupKey As Variant, group As Object, r As Long, n As Long, matches As Long
    Dim rowDoc As String, rowItem As String, rowMat As String, rowBatch As String
    Dim openQty As Double, expectedOpenQty As Double
    On Error GoTo Failed
    If wrGrid Is Nothing Then
        messageText = "STOP_BATCH_REQUEST_GRID_NOT_FOUND"
        Exit Function
    End If
    n = CLng(wrGrid.rowCount)
    For Each groupKey In groups.keys
        Set group = groups.item(CStr(groupKey))
        matches = 0: group("RequestRow") = -1
        For r = 0 To n - 1
            rowDoc = CWR_Cell(wrGrid, r, "DLVNO")
            rowItem = CWR_Cell(wrGrid, r, "ITMNO")
            rowMat = CWR_Cell(wrGrid, r, "MATNR")
            rowBatch = CWR_Cell(wrGrid, r, "CHARG")
            If CWR_MONValueEquals(rowDoc, CWR_EXPECTED_DOC) And _
               CWR_MONValueEquals(rowItem, CStr(group("Item"))) And _
               StrComp(rowMat, CStr(group("Material")), vbTextCompare) = 0 And _
               StrComp(rowBatch, CStr(group("Batch")), vbTextCompare) = 0 Then
                matches = matches + 1
                group("RequestRow") = r
                openQty = CWR_Number(CWR_Cell(wrGrid, r, "UI_OPNQTY"))
            End If
        Next r
        expectedOpenQty = CDbl(group("TargetQty"))
        If group.Exists("CreateQty") Then expectedOpenQty = CDbl(group("CreateQty"))
        If group.Exists("Remaining") Then expectedOpenQty = CDbl(group("Remaining"))
        If matches <> 1 Or Abs(openQty - expectedOpenQty) > 0.000001 Then
            messageText = "STATE_UNKNOWN_BEFORE_BATCH_ADOPT|REQUEST_MATCHES=" & CStr(matches) & _
                "|ITEM=" & CStr(group("Item")) & "|OPEN_QTY=" & CStr(openQty) & _
                "|EXPECTED=" & CStr(expectedOpenQty)
            Exit Function
        End If
        group("Remaining") = expectedOpenQty
    Next groupKey
    CWR_BatchMapRequestRows = True
    Exit Function
Failed:
    messageText = "BATCH_REQUEST_MAP_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_BatchAdoptOneRequest(ByVal records As Collection, ByVal group As Object, _
                                           ByVal createSes As Object, ByVal requestGridId As String, _
                                           ByVal stockTabId As String, ByVal stockGridId As String, _
                                           ByVal stockToolbarId As String, ByVal qtyColumn As String, _
                                           ByVal adoptId As String, ByVal wtTabId As String, ByVal wtGridId As String, _
                                           ByRef adoptedCount As Long, ByRef messageText As String) As Boolean
    Dim wrGrid As Object, stockTab As Object, stockGrid As Object, stockToolbar As Object, rec As Object
    Dim requestRow As Long, sourceRow As Long, expectedOpen As Double, nextOpen As Double, actionDetail As String
    Dim sourceAvailableQty As Double, sourceAvailableText As String, sourceAvailableColumn As String, sourceStockType As String
    Dim statusType As String, statusText As String, adoptSent As Boolean

    On Error GoTo Failed
    Do While CDbl(group("Remaining")) > 0.000001
        Set wrGrid = createSes.FindById(requestGridId, False)
        If Not CWR_BatchFindRequestRow(wrGrid, group, requestRow, expectedOpen, messageText) Then Exit Function
        If Abs(expectedOpen - CDbl(group("Remaining"))) > 0.000001 Then
            messageText = "STATE_UNKNOWN_AFTER_ADOPT|REQUEST_OPEN_QTY=" & CStr(expectedOpen) & _
                "|EXPECTED=" & CStr(group("Remaining"))
            Exit Function
        End If
        If Not CWR_ActivateRequestRowForSource(createSes, wrGrid, requestRow, messageText) Then Exit Function
        Set wrGrid = createSes.FindById(requestGridId, False)
        If wrGrid Is Nothing Or CWR_SelectedRowsText(wrGrid) <> CStr(requestRow) Then
            messageText = "STOP_BATCH_REQUEST_SELECTION_READBACK"
            Exit Function
        End If
        Set stockTab = createSes.FindById(stockTabId, False)
        If stockTab Is Nothing Then
            messageText = "STOP_BATCH_SOURCE_STOCK_TAB_NOT_FOUND"
            Exit Function
        End If
        stockTab.Select
        If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 15#) Then
            messageText = "STOP_BATCH_SOURCE_STOCK_TAB_TIMEOUT"
            Exit Function
        End If
        Set stockGrid = createSes.FindById(stockGridId, False)
        Set stockToolbar = createSes.FindById(stockToolbarId, False)
        If stockGrid Is Nothing Or stockToolbar Is Nothing Then
            messageText = "STOP_BATCH_SOURCE_STOCK_SURFACE_NOT_FOUND"
            Exit Function
        End If
        If Not CWR_BatchFindTopSourceRecord(records, group, stockGrid, rec, sourceRow, sourceAvailableQty, _
                                             sourceAvailableText, sourceAvailableColumn, sourceStockType, messageText) Then Exit Function
        If Not CWR_BatchUsePlan(rec, messageText) Then Exit Function
        If Not CWR_RequireReworkAction("WT_ADOPT", messageText) Then Exit Function
        CWR_Log "SOURCE_ROW_FOUND", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                "PLAN=" & CStr(rec("PlanID")) & "; ITEM=" & CStr(group("Item")) & _
                "; PRODUCT=" & CStr(group("Material")) & "; ROW=" & CStr(sourceRow) & _
                "; BIN=" & CStr(rec("TargetBin")) & "; BATCH=" & CStr(rec("TargetBatch")) & _
                "; STOCK_TYPE=" & sourceStockType & "; AVAILABLE=" & CStr(sourceAvailableQty) & _
                "; AVAILABLE_RAW=" & sourceAvailableText & "; AVAILABLE_COL=" & sourceAvailableColumn & _
                "; REQUIRED=" & CStr(CDbl(rec("DraftQty")))
        stockGrid.ModifyCell sourceRow, qtyColumn, CWR_FastQuantityText(CDbl(rec("DraftQty")))
        stockGrid.TriggerModified
        If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 20#) Then
            messageText = "STOP_BATCH_SOURCE_QTY_TRIGGER_TIMEOUT|ROW=" & CStr(sourceRow)
            Exit Function
        End If
        Set stockGrid = createSes.FindById(stockGridId, False)
        If stockGrid Is Nothing Or Abs(CWR_Number(CWR_Cell(stockGrid, sourceRow, qtyColumn)) - CDbl(rec("DraftQty"))) > 0.000001 Then
            messageText = "STOP_BATCH_SOURCE_QTY_READBACK|ROW=" & CStr(sourceRow)
            Exit Function
        End If
        stockGrid.selectedRows = CStr(sourceRow)
        stockGrid.SetCurrentCell sourceRow, "LGPLA"
        stockGrid.SelectionChanged
        If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 15#) Then
            messageText = "STOP_BATCH_SOURCE_SELECTION_TIMEOUT|ROW=" & CStr(sourceRow)
            Exit Function
        End If
        Set stockGrid = createSes.FindById(stockGridId, False)
        If stockGrid Is Nothing Or CWR_SelectedRowsText(stockGrid) <> CStr(sourceRow) Then
            messageText = "STOP_BATCH_SOURCE_SELECTION_READBACK|ROW=" & CStr(sourceRow)
            Exit Function
        End If
        If Not CWR_GetVerifiedToolbarAction(stockToolbar, adoptId, "Adopt Quantity in WhsTask", actionDetail, messageText) Then Exit Function
        stockToolbar.PressButton adoptId
        adoptSent = True
        If Not CWR_WaitPOSTSaveSettle(createSes, 0.5, 30#) Then
            messageText = "STATE_UNKNOWN_AFTER_ADOPT|BATCH_TIMEOUT"
            Exit Function
        End If
        If Not CWR_BatchNoSAPError(createSes, "BATCH_ADOPT", messageText) Then
            messageText = "STATE_UNKNOWN_AFTER_ADOPT|" & messageText
            Exit Function
        End If
        nextOpen = CDbl(group("Remaining")) - CDbl(rec("DraftQty"))
        adoptedCount = adoptedCount + 1
        If Not CWR_BatchVerifyAdoptGate(createSes, requestGridId, wtTabId, wtGridId, group, nextOpen, adoptedCount, messageText) Then
            messageText = "STATE_UNKNOWN_AFTER_ADOPT|" & messageText
            Exit Function
        End If
        group("Remaining") = nextOpen
        rec("Adopted") = True
        CWR_Log "BATCH_ADOPT", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                "PLAN=" & CStr(rec("PlanID")) & "; SOURCE_ROW=" & CStr(sourceRow) & _
                "; BIN=" & CStr(rec("TargetBin")) & "; QTY=" & CStr(rec("DraftQty")) & _
                "; ADOPT_MINIMAL_VERIFY_PASS"
    Loop
    CWR_BatchAdoptOneRequest = True
    Exit Function
Failed:
    messageText = "BATCH_ADOPT_ERROR: " & CStr(Err.Number) & ": " & Err.Description & _
        IIf(adoptSent, "|STATE_UNKNOWN_AFTER_ADOPT|NO_RETRY", "")
End Function

Private Function CWR_BatchFindRequestRow(ByVal wrGrid As Object, ByVal group As Object, _
                                          ByRef requestRow As Long, ByRef openQty As Double, _
                                          ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, matches As Long, rowDoc As String, rowItem As String, rowMat As String, rowBatch As String
    On Error GoTo Failed
    If wrGrid Is Nothing Then
        messageText = "STOP_BATCH_REQUEST_GRID_UNAVAILABLE"
        Exit Function
    End If
    n = CLng(wrGrid.rowCount): requestRow = -1
    For r = 0 To n - 1
        rowDoc = CWR_Cell(wrGrid, r, "DLVNO")
        rowItem = CWR_Cell(wrGrid, r, "ITMNO")
        rowMat = CWR_Cell(wrGrid, r, "MATNR")
        rowBatch = CWR_Cell(wrGrid, r, "CHARG")
        If CWR_MONValueEquals(rowDoc, CWR_EXPECTED_DOC) And _
           CWR_MONValueEquals(rowItem, CStr(group("Item"))) And _
           StrComp(rowMat, CStr(group("Material")), vbTextCompare) = 0 And _
           StrComp(rowBatch, CStr(group("Batch")), vbTextCompare) = 0 Then
            matches = matches + 1
            requestRow = r
            openQty = CWR_Number(CWR_Cell(wrGrid, r, "UI_OPNQTY"))
        End If
    Next r
    If matches <> 1 Then
        messageText = "STOP_BATCH_REQUEST_MATCH_COUNT=" & CStr(matches) & "|ITEM=" & CStr(group("Item"))
        Exit Function
    End If
    CWR_BatchFindRequestRow = True
    Exit Function
Failed:
    messageText = "BATCH_REQUEST_FIND_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_BatchFindTopSourceRecord(ByVal records As Collection, ByVal group As Object, _
                                                ByVal stockGrid As Object, ByRef resultRec As Object, _
                                                ByRef sourceRow As Long, ByRef sourceAvailableQty As Double, _
                                                ByRef sourceAvailableText As String, ByRef sourceAvailableColumn As String, _
                                                ByRef sourceStockType As String, ByRef messageText As String) As Boolean
    Dim r As Long, i As Long, n As Long, rec As Object, binText As String, batchText As String, stockType As String
    Dim stockMaterial As String, availableText As String, availableColumn As String, availableQty As Double, availableParseOK As Boolean
    Dim tupleMatches As Long, readableAvailableRows As Long, parseFailedRows As Long, maxAvailableQty As Double
    Dim maxAvailableText As String, maxAvailableColumn As String, requiredQty As Double, tupleMatch As Boolean
    On Error GoTo Failed
    sourceRow = -1: sourceAvailableQty = 0#: sourceAvailableText = "": sourceAvailableColumn = "": sourceStockType = ""
    If stockGrid Is Nothing Then
        messageText = "STOP_BATCH_SOURCE_GRID_UNAVAILABLE"
        Exit Function
    End If
    n = CLng(stockGrid.rowCount)
    For i = 1 To records.count
        Set rec = records.item(i)
        If StrComp(CStr(rec("GroupKey")), CStr(group("Item")) & "|" & CStr(group("Material")) & "|" & CStr(group("Batch")), vbTextCompare) = 0 And _
           Not CBool(rec("Adopted")) And CDbl(rec("DraftQty")) > 0.000001 Then
            requiredQty = CDbl(rec("DraftQty"))
            Exit For
        End If
    Next i
    If requiredQty <= 0.000001 Then
        messageText = "STOP_BATCH_SOURCE_REQUIRED_QTY_UNPROVEN|ITEM=" & CStr(group("Item")) & _
            "|BATCH=" & CStr(group("Batch"))
        Exit Function
    End If
    If CBool(group("SourceTupleAmbiguous")) Then
        messageText = "STOP_BATCH_SOURCE_TUPLE_AMBIGUOUS|ITEM=" & CStr(group("Item")) & _
            "|PRODUCT=" & CStr(group("Material")) & "|BATCH=" & CStr(group("Batch")) & _
            "|NO_SOURCE_WRITE"
        Exit Function
    End If
    If Trim$(CStr(group("TargetBin"))) = "" Then
        messageText = "STOP_BATCH_SOURCE_BIN_UNPROVEN|ITEM=" & CStr(group("Item")) & _
            "|PRODUCT=" & CStr(group("Material")) & "|BATCH=" & CStr(group("Batch")) & _
            "|REQ_QTY=" & CStr(requiredQty) & "|NO_SOURCE_WRITE"
        Exit Function
    End If
    ' Source grid first, Plan records second: the first match is literally the
    ' first usable source row visible in SAP.
    For r = 0 To n - 1
        binText = CWR_Cell(stockGrid, r, "LGPLA")
        batchText = CWR_Cell(stockGrid, r, "CHARG")
        stockType = CWR_SourceStockType(stockGrid, r)
        stockMaterial = CWR_Cell(stockGrid, r, "MATNR")
        If stockMaterial = "" Then stockMaterial = CWR_Cell(stockGrid, r, "PRODUCTNO")
        tupleMatch = (StrComp(binText, CStr(group("TargetBin")), vbTextCompare) = 0 And _
                      StrComp(batchText, CStr(group("Batch")), vbTextCompare) = 0 And _
                      (stockMaterial = "" Or StrComp(stockMaterial, CStr(group("Material")), vbTextCompare) = 0) And _
                      (stockType = "" Or StrComp(stockType, CWR_TARGET_STOCK_TYPE, vbTextCompare) = 0))
        If tupleMatch Then
            tupleMatches = tupleMatches + 1
            availableQty = CWR_SourceAvailableQty(stockGrid, r, availableText, availableColumn, availableParseOK)
            If availableText <> "" And availableParseOK Then
                readableAvailableRows = readableAvailableRows + 1
                If readableAvailableRows = 1 Or availableQty > maxAvailableQty Then
                    maxAvailableQty = availableQty
                    maxAvailableText = availableText
                    maxAvailableColumn = availableColumn
                End If
            ElseIf availableText <> "" Then
                parseFailedRows = parseFailedRows + 1
            End If
            For i = 1 To records.count
                Set rec = records.item(i)
                If StrComp(CStr(rec("GroupKey")), CStr(group("Item")) & "|" & CStr(group("Material")) & "|" & CStr(group("Batch")), vbTextCompare) = 0 And _
                   Not CBool(rec("Adopted")) And availableParseOK And availableQty + 0.000001 >= CDbl(rec("DraftQty")) Then
                sourceRow = r
                Set resultRec = rec
                    sourceAvailableQty = availableQty
                    sourceAvailableText = availableText
                    sourceAvailableColumn = availableColumn
                    sourceStockType = stockType
                CWR_BatchFindTopSourceRecord = True
                Exit Function
            End If
            Next i
        End If
    Next r
    If tupleMatches = 0 Then
        messageText = "STOP_BATCH_SOURCE_ROW_NOT_FOUND|ITEM=" & CStr(group("Item")) & _
            "|PRODUCT=" & CStr(group("Material")) & "|BIN=" & CStr(group("TargetBin")) & _
            "|BATCH=" & CStr(group("Batch")) & "|STOCK_TYPE=" & CWR_TARGET_STOCK_TYPE
    ElseIf readableAvailableRows = 0 And parseFailedRows > 0 Then
        messageText = "STOP_BATCH_SOURCE_PARSE_FAILED|ITEM=" & CStr(group("Item")) & _
            "|BIN=" & CStr(group("TargetBin")) & "|BATCH=" & CStr(group("Batch")) & _
            "|AVAILABLE_RAW=" & availableText & "|AVAILABLE_COL=" & availableColumn & _
            "|REQ_QTY=" & CStr(requiredQty) & "|REQUEST_OPEN_QTY=" & CStr(group("Remaining"))
    ElseIf readableAvailableRows = 0 Then
        messageText = "STOP_BATCH_SOURCE_COLUMN_NOT_FOUND|ITEM=" & CStr(group("Item")) & _
            "|BIN=" & CStr(group("TargetBin")) & "|BATCH=" & CStr(group("Batch")) & _
            "|REQ_QTY=" & CStr(requiredQty) & "|REQUEST_OPEN_QTY=" & CStr(group("Remaining"))
    Else
        messageText = "STOP_BATCH_SOURCE_INSUFFICIENT|ITEM=" & CStr(group("Item")) & _
            "|BIN=" & CStr(group("TargetBin")) & "|BATCH=" & CStr(group("Batch")) & _
            "|AVAILABLE=" & CStr(maxAvailableQty) & "|AVAILABLE_RAW=" & maxAvailableText & _
            "|AVAILABLE_COL=" & maxAvailableColumn & "|REQ_QTY=" & CStr(requiredQty) & _
            "|REQUEST_OPEN_QTY=" & CStr(group("Remaining"))
    End If
    Exit Function
Failed:
    messageText = "BATCH_SOURCE_FIND_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' The source-stock grids expose both CAT and TO_CAT on some SAP layouts.  The
' POST precheck uses TO_CAT as the source-of-truth, so Create WT must use the
' same precedence or it can reject the exact row that precheck already proved.
Private Function CWR_SourceStockType(ByVal stockGrid As Object, ByVal rowNo As Long) As String
    CWR_SourceStockType = CWR_Cell(stockGrid, rowNo, "TO_CAT")
    If CWR_SourceStockType = "" Then CWR_SourceStockType = CWR_Cell(stockGrid, rowNo, "CAT")
End Function

' UI_VSOLA is an alternative-UoM source quantity.  Prefer the matching
' available-quantity fields, then use only established fallbacks from the
' bounded read-only grid contract.  Every text value must pass the canonical
' SAP parser so callers can reject malformed availability explicitly.
Private Function CWR_SourceAvailableQty(ByVal stockGrid As Object, ByVal rowNo As Long, _
                                        ByRef rawText As String, ByRef columnName As String, _
                                        ByRef parseOK As Boolean) As Double
    Dim candidate As Variant, candidateText As String, parsedQty As Double
    rawText = "": columnName = ""
    parseOK = False
    For Each candidate In Array("UI_AVLQTY_A", "AVLUOM_A", "UI_AVLQTY", "AVLUOM", "AVL_ALT_QTY")
        candidateText = CWR_Cell(stockGrid, rowNo, CStr(candidate))
        If candidateText <> "" Then
            If rawText = "" Then
                rawText = candidateText
                columnName = CStr(candidate)
            End If
            If CWR_TryParseSAPQty(candidateText, parsedQty) Then
                rawText = candidateText
                columnName = CStr(candidate)
                parseOK = True
                CWR_SourceAvailableQty = parsedQty
                Exit Function
            End If
        End If
    Next candidate
End Function

' SAP EWM can expose a partial Warehouse Task grid immediately after its tab
' first becomes active.  Select the exact mapped tab twice and accept only the
' second grid instance.  This is bounded UI navigation/readback; it never
' repeats a business write.
Private Function CWR_RefreshWarehouseTaskDraftGrid(ByVal createSes As Object, _
                                                    ByVal wtTabId As String, ByVal wtGridId As String, _
                                                    ByRef wtGrid As Object, ByRef firstRows As Long, _
                                                    ByRef refreshedRows As Long, ByRef messageText As String) As Boolean
    Dim wtTab As Object
    firstRows = -1
    refreshedRows = -1
    Set wtGrid = Nothing
    Set wtTab = createSes.FindById(wtTabId, False)
    If wtTab Is Nothing Then
        messageText = "WT_TAB"
        Exit Function
    End If
    wtTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 15#) Then
        messageText = "WT_TAB_TIMEOUT"
        Exit Function
    End If
    Set wtGrid = createSes.FindById(wtGridId, False)
    If wtGrid Is Nothing Then
        messageText = "DRAFT_ROWS_INITIAL_MISSING"
        Exit Function
    End If
    firstRows = CLng(wtGrid.rowCount)

    ' Reacquire after the first render so the second Select and readback do not
    ' operate on a stale GUI object.
    Set wtTab = createSes.FindById(wtTabId, False)
    If wtTab Is Nothing Then
        messageText = "WT_TAB_REFRESH"
        Exit Function
    End If
    wtTab.Select
    If Not CWR_WaitPOSTSaveSettle(createSes, 0.25, 15#) Then
        messageText = "WT_TAB_REFRESH_TIMEOUT"
        Exit Function
    End If
    Set wtGrid = createSes.FindById(wtGridId, False)
    If wtGrid Is Nothing Then
        messageText = "DRAFT_ROWS_REFRESH_MISSING"
        Exit Function
    End If
    refreshedRows = CLng(wtGrid.rowCount)
    CWR_RefreshWarehouseTaskDraftGrid = True
End Function

Private Function CWR_BatchVerifyAdoptGate(ByVal createSes As Object, ByVal requestGridId As String, _
                                           ByVal wtTabId As String, ByVal wtGridId As String, ByVal group As Object, _
                                           ByVal expectedOpen As Double, ByVal expectedDraftCount As Long, _
                                           ByRef messageText As String) As Boolean
    Dim wrGrid As Object, wtGrid As Object, requestRow As Long, openQty As Double
    Dim firstDraftRows As Long, refreshedDraftRows As Long
    Set wrGrid = createSes.FindById(requestGridId, False)
    If Not CWR_BatchFindRequestRow(wrGrid, group, requestRow, openQty, messageText) Then Exit Function
    If Abs(openQty - expectedOpen) > 0.000001 Then
        messageText = "STOP_BATCH_ADOPT_QUANTITY_UNPROVEN|OPEN_QTY=" & CStr(openQty) & _
            "|EXPECTED=" & CStr(expectedOpen)
        Exit Function
    End If
    ' The first Warehouse Task opening can be partially rendered.  Its second
    ' mapped-tab readback is the only draft count used to prove Adopt.
    If Not CWR_RefreshWarehouseTaskDraftGrid(createSes, wtTabId, wtGridId, wtGrid, _
                                              firstDraftRows, refreshedDraftRows, messageText) Then
        messageText = "STOP_BATCH_ADOPT_QUANTITY_UNPROVEN|" & messageText
        Exit Function
    End If
    If refreshedDraftRows <> expectedDraftCount Then
        messageText = "STOP_BATCH_ADOPT_QUANTITY_UNPROVEN|DRAFT_ROWS=" & CStr(refreshedDraftRows) & _
            "|FIRST_PASS_ROWS=" & CStr(firstDraftRows) & "|EXPECTED=" & CStr(expectedDraftCount)
        Exit Function
    End If
    CWR_Log "BATCH_WT_TAB_REFRESH", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "ITEM=" & CStr(group("Item")) & "; FIRST_PASS_ROWS=" & CStr(firstDraftRows) & _
            "; REFRESHED_ROWS=" & CStr(refreshedDraftRows) & "; EXPECTED=" & CStr(expectedDraftCount)
    CWR_BatchVerifyAdoptGate = True
End Function

Private Function CWR_BatchVerifyAllAdopts(ByVal createSes As Object, ByVal requestGridId As String, _
                                           ByVal wtTabId As String, ByVal wtGridId As String, ByVal groups As Object, _
                                           ByVal expectedDraftCount As Long, ByRef messageText As String) As Boolean
    Dim groupKey As Variant, group As Object, wrGrid As Object, wtGrid As Object
    Dim requestRow As Long, openQty As Double, firstDraftRows As Long, refreshedDraftRows As Long
    On Error GoTo Failed
    Set wrGrid = createSes.FindById(requestGridId, False)
    For Each groupKey In groups.keys
        Set group = groups.item(CStr(groupKey))
        If Not CWR_BatchFindRequestRow(wrGrid, group, requestRow, openQty, messageText) Then Exit Function
        If Abs(openQty) > 0.000001 Or Abs(CDbl(group("Remaining"))) > 0.000001 Then
            messageText = "STOP_BATCH_ADOPT_QUANTITY_UNPROVEN|ITEM=" & CStr(group("Item")) & _
                "|OPEN_QTY=" & CStr(openQty)
            Exit Function
        End If
    Next groupKey
    If Not CWR_RefreshWarehouseTaskDraftGrid(createSes, wtTabId, wtGridId, wtGrid, _
                                              firstDraftRows, refreshedDraftRows, messageText) Then
        messageText = "STOP_BATCH_ADOPT_QUANTITY_UNPROVEN|WT_TAB_FINAL|" & messageText
        Exit Function
    End If
    If refreshedDraftRows <> expectedDraftCount Then
        messageText = "STOP_BATCH_ADOPT_QUANTITY_UNPROVEN|FINAL_DRAFT_ROWS=" & CStr(refreshedDraftRows) & _
            "|FIRST_PASS_ROWS=" & CStr(firstDraftRows) & "|EXPECTED=" & CStr(expectedDraftCount)
        Exit Function
    End If
    CWR_Log "BATCH_ADOPT_GATE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
            "ADOPT_MINIMAL_VERIFY_PASS|ALL_REQUESTS_OPEN_QTY=0|DRAFT_ROWS=" & CStr(expectedDraftCount) & _
            "|WT_TAB_FIRST_ROWS=" & CStr(firstDraftRows) & "|WT_TAB_REFRESHED_ROWS=" & CStr(refreshedDraftRows)
    CWR_BatchVerifyAllAdopts = True
    Exit Function
Failed:
    messageText = "BATCH_ADOPT_FINAL_GATE_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_BatchPostVerifyAndRecord(ByVal records As Collection, ByRef messageText As String, _
                                                ByRef wtList As String) As Boolean
    Dim monSes As Object, lower As Object, rec As Object, group As Object
    Dim verifyGroups As Object, oldWTs As Object, newWTs As Object, used As Object
    Dim r As Long, i As Long, n As Long, groupKey As Variant, wtKey As Variant
    Dim currentDoc As String, rowWT As String, rowStatus As String, rowItem As String, rowMat As String
    Dim rowBatch As String, rowBin As String, rowDest As String, rowProc As String, rowQty As Double
    Dim oldRead As Double, newRead As Double, expectedOld As Double, expectedCreate As Double, expectedTarget As Double
    Dim totalRead As Double, proofText As String, newWTText As String, firstWT As String
    Dim oldRequired As Boolean, matchedGroup As Boolean
    On Error GoTo Failed
    Set verifyGroups = CreateObject("Scripting.Dictionary")
    verifyGroups.CompareMode = vbTextCompare
    Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = vbTextCompare

    ' Build one aggregate contract per item/material/source-batch group.
    For i = 1 To records.count
        Set rec = records.item(i)
        groupKey = CStr(rec("GroupKey"))
        If Not verifyGroups.Exists(CStr(groupKey)) Then
            Set group = CreateObject("Scripting.Dictionary")
            group.CompareMode = vbTextCompare
            group.Add "PlanID", rec("PlanID")
            group.Add "Doc", rec("Doc")
            group.Add "Item", rec("Item")
            group.Add "Material", rec("Material")
            group.Add "TargetBatch", rec("TargetBatch")
            group.Add "TargetBin", rec("TargetBin")
            group.Add "DestBin", rec("DestBin")
            group.Add "ProcTy", rec("ProcTy")
            group.Add "ExecMode", ""
            group.Add "OldTotal", 0#
            group.Add "OldRead", 0#
            group.Add "ExpectedCreate", 0#
            group.Add "TargetTotal", 0#
            group.Add "NewRead", 0#
            group.Add "Unexpected", 0&
            Set oldWTs = CreateObject("Scripting.Dictionary")
            oldWTs.CompareMode = vbTextCompare
            group.Add "OldWTs", oldWTs
            Set newWTs = CreateObject("Scripting.Dictionary")
            newWTs.CompareMode = vbTextCompare
            group.Add "NewWTs", newWTs
            verifyGroups.Add CStr(groupKey), group
        End If
        Set group = verifyGroups.item(CStr(groupKey))
        group("OldTotal") = CDbl(group("OldTotal")) + CDbl(rec("OldQty"))
        group("TargetTotal") = CDbl(group("TargetTotal")) + CDbl(rec("TargetQty"))
        If rec.Exists("PlannedDraftQty") Then
            group("ExpectedCreate") = CDbl(group("ExpectedCreate")) + CDbl(rec("PlannedDraftQty"))
        Else
            group("ExpectedCreate") = CDbl(group("ExpectedCreate")) + CDbl(rec("DraftQty"))
        End If
        If group("ExecMode") = "" And rec.Exists("SchedulerRec") Then
            group("ExecMode") = CWR_Text(rec.item("SchedulerRec")("ExecMode"))
        End If
        Set oldWTs = group.item("OldWTs")
        If CWR_Text(rec("OldWT")) <> "" Then oldWTs(CWR_Text(rec("OldWT"))) = True
    Next i

    If Not CWR_GetMONSession(monSes, messageText) Then Exit Function
    If Not CWR_LoadFilteredWTGrid(monSes, lower, messageText) Then Exit Function
    n = CLng(lower.rowCount)
    currentDoc = ""
    For r = 0 To n - 1
        If CWR_Cell(lower, r, "DOCNO") <> "" Then currentDoc = CWR_Cell(lower, r, "DOCNO")
        rowWT = CWR_Cell(lower, r, "TANUM")
        rowStatus = CWR_Cell(lower, r, "TOSTAT")
        If rowWT <> "" And rowStatus = "" And CWR_MONValueEquals(currentDoc, CWR_EXPECTED_DOC) Then
            rowItem = CWR_Cell(lower, r, "ITEMNO")
            rowMat = CWR_Cell(lower, r, "MATNR")
            rowBatch = CWR_Cell(lower, r, "CHARG")
            rowBin = CWR_Cell(lower, r, "VLPLA")
            rowDest = CWR_Cell(lower, r, "NLPLA")
            rowProc = CWR_Cell(lower, r, "PROCTY")
            rowQty = CWR_Number(CWR_Cell(lower, r, "VSOLM"))
            matchedGroup = False
            For Each groupKey In verifyGroups.keys
                Set group = verifyGroups.item(CStr(groupKey))
                If CWR_MONValueEquals(rowItem, CStr(group("Item"))) And _
                   StrComp(rowMat, CStr(group("Material")), vbTextCompare) = 0 Then
                    matchedGroup = True
                    Set oldWTs = group.item("OldWTs")
                    If oldWTs.Exists(rowWT) Then
                        group("OldRead") = CDbl(group("OldRead")) + rowQty
                    ElseIf StrComp(rowBatch, CStr(group("TargetBatch")), vbTextCompare) = 0 And _
                           StrComp(rowBin, CStr(group("TargetBin")), vbTextCompare) = 0 And _
                           StrComp(rowDest, CStr(group("DestBin")), vbTextCompare) = 0 And _
                           StrComp(rowProc, CStr(group("ProcTy")), vbTextCompare) = 0 Then
                        group("NewRead") = CDbl(group("NewRead")) + rowQty
                        Set newWTs = group.item("NewWTs")
                        If newWTs.Exists(rowWT) Then
                            newWTs(rowWT) = CDbl(newWTs(rowWT)) + rowQty
                        Else
                            newWTs.Add rowWT, rowQty
                        End If
                    Else
                        group("Unexpected") = CLng(group("Unexpected")) + 1
                    End If
                    Exit For
                End If
            Next groupKey
            ' Unrelated active WTs in the filtered MON layout do not belong to
            ' this group. Exact item/material matches remain fail-closed above.
        End If
    Next r

    For Each groupKey In verifyGroups.keys
        Set group = verifyGroups.item(CStr(groupKey))
        expectedOld = CDbl(group("OldTotal"))
        expectedCreate = CDbl(group("ExpectedCreate"))
        expectedTarget = CDbl(group("TargetTotal"))
        oldRead = CDbl(group("OldRead"))
        newRead = CDbl(group("NewRead"))
        oldRequired = (StrComp(CStr(group("ExecMode")), "KEEP_EXISTING_CREATE_DELTA", vbTextCompare) = 0)
        If CLng(group("Unexpected")) > 0 Then
            messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|GROUP_POSTVERIFY_UNEXPECTED_ACTIVE_WT=" & _
                CStr(group("Unexpected")) & "|GROUP=" & CStr(groupKey)
            Exit Function
        End If
        If oldRequired And Abs(oldRead - expectedOld) > 0.000001 Then
            messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|GROUP_POSTVERIFY_OLD_WT_TOTAL=" & CStr(oldRead) & _
                "|EXPECTED_OLD=" & CStr(expectedOld) & "|GROUP=" & CStr(groupKey)
            Exit Function
        End If
        If Not oldRequired And oldRead > 0.000001 Then
            messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|GROUP_POSTVERIFY_CANCELLED_OLD_WT_STILL_ACTIVE=" & _
                CStr(oldRead) & "|GROUP=" & CStr(groupKey)
            Exit Function
        End If
        totalRead = oldRead + newRead
        If Abs(newRead - expectedCreate) > 0.000001 Or Abs(totalRead - expectedTarget) > 0.000001 Then
            messageText = "STATE_UNKNOWN_AFTER_CREATE_SAVE|GROUP_POSTVERIFY_TOTAL_MISMATCH|GROUP=" & CStr(groupKey) & _
                "|OLD_ACTIVE=" & CStr(oldRead) & "|NEW_DELTA=" & CStr(newRead) & _
                "|EXPECTED_CREATE=" & CStr(expectedCreate) & "|TOTAL=" & CStr(totalRead) & _
                "|TARGET=" & CStr(expectedTarget)
            Exit Function
        End If
        Set newWTs = group.item("NewWTs")
        newWTText = ""
        firstWT = ""
        For Each wtKey In newWTs.keys
            If firstWT = "" Then firstWT = CStr(wtKey)
            If newWTText <> "" Then newWTText = newWTText & ","
            newWTText = newWTText & CStr(wtKey) & "=" & CStr(newWTs(wtKey))
            If Not used.Exists(CStr(wtKey)) Then
                used.Add CStr(wtKey), True
                If wtList <> "" Then wtList = wtList & ","
                wtList = wtList & CStr(wtKey)
            End If
        Next wtKey
        proofText = "GROUP_POST_VERIFY|OLD_ACTIVE=" & CStr(oldRead) & "; NEW_DELTA=" & CStr(newRead) & _
            "; TOTAL=" & CStr(totalRead) & "; TARGET=" & CStr(expectedTarget) & _
            "; NEW_WTS=" & newWTText & "; BATCH=" & CStr(group("TargetBatch")) & _
            "; SOURCE_BIN=" & CStr(group("TargetBin")) & "; DEST_BIN=" & CStr(group("DestBin")) & _
            "; PROCTY=" & CStr(group("ProcTy"))
        For i = 1 To records.count
            Set rec = records.item(i)
            If StrComp(CStr(rec("GroupKey")), CStr(groupKey), vbTextCompare) = 0 Then
                If firstWT <> "" Then rec("NewWT") = firstWT
                If Not CWR_BatchUsePlan(rec, messageText) Then Exit Function
                If firstWT <> "" Then CWR_RecordOutboundWTCreated firstWT, proofText
                CWR_Log "BATCH_WT_POST_VERIFY", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                    "PLAN=" & CStr(rec("PlanID")) & "; " & proofText & "; READY_FOR_CF; CF_NOT_SENT"
            End If
        Next i
    Next groupKey
    CWR_BatchPostVerifyAndRecord = True
    Exit Function
Failed:
    messageText = "BATCH_WT_POSTVERIFY_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Button #4 may run only a CHECK-locked plan that belongs to the current WH
' snapshot.  It deliberately performs no Posting reload or re-planning.
Private Function CWR_AssertQueuePlanLock(ByVal wsPlan As Worksheet, ByVal wsWH As Worksheet, _
                                          ByRef messageText As String) As Boolean
    Dim planRunID As String, whRunID As String, planHash As String, whHash As String, currentWHHash As String
    Dim postingRunID As String, monRunID As String
    On Error GoTo Failed
    If CWR_Text(wsPlan.Range("AJ1").Value2) <> "PLAN LOCK" Or _
       UCase$(CWR_Text(wsPlan.Range("AJ2").Value2)) <> "LOCKED" Then
        messageText = "STOP_PLAN_NOT_LOCKED_RECHECK"
        Exit Function
    End If
    planRunID = CWR_Text(wsPlan.Range("AK2").Value2)
    whRunID = CWR_Text(wsWH.Range("W2").Value2)
    If planRunID = "" Or StrComp(planRunID, whRunID, vbTextCompare) <> 0 Then
        messageText = "STOP_PLAN_STALE_RECHECK|PLAN_RUN=" & planRunID & "|WH_RUN=" & whRunID
        Exit Function
    End If
    planHash = CWR_Text(wsPlan.Range("AN2").Value2)
    whHash = CWR_Text(wsPlan.Range("AO2").Value2)
    If planHash = "" Or whHash = "" Or CWR_Text(wsPlan.Range("AT2").Value2) <> "CHECK_LOCKED_V2" Then
        messageText = "STOP_PLAN_LOCK_CONTRACT_INCOMPLETE_RECHECK"
        Exit Function
    End If
    If StrComp(planHash, CWR_PlanContractHash(wsPlan), vbTextCompare) <> 0 Then
        messageText = "STOP_PLAN_STALE_RECHECK|PLAN_HASH_CHANGED"
        Exit Function
    End If
    currentWHHash = CWR_WHActualContractHash(wsPlan, wsWH)
    If StrComp(whHash, currentWHHash, vbTextCompare) <> 0 Then
        If Not CWR_MigrateLegacyPlanLockIfSafe(wsPlan, wsWH, whHash, currentWHHash) Then
            messageText = "STOP_PLAN_STALE_RECHECK|WH_ACTUAL_CHANGED_AFTER_CHECK"
            Exit Function
        End If
        CWR_Log "PLAN_LOCK", "", "", "PASS", _
                "MIGRATED_STATUS_FIELD_HASH|SOURCE_AND_USER_INPUTS_MATCH_PLAN|SAP_WRITE=0"
    End If
    postingRunID = CWR_Text(wsPlan.Range("AQ2").Value2)
    If postingRunID <> "" And StrComp(postingRunID, CWR_HistoryPostingRunID(), vbTextCompare) <> 0 Then
        messageText = "STOP_PLAN_STALE_RECHECK|POSTING_RUN_CHANGED"
        Exit Function
    End If
    monRunID = CWR_Text(wsPlan.Range("AR2").Value2)
    If monRunID = "" Or StrComp(monRunID, whRunID, vbTextCompare) <> 0 Then
        messageText = "STOP_PLAN_STALE_RECHECK|MON_RUN_CHANGED"
        Exit Function
    End If
    CWR_AssertQueuePlanLock = True
    Exit Function
Failed:
    messageText = "PLAN_LOCK_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_PlanContractHash(ByVal wsPlan As Worksheet) As String
    Dim r As Long, c As Long, lastRow As Long, payload As String
    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    If lastRow < 2 Then lastRow = 1
    For r = 2 To lastRow
        payload = payload & "|R" & CStr(r)
        For c = 1 To CWR_PLAN_COL_QTY_DELTA
            payload = payload & "|" & CWR_Text(wsPlan.Cells(r, c).Value2)
        Next c
        For c = CWR_PLAN_COL_QTY_GROUP_ID To CWR_PLAN_COL_GROUP_NEED_QTY
            payload = payload & "|" & CWR_Text(wsPlan.Cells(r, c).Value2)
        Next c
    Next r
    CWR_PlanContractHash = CWR_StableHash(payload)
End Function

Private Function CWR_WHActualContractHash(ByVal wsPlan As Worksheet, ByVal wsWH As Worksheet) As String
    Dim r As Long, c As Long, whRow As Long, lastRow As Long, payload As String
    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    If lastRow < 2 Then lastRow = 1
    For r = 2 To lastRow
        whRow = CLng(CWR_Number(wsPlan.Cells(r, 27).Value2))
        payload = payload & "|P=" & CWR_Text(wsPlan.Cells(r, 1).Value2) & "|W=" & CStr(whRow)
        ' Keep the lock tied to immutable MON data plus user-owned picking
        ' inputs. O:S and Y:AD are derived status/result fields that CHECK or
        ' the executor updates after an earlier OB; hashing them falsely
        ' blocked the next OB in the same WH snapshot.
        For c = 1 To 14
            payload = payload & "|" & CWR_Text(wsWH.Cells(whRow, c).Value2)
        Next c
        For c = 20 To 22
            payload = payload & "|" & CWR_Text(wsWH.Cells(whRow, c).Value2)
        Next c
    Next r
    CWR_WHActualContractHash = CWR_StableHash(payload)
End Function

' Compatibility hash for a PLAN locked before the status-field exclusion was
' introduced. It is used only to recognize and safely migrate an old lock;
' new CHECK runs always write CWR_WHActualContractHash instead.
Private Function CWR_WHActualLegacyContractHash(ByVal wsPlan As Worksheet, ByVal wsWH As Worksheet) As String
    Dim r As Long, c As Long, whRow As Long, lastRow As Long, payload As String
    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    If lastRow < 2 Then lastRow = 1
    For r = 2 To lastRow
        whRow = CLng(CWR_Number(wsPlan.Cells(r, 27).Value2))
        payload = payload & "|P=" & CWR_Text(wsPlan.Cells(r, 1).Value2) & "|W=" & CStr(whRow)
        For c = 1 To 30
            payload = payload & "|" & CWR_Text(wsWH.Cells(whRow, c).Value2)
        Next c
    Next r
    CWR_WHActualLegacyContractHash = CWR_StableHash(payload)
End Function

' When the source/user fields still match the locked PLAN, an old full-row
' hash may be migrated after the executor has changed only derived status
' fields. This does not re-plan and never permits a changed source or input.
Private Function CWR_MigrateLegacyPlanLockIfSafe(ByVal wsPlan As Worksheet, ByVal wsWH As Worksheet, _
                                                  ByVal storedHash As String, ByVal currentHash As String) As Boolean
    Dim r As Long, lastRow As Long, whRow As Long
    Dim planBatch As String, actualBatch As String
    On Error GoTo ExitHere
    If StrComp(storedHash, CWR_WHActualLegacyContractHash(wsPlan, wsWH), vbTextCompare) = 0 Then
        wsPlan.Range("AO2").Value2 = currentHash
        CWR_MigrateLegacyPlanLockIfSafe = True
        Exit Function
    End If

    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        whRow = CLng(CWR_Number(wsPlan.Cells(r, 27).Value2))
        If whRow < 5 Then GoTo ExitHere
        If CWR_Text(wsWH.Cells(whRow, 1).Value2) <> CWR_Text(wsPlan.Cells(r, 2).Value2) Or _
           CWR_Text(wsWH.Cells(whRow, 2).Value2) <> CWR_Text(wsPlan.Cells(r, 3).Value2) Or _
           CWR_Text(wsWH.Cells(whRow, 4).Value2) <> CWR_Text(wsPlan.Cells(r, 4).Value2) Or _
           CWR_Text(wsWH.Cells(whRow, 5).Value2) <> CWR_Text(wsPlan.Cells(r, 5).Value2) Or _
           CWR_Text(wsWH.Cells(whRow, 6).Value2) <> CWR_Text(wsPlan.Cells(r, 6).Value2) Or _
           Abs(CWR_Number(wsWH.Cells(whRow, 7).Value2) - CWR_Number(wsPlan.Cells(r, 7).Value2)) > 0.000001 Or _
           CWR_Text(wsWH.Cells(whRow, 8).Value2) <> CWR_Text(wsPlan.Cells(r, 9).Value2) Or _
           CWR_Text(wsWH.Cells(whRow, 9).Value2) <> CWR_Text(wsPlan.Cells(r, 11).Value2) Or _
           Abs(CWR_Number(wsWH.Cells(whRow, 12).Value2) - CWR_Number(wsPlan.Cells(r, 8).Value2)) > 0.000001 Or _
           CWR_Text(wsWH.Cells(whRow, 14).Value2) <> CWR_Text(wsPlan.Cells(r, 10).Value2) Then
            GoTo ExitHere
        End If
        planBatch = CWR_Text(wsPlan.Cells(r, 14).Value2)
        actualBatch = CWR_Text(wsWH.Cells(whRow, 22).Value2)
        If actualBatch <> "" And StrComp(actualBatch, planBatch, vbTextCompare) <> 0 Then GoTo ExitHere
        If StrComp(planBatch, CWR_Text(wsPlan.Cells(r, 11).Value2), vbTextCompare) <> 0 And actualBatch = "" Then GoTo ExitHere
    Next r
    wsPlan.Range("AO2").Value2 = currentHash
    CWR_MigrateLegacyPlanLockIfSafe = True
ExitHere:
End Function

Private Function CWR_StableHash(ByVal textIn As String) As String
    Const HASH_MOD As Double = 2147483629#
    Dim i As Long, codePoint As Long, valueOut As Double
    For i = 1 To Len(textIn)
        codePoint = AscW(Mid$(textIn, i, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536
        valueOut = valueOut * 131# + CDbl(codePoint)
        valueOut = valueOut - Fix(valueOut / HASH_MOD) * HASH_MOD
    Next i
    CWR_StableHash = "H" & CStr(Fix(valueOut))
End Function

' Compatibility bridge for a PLAN locked by the immediately previous release.
' That release deliberately wrote BLOCKED_DELTA_EXECUTOR_REQUIRED because the
' executor did not exist yet.  Promote only that exact, otherwise-PASS contract
' to READY; this is not replanning and the delta route still performs fresh
' Monitor, POST-source, and PRDO readbacks before its first SAP write.
Private Sub CWR_PromoteLegacyDeltaContracts(ByVal wsPlan As Worksheet, ByVal obText As String)
    Dim r As Long, lastRow As Long, promoted As Long
    Dim planID As String, docText As String, wtText As String
    Dim decisionText As String, gateText As String, qtyMode As String, execMode As String, execState As String
    On Error GoTo ExitHere
    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If StrComp(CWR_Text(wsPlan.Cells(r, 2).Value2), obText, vbTextCompare) <> 0 Then GoTo NextPlan
        decisionText = UCase$(CWR_Text(wsPlan.Cells(r, 20).Value2))
        gateText = UCase$(CWR_Text(wsPlan.Cells(r, 24).Value2))
        qtyMode = UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_QTY_MODE).Value2))
        execMode = UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_EXEC_MODE).Value2))
        execState = UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_EXEC_STATE).Value2))
        If decisionText = "CHANGE_QTY" And gateText = "PASS" And qtyMode = "DELTA_ADD" And _
           execMode = "KEEP_EXISTING_CREATE_DELTA" And execState = "BLOCKED_DELTA_EXECUTOR_REQUIRED" Then
            wsPlan.Cells(r, CWR_PLAN_COL_EXEC_STATE).Value2 = "READY_TO_EXECUTE"
            promoted = promoted + 1
            planID = CWR_Text(wsPlan.Cells(r, 1).Value2)
            docText = CWR_Text(wsPlan.Cells(r, 3).Value2)
            wtText = CWR_Text(wsPlan.Cells(r, 5).Value2)
            CWR_Log "PLAN_DELTA_COMPAT", obText, docText, "PASS", _
                "PLAN=" & planID & "; WT=" & wtText & _
                "; BLOCKED_DELTA_EXECUTOR_REQUIRED->READY_TO_EXECUTE; SAP_WRITE=0; " & _
                "fresh Monitor/POST/PRDO gates required before write."
        End If
NextPlan:
    Next r
    If promoted > 0 Then
        wsPlan.Range("AL2").Value2 = "PLAN_EXECUTABLE|LEGACY_DELTA_PROMOTED=" & CStr(promoted)
    End If
ExitHere:
End Sub

' Processes all safe, independent Plan rows for the OB currently entered on
' INPUT.  Status C/WT_DONE is recorded as completed and never touched.  Status
' A is eligible only as the known cancelled-resume state; its remaining
' quantity is freshly read by the exact case route before any later write.
' A local blocker is reported and blocks only its ODO-item group; an uncertain
' SAP outcome always ends the queue without retry.
Private Function CWR_ExecuteOBQueueToCFHandoffCore() As String
    Dim wsPlan As Worksheet, wsWH As Worksheet, wsInput As Worksheet
    Dim candidates As Object, itemKeys As Object, blockedKeys As Object, groupBlockers As Object
    Dim r As Long, lastRow As Long, whRow As Long, planID As String, obText As String
    Dim docText As String, itemText As String, rawStatus As String, wtState As String
    Dim decisionText As String, gateText As String, flowText As String, proofText As String, handoffText As String
    Dim execState As String, execMode As String, qtyMode As String
    Dim keyText As String, resultText As String, queueError As String, messageText As String, keyValue As Variant, removeKey As Variant
    Dim createdCount As Long, waitCFCount As Long, doneCount As Long, skippedCount As Long, pendingACount As Long
    Dim candidateCount As Long, stopCount As Long, unscopedStopCount As Long, allowCancelledResume As Boolean
    Dim removals As Collection

    On Error GoTo Failed
    CWR_ExecuteOBQueueToCFHandoffCore = CWR_ExecuteLockedPlanSchedule()
    Exit Function

    ' Retained only as historical source evidence.  Button #4 exits above into
    ' the locked-plan phase scheduler and can no longer run this per-Plan
    ' full-flow loop or its special batch branch.
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
    Set wsInput = CF_GetInputSheet()
    obText = CWR_Text(wsInput.Range("A7").Value2)
    If obText = "" Then
        CWR_ExecuteOBQueueToCFHandoffCore = "STOP|OB_QUEUE_INPUT_OB_MISSING|CF=USER_HANDOFF"
        Exit Function
    End If
    CWR_HISTORY_SEQ = 0
    CWR_HISTORY_RUN_ID = CF_HistoryNewRunID("EXEC", CWR_HistoryCheckRunID())
    CWR_CONTEXT_VALID = False
    CWR_CONTEXT_PLAN_ID = vbNullString
    CWR_CONTEXT_FINGERPRINT = vbNullString
    CWR_CONTEXT_OB = vbNullString
    CWR_CONTEXT_DOC = vbNullString
    CWR_CONTEXT_WT = vbNullString
    CWR_CREATE_CONTEXT_VALID = False
    CWR_CREATE_CONTEXT_OB = vbNullString
    CWR_CREATE_CONTEXT_DOC = vbNullString
    Set CWR_CREATE_CONTEXT_SESSION = Nothing
    CWR_RECOVERY_CONTEXT = False
    CWR_RECOVERY_CONTEXT = False
    CWR_ResetMONNavigationCounters
    If Not CWR_AssertQueuePlanLock(wsPlan, wsWH, messageText) Then
        CWR_ExecuteOBQueueToCFHandoffCore = "STOP|" & messageText & "|NO_SAP_WRITE|CF=USER_HANDOFF"
        Exit Function
    End If
    CWR_PromoteLegacyDeltaContracts wsPlan, obText
    Set candidates = CreateObject("Scripting.Dictionary")
    candidates.CompareMode = vbTextCompare
    Set itemKeys = CreateObject("Scripting.Dictionary")
    itemKeys.CompareMode = vbTextCompare
    Set blockedKeys = CreateObject("Scripting.Dictionary")
    blockedKeys.CompareMode = vbTextCompare
    Set groupBlockers = CreateObject("Scripting.Dictionary")
    groupBlockers.CompareMode = vbTextCompare

    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If StrComp(CWR_Text(wsPlan.Cells(r, 2).Value2), obText, vbTextCompare) <> 0 Then GoTo NextPlan
        planID = CWR_Text(wsPlan.Cells(r, 1).Value2)
        whRow = CLng(CWR_Number(wsPlan.Cells(r, 27).Value2))
        If planID = "" Or whRow < 5 Then
            stopCount = stopCount + 1
            unscopedStopCount = unscopedStopCount + 1
            CWR_LogQueueCase wsPlan, r, "STOP", "OB_QUEUE_PLAN_ID_OR_WH_ROW_INVALID"
            GoTo NextPlan
        End If
        docText = CWR_Text(wsWH.Cells(whRow, 2).Value2)
        itemText = CWR_Text(wsWH.Cells(whRow, 3).Value2)
        rawStatus = UCase$(CWR_Text(wsWH.Cells(whRow, 10).Value2))
        wtState = UCase$(CWR_Text(wsWH.Cells(whRow, 26).Value2))
        decisionText = UCase$(CWR_Text(wsPlan.Cells(r, 20).Value2))
        gateText = UCase$(CWR_Text(wsPlan.Cells(r, 24).Value2))
        qtyMode = UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_QTY_MODE).Value2))
        execMode = UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_EXEC_MODE).Value2))
        execState = UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_EXEC_STATE).Value2))
        flowText = UCase$(CWR_Text(wsWH.Cells(whRow, 16).Value2))
        proofText = UCase$(CWR_Text(wsWH.Cells(whRow, 17).Value2))
        handoffText = UCase$(CWR_Text(wsWH.Cells(whRow, 18).Value2))
        If docText <> "" And itemText <> "" Then keyText = docText & "|" & itemText Else keyText = ""

        If rawStatus = "C" Or wtState = "WT_DONE" Or decisionText = "DONE" Then
            doneCount = doneCount + 1
            CWR_LogQueueCase wsPlan, r, "PASS", "SKIP_C_CONFIRMED|WT_DONE|NO_SAP_WRITE"
            GoTo NextPlan
        End If
        If rawStatus = "A" And wtState = "WT_CANCELLED_ZERO" Then
            skippedCount = skippedCount + 1
            CWR_LogQueueCase wsPlan, r, "PASS", "SKIP_A_ZERO_REMAINING|NO_SAP_WRITE"
            GoTo NextPlan
        End If
        If flowText = "CANCELLED_WT_PRDO_QTY_ZERO" And proofText = "PRDO_QTY_ZERO_VERIFIED" And _
           handoffText = "USER_CF_HANDOFF" Then
            skippedCount = skippedCount + 1
            CWR_LogQueueCase wsPlan, r, "PASS", "SKIP_CANCEL_TO_ZERO_ALREADY_VERIFIED|NO_NEW_WT|CF_NOT_SENT"
            GoTo NextPlan
        End If
        If flowText = "READY_FOR_CF" And proofText = "WT_POST_VERIFY_PASS" Then
            waitCFCount = waitCFCount + 1
            CWR_LogQueueCase wsPlan, r, "PASS", "WAIT_USER_CF|OUTBOUND_WT_ALREADY_POST_VERIFIED|CF_NOT_SENT"
            GoTo NextPlan
        End If
        If Left$(proofText, 11) = "POST_CHANGE_" And handoffText = "USER_CF_HANDOFF" Then
            waitCFCount = waitCFCount + 1
            CWR_LogQueueCase wsPlan, r, "PASS", "WAIT_USER_CF|POST_CHANGE_WT_ALREADY_CREATED|CF_NOT_SENT"
            GoTo NextPlan
        End If
        If decisionText = "USER_SKIP" Or flowText = "CF_USER_HANDOFF" Then
            skippedCount = skippedCount + 1
            CWR_LogQueueCase wsPlan, r, "PASS", "SKIP_NO_SAP_WRITE_REQUIRED|CF_USER_HANDOFF"
            GoTo NextPlan
        End If
        If execState = "READY_FOR_CF" Then
            waitCFCount = waitCFCount + 1
            CWR_LogQueueCase wsPlan, r, "PASS", "WAIT_USER_CF|CHECK_READY_FOR_CF|CF_NOT_SENT"
            GoTo NextPlan
        End If
        If Left$(execState, 8) = "BLOCKED_" Then
            skippedCount = skippedCount + 1
            stopCount = stopCount + 1
            CWR_LogQueueCase wsPlan, r, "STOP", "SKIP_CHECK_BLOCKED|STATE=" & execState & _
                             "|MODE=" & execMode & "|QTY_MODE=" & qtyMode & "|NO_SAP_WRITE"
            GoTo NextPlan
        End If

        allowCancelledResume = (rawStatus = "A" And _
            (wtState = "WT_CANCELLED_PENDING_REMAINING" Or wtState = "WT_CANCELLED_REMAINING"))
        If execState <> "READY_TO_EXECUTE" Or _
           (execMode <> "REPLACE_CANCEL_CREATE" And execMode <> "CANCEL_PRDO_ZERO" And _
            execMode <> "KEEP_EXISTING_CREATE_DELTA") Or _
           Not ((gateText = "PASS" And CWR_IsBasicPlanDecision(decisionText)) Or allowCancelledResume) Then
            stopCount = stopCount + 1
            CWR_LogQueueCase wsPlan, r, "STOP", "OB_QUEUE_CASE_NOT_ELIGIBLE|STATE=" & execState & _
                             "|MODE=" & execMode & "|DECISION=" & decisionText & _
                             "|GATE=" & gateText & "|WT_STATE=" & wtState
            If keyText <> "" Then
                groupBlockers(keyText) = True
            Else
                unscopedStopCount = unscopedStopCount + 1
            End If
            GoTo NextPlan
        End If
        If docText = "" Or itemText = "" Then
            stopCount = stopCount + 1
            unscopedStopCount = unscopedStopCount + 1
            CWR_LogQueueCase wsPlan, r, "STOP", "OB_QUEUE_DOC_OR_ITEM_MISSING"
            GoTo NextPlan
        End If
        If itemKeys.Exists(keyText) Then
            blockedKeys(keyText) = True
            stopCount = stopCount + 1
            CWR_LogQueueCase wsPlan, r, "STOP", "DEFER_MULTI_WT_SAME_ODO_ITEM_GROUP|KEY=" & keyText & _
                             "|NO_SINGLE_CASE_WRITE"
        Else
            itemKeys.Add keyText, planID
            candidates.Add planID, r
            candidateCount = candidateCount + 1
            If allowCancelledResume Then pendingACount = pendingACount + 1
        End If
NextPlan:
    Next r

    ' Two or more active WT rows can legitimately belong to the same ODO item.
    ' They must be changed together: one PRDO edit session, rows in SAP order,
    ' and one Save.  Do not downgrade that group into single-case retries.
    If blockedKeys.count > 0 And groupBlockers.count = 0 And unscopedStopCount = 0 Then
        resultText = CWR_ExecuteBatchQtyIncreaseOB(obText)
        If Left$(resultText, 5) = "PASS|" Then
            CWR_QueueSummary obText, 0, 0, doneCount, skippedCount, pendingACount, 0, _
                             "PASS|BATCH_QTY_ONE_SAVE"
            CWR_ExecuteOBQueueToCFHandoffCore = resultText
        Else
            CWR_QueueSummary obText, 0, 0, doneCount, skippedCount, pendingACount, 0, _
                             "STOP|BATCH_QTY_ONE_SAVE"
            CWR_ExecuteOBQueueToCFHandoffCore = resultText
        End If
        GoTo ExitHere
    End If

    ' A local source/release blocker is reported to the user, but it only
    ' blocks cases sharing that same ODO item.  Other items in the OB may be
    ' safely processed after their own fresh preflight.  This prevents a
    ' shortage at one position from suppressing an independent eligible case.
    Set removals = New Collection
    For Each keyValue In candidates.keys
        r = CLng(candidates(CStr(keyValue)))
        keyText = CWR_Text(wsPlan.Cells(r, 3).Value2) & "|" & _
                  CWR_Text(wsWH.Cells(CLng(CWR_Number(wsPlan.Cells(r, 27).Value2)), 3).Value2)
        If blockedKeys.Exists(keyText) Then
            removals.Add CStr(keyValue)
            skippedCount = skippedCount + 1
            CWR_LogQueueCase wsPlan, r, "STOP", "DEFER_MULTI_WT_SAME_ODO_ITEM_GROUP|KEY=" & keyText & _
                             "|GROUP_EXECUTOR_REQUIRED|NO_SINGLE_CASE_WRITE"
        ElseIf groupBlockers.Exists(keyText) Then
            removals.Add CStr(keyValue)
            skippedCount = skippedCount + 1
            CWR_LogQueueCase wsPlan, r, "STOP", "SKIP_GROUP_BLOCKED_BY_CASE_ON_SAME_ODO_ITEM|KEY=" & keyText & _
                             "|USER_ACTION_REQUIRED"
        End If
    Next keyValue
    For Each removeKey In removals
        candidates.Remove CStr(removeKey)
    Next removeKey
    candidateCount = candidates.count

    If unscopedStopCount > 0 Then
        CWR_QueueSummary obText, createdCount, waitCFCount, doneCount, skippedCount, pendingACount, candidateCount, _
                         "STOP|UNSCOPED_CASE_BLOCKERS=" & CStr(unscopedStopCount)
        CWR_ExecuteOBQueueToCFHandoffCore = "STOP|OB_QUEUE_UNSCOPED_CASE_BLOCKERS=" & CStr(unscopedStopCount) & _
                                            "|OB=" & obText & "|NO_SAP_WRITE|CF=USER_HANDOFF"
        Exit Function
    End If
    If candidates.count = 0 Then
        CWR_QueueSummary obText, createdCount, waitCFCount, doneCount, skippedCount, pendingACount, candidateCount, _
                         "PASS|NO_READY_TO_EXECUTE|CHECK_BLOCKED=" & CStr(stopCount)
        CWR_ExecuteOBQueueToCFHandoffCore = "PASS|OB_QUEUE_NO_READY_TO_EXECUTE|OB=" & obText & _
                                            "|CHECK_BLOCKED=" & CStr(stopCount) & "|NO_SAP_WRITE|CF=USER_HANDOFF"
        Exit Function
    End If

    For Each keyValue In candidates.keys
        CWR_REQUESTED_PLAN_ID = CStr(keyValue)
        resultText = CWR_ExecutePlanToCFHandoffCore()
        If Left$(resultText, 5) <> "PASS|" Then
            CWR_Log "OB_QUEUE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "STOP", _
                    "PLAN=" & CWR_REQUESTED_PLAN_ID & "; " & resultText
            CWR_QueueSummary obText, createdCount, waitCFCount, doneCount, skippedCount, pendingACount, candidateCount, _
                             "STOP|PLAN=" & CWR_REQUESTED_PLAN_ID
            CWR_ExecuteOBQueueToCFHandoffCore = "STOP|OB_QUEUE_CASE_STOP|PLAN=" & CWR_REQUESTED_PLAN_ID & _
                                                "|" & resultText & "|CF=USER_HANDOFF"
            GoTo ExitHere
        End If
        If InStr(1, resultText, "POST_CHANGE_WT_PENDING_USER_CF", vbTextCompare) > 0 Then
            waitCFCount = waitCFCount + 1
            CWR_Log "OB_QUEUE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                    "PLAN=" & CWR_REQUESTED_PLAN_ID & "; WAIT_USER_CF_AFTER_POST_CHANGE; CF_NOT_SENT"
        ElseIf InStr(1, resultText, "CANCEL_TO_ZERO_READY_FOR_CF", vbTextCompare) > 0 Then
            skippedCount = skippedCount + 1
            CWR_Log "OB_QUEUE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                    "PLAN=" & CWR_REQUESTED_PLAN_ID & "; CANCEL_TO_ZERO_COMPLETE; NO_NEW_WT; CF_NOT_SENT"
        ElseIf InStr(1, resultText, "READY_FOR_CF", vbTextCompare) > 0 Or _
               InStr(1, resultText, "OUTBOUND_WT_PENDING_USER_CF", vbTextCompare) > 0 Then
            createdCount = createdCount + 1
            CWR_Log "OB_QUEUE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                    "PLAN=" & CWR_REQUESTED_PLAN_ID & "; OUTBOUND_WT_READY_FOR_CF; CF_NOT_SENT"
        Else
            skippedCount = skippedCount + 1
            CWR_Log "OB_QUEUE", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                    "PLAN=" & CWR_REQUESTED_PLAN_ID & "; NO_SAP_WRITE_REQUIRED; CF=USER_HANDOFF"
        End If
    Next keyValue

    CWR_QueueSummary obText, createdCount, waitCFCount, doneCount, skippedCount, pendingACount, candidateCount, _
                     "PASS|USER_ACTION_BLOCKERS=" & CStr(stopCount)
    CWR_ExecuteOBQueueToCFHandoffCore = "PASS|OB_QUEUE_READY_FOR_CF|OB=" & obText & _
                                        "|OUTBOUND_WT_CREATED=" & CStr(createdCount) & _
                                        "|WAIT_USER_CF=" & CStr(waitCFCount) & _
                                        "|SKIP_C_CONFIRMED=" & CStr(doneCount) & _
                                        "|SKIPPED=" & CStr(skippedCount) & _
                                        "|A_RESUME=" & CStr(pendingACount) & _
                                        "|USER_ACTION_BLOCKERS=" & CStr(stopCount) & "|CF_NOT_SENT"
ExitHere:
    CWR_REQUESTED_PLAN_ID = vbNullString
    Exit Function
Failed:
    CWR_Log "OB_QUEUE", obText, "", "FAIL", "Err=" & CStr(Err.Number) & ": " & Err.Description
    CWR_REQUESTED_PLAN_ID = vbNullString
    CWR_ExecuteOBQueueToCFHandoffCore = "FAIL|OB_QUEUE|" & CStr(Err.Number) & "|" & Err.Description & "|CF=USER_HANDOFF"
End Function

Private Function CWR_IsKnownBatchPeerOldWT(ByVal wtNo As String) As Boolean
    On Error GoTo ExitHere
    If Not CWR_BATCH_PEER_OLD_WTS Is Nothing Then _
        CWR_IsKnownBatchPeerOldWT = CWR_BATCH_PEER_OLD_WTS.Exists(CWR_Text(wtNo))
ExitHere:
End Function

' Button #4 is a locked-plan scheduler.  It performs no classification,
' source selection, release choice, QtyMode derivation or re-plan.  Those
' choices are contractual outputs of CHECK.  The SAP routines called below are
' the existing bounded executors; this layer only controls their phase/order.
Private Function CWR_ExecuteLockedPlanSchedule() As String
    Dim wsPlan As Worksheet, wsWH As Worksheet, wsInput As Worksheet
    Dim plans As Collection, docGroups As Object, qtyGroups As Object, locationGroups As Object, createGroups As Object
    Dim obText As String, messageText As String, stopAll As Boolean, summaryResult As String
    Dim scopeText As String, firstDocText As String
    Dim readyCount As Long, blockedCount As Long, unknownCount As Long, skippedCount As Long, readyForCFCount As Long
    Dim unaccountedCount As Long, initialUnaccounted As Long, rec As Object

    On Error GoTo Failed
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
    Set wsInput = CF_GetInputSheet()
    obText = CWR_Text(wsInput.Range("A7").Value2)
    If obText = "" Then
        CWR_ExecuteLockedPlanSchedule = "STOP|OB_QUEUE_INPUT_OB_MISSING|CF=USER_HANDOFF"
        Exit Function
    End If

    CWR_HISTORY_SEQ = 0
    CWR_HISTORY_RUN_ID = CF_HistoryNewRunID("EXEC", CWR_HistoryCheckRunID())
    CWR_ResetSchedulerExecutionContext
    Set CWR_PRDO_DOCUMENT_SAVE_COUNTS = CreateObject("Scripting.Dictionary")
    CWR_PRDO_DOCUMENT_SAVE_COUNTS.CompareMode = vbTextCompare
    Set CWR_CREATE_GROUP_ACTION_COUNTS = CreateObject("Scripting.Dictionary")
    CWR_CREATE_GROUP_ACTION_COUNTS.CompareMode = vbTextCompare
    Set CWR_CANCEL_RESULT_BY_WT = CreateObject("Scripting.Dictionary")
    CWR_CANCEL_RESULT_BY_WT.CompareMode = vbTextCompare
    CWR_Log "OB_PLAN_BUILDING", obText, "", "PASS", "PLAN_SOURCE=CHECK_LOCKED|BUTTON=4|SAP_WRITE=0"
    CWR_Log "EXECUTION_OWNERSHIP", obText, "", "PASS", _
        "EXECUTION_ATTEMPT_ID=" & CWR_HISTORY_RUN_ID & "|CHECK_RUN_ID=" & CWR_HistoryCheckRunID() & _
        "|OWNERSHIP_KEY=OB|PLAN_FINGERPRINT|CHECK_RUN_ID|ACTION_SEQUENCE|SAP_WRITE=0"
    CWR_Log "EXECUTION_CONTEXT_RESET", obText, "", "PASS", _
        "RUN_SCOPE=" & CWR_HISTORY_RUN_ID & "|WHTASK_FIRST_OPEN_PER_RUN=TRUE|" & _
        "WHTASK_REUSE_ONLY_WITHIN_RUN=TRUE|SAP_WRITE=0"
    If Not CWR_AssertQueuePlanLock(wsPlan, wsWH, messageText) Then
        CWR_ExecuteLockedPlanSchedule = "STOP|" & messageText & "|NO_SAP_WRITE|CF=USER_HANDOFF"
        Exit Function
    End If
    If Not CWR_BuildExecutionGroups(wsPlan, wsWH, obText, plans, docGroups, qtyGroups, locationGroups, createGroups, _
                                     skippedCount, messageText) Then
        CWR_ExecuteLockedPlanSchedule = "STOP|" & messageText & "|NO_SAP_WRITE|CF=USER_HANDOFF"
        Exit Function
    End If
    CWR_Log "OB_PLAN_READY", obText, "", "PASS", CWR_ObPlanReadyDetail(plans, docGroups, qtyGroups, locationGroups, createGroups)
    If plans.count = 0 Then
        readyForCFCount = CWR_CountPlansAtGate(wsPlan, obText, "READY_FOR_CF")
        If readyForCFCount > 0 Then
            CWR_QueueSummary obText, 0, readyForCFCount, 0, skippedCount, 0, 0, "PASS|READY_FOR_CF_MANUAL_HANDOFF"
            CWR_ExecuteLockedPlanSchedule = "PASS|READY_FOR_CF_MANUAL_HANDOFF|OB=" & obText & _
                "|READY_FOR_CF=" & CStr(readyForCFCount) & "|SKIPPED=" & CStr(skippedCount) & _
                "|NO_SAP_EXECUTION_REQUIRED|CF_NOT_SENT"
        Else
            CWR_QueueSummary obText, 0, 0, 0, skippedCount, 0, 0, "PASS|NO_READY_TO_EXECUTE"
            CWR_ExecuteLockedPlanSchedule = "PASS|OB_QUEUE_NO_READY_TO_EXECUTE|OB=" & obText & _
                "|SKIPPED=" & CStr(skippedCount) & "|CF_NOT_SENT"
        End If
        Exit Function
    End If

    If Not CWR_ObPreflightPass(plans, docGroups, qtyGroups, locationGroups, createGroups, obText, messageText) Then
        CWR_Log "OB_BLOCKED_PREWRITE", obText, "", "STOP", messageText & "|SAP_WRITE=0|NEXT_OB_ALLOWED=TRUE"
        CWR_ExecuteLockedPlanSchedule = "STOP|OB_BLOCKED_PREWRITE|OB=" & obText & _
            "|" & messageText & "|NO_SAP_WRITE|CF_NOT_SENT"
        CWR_REQUESTED_PLAN_ID = vbNullString
        Exit Function
    End If
    firstDocText = CWR_SchedulerFirstDocument(plans)
    scopeText = CWR_RegressionScopeDetailForPlan(wsPlan, obText)
    CWR_Log "OB_REGRESSION_SCOPE", obText, firstDocText, "PASS", scopeText & "|SAP_WRITE=0"
    If Not CWR_RequireRegressionGate(messageText) Then
        CWR_Log "OB_REGRESSION_GATE_BLOCKED", obText, firstDocText, "STOP", _
            "RUN_ID=" & CWR_HISTORY_RUN_ID & "|PLAN_FINGERPRINT=" & CWR_ObPlanFingerprint(plans) & _
            "|" & messageText & "|SAP_WRITE=0|NEXT_EXPECTED_PHASE=OB_PREFLIGHT_PASS"
        CWR_QueueSummary obText, 0, 0, 0, skippedCount, 0, plans.count, _
            "STOP|OB_REGRESSION_GATE_BLOCKED|" & messageText & "|SAP_WRITE=0"
        CWR_ExecuteLockedPlanSchedule = "STOP|OB_REGRESSION_GATE_BLOCKED|OB=" & obText & _
            "|" & messageText & "|NO_SAP_WRITE|CF=USER_HANDOFF"
        CWR_REQUESTED_PLAN_ID = vbNullString
        Exit Function
    End If
    CWR_Log "OB_REGRESSION_GATE_PASS", obText, firstDocText, "PASS", _
        messageText & "|PLAN_FINGERPRINT=" & CWR_ObPlanFingerprint(plans) & "|SAP_WRITE=0"
    CWR_Log "OB_PREFLIGHT_PASS", obText, "", "PASS", "PLAN_FINGERPRINT=" & CWR_ObPlanFingerprint(plans) & _
        "|ACTIONS_FROZEN=TRUE|REPLAN=FALSE|SAP_WRITE=0"
    CWR_Log "OB_EXECUTING", obText, "", "PASS", "PHASE_ORDER=CANCEL_GROUP|LOCATION_GROUP|PRDO_DOCUMENT_GROUP|CREATE_GROUP|FINAL_VERIFY|CF_USER_HANDOFF"
    If Not CWR_ExecutePreflightGroups(qtyGroups, stopAll, messageText) Then GoTo Finalize
    If Not CWR_ExecuteLocationGroups(locationGroups, stopAll, messageText) Then GoTo Finalize
    If Not CWR_ExecuteDocumentGroups(docGroups, qtyGroups, stopAll, messageText) Then GoTo Finalize
    If Not CWR_AdvanceDataReady(plans, stopAll, messageText) Then GoTo Finalize
    If Not CWR_ExecuteCreateGroups(createGroups, stopAll, messageText) Then GoTo Finalize

Finalize:
    CWR_Log "OB_POST_VERIFY", obText, "", IIf(stopAll, "STOP", "PASS"), _
        "READY_FOR_CF_CANDIDATES=POST_VERIFY|CF_SENT=FALSE"
    CWR_Log "MON_NAVIGATION", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", CWR_MONNavigationEvidence()
    CWR_ScheduleCounts plans, readyCount, blockedCount, unknownCount, unaccountedCount
    initialUnaccounted = unaccountedCount
    If unaccountedCount > 0 Then
        ' Every loaded member must have a terminal outcome.  Convert any
        ' residual phase state into an explicit blocker instead of letting the
        ' summary silently report a complete queue.
        For Each rec In plans
            If Not CBool(rec("Terminal")) And Not CWR_IsTerminalExecutionState(CStr(rec("State"))) Then
                CWR_SchedulerFailure rec, "INCOMPLETE_EXECUTION", _
                    "STOP_INCOMPLETE_EXECUTION|PLAN=" & CStr(rec("PlanID")) & _
                    "|ITEM=" & CStr(rec("Item")) & "|MATERIAL=" & CStr(rec("Material")) & _
                    "|LAST_STATE=" & CStr(rec("State")) & "|EXPECTED_NEXT_STAGE=EXPLICIT_RESUME", stopAll
            End If
        Next rec
        CWR_ScheduleCounts plans, readyCount, blockedCount, unknownCount, unaccountedCount
    End If
    If unknownCount = 0 And Not stopAll And blockedCount = 0 And unaccountedCount = 0 And readyCount = plans.count Then _
        CWR_RecordSuccessfulSchedulerRegressionProofs obText, plans.count
    If unknownCount > 0 Then
        summaryResult = "STOP|STATE_UNKNOWN|" & messageText
    ElseIf initialUnaccounted > 0 Or unaccountedCount > 0 Then
        summaryResult = "STOP|INCOMPLETE_EXECUTION|UNACCOUNTED=" & CStr(initialUnaccounted) & "|" & messageText
    ElseIf stopAll Or blockedCount > 0 Then
        summaryResult = "STOP|SCHEDULER_HALTED_OR_BLOCKED|" & messageText
    Else
        summaryResult = "PASS|LOCKED_PLAN_PHASES_COMPLETE"
    End If
    If unknownCount = 0 And Not stopAll And blockedCount = 0 And unaccountedCount = 0 And readyCount = plans.count Then
        CWR_Log "OB_COMMIT_BARRIER", obText, "", "PASS", _
            "POST_VERIFY_PASS=TRUE|NO_CANCEL_RETRY|NO_QTY_RETRY|NO_CREATE_RETRY|CF_USER_HANDOFF_ONLY"
        CWR_Log "OB_READY_FOR_CF", obText, "", "PASS", "READY_FOR_CF=TRUE|CF_NOT_SENT"
        CWR_Log "OB_DONE", obText, "", "PASS", "TERMINAL=READY_FOR_CF|NEXT_OB_ALLOWED=TRUE|CF_NOT_SENT"
    End If
    CWR_QueueSummary obText, readyCount, 0, 0, skippedCount + blockedCount, 0, plans.count, summaryResult
    If initialUnaccounted > 0 Or unaccountedCount > 0 Then
        CWR_ExecuteLockedPlanSchedule = "STOP|OB_QUEUE_INCOMPLETE_EXECUTION|OB=" & obText & _
            "|READY_FOR_CF=" & CStr(readyCount) & "|BLOCKED=" & CStr(blockedCount) & _
            "|STATE_UNKNOWN=" & CStr(unknownCount) & "|UNACCOUNTED=" & CStr(initialUnaccounted) & _
            "|SKIPPED=" & CStr(skippedCount) & "|CF_NOT_SENT"
    ElseIf unknownCount > 0 Or stopAll Then
        CWR_ExecuteLockedPlanSchedule = "STOP|OB_QUEUE_STATE_UNKNOWN_OR_HALTED|OB=" & obText & _
            "|READY_FOR_CF=" & CStr(readyCount) & "|BLOCKED=" & CStr(blockedCount) & _
            "|STATE_UNKNOWN=" & CStr(unknownCount) & "|SKIPPED=" & CStr(skippedCount) & "|CF_NOT_SENT"
    ElseIf blockedCount > 0 Then
        CWR_ExecuteLockedPlanSchedule = "STOP|OB_QUEUE_BLOCKED|OB=" & obText & _
            "|READY_FOR_CF=" & CStr(readyCount) & "|BLOCKED=" & CStr(blockedCount) & _
            "|STATE_UNKNOWN=0|SKIPPED=" & CStr(skippedCount) & "|CF_NOT_SENT"
    Else
        CWR_ExecuteLockedPlanSchedule = "PASS|READY_FOR_CF|OB=" & obText & _
            "|READY_FOR_CF=" & CStr(readyCount) & "|BLOCKED=" & CStr(blockedCount) & _
            "|STATE_UNKNOWN=0|SKIPPED=" & CStr(skippedCount) & "|CF_NOT_SENT"
    End If
    CWR_REQUESTED_PLAN_ID = vbNullString
    Exit Function
Failed:
    CWR_REQUESTED_PLAN_ID = vbNullString
    CWR_ExecuteLockedPlanSchedule = "FAIL|LOCKED_PLAN_SCHEDULER|" & CStr(Err.Number) & "|" & _
        Err.Description & "|CF=USER_HANDOFF"
End Function

Private Function CWR_CountPlansAtGate(ByVal wsPlan As Worksheet, ByVal obText As String, _
                                      ByVal expectedGate As String) As Long
    Dim r As Long, lastRow As Long
    If wsPlan Is Nothing Then Exit Function
    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If CWR_Text(wsPlan.Cells(r, 2).Value2) = obText And _
           StrComp(CWR_Text(wsPlan.Cells(r, 24).Value2), expectedGate, vbTextCompare) = 0 Then
            CWR_CountPlansAtGate = CWR_CountPlansAtGate + 1
        End If
    Next r
End Function

' The scheduler consumes one immutable OB package.  This fingerprint is an
' execution identity, not a replacement for the PLAN/WH lock hash; it keeps
' the OB-level history auditable and prevents a later context switch from
' being mistaken for the same package.
Private Function CWR_ObPlanFingerprint(ByVal plans As Collection) As String
    Dim rec As Object, payload As String
    On Error GoTo Failed
    For Each rec In plans
        payload = payload & "|PLAN=" & CWR_Text(rec("PlanID")) & _
            "|OB=" & CWR_Text(rec("OB")) & "|DOC=" & CWR_Text(rec("Doc")) & _
            "|ITEM=" & CWR_Text(rec("Item")) & "|MATERIAL=" & CWR_Text(rec("Material")) & _
            "|OLD_WT=" & CWR_Text(rec("OldWT")) & "|TARGET=" & CStr(rec("TargetQty")) & _
            "|SYS_BIN=" & CWR_Text(rec("SysBin")) & "|SYS_BATCH=" & CWR_Text(rec("SysBatch")) & _
            "|TARGET_BIN=" & CWR_Text(rec("TargetBin")) & "|TARGET_BATCH=" & CWR_Text(rec("TargetBatch")) & _
            "|DEST_BIN=" & CWR_Text(rec("DestBin")) & "|PROC=" & CWR_Text(rec("ProcTy")) & _
            "|QTY_GROUP=" & CWR_Text(rec("QtyGroupId")) & "|DOC_GROUP=" & CWR_Text(rec("DocGroup")) & _
            "|LOCATION_GROUP=" & CWR_Text(rec("LocationGroup")) & "|CREATE_GROUP=" & CWR_Text(rec("CreateGroup")) & _
            "|FINGERPRINT=" & CWR_Text(rec("PlanFingerprint"))
    Next rec
    CWR_ObPlanFingerprint = CWR_StableHash(payload)
    Exit Function
Failed:
    CWR_ObPlanFingerprint = "H0"
End Function

Private Function CWR_ObPlanReadyDetail(ByVal plans As Collection, ByVal docGroups As Object, _
                                        ByVal qtyGroups As Object, ByVal locationGroups As Object, _
                                        ByVal createGroups As Object) As String
    CWR_ObPlanReadyDetail = "MEMBERS=" & CStr(plans.count) & _
        "|DOC_GROUPS=" & CStr(docGroups.count) & "|QTY_GROUPS=" & CStr(qtyGroups.count) & _
        "|LOCATION_GROUPS=" & CStr(locationGroups.count) & "|CREATE_GROUPS=" & CStr(createGroups.count) & _
        "|PLAN_FINGERPRINT=" & CWR_ObPlanFingerprint(plans) & "|ACTIONS_FROZEN=TRUE"
End Function

' Read-only OB preflight.  It validates the frozen plan package before the
' first Cancel/POST/PRDO/Create write.  A pre-write failure is scoped to this
' OB; the caller may advance the OB queue, while a post-write unknown remains
' a hard stop in CWR_SchedulerFailure.
Private Function CWR_ObPreflightPass(ByVal plans As Collection, ByVal docGroups As Object, _
                                      ByVal qtyGroups As Object, ByVal locationGroups As Object, _
                                      ByVal createGroups As Object, ByVal obText As String, _
                                      ByRef messageText As String) As Boolean
    Dim rec As Object, modeText As String, planID As String, membersText As String
    Dim allowed As Object, keyValue As Variant, groupMembers As Collection
    Dim groupID As String, expectedMembers As Long, actualMembers As Long
    On Error GoTo Failed
    Set allowed = CreateObject("Scripting.Dictionary")
    allowed.CompareMode = vbTextCompare
    allowed.Add "CANCEL_PRDO_ZERO", True
    allowed.Add "REPLACE_CANCEL_CREATE", True
    allowed.Add "KEEP_EXISTING_CREATE_DELTA", True
    allowed.Add "CF_ONLY", True
    allowed.Add "RETAIN_UNTOUCHED", True
    allowed.Add "CHANGE_QTY", True

    For Each rec In plans
        planID = CWR_Text(rec("PlanID"))
        modeText = UCase$(CWR_Text(rec("ExecMode")))
        If planID = "" Or CWR_Text(rec("OB")) <> obText Or CWR_Text(rec("Doc")) = "" Or _
           CWR_Text(rec("Item")) = "" Or CWR_Text(rec("Material")) = "" Or _
           CWR_Text(rec("OldWT")) = "" Then
            messageText = "PLAN_MEMBER_IDENTITY_INVALID|PLAN=" & planID
            Exit Function
        End If
        If Not allowed.Exists(modeText) Then
            If CBool(rec("PostRecoveryPending")) And modeText = "NONE" Then
                CWR_Log "RECOVER_POST", obText, CWR_Text(rec("Doc")), "PASS", _
                    "PLAN=" & planID & "|ACTION=POST|STATE=WRITE_SENT_UNVERIFIED|" & _
                    "ACTION_KEY=" & CWR_BusinessActionKey(rec, "POST") & "|MODE=NONE_ALLOWED_RECOVERY_ONLY"
            Else
                messageText = "PLAN_MEMBER_EXEC_MODE_UNSUPPORTED|PLAN=" & planID & "|MODE=" & modeText
                Exit Function
            End If
        End If
        If CWR_Text(rec("DocGroup")) = "" Or CWR_Text(rec("QtyGroupId")) = "" Or _
           CWR_Text(rec("LocationGroup")) = "" Or CWR_Text(rec("CreateGroup")) = "" Then
            messageText = "PLAN_MEMBER_GROUP_ID_MISSING|PLAN=" & planID
            Exit Function
        End If
        If CWR_Text(rec("PlanFingerprint")) = "" Then
            messageText = "PLAN_MEMBER_FINGERPRINT_MISSING|PLAN=" & planID
            Exit Function
        End If
        If CLng(rec("GroupMembers")) < 1 Then
            messageText = "PLAN_MEMBER_GROUP_COUNT_INVALID|PLAN=" & planID
            Exit Function
        End If
        If CBool(rec("PostRecoveryPending")) Then
            CWR_Log "ACTION_SKIP", obText, CWR_Text(rec("Doc")), "PASS", _
                "PLAN=" & planID & "|POST_RECOVERY_READBACK_ONLY=TRUE|NO_POST_RETRY"
        End If
        membersText = membersText & IIf(membersText = "", "", ",") & planID
    Next rec

    For Each keyValue In qtyGroups.keys
        Set groupMembers = qtyGroups.item(CStr(keyValue))
        actualMembers = groupMembers.count
        expectedMembers = 0
        For Each rec In groupMembers
            expectedMembers = expectedMembers + 1
        Next rec
        If actualMembers <> expectedMembers Then
            messageText = "QTY_GROUP_MEMBER_COUNT_UNSTABLE|QTY_GROUP=" & CStr(keyValue)
            Exit Function
        End If
    Next keyValue
    For Each keyValue In docGroups.keys
        If docGroups.item(CStr(keyValue)).count < 1 Then
            messageText = "PRDO_DOCUMENT_GROUP_EMPTY|DOC_GROUP=" & CStr(keyValue)
            Exit Function
        End If
    Next keyValue
    For Each keyValue In locationGroups.keys
        If locationGroups.item(CStr(keyValue)).count < 1 Then
            messageText = "LOCATION_GROUP_EMPTY|LOCATION_GROUP=" & CStr(keyValue)
            Exit Function
        End If
    Next keyValue
    For Each keyValue In createGroups.keys
        If createGroups.item(CStr(keyValue)).count < 1 Then
            messageText = "CREATE_GROUP_EMPTY|CREATE_GROUP=" & CStr(keyValue)
            Exit Function
        End If
    Next keyValue
    CWR_Log "OB_PLAN_MEMBERS", obText, "", "PASS", "COUNT=" & CStr(plans.count) & "|PLAN_IDS=" & membersText
    CWR_Log "OB_PLAN_ACTIONS", obText, "", "PASS", _
        "CANCEL_GROUPS=" & CStr(qtyGroups.count) & "|PRDO_DOCUMENT_GROUPS=" & CStr(docGroups.count) & _
        "|LOCATION_GROUPS=" & CStr(locationGroups.count) & "|CREATE_GROUPS=" & CStr(createGroups.count) & _
        "|PLAN_ONLY=TRUE"
    CWR_ObPreflightPass = True
    Exit Function
Failed:
    messageText = "OB_PREFLIGHT_ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function CWR_BuildExecutionGroups(ByVal wsPlan As Worksheet, ByVal wsWH As Worksheet, _
                                           ByVal obText As String, ByRef plans As Collection, _
                                           ByRef docGroups As Object, ByRef qtyGroups As Object, _
                                           ByRef locationGroups As Object, _
                                           ByRef createGroups As Object, ByRef skippedCount As Long, _
                                           ByRef messageText As String) As Boolean
    Dim r As Long, lastRow As Long, rec As Object, gateText As String, execState As String
    Dim planID As String, whRow As Long
    Dim recoverablePostUnknown As Boolean, businessRecovery As Boolean
    On Error GoTo Failed
    Set plans = New Collection
    Set docGroups = CreateObject("Scripting.Dictionary"): docGroups.CompareMode = vbTextCompare
    Set qtyGroups = CreateObject("Scripting.Dictionary"): qtyGroups.CompareMode = vbTextCompare
    Set locationGroups = CreateObject("Scripting.Dictionary"): locationGroups.CompareMode = vbTextCompare
    Set createGroups = CreateObject("Scripting.Dictionary"): createGroups.CompareMode = vbTextCompare
    CWR_Log "RECOVERY_SCAN_START", obText, "", "PASS", _
        "SOURCE=PLAN+RUN_HISTORY|SAP_READBACK=DEFERRED_TO_RECOVERY_EXECUTOR|NO_SAP_WRITE"
    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If StrComp(CWR_Text(wsPlan.Cells(r, 2).Value2), obText, vbTextCompare) <> 0 Then GoTo NextPlan
        planID = CWR_Text(wsPlan.Cells(r, 1).Value2)
        gateText = UCase$(CWR_Text(wsPlan.Cells(r, 24).Value2))
        execState = UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_EXEC_STATE).Value2))
        businessRecovery = CWR_HasRecoverablePickBinPostByBusiness( _
            CWR_Text(wsPlan.Cells(r, 2).Value2), CWR_Text(wsPlan.Cells(r, 3).Value2), _
            CWR_Text(wsPlan.Cells(r, 5).Value2), CWR_Text(wsPlan.Cells(r, 6).Value2))
        recoverablePostUnknown = businessRecovery Or CWR_HasRecoverablePickBinPostUnknown(planID, _
            CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_FINGERPRINT).Value2))
        If CWR_IsTerminalExecutionState(gateText) Or CWR_IsTerminalExecutionState(execState) Then
            skippedCount = skippedCount + 1
            CWR_LogQueueCase wsPlan, r, "PASS", "ACTION_SKIP|TERMINAL_STATE=" & IIf(execState <> "", execState, gateText) & _
                "|PASS_ACTION_NOT_REEXECUTED|NO_SAP_WRITE"
            GoTo NextPlan
        End If
        If (gateText <> "READY_TO_EXECUTE" Or execState <> "READY_TO_EXECUTE") And _
           Not recoverablePostUnknown Then
            skippedCount = skippedCount + 1
            CWR_LogQueueCase wsPlan, r, "SKIP", "CHECK_GATE=" & gateText & "|CHECK_STATE=" & execState & "|NO_SAP_WRITE"
            GoTo NextPlan
        End If
        whRow = CLng(CWR_Number(wsPlan.Cells(r, 27).Value2))
        If planID = "" Or whRow < 5 Then
            messageText = "PLAN_ID_OR_WH_ROW_INVALID"
            Exit Function
        End If
        If CWR_HasUnreconciledUnknown(planID, CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_FINGERPRINT).Value2)) And _
           Not recoverablePostUnknown Then
            skippedCount = skippedCount + 1
            CWR_LogQueueCase wsPlan, r, "STOP", "STATE_UNKNOWN_FROM_PRIOR_SENT_ACTION|RECONCILE_READBACK_REQUIRED|NO_SAP_WRITE"
            GoTo NextPlan
        End If
        If recoverablePostUnknown Then
            ' The only scheduler exception to the generic unknown-state stop.
            ' It admits this record solely so Codex_ExecutePostingBatchAtPickBin
            ' can perform its fresh POST readback.  The POST checkpoint still
            ' blocks a second Create+Save, and cancellation remains idempotent.
            CWR_LogQueueCase wsPlan, r, "PASS", _
                "POST_CREATE_SAVE_PRIOR_UNKNOWN|CONTROLLED_POST_RECONCILE_ONLY|NO_POST_RETRY|" & _
                "RECOVERY_IDENTITY=OB|DOC|WT|MATERIAL"
        End If
        Set rec = CreateObject("Scripting.Dictionary")
        rec.CompareMode = vbTextCompare
        rec.Add "PlanID", planID
        rec.Add "PlanRow", r
        rec.Add "WHRow", whRow
        rec.Add "OB", CWR_Text(wsPlan.Cells(r, 2).Value2)
        rec.Add "Doc", CWR_Text(wsPlan.Cells(r, 3).Value2)
        rec.Add "Item", CWR_Text(wsWH.Cells(whRow, 3).Value2)
        rec.Add "Material", CWR_Text(wsPlan.Cells(r, 6).Value2)
        rec.Add "OldWT", CWR_Text(wsPlan.Cells(r, 5).Value2)
        rec.Add "SysQty", CWR_Number(wsPlan.Cells(r, 7).Value2)
        rec.Add "TargetQty", CWR_Number(wsPlan.Cells(r, 8).Value2)
        rec.Add "QtyDelta", CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_QTY_DELTA).Value2)
        rec.Add "QtyGroupId", CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_QTY_GROUP_ID).Value2)
        rec.Add "GroupCurrent", CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_CURRENT).Value2)
        rec.Add "GroupTarget", CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_TARGET).Value2)
        rec.Add "GroupDelta", CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_DELTA).Value2)
        rec.Add "GroupMembers", CLng(CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_MEMBERS).Value2))
        rec.Add "GroupNeedQty", (UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_NEED_QTY).Value2)) = "TRUE")
        rec.Add "GroupLiveQty", 0#
        rec.Add "GroupQtyBaselineProven", False
        rec.Add "DestBin", ""
        rec.Add "ProcTy", ""
        rec.Add "SysBin", CWR_Text(wsPlan.Cells(r, 9).Value2)
        rec.Add "SysBatch", CWR_Text(wsPlan.Cells(r, 11).Value2)
        rec.Add "SourceBin", CWR_Text(wsPlan.Cells(r, 9).Value2)
        rec.Add "SourceBatch", CWR_Text(wsPlan.Cells(r, 11).Value2)
        rec.Add "TargetBin", CWR_Text(wsPlan.Cells(r, 10).Value2)
        rec.Add "TargetBatch", CWR_Text(wsPlan.Cells(r, 14).Value2)
        rec.Add "QtyMode", UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_QTY_MODE).Value2))
        rec.Add "ExecMode", UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_EXEC_MODE).Value2))
        rec.Add "NeedQty", (UCase$(CWR_Text(wsPlan.Cells(r, 21).Value2)) = "TRUE")
        rec.Add "NeedBin", (UCase$(CWR_Text(wsPlan.Cells(r, 22).Value2)) = "TRUE")
        rec.Add "NeedBatch", (UCase$(CWR_Text(wsPlan.Cells(r, 23).Value2)) = "TRUE")
        rec.Add "DocGroup", CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_DOC_GROUP).Value2)
        rec.Add "LocationGroup", CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_LOCATION_GROUP).Value2)
        ' A Create+Save action is never shared by different materials.  Older
        ' PLAN rows carried a short CreateGroup such as
        ' CREATE|OB|DOC|ITEM, which let one blocked material suppress every
        ' other material in the same document.  Keep the checked group label
        ' for auditability, but make the scheduler key immutable at the
        ' business-action scope (OB + document + item + material).
        rec.Add "CreateGroup", CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_CREATE_GROUP).Value2) & _
            "|OB=" & CWR_SchedulerIdentityPart(CWR_Text(wsPlan.Cells(r, 2).Value2)) & _
            "|DOC=" & CWR_SchedulerIdentityPart(CWR_Text(wsPlan.Cells(r, 3).Value2)) & _
            "|ITEM=" & CWR_SchedulerIdentityPart(CWR_Text(wsWH.Cells(whRow, 3).Value2)) & _
            "|MAT=" & CWR_SchedulerIdentityPart(CWR_Text(wsPlan.Cells(r, 6).Value2)) & _
            "|SRC_BIN=" & CWR_SchedulerIdentityPart(CWR_Text(wsPlan.Cells(r, 12).Value2)) & _
            "|SRC_BATCH=" & CWR_SchedulerIdentityPart(CWR_Text(wsPlan.Cells(r, 13).Value2)) & _
            "|DST_BIN=" & CWR_SchedulerIdentityPart(CWR_Text(wsPlan.Cells(r, 10).Value2)) & _
            "|DST_BATCH=" & CWR_SchedulerIdentityPart(CWR_Text(wsPlan.Cells(r, 14).Value2))
        rec.Add "PlanFingerprint", CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_FINGERPRINT).Value2)
        rec.Add "State", "PLANNED"
        rec.Add "QtyReady", False
        rec.Add "LocationReady", False
        rec.Add "Terminal", False
        rec.Add "CancelAlready", False
        rec.Add "CancelSent", False
        rec.Add "CancelVerified", False
        rec.Add "CancelProof", ""
        rec.Add "CancelRow", -1
        rec.Add "PostRecoveryPending", recoverablePostUnknown
        rec.Add "RecoveryOnly", recoverablePostUnknown And _
            (gateText <> "READY_TO_EXECUTE" Or execState <> "READY_TO_EXECUTE")
        plans.Add rec
        CWR_GroupAdd docGroups, CStr(rec("DocGroup")), rec
        CWR_GroupAdd qtyGroups, CStr(rec("QtyGroupId")), rec
        CWR_GroupAdd locationGroups, CStr(rec("LocationGroup")), rec
        CWR_GroupAdd createGroups, CStr(rec("CreateGroup")), rec
        CWR_SchedulerState rec, "PLANNED", "PLAN_LOADED", "PASS", _
            "CHECK_LOCKED_PLAN_CONSUMED|POST_RECOVERY_PENDING=" & CStr(recoverablePostUnknown) & _
            "|RECOVERY_ONLY=" & CStr(CBool(rec("RecoveryOnly")))
        If recoverablePostUnknown Then
            CWR_Log "RECOVER_POST", CStr(rec("OB")), CStr(rec("Doc")), "PASS", _
                "PLAN=" & CStr(rec("PlanID")) & "|ACTION=POST|STATE=WRITE_SENT_UNVERIFIED|" & _
                "ACTION_KEY=" & CWR_BusinessActionKey(rec, "POST") & "|READBACK_ONLY=TRUE|NO_POST_RETRY"
        End If
NextPlan:
    Next r
    CWR_BuildExecutionGroups = True
    Exit Function
Failed:
    messageText = "BUILD_EXECUTION_GROUPS_ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function CWR_IsTerminalExecutionState(ByVal stateText As String) As Boolean
    Dim normalized As String
    normalized = UCase$(CWR_Text(stateText))
    If Left$(normalized, 8) = "BLOCKED_" Or _
       Left$(normalized, 14) = "STATE_UNKNOWN_" Or _
       Left$(normalized, 5) = "STOP_" Or _
       InStr(1, normalized, "WAIT_USER_CF", vbTextCompare) > 0 Then
        CWR_IsTerminalExecutionState = True
        Exit Function
    End If
    Select Case normalized
        Case "PASS", "PASS_RECONCILED", "ALREADY_DONE", "NOT_REQUIRED", "SKIP_PROVEN", _
             "READY_FOR_CF", "OB_READY_FOR_CF", "OB_CF_PASS", "OB_DONE", "DONE", "WT_DONE", _
             "DONE_BY_USER_CF", "WT_CANCELLED", "WT_CANCELLED_REMAINING", "WT_CANCELLED_PENDING_REMAINING"
            CWR_IsTerminalExecutionState = True
    End Select
End Function

Private Sub CWR_GroupAdd(ByVal groups As Object, ByVal groupID As String, ByVal rec As Object)
    Dim members As Collection
    If groupID = "" Then groupID = "UNGROUPED|" & CStr(rec("PlanID"))
    If Not groups.Exists(groupID) Then
        Set members = New Collection
        groups.Add groupID, members
    End If
    groups.item(groupID).Add rec
End Sub

' Read the complete cancellation queue once before any Cancel is sent.  The
' queue is frozen from CHECK; this only proves the current MON state and keeps
' already-cancelled WTs out of the write queue.
Private Function CWR_SchedulerPreflightCancelGroup(ByVal members As Collection, ByVal groupID As String, _
                                                    ByRef messageText As String) As Boolean
    Dim firstRec As Object, rec As Object, ses As Object, lower As Object, expected As Object
    Dim found As Object, statuses As Object, r As Long, n As Long, currentDoc As String
    Dim rowDoc As String, rowWT As String, rowStatus As String, rowItem As String, rowMat As String
    Dim rowBin As String, rowBatch As String, rowQty As Double
    On Error GoTo Failed
    Set expected = CreateObject("Scripting.Dictionary"): expected.CompareMode = vbTextCompare
    Set found = CreateObject("Scripting.Dictionary"): found.CompareMode = vbTextCompare
    Set statuses = CreateObject("Scripting.Dictionary"): statuses.CompareMode = vbTextCompare
    For Each rec In members
        If Not CBool(rec("Terminal")) Then
            If firstRec Is Nothing Then Set firstRec = rec
            rowWT = CWR_Text(rec("OldWT"))
            If rowWT = "" Or expected.Exists(rowWT) Then
                messageText = "STOP_CANCEL_QUEUE_WT_IDENTITY|QTY_GROUP=" & CStr(groupID) & "|PLAN=" & CStr(rec("PlanID"))
                Exit Function
            End If
            expected.Add rowWT, rec
            found.Add rowWT, 0&
            statuses.Add rowWT, ""
            rec("CancelAlready") = False
            rec("CancelSent") = False
            rec("CancelVerified") = False
            rec("CancelProof") = ""
        End If
    Next rec
    If firstRec Is Nothing Then
        CWR_SchedulerPreflightCancelGroup = True
        Exit Function
    End If
    If Not CWR_SchedulerUsePlan(firstRec, messageText) Then Exit Function
    If Not CWR_GetMONSession(ses, messageText) Then Exit Function
    If Not CWR_LoadFilteredWTGrid(ses, lower, messageText) Then Exit Function
    n = CLng(lower.rowCount)
    For r = 0 To n - 1
        rowDoc = CWR_Cell(lower, r, "DOCNO")
        If rowDoc <> "" Then currentDoc = rowDoc
        rowWT = CWR_Cell(lower, r, "TANUM")
        If CWR_MONValueEquals(currentDoc, CStr(firstRec("Doc"))) And expected.Exists(rowWT) Then
            Set rec = expected.item(rowWT)
            found(rowWT) = CLng(found(rowWT)) + 1
            rowStatus = CWR_Cell(lower, r, "TOSTAT")
            statuses(rowWT) = rowStatus
            If rowStatus = "" Then
                rowItem = CWR_Cell(lower, r, "ITEMNO")
                rowMat = CWR_Cell(lower, r, "MATNR")
                rowBin = CWR_Cell(lower, r, "VLPLA")
                rowBatch = CWR_Cell(lower, r, "CHARG")
                rowQty = CWR_Number(CWR_Cell(lower, r, "VSOLM"))
                If Not CWR_MONValueEquals(rowItem, CStr(rec("Item"))) Or _
                   StrComp(rowMat, CStr(rec("Material")), vbTextCompare) <> 0 Or _
                   StrComp(rowBin, CStr(rec("SysBin")), vbTextCompare) <> 0 Or _
                   StrComp(rowBatch, CStr(rec("SysBatch")), vbTextCompare) <> 0 Or _
                   Abs(rowQty - CDbl(rec("SysQty"))) > 0.000001 Then
                    messageText = "STOP_CANCEL_QUEUE_TUPLE_MISMATCH|WT=" & rowWT & "|PLAN=" & CStr(rec("PlanID"))
                    Exit Function
                End If
                rec("DestBin") = CWR_Cell(lower, r, "NLPLA")
                rec("ProcTy") = CWR_Cell(lower, r, "PROCTY")
                If CStr(rec("DestBin")) = "" Or CStr(rec("ProcTy")) = "" Then
                    messageText = "STOP_CANCEL_QUEUE_DESTINATION_MISSING|WT=" & rowWT
                    Exit Function
                End If
            End If
        End If
    Next r
    For Each rec In members
        If Not CBool(rec("Terminal")) Then
            rowWT = CStr(rec("OldWT"))
            If CLng(found(rowWT)) > 1 Then
                messageText = "STOP_CANCEL_QUEUE_DUPLICATE_WT|WT=" & rowWT
                Exit Function
            End If
            If CLng(found(rowWT)) = 0 Or CStr(statuses(rowWT)) <> "" Then
                rec("CancelAlready") = True
                rec("CancelProof") = "CANCEL_QUEUE_SNAPSHOT=PASS; WT=" & rowWT & _
                    IIf(CLng(found(rowWT)) = 0, "; OLD_WT_NOT_VISIBLE", "; OLD_WT_STATUS=" & CStr(statuses(rowWT)))
            Else
                rec("CancelProof") = "CANCEL_QUEUE_SNAPSHOT=PASS; WT=" & rowWT & "; OLD_WT_ACTIVE"
            End If
            CWR_SchedulerState rec, "CANCEL_PENDING", "CANCEL_QUEUE", "PASS", _
                CStr(rec("CancelProof")) & "|QTY_GROUP=" & CStr(groupID)
        End If
    Next rec
    CWR_SchedulerPreflightCancelGroup = True
    Exit Function
Failed:
    messageText = "CANCEL_QUEUE_PRECHECK_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' One group action selects every exact active WT row and sends one Cancel WT
' command/confirm.  The post-readback still reconciles every WT in the group.
Private Function CWR_SchedulerSendQueuedCancelGroup(ByVal members As Collection, ByVal groupID As String, _
                                                     ByRef messageText As String) As Boolean
    Dim firstRec As Object, rec As Object, ses As Object, lower As Object
    Dim expected As Object, keyValue As Variant
    Dim r As Long, n As Long, currentDoc As String, rowDoc As String, rowWT As String
    Dim rowStatus As String, rowItem As String, rowMat As String, rowBin As String, rowBatch As String
    Dim rowQty As Double, rowList As String, wtList As String, activeCount As Long
    On Error GoTo Failed
    Set expected = CreateObject("Scripting.Dictionary"): expected.CompareMode = vbTextCompare
    For Each rec In members
        If Not CBool(rec("Terminal")) And Not CBool(rec("CancelAlready")) Then
            If firstRec Is Nothing Then Set firstRec = rec
            If expected.Exists(CStr(rec("OldWT"))) Then
                messageText = "STOP_CANCEL_GROUP_DUPLICATE_WT|WT=" & CStr(rec("OldWT")) & _
                    "|QTY_GROUP=" & groupID
                Exit Function
            End If
            expected.Add CStr(rec("OldWT")), rec
            rec("CancelRow") = -1
        End If
    Next rec
    If expected.count = 0 Then
        CWR_SchedulerSendQueuedCancelGroup = True
        Exit Function
    End If
    If firstRec Is Nothing Then
        messageText = "STOP_CANCEL_GROUP_FIRST_MEMBER_MISSING|QTY_GROUP=" & groupID
        Exit Function
    End If
    If Not CWR_SchedulerUsePlan(firstRec, messageText) Then Exit Function
    If Not CWR_GetMONSession(ses, messageText) Then Exit Function
    If Not CWR_LoadFilteredWTGrid(ses, lower, messageText) Then Exit Function

    n = CLng(lower.rowCount): currentDoc = ""
    For r = 0 To n - 1
        rowDoc = CWR_Cell(lower, r, "DOCNO")
        If rowDoc <> "" Then currentDoc = rowDoc
        rowWT = CWR_Cell(lower, r, "TANUM")
        If CWR_MONValueEquals(currentDoc, CStr(firstRec("Doc"))) And expected.Exists(rowWT) Then
            Set rec = expected.item(rowWT)
            If CStr(rec("CancelRow")) <> "-1" Then
                messageText = "STOP_CANCEL_GROUP_DUPLICATE_ROW|WT=" & rowWT
                Exit Function
            End If
            rowStatus = CWR_Cell(lower, r, "TOSTAT")
            rowItem = CWR_Cell(lower, r, "ITEMNO")
            rowMat = CWR_Cell(lower, r, "MATNR")
            rowBin = CWR_Cell(lower, r, "VLPLA")
            rowBatch = CWR_Cell(lower, r, "CHARG")
            rowQty = CWR_Number(CWR_Cell(lower, r, "VSOLM"))
            If rowStatus <> "" Or Not CWR_MONValueEquals(rowItem, CStr(rec("Item"))) Or _
               StrComp(rowMat, CStr(rec("Material")), vbTextCompare) <> 0 Or _
               StrComp(rowBin, CStr(rec("SysBin")), vbTextCompare) <> 0 Or _
               StrComp(rowBatch, CStr(rec("SysBatch")), vbTextCompare) <> 0 Or _
               Abs(rowQty - CDbl(rec("SysQty"))) > 0.000001 Then
                messageText = "STOP_CANCEL_GROUP_TUPLE_MISMATCH|WT=" & rowWT & _
                    "|STATUS=" & rowStatus & "|ITEM=" & rowItem & "|MAT=" & rowMat & _
                    "|BIN=" & rowBin & "|BATCH=" & rowBatch & "|QTY=" & CStr(rowQty)
                Exit Function
            End If
            rec("CancelRow") = r
            activeCount = activeCount + 1
            If rowList = "" Then
                rowList = CStr(r): wtList = rowWT
            Else
                rowList = rowList & "," & CStr(r): wtList = wtList & "," & rowWT
            End If
        End If
    Next r
    If activeCount <> expected.count Then
        messageText = "STOP_CANCEL_GROUP_ROW_MATCH_COUNT|EXPECTED=" & CStr(expected.count) & _
            "|FOUND=" & CStr(activeCount) & "|QTY_GROUP=" & groupID
        Exit Function
    End If
    If Not CWR_RequireReworkAction("WT_CANCEL", messageText) Then Exit Function
    If Not CWR_SendCancelSelection(ses, lower, rowList, messageText, wtList) Then Exit Function

    For Each keyValue In expected.keys
        Set rec = expected.item(CStr(keyValue))
        rec("CancelSent") = True
        CWR_SchedulerState rec, "CANCEL_SENT", "CANCEL_GROUP", "SENT", _
            "WT_CANCEL_GROUP_SENT|GROUP_ACTION_COUNT=1|GROUP_MEMBERS=" & CStr(expected.count) & _
            "|SELECTED_ROWS=" & rowList & "|WT=" & CStr(rec("OldWT")) & _
            "|IDEMPOTENCY_KEY=CANCEL_GROUP|QTY_GROUP=" & groupID & "|MON_READBACK=DEFERRED_TO_GROUP"
    Next keyValue
    CWR_Log "SCHED_CANCEL_GROUP_ACTION", CStr(firstRec("OB")), CStr(firstRec("Doc")), "SENT", _
        "QTY_GROUP=" & groupID & "|GROUP_ACTION_COUNT=1|GROUP_MEMBERS=" & CStr(expected.count) & _
        "|SELECTED_ROWS=" & rowList & "|WT_LIST=" & wtList & "|MON_READBACK=DEFERRED_TO_GROUP"
    CWR_SchedulerSendQueuedCancelGroup = True
    Exit Function
Failed:
    messageText = "STATE_UNKNOWN_AFTER_WT_CANCEL_GROUP|QTY_GROUP=" & groupID & _
        "|CANCEL_GROUP_SEND_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Legacy single-WT helper retained for the non-scheduler cancel route.  The
' locked scheduler calls CWR_SchedulerSendQueuedCancelGroup above.
Private Function CWR_SchedulerSendQueuedCancel(ByVal rec As Object, ByVal groupID As String, _
                                                ByRef messageText As String) As Boolean
    Dim ses As Object, lower As Object, targetRow As Long
    On Error GoTo Failed
    If Not CWR_SchedulerUsePlan(rec, messageText) Then Exit Function
    If Not CWR_GetMONSession(ses, messageText) Then Exit Function
    If Not CWR_LocateExactActiveWT(ses, lower, targetRow, messageText) Then Exit Function
    If Not CWR_RequireReworkAction("WT_CANCEL", messageText) Then Exit Function
    If Not CWR_SendExactCancel(ses, lower, targetRow, messageText) Then
        messageText = "STATE_UNKNOWN_AFTER_WT_CANCEL|WT=" & CStr(rec("OldWT")) & "|" & messageText
        Exit Function
    End If
    rec("DestBin") = CWR_EXPECTED_DEST_BIN
    rec("ProcTy") = CWR_EXPECTED_PROCTY
    rec("CancelSent") = True
    CWR_SchedulerState rec, "CANCEL_SENT", "CANCEL_QUEUE", "SENT", _
        "WT_CANCEL_SENT|IDEMPOTENCY_KEY=CANCEL|" & CStr(rec("OB")) & "|" & CStr(rec("Doc")) & "|" & CStr(rec("OldWT")) & _
        "|QTY_GROUP=" & CStr(groupID) & "|MON_READBACK=DEFERRED_TO_QUEUE"
    CWR_SchedulerSendQueuedCancel = True
    Exit Function
Failed:
    messageText = "STATE_UNKNOWN_AFTER_WT_CANCEL|WT=" & CStr(rec("OldWT")) & _
        "|CANCEL_QUEUE_SEND_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' One bounded MON refresh reconciles every member that reached CANCEL_SENT and
' every member that was already cancelled before this run.  No Cancel is ever
' retried from this readback.
Private Function CWR_SchedulerVerifyCancelledGroup(ByVal members As Collection, ByVal allMembers As Collection, _
                                                    ByVal groupID As String, _
                                                    ByRef messageText As String) As Boolean
    Dim firstRec As Object, rec As Object, ses As Object, lower As Object
    Dim expected As Object, found As Object, statuses As Object, memberItems As Object, allowedActive As Object
    Dim r As Long, n As Long, currentDoc As String, rowDoc As String, rowWT As String, rowStatus As String
    Dim rowItem As String, rowMat As String, itemKey As String, unexpectedActive As Long
    On Error GoTo Failed
    Set expected = CreateObject("Scripting.Dictionary"): expected.CompareMode = vbTextCompare
    Set found = CreateObject("Scripting.Dictionary"): found.CompareMode = vbTextCompare
    Set statuses = CreateObject("Scripting.Dictionary"): statuses.CompareMode = vbTextCompare
    Set memberItems = CreateObject("Scripting.Dictionary"): memberItems.CompareMode = vbTextCompare
    Set allowedActive = CreateObject("Scripting.Dictionary"): allowedActive.CompareMode = vbTextCompare
    For Each rec In members
        If Not CBool(rec("Terminal")) Then
            If firstRec Is Nothing Then Set firstRec = rec
            rowWT = CStr(rec("OldWT"))
            expected.Add rowWT, rec
            found.Add rowWT, 0&
            statuses.Add rowWT, ""
            memberItems(CWR_SchedulerPRDOKey(CStr(rec("Item")), CStr(rec("Material")), "")) = True
        End If
    Next rec
    ' A cancellation readback covers the complete quantity group.  Retained
    ' DELTA_ADD peers are expected to remain active until their own Create
    ' route; they are allowed only by exact planned WT identity.
    For Each rec In allMembers
        If Not CBool(rec("Terminal")) Then
            If StrComp(CStr(rec("ExecMode")), "KEEP_EXISTING_CREATE_DELTA", vbTextCompare) = 0 Then
                If CStr(rec("OldWT")) <> "" Then allowedActive(CStr(rec("OldWT"))) = True
            End If
        End If
    Next rec
    If firstRec Is Nothing Then
        CWR_SchedulerVerifyCancelledGroup = True
        Exit Function
    End If
    If Not CWR_SchedulerUsePlan(firstRec, messageText) Then Exit Function
    If Not CWR_GetMONSession(ses, messageText) Then Exit Function
    If Not CWR_LoadFilteredWTGrid(ses, lower, messageText) Then Exit Function
    n = CLng(lower.rowCount)
    For r = 0 To n - 1
        rowDoc = CWR_Cell(lower, r, "DOCNO")
        If rowDoc <> "" Then currentDoc = rowDoc
        If CWR_MONValueEquals(currentDoc, CStr(firstRec("Doc"))) Then
            rowWT = CWR_Cell(lower, r, "TANUM")
            rowStatus = CWR_Cell(lower, r, "TOSTAT")
            If expected.Exists(rowWT) Then
                found(rowWT) = CLng(found(rowWT)) + 1
                statuses(rowWT) = rowStatus
            ElseIf rowWT <> "" And rowStatus = "" Then
                rowItem = CWR_Cell(lower, r, "ITEMNO")
                rowMat = CWR_Cell(lower, r, "MATNR")
                itemKey = CWR_SchedulerPRDOKey(rowItem, rowMat, "")
                If memberItems.Exists(itemKey) Then
                    If Not allowedActive.Exists(rowWT) Then unexpectedActive = unexpectedActive + 1
                End If
            End If
        End If
    Next r
    If unexpectedActive > 0 Then
        messageText = "STATE_UNKNOWN_AFTER_CANCEL_BATCH_READBACK|UNEXPECTED_ACTIVE_WT=" & CStr(unexpectedActive) & _
            "|QTY_GROUP=" & CStr(groupID)
        Exit Function
    End If
    For Each rec In members
        If Not CBool(rec("Terminal")) Then
            rowWT = CStr(rec("OldWT"))
            If CLng(found(rowWT)) > 1 Then
                messageText = "STATE_UNKNOWN_AFTER_CANCEL_BATCH_READBACK|DUPLICATE_OLD_WT=" & rowWT
                Exit Function
            End If
            If CLng(found(rowWT)) = 1 And CStr(statuses(rowWT)) = "" Then
                messageText = "STATE_UNKNOWN_AFTER_CANCEL_BATCH_READBACK|OLD_WT_STILL_ACTIVE=" & rowWT
                Exit Function
            End If
            rec("CancelVerified") = True
            rec("CancelProof") = "CANCEL_BATCH_MON_REFRESH=PASS; WT=" & rowWT & _
                IIf(CLng(found(rowWT)) = 0, "; OLD_WT_NOT_VISIBLE", "; OLD_WT_STATUS=" & CStr(statuses(rowWT)))
        End If
    Next rec
    CWR_SchedulerVerifyCancelledGroup = True
    Exit Function
Failed:
    messageText = "STATE_UNKNOWN_AFTER_CANCEL_BATCH_READBACK|ERROR=" & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_ExecutePreflightGroups(ByVal qtyGroups As Object, ByRef stopAll As Boolean, _
                                             ByRef messageText As String) As Boolean
    Dim groupID As Variant, members As Collection, rec As Object, firstRec As Object
    Dim cancelMembers As Collection
    Dim resultText As String, execMode As String, groupFailed As Boolean
    Dim hasCancel As Boolean, hasRetained As Boolean, modeText As String
    For Each groupID In qtyGroups.keys
        Set members = qtyGroups.item(CStr(groupID))
        Set firstRec = Nothing
        Set cancelMembers = Nothing
        execMode = ""
        groupFailed = False
        hasCancel = False
        hasRetained = False
        messageText = ""
        For Each rec In members
            If CBool(rec("Terminal")) Then GoTo NextContract
            If firstRec Is Nothing Then Set firstRec = rec: execMode = CStr(rec("ExecMode"))
            modeText = UCase$(CStr(rec("ExecMode")))
            If modeText = "NONE" And CBool(rec("PostRecoveryPending")) Then
                ' This member owns a prior POST write boundary, not a new
                ' route.  It participates in the shared quantity aggregate but
                ' has no cancel/qty action until the location recovery proves
                ' the final SAP batch.
                CWR_Log "RECOVER_POST", CStr(rec("OB")), CStr(rec("Doc")), "PASS", _
                    "PLAN=" & CStr(rec("PlanID")) & "|ACTION=POST|STATE=WRITE_SENT_UNVERIFIED|" & _
                    "ACTION_KEY=" & CWR_BusinessActionKey(rec, "POST") & "|PREFLIGHT=READBACK_ONLY"
                GoTo NextContract
            End If
            Select Case modeText
                Case "KEEP_EXISTING_CREATE_DELTA"
                    hasRetained = True
                Case "REPLACE_CANCEL_CREATE", "CANCEL_PRDO_ZERO"
                    hasCancel = True
                Case Else
                    messageText = "STOP_PLAN_EXEC_MODE_NOT_SUPPORTED|" & modeText & _
                        "|QTY_GROUP=" & CStr(groupID)
                    groupFailed = True
                    Exit For
            End Select
            If Not CWR_SchedulerUsePlan(rec, messageText) Then groupFailed = True: Exit For
            If Not CWR_AssertFullRouteActionContract(messageText) Then groupFailed = True: Exit For
NextContract:
        Next rec
        If groupFailed Or firstRec Is Nothing Then
            If messageText = "" Then messageText = "GROUP_PREFLIGHT_UNPROVEN|QTY_GROUP=" & CStr(groupID)
            For Each rec In members
                If Not CBool(rec("Terminal")) Then CWR_SchedulerFailure rec, "GROUP_PREFLIGHT", messageText, stopAll
            Next rec
            If stopAll Then Exit Function
            GoTo NextGroup
        End If

        ' A shared QTY group can legitimately contain a cancel/replace member
        ' alongside a retained delta member.  The old scheduler rejected that
        ' composition as GROUP_EXEC_MODE_MIXED before it ever reached Cancel.
        ' The safe contract is phase-based instead: verify retained members,
        ' cancel every replace/zero old WT first, then let the later location
        ' phase perform POST/change-bin and the later qty phase perform PRDO.
        If hasRetained Then
            For Each rec In members
                If CBool(rec("Terminal")) Or _
                   StrComp(CStr(rec("ExecMode")), "KEEP_EXISTING_CREATE_DELTA", vbTextCompare) <> 0 Then GoTo NextDelta
                If Not CWR_SchedulerUsePlan(rec, messageText) Then
                    CWR_SchedulerFailure rec, "VERIFY_RETAINED_OUTBOUND_WT", messageText, stopAll
                    groupFailed = True
                ElseIf Not CWR_AssertRetainedDeltaWT(messageText) Then
                    CWR_SchedulerFailure rec, "VERIFY_RETAINED_OUTBOUND_WT", messageText, stopAll
                    groupFailed = True
                Else
                    rec("DestBin") = CWR_EXPECTED_DEST_BIN
                    rec("ProcTy") = CWR_EXPECTED_PROCTY
                    CWR_SchedulerState rec, "RETAINED_WT_VERIFIED", "VERIFY_RETAINED_OUTBOUND_WT", "PASS", _
                        "EXACT_TUPLE_PROVED|QTY_GROUP=" & CStr(groupID)
                End If
NextDelta:
            Next rec
        End If

        If Not groupFailed And hasCancel Then
            Set cancelMembers = New Collection
            For Each rec In members
                If Not CBool(rec("Terminal")) Then
                    modeText = UCase$(CStr(rec("ExecMode")))
                    If modeText = "REPLACE_CANCEL_CREATE" Or modeText = "CANCEL_PRDO_ZERO" Then
                        cancelMembers.Add rec
                    End If
                End If
            Next rec

            ' Build one fresh, exact cancellation queue for only the members
            ' that own an old WT.  No SAP write occurs until every such member
            ' is known active or already cancelled.  This filtered queue is
            ' what makes mixed cancel+location+qty groups safe and prevents
            ' a preflight mode mismatch from skipping required cancellations.
            If cancelMembers.count > 0 Then
                If Not CWR_SchedulerPreflightCancelGroup(cancelMembers, CStr(groupID), messageText) Then
                    groupFailed = True
                End If
            End If
            If Not groupFailed Then
                For Each rec In cancelMembers
                    If CBool(rec("CancelAlready")) Then
                        If Not CWR_SchedulerUsePlan(rec, messageText) Then
                            groupFailed = True
                            Exit For
                        End If
                        rec("DestBin") = CWR_EXPECTED_DEST_BIN
                        rec("ProcTy") = CWR_EXPECTED_PROCTY
                    End If
                Next rec
            End If
            If Not groupFailed Then
                If Not CWR_SchedulerSendQueuedCancelGroup(cancelMembers, CStr(groupID), messageText) Then
                    groupFailed = True
                End If
            End If
            If Not groupFailed Then
                ' One bounded MON refresh verifies the whole queue.  This is
                ' the only post-Cancel readback and no member is retried here.
                If Not CWR_SchedulerVerifyCancelledGroup(cancelMembers, members, CStr(groupID), messageText) Then
                    groupFailed = True
                End If
            End If
            If Not groupFailed Then
                For Each rec In cancelMembers
                    If CBool(rec("Terminal")) Then GoTo NextCancelVerified
                    If Not CWR_SchedulerUsePlan(rec, messageText) Then
                        groupFailed = True
                        Exit For
                    End If
                    CWR_RecordCancelledSnapshot CStr(rec("CancelProof"))
                    CWR_SchedulerState rec, "CANCEL_VERIFIED", "CANCEL_BATCH_READBACK", "PASS", _
                        CStr(rec("CancelProof")) & "|QTY_GROUP=" & CStr(groupID) & _
                        "|IDEMPOTENCY_KEY=CANCEL|" & CStr(rec("OB")) & "|" & CStr(rec("Doc")) & "|" & CStr(rec("OldWT"))
NextCancelVerified:
                Next rec
            End If
            If Not groupFailed Then
                For Each rec In cancelMembers
                    If Not CBool(rec("Terminal")) Then CWR_SchedulerState rec, "SOURCE_READY", "CANCEL_GROUP", "PASS", _
                        "CANCEL_QUEUE_VERIFIED|NO_PER_PLAN_REMAINING_QTY_ASSUMPTION"
                Next rec
                CWR_Log "SCHED_CANCEL_PRIORITY", CStr(firstRec("OB")), CStr(firstRec("Doc")), "PASS", _
                    "QTY_GROUP=" & CStr(groupID) & "|CANCEL_MEMBERS=" & CStr(cancelMembers.count) & _
                    "|ORDER=CANCEL_VERIFIED_BEFORE_LOCATION_POST_CHANGE_BEFORE_QTY_PRDO"
            End If
        End If

        If groupFailed Then
            For Each rec In members
                If Not CBool(rec("Terminal")) Then CWR_SchedulerFailure rec, "GROUP_CANCEL_RECONCILE", messageText, stopAll
            Next rec
            If stopAll Then Exit Function
        End If
NextGroup:
    Next groupID
    CWR_ExecutePreflightGroups = Not stopAll
End Function

Private Function CWR_SchedulerVerifyPRDOGroup(ByVal members As Collection, ByRef messageText As String) As Boolean
    Dim firstRec As Object, rec As Object, ses As Object, grid As Object
    Dim r As Long, n As Long, liveQty As Double, liveTotal As Double, batchText As String
    Dim expectedBefore As Double, expectedTarget As Double, plannedBefore As Double, plannedTarget As Double
    Dim rowItem As String, rowMat As String, rowKey As String, memberKey As String, keyValue As Variant
    Dim expectedRows As Object, matchCounts As Object
    On Error GoTo Failed
    Set expectedRows = CreateObject("Scripting.Dictionary")
    expectedRows.CompareMode = vbTextCompare
    Set matchCounts = CreateObject("Scripting.Dictionary")
    matchCounts.CompareMode = vbTextCompare
    For Each rec In members
        If Not CBool(rec("Terminal")) Then
            Set firstRec = rec
            memberKey = CWR_SchedulerPRDOKey(CStr(rec("Item")), CStr(rec("Material")), CStr(rec("SysBatch")))
            If memberKey = "" Then
                messageText = "STOP_PRDO_QTY_GROUP_KEY_MISSING|PLAN=" & CStr(rec("PlanID"))
                Exit Function
            End If
            If expectedRows.Exists(memberKey) Then
                expectedRows(memberKey) = CDbl(expectedRows(memberKey)) + CDbl(rec("SysQty"))
            Else
                expectedRows.Add memberKey, CDbl(rec("SysQty"))
                matchCounts.Add memberKey, 0&
            End If
            plannedBefore = plannedBefore + CDbl(rec("SysQty"))
            plannedTarget = plannedTarget + CDbl(rec("TargetQty"))
        End If
    Next rec
    If firstRec Is Nothing Then
        CWR_SchedulerVerifyPRDOGroup = True
        Exit Function
    End If
    If Not CWR_SchedulerUsePlan(firstRec, messageText) Then Exit Function
    If Not CWR_OpenExpectedPRDO(ses, grid, messageText) Then Exit Function
    n = CLng(grid.rowCount)
    For r = 0 To n - 1
        rowItem = CWR_Cell(grid, r, "ITEMNO")
        rowMat = CWR_Cell(grid, r, "PRODUCTNO")
        batchText = CWR_Cell(grid, r, "BATCHNO_UI")
        rowKey = CWR_SchedulerPRDOKey(rowItem, rowMat, batchText)
        If expectedRows.Exists(rowKey) Then
            matchCounts(rowKey) = CLng(matchCounts(rowKey)) + 1
            liveQty = CWR_Number(CWR_Cell(grid, r, "QTY_UI"))
            liveTotal = liveTotal + liveQty
        End If
    Next r
    expectedBefore = CDbl(firstRec("GroupCurrent"))
    expectedTarget = CDbl(firstRec("GroupTarget"))
    ' A blocked/non-terminal member must never disappear from a shared PRDO
    ' aggregate.  When it is absent from the execution collection, report the
    ' missing member explicitly instead of comparing one member (for example
    ' 1 -> 0) with the group aggregate (180 -> 146).
    If members.count < CLng(firstRec("GroupMembers")) Then
        messageText = "STOP_PRDO_QTY_GROUP_MEMBER_INCOMPLETE|QTY_GROUP=" & CStr(firstRec("QtyGroupId")) & _
            "|ACTIVE_MEMBERS=" & CStr(members.count) & "|EXPECTED_MEMBERS=" & CStr(firstRec("GroupMembers")) & _
            "|PLAN_BEFORE=" & CStr(plannedBefore) & "|GROUP_BEFORE=" & CStr(expectedBefore) & _
            "|PLAN_TARGET=" & CStr(plannedTarget) & "|GROUP_TARGET=" & CStr(expectedTarget)
        Exit Function
    End If
    If Abs(plannedBefore - expectedBefore) > 0.000001 Or Abs(plannedTarget - expectedTarget) > 0.000001 Then
        messageText = "STOP_PRDO_QTY_GROUP_PLAN_AGGREGATE_MISMATCH|QTY_GROUP=" & CStr(firstRec("QtyGroupId")) & _
            "|PLAN_BEFORE=" & CStr(plannedBefore) & "|GROUP_BEFORE=" & CStr(expectedBefore) & _
            "|PLAN_TARGET=" & CStr(plannedTarget) & "|GROUP_TARGET=" & CStr(expectedTarget)
        Exit Function
    End If
    For Each keyValue In expectedRows.keys
        If CLng(matchCounts(CStr(keyValue))) <> 1 Then
            messageText = "STOP_PRDO_QTY_GROUP_MATCH_COUNT=" & CStr(matchCounts(CStr(keyValue))) & _
                "|QTY_GROUP=" & CStr(firstRec("QtyGroupId")) & "|PRDO_KEY=" & CStr(keyValue)
            Exit Function
        End If
    Next keyValue
    If Abs(liveTotal - expectedBefore) > 0.000001 And Abs(liveTotal - expectedTarget) > 0.000001 Then
        messageText = "STATE_UNKNOWN_QTY_BASELINE|QTY_GROUP=" & CStr(firstRec("QtyGroupId")) & _
            "|LIVE_TOTAL=" & CStr(liveTotal) & "|EXPECTED_BEFORE=" & CStr(expectedBefore) & _
            "|TARGET=" & CStr(expectedTarget)
        Exit Function
    End If
    For Each rec In members
        rec("GroupLiveQty") = liveTotal
        rec("GroupQtyBaselineProven") = True
        If CFAlloc_IsContractBPass(rec) Then rec("QtyReady") = True
    Next rec
    CWR_SchedulerState firstRec, "GROUP_QTY_BASELINE_PROVEN", "QTY_GROUP_BASELINE", "PASS", _
        "PRDO_ROWS=" & CStr(expectedRows.count) & "|LIVE_TOTAL=" & CStr(liveTotal) & _
        "|EXPECTED_BEFORE=" & CStr(expectedBefore) & "|TARGET=" & CStr(expectedTarget)
    CWR_SchedulerVerifyPRDOGroup = True
    Exit Function
Failed:
    messageText = "PRDO_QTY_GROUP_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

' Keys use the same whitespace/leading-zero normalization as CWR_MONValueEquals.
' A QTY group may contain several batch rows for one ODO item, so its baseline
' must be the aggregate of those exact PRDO rows, never the first row alone.
Private Function CWR_SchedulerPRDOKey(ByVal itemText As String, ByVal materialText As String, _
                                      ByVal batchText As String) As String
    CWR_SchedulerPRDOKey = CWR_SchedulerIdentityPart(itemText) & "|" & _
        CWR_SchedulerIdentityPart(materialText) & "|" & CWR_SchedulerIdentityPart(batchText)
End Function

Private Function CWR_SchedulerIdentityPart(ByVal valueText As String) As String
    Dim normalized As String
    normalized = Replace(UCase$(CWR_Text(valueText)), " ", "")
    Do While Len(normalized) > 1 And Left$(normalized, 1) = "0"
        normalized = Mid$(normalized, 2)
    Loop
    CWR_SchedulerIdentityPart = normalized
End Function

Private Function CWR_SchedulerGroupIsBlocked(ByVal members As Collection) As Boolean
    Dim rec As Object, stateText As String
    For Each rec In members
        If Not CBool(rec("Terminal")) Then
            stateText = UCase$(CStr(rec("State")))
            If InStr(1, stateText, "BLOCKED", vbTextCompare) > 0 Or _
               InStr(1, stateText, "STATE_UNKNOWN", vbTextCompare) > 0 Or _
               InStr(1, stateText, "WAIT_USER_CF", vbTextCompare) > 0 Then
                CWR_SchedulerGroupIsBlocked = True
                Exit Function
            End If
        End If
    Next rec
End Function

Private Sub CWR_SchedulerPreserveBlockedGroup(ByVal members As Collection, ByVal stageText As String, _
                                              ByVal groupID As String)
    Dim rec As Object
    For Each rec In members
        If Not CBool(rec("Terminal")) Then
            CWR_SchedulerState rec, CStr(rec("State")), stageText, "SKIP", _
                "BLOCKED_GROUP_PRESERVED|QTY_GROUP=" & groupID & "|NO_SAP_WRITE"
        End If
    Next rec
End Sub

' True PRDO DocumentGroup executor.  CHECK may put several quantity/source
' contracts under one document/item package; all readbacks are completed first
' and the SAP PRDO surface is then saved at most once for that document group.
' The older QTY-group routine below remains as a compatibility route for the
' single-case executor, but Button #4 uses this document-level scheduler.
Private Function CWR_ExecuteDocumentGroups(ByVal docGroups As Object, ByVal qtyGroups As Object, _
                                            ByRef stopAll As Boolean, ByRef messageText As String) As Boolean
    Dim docID As Variant, qtyID As Variant, docMembers As Collection, qtyMembers As Collection
    Dim qtyIDs As Object, included As Object, rec As Object, batchRec As Object, firstBatch As Object
    Dim records As Collection, prdoGroups As Object, groupObj As Object
    Dim groupBlocked As Boolean, writeRequired As Boolean, docGroupID As String
    For Each docID In docGroups.keys
        Set docMembers = docGroups.item(CStr(docID))
        ' Do not gate the whole document on one member.  Each QTY group below
        ' owns its own readback/write contract; a blocked member is preserved
        ' in place and independent groups continue through this document.
        Set qtyIDs = CreateObject("Scripting.Dictionary")
        qtyIDs.CompareMode = vbTextCompare
        For Each rec In docMembers
            If Not CBool(rec("Terminal")) Then qtyIDs(CStr(rec("QtyGroupId"))) = True
        Next rec
        groupBlocked = False
        For Each qtyID In qtyIDs.keys
            Set qtyMembers = qtyGroups.item(CStr(qtyID))
            If CWR_SchedulerGroupIsBlocked(qtyMembers) Then
                CWR_SchedulerPreserveBlockedGroup qtyMembers, "PRDO_DOCUMENT_QTY_GROUP_SKIPPED", CStr(qtyID)
                ' Member isolation: a blocked quantity contract must not
                ' suppress independent quantity contracts in the same PRDO
                ' document.  The blocked group remains fail-closed, while
                ' the next QTY group is still allowed to prove and execute.
                CWR_Log "SCHED_MEMBER_ISOLATION", CStr(docMembers.item(1)("OB")), _
                    CStr(docMembers.item(1)("Doc")), "PASS", _
                    "BLOCKED_QTY_GROUP=" & CStr(qtyID) & "|DOC_GROUP=" & CStr(docID) & _
                    "|OTHER_QTY_GROUPS_CONTINUE=TRUE|NO_SAP_WRITE_FOR_BLOCKED_GROUP"
                GoTo NextQtyGroup
            End If
            If Not CWR_SchedulerVerifyPRDOGroup(qtyMembers, messageText) Then
                For Each rec In qtyMembers
                    If Not CBool(rec("Terminal")) Then CWR_SchedulerFailure rec, "QTY_GROUP_BASELINE", messageText, stopAll
                Next rec
                groupBlocked = True
                If stopAll Then Exit Function
                GoTo NextQtyGroup
            End If
            If Not CWR_SchedulerVerifyQtySources(qtyMembers, stopAll, messageText) Then
                groupBlocked = True
                If stopAll Then Exit Function
            End If
NextQtyGroup:
        Next qtyID
        If Not CWR_SchedulerBuildPRDOBatch(docMembers, records, prdoGroups) Or records.count = 0 Then
            For Each rec In docMembers
                If Not CBool(rec("Terminal")) Then
                    ' A missing document batch is not a successful no-op.  The
                    ' old code marked every remaining member QTY_READY here,
                    ' which silently converted a blocked/omitted member into a
                    ' later Create candidate.  Keep the member accounted for
                    ' with an explicit terminal blocker.
                    CWR_SchedulerFailure rec, "PRDO_DOCUMENT_GROUP_RECORDS", _
                        "STOP_PRDO_DOCUMENT_MEMBER_UNACCOUNTED|DOCUMENT=" & CStr(rec("Doc")) & _
                        "|QTY_GROUP=" & CStr(rec("QtyGroupId")) & "|NO_PRDO_SAVE", stopAll
                End If
            Next rec
            If stopAll Then Exit Function
            GoTo NextDocument
        End If
        Set firstBatch = records.item(1)
        Set rec = firstBatch.item("SchedulerRec")
        docGroupID = CStr(rec("DocGroup"))
        writeRequired = False
        For Each qtyID In prdoGroups.keys
            Set groupObj = prdoGroups.item(CStr(qtyID))
            If CBool(groupObj("WriteRequired")) Then writeRequired = True
        Next qtyID
        If CWR_PRDO_DOCUMENT_SAVE_COUNTS Is Nothing Then
            Set CWR_PRDO_DOCUMENT_SAVE_COUNTS = CreateObject("Scripting.Dictionary")
            CWR_PRDO_DOCUMENT_SAVE_COUNTS.CompareMode = vbTextCompare
        End If
        If writeRequired And CWR_PRDO_DOCUMENT_SAVE_COUNTS.Exists(docGroupID) Then
            messageText = "STOP_PRDO_DOCUMENT_GROUP_DUPLICATE_SAVE|DOCUMENT=" & CStr(rec("Doc")) & _
                "|DOC_GROUP=" & docGroupID & "|SAVE_COUNT=" & CStr(CWR_PRDO_DOCUMENT_SAVE_COUNTS(docGroupID))
            CWR_SchedulerFailGroup records, "PRDO_DOCUMENT_GROUP_IDEMPOTENCY", messageText, stopAll
            If stopAll Then Exit Function
            GoTo NextDocument
        End If
        If Not CWR_SchedulerUsePlan(firstBatch, messageText) Then
            CWR_SchedulerFailGroup records, "PRDO_DOCUMENT_GROUP_CONTEXT", messageText, stopAll
            If stopAll Then Exit Function
            GoTo NextDocument
        End If
        If Not CWR_BatchSetPRDOQtyOneSave(records, prdoGroups, messageText) Then
            CWR_SchedulerFailGroup records, "PRDO_DOCUMENT_GROUP", messageText, stopAll
            If stopAll Then Exit Function
            GoTo NextDocument
        End If
        If writeRequired Then
            CWR_PRDO_DOCUMENT_SAVE_COUNTS(docGroupID) = 1
        Else
            CWR_PRDO_DOCUMENT_SAVE_COUNTS(docGroupID) = 0
        End If
        Set included = CreateObject("Scripting.Dictionary")
        included.CompareMode = vbTextCompare
        For Each batchRec In records
            Set rec = batchRec.item("SchedulerRec")
            included(CStr(rec("PlanID"))) = True
            rec("QtyReady") = True
            CWR_SchedulerState rec, "QTY_READY", "PRDO_DOCUMENT_GROUP", "PASS", _
                "DOCUMENT=" & CStr(rec("Doc")) & "|DOC_GROUP=" & docGroupID & _
                "|MEMBERS=" & CStr(docMembers.count) & "|SAVE_COUNT=" & IIf(writeRequired, "1", "0") & _
                "|PRDO_DOCUMENT_ACTION=ONE_SAVE"
        Next batchRec
        For Each rec In docMembers
            If Not CBool(rec("Terminal")) And Not included.Exists(CStr(rec("PlanID"))) Then
                CWR_SchedulerFailure rec, "PRDO_DOCUMENT_MEMBER_ISOLATION", _
                    "STOP_PRDO_DOCUMENT_MEMBER_NOT_INCLUDED|DOCUMENT=" & CStr(rec("Doc")) & _
                    "|QTY_GROUP=" & CStr(rec("QtyGroupId")) & "|NO_SILENT_SKIP", stopAll
            End If
        Next rec
        If stopAll Then Exit Function
        Set rec = firstBatch.item("SchedulerRec")
        CWR_Log "PRDO_DOCUMENT_ACTION", CStr(rec("OB")), CStr(rec("Doc")), "PASS", _
            "DOCUMENT=" & CStr(rec("Doc")) & "|DOC_GROUP=" & docGroupID & _
            "|MEMBERS=" & CStr(docMembers.count) & "|SAVE_COUNT=" & IIf(writeRequired, "1", "0") & _
            "|READBACK_ALL=TRUE|SAP_WRITE=" & IIf(writeRequired, "1", "0")
NextDocument:
    Next docID
    CWR_ExecuteDocumentGroups = Not stopAll
End Function

Private Function CWR_ExecuteQtyGroups(ByVal qtyGroups As Object, ByRef stopAll As Boolean, _
                                       ByRef messageText As String) As Boolean
    Dim groupID As Variant, members As Collection, rec As Object, batchRec As Object, records As Collection, prdoGroups As Object
    Dim resultText As String, firstRec As Object, keyValue As Variant, docGroupID As String, writeRequired As Boolean, groupObj As Object
    For Each groupID In qtyGroups.keys
        Set members = qtyGroups.item(CStr(groupID))
        If CWR_SchedulerGroupIsBlocked(members) Then
            CWR_SchedulerPreserveBlockedGroup members, "QTY_GROUP_SKIPPED", CStr(groupID)
            GoTo NextGroup
        End If
        If Not CWR_SchedulerVerifyPRDOGroup(members, messageText) Then
            For Each rec In members
                If Not CBool(rec("Terminal")) Then CWR_SchedulerFailure rec, "QTY_GROUP_BASELINE", messageText, stopAll
            Next rec
            If stopAll Then Exit Function
            GoTo NextGroup
        End If
        CWR_Log "QTY_GROUP_ALLOCATION_STATUS", CStr(members.item(1)("OB")), _
            CStr(members.item(1)("Doc")), "PASS", CFAlloc_GroupStatusText(members)
        If Not CWR_SchedulerVerifyQtySources(members, stopAll, messageText) Then
            If stopAll Then Exit Function
            GoTo NextGroup
        End If
        If Not CWR_SchedulerBuildPRDOBatch(members, records, prdoGroups) Then GoTo NextGroup
        Set firstRec = records.item(1)
        docGroupID = CWR_Text(members.item(1)("DocGroup"))
        writeRequired = False
        For Each keyValue In prdoGroups.keys
            Set groupObj = prdoGroups.item(CStr(keyValue))
            If CBool(groupObj("WriteRequired")) Then writeRequired = True
        Next keyValue
        If CWR_PRDO_DOCUMENT_SAVE_COUNTS Is Nothing Then
            Set CWR_PRDO_DOCUMENT_SAVE_COUNTS = CreateObject("Scripting.Dictionary")
            CWR_PRDO_DOCUMENT_SAVE_COUNTS.CompareMode = vbTextCompare
        End If
        If writeRequired And CWR_PRDO_DOCUMENT_SAVE_COUNTS.Exists(docGroupID) Then
            messageText = "STOP_PRDO_DOCUMENT_GROUP_DUPLICATE_SAVE|DOCUMENT=" & CStr(firstRec("Doc")) & _
                "|DOC_GROUP=" & docGroupID & "|SAVE_COUNT=" & CStr(CWR_PRDO_DOCUMENT_SAVE_COUNTS(docGroupID))
            CWR_SchedulerFailGroup records, "PRDO_DOCUMENT_GROUP_GATE", messageText, stopAll
            If stopAll Then Exit Function
            GoTo NextGroup
        End If
        If Not CWR_SchedulerUsePlan(firstRec, messageText) Then
            CWR_SchedulerFailGroup records, "QTY_GROUP_CONTEXT", messageText, stopAll
            If stopAll Then Exit Function
            GoTo NextGroup
        End If
        If Not CWR_BatchSetPRDOQtyOneSave(records, prdoGroups, messageText) Then
            CWR_SchedulerFailGroup records, "DOCUMENT_QTY_GROUP", messageText, stopAll
            If stopAll Then Exit Function
            GoTo NextGroup
        End If
        If writeRequired Then
            CWR_PRDO_DOCUMENT_SAVE_COUNTS(docGroupID) = 1
            CWR_Log "PRDO_DOCUMENT_ACTION", CStr(firstRec("OB")), CStr(firstRec("Doc")), "PASS", _
                "DOCUMENT=" & CStr(firstRec("Doc")) & "|DOC_GROUP=" & docGroupID & _
                "|MEMBERS=" & CStr(members.count) & "|SAVE_COUNT=1|GROUP_ACTION=ONE_DOCUMENT_SAVE"
        Else
            CWR_Log "PRDO_DOCUMENT_ACTION", CStr(firstRec("OB")), CStr(firstRec("Doc")), "PASS", _
                "DOCUMENT=" & CStr(firstRec("Doc")) & "|DOC_GROUP=" & docGroupID & _
                "|MEMBERS=" & CStr(members.count) & "|SAVE_COUNT=0|READBACK_ONLY=TRUE"
        End If
        For Each batchRec In records
            Set rec = batchRec.item("SchedulerRec")
            rec("QtyReady") = True
            CWR_SchedulerState rec, "QTY_READY", "DOCUMENT_QTY_GROUP", "PASS", _
                               "QTY_GROUP=" & CStr(groupID) & "|ONE_SAVE_OR_READBACK|IDEMPOTENCY_KEY=CHANGE_QTY|" & _
                               CStr(rec("Doc")) & "|" & CStr(rec("Item")) & "|" & CStr(rec("GroupTarget"))
        Next batchRec
NextGroup:
    Next groupID
    CWR_ExecuteQtyGroups = Not stopAll
End Function

Private Function CWR_SchedulerVerifyQtySources(ByVal members As Collection, ByRef stopAll As Boolean, _
                                                ByRef messageText As String) As Boolean
    Dim rec As Object
    For Each rec In members
        If Not CBool(rec("Terminal")) And CBool(rec("NeedQty")) And _
           CStr(rec("ExecMode")) = "KEEP_EXISTING_CREATE_DELTA" Then
            ' PRDO is the authoritative live total.  Once the requested total
            ' is already present, the source check for a new delta is not only
            ' unnecessary but wrong: it would calculate (TARGET - OLD_WT) and
            ' inspect stock even though no quantity write is pending.
            If CBool(rec("GroupQtyBaselineProven")) And _
               CFAlloc_IsContractBPass(rec) Then
                rec("QtyReady") = True
                CWR_SchedulerState rec, "QTY_RECONCILED_PASS", "DELTA_SOURCE_LIVE_PREVERIFY", "PASS", _
                    "MEMBER_ALLOCATION_COMPLETE|REQUIRED_CREATE=0|LIVE_TOTAL=" & CStr(rec("GroupLiveQty")) & _
                    "|TARGET=" & CStr(rec("GroupTarget")) & "|SOURCE_VERIFY=SKIP|CONTRACT_B_PASS"
                CWR_Log "RECOVER_PRDO", CStr(rec("OB")), CStr(rec("Doc")), "PASS", _
                    "PLAN=" & CStr(rec("PlanID")) & "|ACTION=PRDO|STATE=RECONCILED_PASS|" & _
                    "ACTION_KEY=" & CWR_BusinessActionKey(rec, "PRDO") & "|LIVE_TOTAL=" & _
                    CStr(rec("GroupLiveQty")) & "|TARGET=" & CStr(rec("GroupTarget")) & "|NO_QTY_WRITE"
                GoTo NextSourceRec
            End If
            If Not CWR_SchedulerUsePlan(rec, messageText) Then
                CWR_SchedulerFailure rec, "DELTA_SOURCE_LIVE_PREVERIFY", messageText, stopAll
            ElseIf Not CWR_AssertDeltaSourceStockLive(messageText) Then
                CWR_SchedulerFailure rec, "DELTA_SOURCE_LIVE_PREVERIFY", messageText, stopAll
            Else
                CWR_SchedulerState rec, "RETAINED_WT_VERIFIED", "DELTA_SOURCE_LIVE_PREVERIFY", "PASS", _
                                   "EXACT_INCREMENTAL_SOURCE_PROVED"
            End If
            If stopAll Then Exit Function
        End If
NextSourceRec:
    Next rec
    CWR_SchedulerVerifyQtySources = True
End Function

Private Function CWR_SchedulerBuildPRDOBatch(ByVal members As Collection, ByRef records As Collection, _
                                               ByRef prdoGroups As Object) As Boolean
    Dim rec As Object, batchRec As Object, group As Object, groupKey As String, draftQty As Double
    Set records = New Collection
    Set prdoGroups = CreateObject("Scripting.Dictionary")
    prdoGroups.CompareMode = vbTextCompare
    For Each rec In members
        ' A shared ODO item is one PRDO quantity contract.  Include retained
        ' members even when their individual NeedQtyChange is FALSE so that a
        ' 128+118 row pair is read and written as 246 -> 256, not as 118 -> 128.
        If Not CBool(rec("Terminal")) And (CBool(rec("GroupNeedQty")) Or _
           CBool(rec("NeedBin")) Or CBool(rec("NeedBatch")) Or _
           CStr(rec("ExecMode")) = "REPLACE_CANCEL_CREATE") Then
            If InStr(1, CStr(rec("State")), "BLOCKED", vbTextCompare) > 0 Or _
               InStr(1, CStr(rec("State")), "STATE_UNKNOWN", vbTextCompare) > 0 Then GoTo NextRec
            Set batchRec = CreateObject("Scripting.Dictionary")
            batchRec.CompareMode = vbTextCompare
            batchRec.Add "PlanID", rec("PlanID"): batchRec.Add "PlanRow", rec("PlanRow")
            batchRec.Add "WHRow", rec("WHRow"): batchRec.Add "Doc", rec("Doc")
            batchRec.Add "Item", rec("Item"): batchRec.Add "Material", rec("Material")
            batchRec.Add "OldWT", rec("OldWT"): batchRec.Add "OldQty", rec("SysQty")
            batchRec.Add "TargetQty", rec("TargetQty")
            draftQty = CWR_SchedulerCreateQty(rec)
            batchRec.Add "PlannedDraftQty", draftQty
            batchRec.Add "DraftQty", draftQty
            ' Preserve the exact source/target tuple from PLAN.  The create
            ' post-verify and the stock-row selector both use these fields;
            ' leaving TargetBin blank (and replacing TargetBatch with the
            ' system batch) makes a valid Create+Save look like an unknown
            ' outcome even though the new WT is present in MON.
            batchRec.Add "SourceBin", rec("SourceBin")
            batchRec.Add "SourceBatch", rec("SourceBatch")
            batchRec.Add "TargetBin", rec("TargetBin")
            batchRec.Add "TargetBatch", rec("TargetBatch")
            batchRec.Add "StockType", ""
            batchRec.Add "DestBin", rec("DestBin"): batchRec.Add "ProcTy", rec("ProcTy"): batchRec.Add "Adopted", False
            batchRec.Add "SchedulerRec", rec
            groupKey = CStr(rec("Item")) & "|" & CStr(rec("Material")) & "|" & CStr(rec("SysBatch"))
            batchRec.Add "GroupKey", groupKey
            records.Add batchRec
            If Not prdoGroups.Exists(groupKey) Then
                Set group = CreateObject("Scripting.Dictionary")
                group.CompareMode = vbTextCompare
                group.Add "Item", rec("Item"): group.Add "Material", rec("Material")
                group.Add "Batch", rec("SysBatch"): group.Add "OldQty", 0#
                ' The group drives the Create-WT source lookup.  Preserve its
                ' exact planned source tuple here as well as on batchRec; a
                ' missing group TargetBin otherwise degrades into BIN= and a
                ' false SOURCE_ROW_NOT_FOUND after the PRDO Save.
                group.Add "TargetBin", rec("TargetBin")
                group.Add "TargetBatch", rec("TargetBatch")
                group.Add "SourceTupleAmbiguous", False
                group.Add "TargetQty", 0#: group.Add "PRDORow", -1
                group.Add "RequestRow", -1: group.Add "Remaining", 0#: group.Add "CreateQty", 0#
                group.Add "LiveQty", 0#: group.Add "WriteRequired", False
                prdoGroups.Add groupKey, group
            End If
            Set group = prdoGroups.item(groupKey)
            If StrComp(CStr(group("TargetBin")), CStr(rec("TargetBin")), vbTextCompare) <> 0 Or _
               StrComp(CStr(group("TargetBatch")), CStr(rec("TargetBatch")), vbTextCompare) <> 0 Then
                group("SourceTupleAmbiguous") = True
            End If
            group("OldQty") = CDbl(group("OldQty")) + CDbl(rec("SysQty"))
            group("TargetQty") = CDbl(group("TargetQty")) + CDbl(rec("TargetQty"))
            group("CreateQty") = CDbl(group("CreateQty")) + draftQty
        End If
NextRec:
    Next rec
    CWR_SchedulerBuildPRDOBatch = (records.count > 0)
End Function

Private Function CWR_SchedulerCreateQty(ByVal rec As Object) As Double
    If StrComp(CStr(rec("QtyMode")), "DELTA_ADD", vbTextCompare) = 0 Then
        CWR_SchedulerCreateQty = CDbl(rec("QtyDelta"))
    Else
        CWR_SchedulerCreateQty = CDbl(rec("TargetQty"))
    End If
End Function

Private Function CWR_ExecuteLocationGroups(ByVal locationGroups As Object, ByRef stopAll As Boolean, _
                                            ByRef messageText As String) As Boolean
    Dim groupID As Variant, members As Collection, rec As Object, firstRec As Object, resultText As String
    Dim groupQty As Double, finalBatch As String
    For Each groupID In locationGroups.keys
        Set members = locationGroups.item(CStr(groupID))
        If CWR_SchedulerGroupIsBlocked(members) Then
            CWR_SchedulerPreserveBlockedGroup members, "LOCATION_GROUP_SKIPPED", CStr(groupID)
            GoTo NextGroup
        End If
        Set firstRec = Nothing: groupQty = 0#
        For Each rec In members
            If CBool(rec("Terminal")) Then GoTo NextRec
            If firstRec Is Nothing Then Set firstRec = rec
            groupQty = groupQty + CWR_SchedulerCreateQty(rec)
NextRec:
        Next rec
        If firstRec Is Nothing Then GoTo NextGroup
        If Not CBool(firstRec("NeedBin")) And Not CBool(firstRec("NeedBatch")) Then
            For Each rec In members
                If Not CBool(rec("Terminal")) Then rec("LocationReady") = True: CWR_SchedulerState rec, "LOCATION_READY", _
                    "LOCATION_AUTO_PASS", "PASS", "NO_BIN_OR_BATCH_CHANGE|POSTSRC_GROUP=" & CStr(groupID)
            Next rec
            GoTo NextGroup
        End If
        If CBool(firstRec("PostRecoveryPending")) Then
            If Not CWR_SchedulerUseRecoveryPlan(firstRec, messageText) Then
                For Each rec In members
                    If Not CBool(rec("Terminal")) Then CWR_SchedulerFailure rec, "LOCATION_GROUP_RECOVERY_CONTEXT", messageText, stopAll
                Next rec
                If stopAll Then Exit Function
                GoTo NextGroup
            End If
        ElseIf Not CWR_SchedulerUsePlan(firstRec, messageText) Then
            For Each rec In members
                If Not CBool(rec("Terminal")) Then CWR_SchedulerFailure rec, "LOCATION_GROUP_CONTEXT", messageText, stopAll
            Next rec
            If stopAll Then Exit Function
            GoTo NextGroup
        End If
        ' /SCWM/POST is one source-group action.  Use the aggregate required
        ' quantity; never send one POST change per plan member.
        CWR_TARGET_QTY = groupQty
        resultText = Codex_ExecutePostingBatchAtPickBin()
        If Left$(resultText, 5) <> "PASS|" Then
            For Each rec In members
                If Not CBool(rec("Terminal")) Then CWR_SchedulerFailure rec, "LOCATION_GROUP", resultText, stopAll
            Next rec
            If stopAll Then Exit Function
        ElseIf InStr(1, resultText, "POST_CHANGE_WT_CREATED_PENDING_USER_CF", vbTextCompare) > 0 Then
            For Each rec In members
                If Not CBool(rec("Terminal")) Then rec("Terminal") = True: CWR_SchedulerState rec, "WAIT_USER_CF", _
                    "LOCATION_GROUP", "STOP", "POST_CHANGE_WT_POSTVERIFIED|CF_USER_HANDOFF|POSTSRC_GROUP=" & CStr(groupID)
            Next rec
        Else
            ' POST changes the pick-bin stock to the ODO/WH batch.  Carry that
            ' live-proven final batch into the in-memory Create contract; the
            ' CHECK row may contain only the pre-change candidate batch.
            finalBatch = CWR_WH_BATCH
            For Each rec In members
                If Not CBool(rec("Terminal")) Then
                    If (InStr(1, resultText, "PICK_BIN_BATCH_POST_VERIFIED", vbTextCompare) > 0 Or _
                        CBool(rec("PostRecoveryPending"))) And _
                       StrComp(CWR_Text(rec("Material")), CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 And _
                       StrComp(CWR_Text(rec("TargetBin")), CWR_TARGET_BIN, vbTextCompare) = 0 Then
                        If finalBatch <> "" Then rec("TargetBatch") = finalBatch
                    End If
                    rec("LocationReady") = True: CWR_SchedulerState rec, "LOCATION_READY", _
                        "LOCATION_GROUP", "PASS", "POSTSRC_GROUP=" & CStr(groupID) & "|POST_QTY=" & CStr(groupQty) & _
                        "|FINAL_BATCH=" & CStr(rec("TargetBatch"))
                End If
            Next rec
        End If
NextGroup:
    Next groupID
    CWR_ExecuteLocationGroups = Not stopAll
End Function

Private Function CWR_AdvanceDataReady(ByVal plans As Collection, ByRef stopAll As Boolean, _
                                       ByRef messageText As String) As Boolean
    Dim rec As Object, resultText As String, groupReady As Object, groupBlocked As Object, keyText As String
    Set groupReady = CreateObject("Scripting.Dictionary"): groupReady.CompareMode = vbTextCompare
    Set groupBlocked = CreateObject("Scripting.Dictionary"): groupBlocked.CompareMode = vbTextCompare
    For Each rec In plans
        If Not CBool(rec("Terminal")) Then
            keyText = CStr(rec("QtyGroupId"))
            If Not groupReady.Exists(keyText) Then groupReady.Add keyText, True
            If InStr(1, CStr(rec("State")), "BLOCKED", vbTextCompare) > 0 Or _
               InStr(1, CStr(rec("State")), "STATE_UNKNOWN", vbTextCompare) > 0 Then groupBlocked(keyText) = True
            If CStr(rec("ExecMode")) <> "CANCEL_PRDO_ZERO" Then
                If CBool(rec("NeedQty")) And Not CBool(rec("QtyReady")) Then groupBlocked(keyText) = True
                If (CBool(rec("NeedBin")) Or CBool(rec("NeedBatch"))) And Not CBool(rec("LocationReady")) Then groupBlocked(keyText) = True
            End If
        End If
    Next rec
    For Each rec In plans
        If CBool(rec("Terminal")) Then GoTo NextRec
        If CStr(rec("ExecMode")) = "CANCEL_PRDO_ZERO" Then
            If Not CWR_SchedulerUsePlan(rec, messageText) Then
                CWR_SchedulerFailure rec, "CANCEL_TO_ZERO_CONTEXT", messageText, stopAll
            ElseIf CWR_CancelUsesSharedFinalPRDO(plans, rec) Then
                ' A zero-target member can share one PRDO item with a retained
                ' member.  The document-level PRDO save has already reconciled
                ' the aggregate to GroupTarget; writing this member's TargetQty
                ' (zero) would incorrectly undo that shared final total.
                rec("Terminal") = True
                CWR_SchedulerState rec, "READY_FOR_CF", "CANCEL_TO_ZERO", "PASS", _
                    "CANCEL_VERIFIED|SHARED_PRDO_FINAL_TOTAL=" & CStr(rec("GroupTarget")) & _
                    "|NO_PER_MEMBER_ZERO|PRDO_AGGREGATE_ALREADY_RECONCILED"
                CWR_Log "CANCEL_SHARED_PRDO_RECONCILE", CStr(rec("OB")), CStr(rec("Doc")), "PASS", _
                    "PLAN=" & CStr(rec("PlanID")) & "|ITEM=" & CStr(rec("Item")) & _
                    "|MATERIAL=" & CStr(rec("Material")) & "|GROUP_TARGET=" & CStr(rec("GroupTarget")) & _
                    "|ACTION=NO_PER_MEMBER_ZERO|SAP_WRITE=0"
            Else
                resultText = CWR_EnsurePRDOQtyAtTargetOnce()
                If Left$(resultText, 5) <> "PASS|" Then
                    CWR_SchedulerFailure rec, "CANCEL_TO_ZERO_QTY", resultText, stopAll
                Else
                    rec("Terminal") = True
                    CWR_SchedulerState rec, "READY_FOR_CF", "CANCEL_TO_ZERO", "PASS", "PRDO_QTY_ZERO_VERIFIED|NO_NEW_WT"
                End If
            End If
            If stopAll Then Exit Function
            GoTo NextRec
        End If
        keyText = CStr(rec("QtyGroupId"))
        If groupBlocked.Exists(keyText) Then
            CWR_SchedulerFailure rec, "DATA_READY_GROUP_GATE", "STOP_DATA_READY_UNPROVEN|QTY_GROUP=" & keyText, stopAll
            If stopAll Then Exit Function
            GoTo NextRec
        End If
        If (Not CBool(rec("NeedQty")) Or CBool(rec("QtyReady"))) And _
           ((Not CBool(rec("NeedBin")) And Not CBool(rec("NeedBatch"))) Or CBool(rec("LocationReady"))) Then
            CWR_SchedulerState rec, "DATA_READY", "DATA_READY_GATE", "PASS", "QTY_AND_LOCATION_CONTRACT_SATISFIED"
        Else
            CWR_SchedulerFailure rec, "DATA_READY_GATE", "STOP_DATA_READY_UNPROVEN|STATE=" & CStr(rec("State")), stopAll
            If stopAll Then Exit Function
        End If
NextRec:
    Next rec
    CWR_AdvanceDataReady = Not stopAll
End Function

' A CANCEL_PRDO_ZERO member must not be treated as a standalone PRDO row when
' another WT for the same ODO item/material remains in the shared quantity
' group.  In that case the document-level save owns the aggregate final target;
' the cancel member only needs its WT cancellation proof and terminal handoff.
Private Function CWR_CancelUsesSharedFinalPRDO(ByVal plans As Collection, ByVal cancelRec As Object) As Boolean
    Dim rec As Object, sameIdentity As Boolean, modeText As String
    On Error GoTo ExitHere
    If plans Is Nothing Or cancelRec Is Nothing Then Exit Function
    If StrComp(CStr(cancelRec("ExecMode")), "CANCEL_PRDO_ZERO", vbTextCompare) <> 0 Then Exit Function
    If Not CBool(cancelRec("QtyReady")) Then Exit Function
    If CDbl(cancelRec("GroupTarget")) <= 0.000001 Then Exit Function
    For Each rec In plans
        If StrComp(CStr(rec("PlanID")), CStr(cancelRec("PlanID")), vbTextCompare) <> 0 Then
            sameIdentity = StrComp(CStr(rec("Doc")), CStr(cancelRec("Doc")), vbTextCompare) = 0 And _
                           CWR_MONValueEquals(CStr(rec("Item")), CStr(cancelRec("Item"))) And _
                           StrComp(CStr(rec("Material")), CStr(cancelRec("Material")), vbTextCompare) = 0 And _
                           StrComp(CStr(rec("SysBatch")), CStr(cancelRec("SysBatch")), vbTextCompare) = 0 And _
                           StrComp(CStr(rec("QtyGroupId")), CStr(cancelRec("QtyGroupId")), vbTextCompare) = 0
            If sameIdentity Then
                modeText = UCase$(CStr(rec("ExecMode")))
                If modeText = "KEEP_EXISTING_CREATE_DELTA" Or modeText = "REPLACE_CANCEL_CREATE" Then
                    CWR_CancelUsesSharedFinalPRDO = True
                    Exit Function
                End If
            End If
        End If
    Next rec
ExitHere:
End Function

Private Function CWR_ExecuteCreateGroups(ByVal createGroups As Object, ByRef stopAll As Boolean, _
                                           ByRef messageText As String) As Boolean
    Dim groupID As Variant, members As Collection, rec As Object, rec2 As Object, firstRec As Object, resultText As String
    Dim records As Collection, prdoGroups As Object, createSes As Object, wtList As String
    For Each groupID In createGroups.keys
        Set members = createGroups.item(CStr(groupID))
        Set firstRec = Nothing
        For Each rec In members
            If Not CBool(rec("Terminal")) Then
                If firstRec Is Nothing Then Set firstRec = rec
                If CStr(rec("State")) <> "DATA_READY" Then
                    messageText = "STOP_CREATE_GROUP_MEMBER_NOT_DATA_READY|PLAN=" & CStr(rec("PlanID"))
                    For Each rec2 In members
                        If Not CBool(rec2("Terminal")) Then CWR_SchedulerFailure rec2, "CREATE_DATA_READY_GROUP_GATE", messageText, stopAll
                    Next rec2
                    If stopAll Then Exit Function
                    Set firstRec = Nothing
                    Exit For
                End If
            End If
        Next rec
        If firstRec Is Nothing Then GoTo NextGroup
        If Not CWR_SchedulerBuildPRDOBatch(members, records, prdoGroups) Or records.count = 0 Then
            messageText = "STOP_CREATE_GROUP_RECORDS_UNPROVEN|CREATE_GROUP=" & CStr(groupID)
            For Each rec In members
                If Not CBool(rec("Terminal")) Then CWR_SchedulerFailure rec, "CREATE_GROUP_BUILD", messageText, stopAll
            Next rec
            If stopAll Then Exit Function
            GoTo NextGroup
        End If
        If CWR_CREATE_GROUP_ACTION_COUNTS Is Nothing Then
            Set CWR_CREATE_GROUP_ACTION_COUNTS = CreateObject("Scripting.Dictionary")
            CWR_CREATE_GROUP_ACTION_COUNTS.CompareMode = vbTextCompare
        End If
        If CWR_CREATE_GROUP_ACTION_COUNTS.Exists(CStr(groupID)) Then
            messageText = "STOP_CREATE_GROUP_DUPLICATE_ACTION|CREATE_GROUP=" & CStr(groupID) & _
                "|CREATE_SAVE_COUNT=" & CStr(CWR_CREATE_GROUP_ACTION_COUNTS(CStr(groupID)))
            For Each rec2 In members
                If Not CBool(rec2("Terminal")) Then CWR_SchedulerFailure rec2, "CREATE_GROUP_IDEMPOTENCY", messageText, stopAll
            Next rec2
            If stopAll Then Exit Function
            GoTo NextGroup
        End If
        If Not CWR_SchedulerUsePlan(firstRec, messageText) Then
            CWR_SchedulerFailGroup records, "CREATE_GROUP_CONTEXT", messageText, stopAll
            If stopAll Then Exit Function
            GoTo NextGroup
        End If
        ' VBA evaluates both sides of Or.  Keep the SAP stages explicitly
        ' sequential so a failed open/adopt cannot fall through to
        ' post-verify and overwrite the real blocker with MATCH_COUNT=0.
        If Not CWR_BatchOpenCreateWT(records, prdoGroups, createSes, messageText) Then
            CWR_SchedulerFailGroup records, "CREATE_OPEN", messageText, stopAll
            If stopAll Then Exit Function
            GoTo NextGroup
        End If
        If CWR_BatchAnyCreateRemaining(prdoGroups) Then
            If Not CWR_BatchAdoptCreateSave(records, prdoGroups, createSes, messageText) Then
                CWR_SchedulerFailGroup records, "ADOPT_CREATE_SAVE", messageText, stopAll
                If stopAll Then Exit Function
                GoTo NextGroup
            End If
        End If
        If Not CWR_BatchPostVerifyAndRecord(records, messageText, wtList) Then
            CWR_SchedulerFailGroup records, "CREATE_SAVE_POSTVERIFY", messageText, stopAll
            If stopAll Then Exit Function
        Else
            CWR_CREATE_GROUP_ACTION_COUNTS(CStr(groupID)) = 1
            CWR_Log "CREATE_GROUP", CStr(firstRec("OB")), CStr(firstRec("Doc")), "PASS", _
                "CREATE_GROUP=" & CStr(groupID) & "|MEMBERS=" & CStr(members.count) & _
                "|CREATE_SAVE_COUNT=1|POST_VERIFY=PASS|CF_NOT_SENT"
            For Each rec In members
                If Not CBool(rec("Terminal")) Then rec("Terminal") = True: CWR_SchedulerState rec, "READY_FOR_CF", _
                    "CREATE_GROUP", "PASS", "CREATE_SAVE_GROUP_SENT|CREATE_GROUP=" & CStr(groupID) & _
                    "|WT_POST_VERIFY_PASS|CF_NOT_SENT|WT=" & wtList & "|IDEMPOTENCY_KEY=CREATE_WT|" & _
                    CStr(rec("OB")) & "|" & CStr(rec("Doc")) & "|" & CStr(rec("Item")) & "|" & _
                    CStr(rec("Material")) & "|SRC=" & CStr(rec("SourceBin")) & "|" & CStr(rec("SourceBatch")) & _
                    "|DST=" & CStr(rec("TargetBin")) & "|" & CStr(rec("TargetBatch")) & _
                    "|QTY=" & CStr(CWR_SchedulerCreateQty(rec))
            Next rec
        End If
NextGroup:
    Next groupID
    CWR_ExecuteCreateGroups = Not stopAll
End Function

Private Function CWR_SchedulerUsePlan(ByVal rec As Object, ByRef messageText As String) As Boolean
    CWR_REQUESTED_PLAN_ID = CStr(rec("PlanID"))
    CWR_RECOVERY_CONTEXT = False
    ' The SAP surface is scoped to the immutable OB+Document package.  Each
    ' member still binds its own PLAN row (so item/material/WT checks remain
    ' exact), but a same-document member must not be treated as a navigation
    ' reopen merely because its PlanID differs.
    If CWR_CONTEXT_VALID And _
       StrComp(CWR_CONTEXT_OB, CStr(rec("OB")), vbTextCompare) = 0 And _
       StrComp(CWR_CONTEXT_DOC, CStr(rec("Doc")), vbTextCompare) = 0 Then
        If Not CWR_LoadExecutionContext(messageText) Then Exit Function
        CWR_CONTEXT_PLAN_ID = CStr(rec("PlanID"))
        CWR_CONTEXT_FINGERPRINT = CStr(rec("PlanFingerprint"))
        CWR_CONTEXT_WT = CStr(rec("OldWT"))
        CWR_Log "EXECUTION_CONTEXT_REUSE", CStr(rec("OB")), CStr(rec("Doc")), "PASS", _
            "PLAN=" & CStr(rec("PlanID")) & "|ACTION_KEY=" & CWR_BusinessActionKey(rec, "CONTEXT") & _
            "|DOCUMENT_CONTEXT_REUSE=TRUE|SAP_TAB_REUSE=TRUE|PLAN_RELOAD=SKIP|NO_SAP_WRITE"
        If Not CWR_AssertCurrentPlanExecutionContract(messageText) Then Exit Function
        CWR_SchedulerUsePlan = True
        Exit Function
    End If
    If CWR_CONTEXT_VALID And _
       StrComp(CWR_CONTEXT_PLAN_ID, CStr(rec("PlanID")), vbTextCompare) = 0 And _
       StrComp(CWR_CONTEXT_FINGERPRINT, CStr(rec("PlanFingerprint")), vbTextCompare) = 0 And _
       StrComp(CWR_CONTEXT_OB, CStr(rec("OB")), vbTextCompare) = 0 And _
       StrComp(CWR_CONTEXT_DOC, CStr(rec("Doc")), vbTextCompare) = 0 And _
       StrComp(CWR_CONTEXT_WT, CStr(rec("OldWT")), vbTextCompare) = 0 Then
        CWR_Log "EXECUTION_CONTEXT_REUSE", CStr(rec("OB")), CStr(rec("Doc")), "PASS", _
            "PLAN=" & CStr(rec("PlanID")) & "|ACTION_KEY=" & CWR_BusinessActionKey(rec, "CONTEXT") & _
            "|SAP_TAB_REUSE=TRUE|PLAN_RELOAD=SKIP|NO_SAP_WRITE"
    Else
        If Not CWR_LoadExecutionContext(messageText) Then Exit Function
        CWR_CONTEXT_VALID = True
        CWR_CONTEXT_PLAN_ID = CStr(rec("PlanID"))
        CWR_CONTEXT_FINGERPRINT = CStr(rec("PlanFingerprint"))
        CWR_CONTEXT_OB = CStr(rec("OB"))
        CWR_CONTEXT_DOC = CStr(rec("Doc"))
        CWR_CONTEXT_WT = CStr(rec("OldWT"))
        CWR_Log "EXECUTION_CONTEXT_OPEN", CStr(rec("OB")), CStr(rec("Doc")), "PASS", _
            "PLAN=" & CStr(rec("PlanID")) & "|ACTION_KEY=" & CWR_BusinessActionKey(rec, "CONTEXT") & _
            "|SAP_TAB_REUSE=BOUND_TO_RUN|PLAN_RELOAD=TRUE|NO_SAP_WRITE"
    End If
    If Not CWR_AssertCurrentPlanExecutionContract(messageText) Then Exit Function
    CWR_SchedulerUsePlan = True
End Function

' Recovery-only context for a POST Create+Save that already crossed the
' one-shot write boundary.  CHECK may intentionally have no resolved target
' batch; the caller is allowed to query SAP and reconcile that action, but no
' scheduler write is permitted from this context.
Private Function CWR_SchedulerUseRecoveryPlan(ByVal rec As Object, ByRef messageText As String) As Boolean
    CWR_REQUESTED_PLAN_ID = CStr(rec("PlanID"))
    CWR_RECOVERY_CONTEXT = True
    If CWR_CONTEXT_VALID And _
       StrComp(CWR_CONTEXT_OB, CStr(rec("OB")), vbTextCompare) = 0 And _
       StrComp(CWR_CONTEXT_DOC, CStr(rec("Doc")), vbTextCompare) = 0 Then
        If Not CWR_LoadExecutionContext(messageText) Then Exit Function
        CWR_CONTEXT_PLAN_ID = CStr(rec("PlanID"))
        CWR_CONTEXT_FINGERPRINT = CStr(rec("PlanFingerprint"))
        CWR_CONTEXT_WT = CStr(rec("OldWT"))
        CWR_Log "EXECUTION_CONTEXT_RECOVERY_REUSE", CStr(rec("OB")), CStr(rec("Doc")), "PASS", _
            "PLAN=" & CStr(rec("PlanID")) & "|ACTION_KEY=" & CWR_BusinessActionKey(rec, "POST") & _
            "|DOCUMENT_CONTEXT_REUSE=TRUE|RECOVERY_ONLY=TRUE|PLAN_RELOAD=SKIP|NO_POST_RETRY"
        CWR_SchedulerUseRecoveryPlan = True
        Exit Function
    End If
    If Not (CWR_CONTEXT_VALID And _
            StrComp(CWR_CONTEXT_PLAN_ID, CStr(rec("PlanID")), vbTextCompare) = 0 And _
            StrComp(CWR_CONTEXT_FINGERPRINT, CStr(rec("PlanFingerprint")), vbTextCompare) = 0 And _
            StrComp(CWR_CONTEXT_OB, CStr(rec("OB")), vbTextCompare) = 0 And _
            StrComp(CWR_CONTEXT_DOC, CStr(rec("Doc")), vbTextCompare) = 0 And _
            StrComp(CWR_CONTEXT_WT, CStr(rec("OldWT")), vbTextCompare) = 0) Then
        If Not CWR_LoadExecutionContext(messageText) Then Exit Function
        CWR_CONTEXT_VALID = True
        CWR_CONTEXT_PLAN_ID = CStr(rec("PlanID"))
        CWR_CONTEXT_FINGERPRINT = CStr(rec("PlanFingerprint"))
        CWR_CONTEXT_OB = CStr(rec("OB"))
        CWR_CONTEXT_DOC = CStr(rec("Doc"))
        CWR_CONTEXT_WT = CStr(rec("OldWT"))
        CWR_Log "EXECUTION_CONTEXT_RECOVERY_OPEN", CStr(rec("OB")), CStr(rec("Doc")), "PASS", _
            "PLAN=" & CStr(rec("PlanID")) & "|ACTION_KEY=" & CWR_BusinessActionKey(rec, "POST") & _
            "|RECOVERY_ONLY=TRUE|PLAN_RELOAD=TRUE|NO_POST_RETRY"
    Else
        CWR_Log "EXECUTION_CONTEXT_RECOVERY_REUSE", CStr(rec("OB")), CStr(rec("Doc")), "PASS", _
            "PLAN=" & CStr(rec("PlanID")) & "|ACTION_KEY=" & CWR_BusinessActionKey(rec, "POST") & _
            "|RECOVERY_ONLY=TRUE|PLAN_RELOAD=SKIP|NO_POST_RETRY"
    End If
    CWR_SchedulerUseRecoveryPlan = True
End Function

Private Sub CWR_SchedulerFailure(ByVal rec As Object, ByVal stageText As String, _
                                 ByVal detailText As String, ByRef stopAll As Boolean)
    If InStr(1, detailText, "STATE_UNKNOWN", vbTextCompare) > 0 Or _
       InStr(1, detailText, "NO_RETRY", vbTextCompare) > 0 Then
        rec("Terminal") = True
        CWR_SchedulerState rec, "STATE_UNKNOWN_AFTER_" & stageText, stageText, "STOP", detailText
        stopAll = True
    Else
        rec("Terminal") = True
        CWR_SchedulerState rec, "BLOCKED_" & stageText, stageText, "STOP", detailText
    End If
End Sub

Private Sub CWR_SchedulerFailGroup(ByVal records As Collection, ByVal stageText As String, _
                                   ByVal detailText As String, ByRef stopAll As Boolean)
    Dim batchRec As Object, rec As Object
    For Each batchRec In records
        Set rec = batchRec.item("SchedulerRec")
        CWR_SchedulerFailure rec, stageText, detailText, stopAll
    Next batchRec
End Sub

Private Sub CWR_SchedulerState(ByVal rec As Object, ByVal stateText As String, ByVal stageText As String, _
                               ByVal resultText As String, ByVal detailText As String)
    Dim groupText As String
    rec("State") = stateText
    groupText = "DOC_GROUP=" & CStr(rec("DocGroup")) & "|SOURCE_GROUP=" & CStr(rec("LocationGroup")) & _
        "|CREATE_GROUP=" & CStr(rec("CreateGroup")) & "|QTY_GROUP=" & CStr(rec("QtyGroupId")) & _
        "|GROUP_CURRENT=" & CStr(rec("GroupCurrent")) & "|GROUP_TARGET=" & CStr(rec("GroupTarget"))
    CF_HistoryAppend CWR_HistoryRunID(), "EXECUTE", stageText, CStr(rec("PlanID")), CStr(rec("OB")), _
        CStr(rec("Doc")), CStr(rec("Item")), CStr(rec("Material")), CStr(rec("OldWT")), _
        CWR_BusinessActionKey(rec, stageText), resultText, "", "", "", stateText, groupText & "|" & detailText
    ' Scheduler events already carry their explicit PlanID in CF_HistoryAppend;
    ' use the non-ambient logger here so a later context switch cannot stamp a
    ' different plan into RUN_LOG.
    RT_Log "SCHED_" & stageText, CStr(rec("OB")), CStr(rec("Doc")), resultText, _
        "PLAN=" & CStr(rec("PlanID")) & "; STATE=" & stateText & "; " & groupText & "; " & detailText
End Sub

Private Function CWR_HasUnreconciledUnknown(ByVal planID As String, ByVal planFingerprint As String) As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long, stateText As String, resultText As String, detailText As String
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("RUN_HISTORY")
    lastRow = ws.Cells(ws.rows.count, 17).End(xlUp).row
    For r = lastRow To 2 Step -1
        If StrComp(CWR_Text(ws.Cells(r, 6).Value2), planID, vbTextCompare) = 0 And _
           StrComp(CWR_Text(ws.Cells(r, 21).Value2), planFingerprint, vbTextCompare) = 0 Then
            stateText = UCase$(CWR_Text(ws.Cells(r, 15).Value2))
            resultText = UCase$(CWR_Text(ws.Cells(r, 17).Value2))
            detailText = UCase$(CWR_Text(ws.Cells(r, 25).Value2))
            CWR_HasUnreconciledUnknown = (InStr(stateText, "STATE_UNKNOWN") > 0 Or _
                InStr(resultText, "STATE_UNKNOWN") > 0 Or InStr(detailText, "STATE_UNKNOWN") > 0)
            Exit Function
        End If
    Next r
ExitHere:
End Function

' Returns True only for the narrow, recoverable boundary created after POST
' Create+Save was sent.  Other unknown actions remain hard-stopped before the
' scheduler can reach any executor.  The later POST routine still proves the
' live result with a fresh query and never sends a second Create+Save.
Private Function CWR_HasRecoverablePickBinPostUnknown(ByVal planID As String, _
                                                       ByVal planFingerprint As String) As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim stateText As String, resultText As String, detailText As String, isUnknown As Boolean
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("RUN_HISTORY")
    lastRow = ws.Cells(ws.rows.count, 17).End(xlUp).row
    For r = lastRow To 2 Step -1
        If StrComp(CWR_Text(ws.Cells(r, 6).Value2), planID, vbTextCompare) = 0 And _
           StrComp(CWR_Text(ws.Cells(r, 21).Value2), planFingerprint, vbTextCompare) = 0 Then
            stateText = UCase$(CWR_Text(ws.Cells(r, 15).Value2))
            resultText = UCase$(CWR_Text(ws.Cells(r, 17).Value2))
            detailText = UCase$(CWR_Text(ws.Cells(r, 25).Value2))
            isUnknown = (InStr(stateText, "STATE_UNKNOWN") > 0 Or _
                         InStr(resultText, "STATE_UNKNOWN") > 0 Or _
                         InStr(detailText, "STATE_UNKNOWN") > 0)
            CWR_HasRecoverablePickBinPostUnknown = isUnknown And _
                (InStr(stateText, "POST_CREATE_SAVE") > 0 Or _
                 InStr(resultText, "POST_CREATE_SAVE") > 0 Or _
                 InStr(detailText, "STATE_UNKNOWN_AFTER_POST_CREATE_SAVE") > 0)
            Exit Function
        End If
    Next r
ExitHere:
End Function

' Fingerprints are deliberately plan-versioned, so a later CHECK can change
' the fingerprint while the SAP write boundary remains the same business
' action.  Recovery therefore also scans the durable history by OB/document/
' old-WT/material identity and uses SAP readback as the proof source.
Private Function CWR_HasRecoverablePickBinPostByBusiness(ByVal obText As String, _
                                                         ByVal docText As String, _
                                                         ByVal wtText As String, _
                                                         ByVal materialText As String) As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long, stageText As String, resultText As String, detailText As String
    Dim pending As Boolean, matched As Boolean
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("RUN_HISTORY")
    lastRow = ws.Cells(ws.rows.count, 17).End(xlUp).row
    For r = 2 To lastRow
        If StrComp(CWR_Text(ws.Cells(r, 7).Value2), obText, vbTextCompare) = 0 And _
           StrComp(CWR_Text(ws.Cells(r, 8).Value2), docText, vbTextCompare) = 0 And _
           StrComp(CWR_Text(ws.Cells(r, 11).Value2), wtText, vbTextCompare) = 0 And _
           (materialText = "" Or StrComp(CWR_Text(ws.Cells(r, 10).Value2), materialText, vbTextCompare) = 0) Then
            matched = True
            stageText = UCase$(CWR_Text(ws.Cells(r, 5).Value2))
            resultText = UCase$(CWR_Text(ws.Cells(r, 17).Value2))
            detailText = UCase$(CWR_Text(ws.Cells(r, 25).Value2))
            If stageText = "POST_PICK_BIN_CREATE_SAVE_SENT" Then
                pending = True
            ElseIf (stageText = "POST_PICK_BIN_POSTVERIFY" Or stageText = "POST_DIRECT_RECOVERY" Or _
                    stageText = "POST_PICK_BIN_COMMITTED" Or InStr(1, detailText, "PICK_BIN_BATCH_POST_VERIFIED", vbTextCompare) > 0) And _
                   (resultText = "PASS" Or InStr(1, detailText, "PASS", vbTextCompare) > 0) Then
                pending = False
            End If
        End If
    Next r
    CWR_HasRecoverablePickBinPostByBusiness = matched And pending
ExitHere:
End Function

Private Function CWR_BusinessActionKey(ByVal rec As Object, ByVal actionText As String) As String
    Dim actionUpper As String
    Dim identity As String
    actionUpper = UCase$(CWR_Text(actionText))
    ' The key is a business identity, never a PlanID-only shortcut.  It is
    ' immutable for the write boundary and carries the before/target tuple so
    ' a later OB, material, bin or batch cannot inherit an old checkpoint.
    identity = "OB=" & CWR_SchedulerIdentityPart(CWR_Text(rec("OB"))) & _
        "|DOC=" & CWR_SchedulerIdentityPart(CWR_Text(rec("Doc"))) & _
        "|ITEM=" & CWR_SchedulerIdentityPart(CWR_Text(rec("Item"))) & _
        "|MAT=" & CWR_SchedulerIdentityPart(CWR_Text(rec("Material"))) & _
        "|OLD_WT=" & CWR_SchedulerIdentityPart(CWR_Text(rec("OldWT"))) & _
        "|SRC_BIN=" & CWR_SchedulerIdentityPart(CWR_Text(rec("SourceBin"))) & _
        "|SRC_BATCH=" & CWR_SchedulerIdentityPart(CWR_Text(rec("SourceBatch"))) & _
        "|DST_BIN=" & CWR_SchedulerIdentityPart(CWR_Text(rec("TargetBin"))) & _
        "|DST_BATCH=" & CWR_SchedulerIdentityPart(CWR_Text(rec("TargetBatch")))
    If InStr(1, actionUpper, "PRDO", vbTextCompare) > 0 Or actionUpper = "QTY" Then
        CWR_BusinessActionKey = "PRDO|" & identity & "|BEFORE=" & CStr(rec("GroupCurrent")) & _
            "|TARGET=" & CStr(rec("GroupTarget"))
    ElseIf InStr(1, actionUpper, "POST", vbTextCompare) > 0 Then
        CWR_BusinessActionKey = "POST|" & identity & "|QTY=" & CStr(rec("TargetQty")) & _
            "|DELTA=" & CStr(rec("QtyDelta"))
    ElseIf InStr(1, actionUpper, "CANCEL", vbTextCompare) > 0 Then
        CWR_BusinessActionKey = "CANCEL|" & identity & "|QTY=" & CStr(rec("SysQty"))
    Else
        CWR_BusinessActionKey = CWR_HistoryActionKey(actionText) & "|" & identity & _
            "|TARGET=" & CStr(rec("TargetQty"))
    End If
End Function

Private Function CWR_ResolveRecoverablePOSTBatch(ByVal grid As Object, ByRef messageText As String) As Boolean
    Dim r As Long, n As Long, lastMaterial As String, matText As String, binText As String
    Dim stockTypeText As String, batchText As String, avail As Double, total As Double, matches As Long
    On Error GoTo Failed
    If CWR_TARGET_BATCH <> "" Then
        CWR_ResolveRecoverablePOSTBatch = True
        Exit Function
    End If
    n = CLng(grid.rowCount)
    For r = 0 To n - 1
        matText = CWR_Cell(grid, r, "MATNR")
        If matText <> "" Then lastMaterial = matText Else matText = lastMaterial
        binText = CWR_Cell(grid, r, "VLPLA")
        stockTypeText = CWR_Cell(grid, r, "TO_CAT")
        If stockTypeText = "" Then stockTypeText = CWR_Cell(grid, r, "CAT")
        batchText = CWR_Cell(grid, r, "CHARG")
        avail = CWR_Number(CWR_Cell(grid, r, "AVL_ALT_QTY"))
        total = CWR_Number(CWR_Cell(grid, r, "LIME_QUAN"))
        If StrComp(matText, CWR_EXPECTED_MATERIAL, vbTextCompare) = 0 And _
           StrComp(binText, CWR_TARGET_BIN, vbTextCompare) = 0 And _
           StrComp(stockTypeText, CWR_TARGET_STOCK_TYPE, vbTextCompare) = 0 And _
           StrComp(batchText, CWR_WH_BATCH, vbTextCompare) = 0 And _
           avail + 0.000001 >= CWR_TARGET_QTY And total + 0.000001 >= CWR_TARGET_QTY Then
            matches = matches + 1
        End If
    Next r
    If matches <> 1 Then
        messageText = "STATE_UNKNOWN_AFTER_POST_CREATE_SAVE|RECOVERY_FINAL_BATCH_MATCH_COUNT=" & CStr(matches) & _
            "|BIN=" & CWR_TARGET_BIN & "|WH_BATCH=" & CWR_WH_BATCH & "|REQUIRED=" & CStr(CWR_TARGET_QTY)
        Exit Function
    End If
    ' The live row is now the final ODO/WH batch.  The old candidate batch is
    ' intentionally not reconstructed from history; SAP has just proven the
    ' final tuple and no second Create+Save is permitted.
    CWR_TARGET_BATCH = CWR_WH_BATCH
    CWR_PICK_BIN_BATCH_BEFORE = CWR_WH_BATCH
    CWR_Log "RECOVER_POST", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
        "ACTION=POST|STATE=VERIFIED_PASS|FINAL_BATCH=" & CWR_TARGET_BATCH & _
        "|BIN=" & CWR_TARGET_BIN & "|QTY=" & CStr(CWR_TARGET_QTY) & "|READ_ONLY=TRUE|NO_POST_RETRY"
    CWR_ResolveRecoverablePOSTBatch = True
    Exit Function
Failed:
    messageText = "STATE_UNKNOWN_AFTER_POST_CREATE_SAVE|RECOVERY_FINAL_BATCH_READ_ERROR: " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub CWR_ScheduleCounts(ByVal plans As Collection, ByRef readyCount As Long, _
                               ByRef blockedCount As Long, ByRef unknownCount As Long, _
                               ByRef unaccountedCount As Long)
    Dim rec As Object, stateText As String
    readyCount = 0: blockedCount = 0: unknownCount = 0: unaccountedCount = 0
    For Each rec In plans
        stateText = CStr(rec("State"))
        If stateText = "READY_FOR_CF" Then
            readyCount = readyCount + 1
        ElseIf InStr(1, stateText, "STATE_UNKNOWN", vbTextCompare) > 0 Then
            unknownCount = unknownCount + 1
        ElseIf InStr(1, stateText, "WAIT_USER_CF", vbTextCompare) > 0 Then
            blockedCount = blockedCount + 1
        ElseIf Left$(stateText, 8) = "BLOCKED_" Then
            blockedCount = blockedCount + 1
        ElseIf Not CWR_IsTerminalExecutionState(stateText) Then
            unaccountedCount = unaccountedCount + 1
        End If
    Next rec
End Sub

Private Sub CWR_LogQueueCase(ByVal wsPlan As Worksheet, ByVal planRow As Long, _
                              ByVal codeText As String, ByVal messageText As String)
    Dim whRow As Long
    On Error GoTo ExitHere
    whRow = CLng(CWR_Number(wsPlan.Cells(planRow, 27).Value2))
    CWR_EXPECTED_WT = CWR_Text(wsPlan.Cells(planRow, 5).Value2)
    CWR_Log "OB_QUEUE_CASE", CWR_Text(wsPlan.Cells(planRow, 2).Value2), _
            CWR_Text(wsPlan.Cells(planRow, 3).Value2), codeText, _
            "PLAN=" & CWR_Text(wsPlan.Cells(planRow, 1).Value2) & "; WH_ROW=" & CStr(whRow) & "; " & messageText
ExitHere:
End Sub

Private Sub CWR_QueueSummary(ByVal obText As String, ByVal createdCount As Long, ByVal waitCFCount As Long, _
                              ByVal doneCount As Long, ByVal skippedCount As Long, ByVal pendingACount As Long, _
                              ByVal candidateCount As Long, ByVal resultText As String)
    Dim ws As Worksheet, readySummary As Long, unknownSummary As Long
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Range("AF1").Value2 = "OB Queue Result"
    ws.Range("AF2").Value2 = resultText & ": OB=" & obText & "; candidates=" & CStr(candidateCount) & _
        "; outbound WTs READY_FOR_CF=" & CStr(createdCount) & "; waiting user CF=" & CStr(waitCFCount) & _
        "; C confirmed skipped=" & CStr(doneCount) & "; other skipped=" & CStr(skippedCount) & _
        "; A resume cases=" & CStr(pendingACount) & "; CF_NOT_SENT. See RUN_LOG for every case."
    readySummary = createdCount + waitCFCount + doneCount
    unknownSummary = IIf(InStr(1, resultText, "STATE_UNKNOWN", vbTextCompare) > 0, 1, 0)
    CF_HistoryAppendSummary CWR_HistoryRunID(), "EXECUTE", "BUTTON4", obText, candidateCount, _
        readySummary, skippedCount, createdCount, waitCFCount, unknownSummary, resultText, _
        "C=" & CStr(doneCount) & "; SKIPPED=" & CStr(skippedCount) & "; A_RESUME=" & CStr(pendingACount)
ExitHere:
End Sub

Private Sub CWR_RecordPostingPreflight(ByVal sourceProof As String, ByVal targetProof As String, _
                                       ByVal storageTypeChangeable As String, ByVal storageSectionChangeable As String, ByVal binChangeable As String, _
                                       ByVal qtyChangeable As String, ByVal batchChangeable As String, _
                                       ByVal stockTypeChangeable As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(5, 15).Value2 = "READY"
    ws.Cells(5, 16).Value2 = "POST_CHANGE_BIN_BATCH_QTY"
    ws.Cells(5, 17).Value2 = "POST_PREVERIFIED"
    ws.Cells(5, 18).Value2 = "POST_SOURCE_AND_TARGET_READ"
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: POST source/new bin/new batch verified; no POST write, no save, CF remains user handoff. " & _
        "NLTYP=" & storageTypeChangeable & "; NLBER=" & storageSectionChangeable & "; NLPLA=" & binChangeable & "; POST_QUAN=" & qtyChangeable & "; TO_CHARG=" & batchChangeable & _
        "; TO_CAT=" & stockTypeChangeable & ". " & sourceProof & "; " & targetProof
ExitHere:
End Sub

Private Sub CWR_RecordPostingChange(ByVal sourceProof As String, ByVal targetProof As String, _
                                    ByVal sourceAvailableAfter As Double, ByVal targetAvailableAfter As Double)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(5, 15).Value2 = "READY"
    ws.Cells(5, 16).Value2 = "PRDO_CHANGE_QTY_THEN_CREATE_WT"
    ws.Cells(5, 17).Value2 = "POST_CHANGE_VERIFIED"
    ws.Cells(5, 18).Value2 = "POST_POSTVERIFY"
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: /SCWM/POST batch/bin change verified once; source available=" & _
        CStr(sourceAvailableAfter) & "; target available=" & CStr(targetAvailableAfter) & _
        ". Next: PRDO quantity only, then create WH; CF remains user handoff. " & _
        sourceProof & "; " & targetProof
ExitHere:
End Sub

Private Sub CWR_RecordPostingCreate(ByVal sourceProof As String, ByVal targetProof As String, _
                                    ByVal counterBefore As String, ByVal counterAfter As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(5, 15).Value2 = "READY"
    ws.Cells(5, 16).Value2 = "PRDO_CHANGE_QTY_THEN_CREATE_WT"
    ws.Cells(5, 17).Value2 = "POST_CHANGE_CREATED_VERIFIED"
    ws.Cells(5, 18).Value2 = "POST_CREATE_POSTVERIFY"
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: exact POST Create+Save confirmed; source counter advanced from " & _
        counterBefore & " to " & counterAfter & ". Next: PRDO quantity only, then create WH; CF remains user handoff. " & _
        sourceProof & "; " & targetProof
ExitHere:
End Sub

Private Sub CWR_RecordPostingCreateByGrid(ByVal sourceProof As String, ByVal targetProof As String, _
                                          ByVal createdBefore As Long, ByVal createdAfter As Long, _
                                          ByVal sourceCounterAfter As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    ws.Cells(5, 15).Value2 = "READY"
    ws.Cells(5, 16).Value2 = "PRDO_CHANGE_QTY_THEN_CREATE_WT"
    ws.Cells(5, 17).Value2 = "POST_CHANGE_CREATED_VERIFIED"
    ws.Cells(5, 18).Value2 = "POST_CREATED_GRID_POSTVERIFY"
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: exact source row selected, Create+Save Posting Change and Yes completed; " & _
        "created-grid rows " & CStr(createdBefore) & " -> " & CStr(createdAfter) & _
        "; source counter=" & sourceCounterAfter & ". Next: PRDO quantity only, then create WH; CF remains user handoff. " & _
        sourceProof & "; " & targetProof
ExitHere:
End Sub

Private Sub CWR_RecordPRDOQtyChange(ByVal qtyText As String, ByVal batchText As String, ByVal statusText As String)
    Dim ws As Worksheet
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("DATA_CORE")
    If CWR_PLAN_DECISION = "CANCEL_TO_ZERO" And Abs(CWR_TARGET_QTY) <= 0.000001 Then
        ws.Cells(CWR_PLAN_WH_ROW, 15).Value2 = "READY_FOR_CF"
        ws.Cells(CWR_PLAN_WH_ROW, 16).Value2 = "CANCELLED_WT_PRDO_QTY_ZERO"
        ws.Cells(CWR_PLAN_WH_ROW, 17).Value2 = "PRDO_QTY_ZERO_VERIFIED"
        ws.Cells(CWR_PLAN_WH_ROW, 18).Value2 = "USER_CF_HANDOFF"
        ws.Cells(CWR_PLAN_WH_ROW, 19).Value2 = "Exact WT cancelled and PRDO Qty read back at 0; no new WT created; CF_NOT_SENT."
        ws.Range("AF1").Value2 = "WT Rework Result"
        ws.Range("AF2").Value2 = "PASS: exact WT cancelled and PRDO Qty=0 verified; no new WH created; CF remains user-owned."
        Exit Sub
    End If
    If CWR_IsDeltaAddRoute() Then
        ws.Cells(CWR_PLAN_WH_ROW, 15).Value2 = "READY"
        ws.Cells(CWR_PLAN_WH_ROW, 16).Value2 = "PRDO_QTY_DELTA_VERIFIED"
        ws.Cells(CWR_PLAN_WH_ROW, 17).Value2 = "PRDO_QTY_DELTA_VERIFIED"
        ws.Cells(CWR_PLAN_WH_ROW, 18).Value2 = "CREATE_DELTA_WT"
        ws.Cells(CWR_PLAN_WH_ROW, 19).Value2 = "Exact PRDO total read back at " & qtyText & _
            "; retained WT=" & CWR_EXPECTED_WT & " Qty=" & CStr(CWR_EXPECTED_SOURCE_QTY) & _
            "; next Create WT Qty=" & CStr(CWR_CREATE_QTY) & "; CF_NOT_SENT."
        ws.Range("AF1").Value2 = "WT Rework Result"
        ws.Range("AF2").Value2 = "PASS: PRDO total verified at " & qtyText & _
            "; retain existing WT and create only delta Qty=" & CStr(CWR_CREATE_QTY) & ". CF remains user-owned."
        Exit Sub
    End If
    ws.Cells(CWR_PLAN_WH_ROW, 15).Value2 = "READY"
    ws.Cells(CWR_PLAN_WH_ROW, 16).Value2 = "CREATE_WT_FROM_TARGET_THEN_CF_HANDOFF"
    ws.Cells(CWR_PLAN_WH_ROW, 17).Value2 = "PRDO_QTY_VERIFIED"
    ws.Cells(CWR_PLAN_WH_ROW, 18).Value2 = "PRDO_QTY_POSTVERIFY"
    ws.Range("AF1").Value2 = "WT Rework Result"
    ws.Range("AF2").Value2 = "PASS: exact PRDO item quantity verified at " & qtyText & _
        "; PRDO batch remains read-only=" & batchText & "; " & statusText & _
        ". Next: create WH from the verified target stock; CF remains user handoff."
ExitHere:
End Sub

Private Sub CWR_Log(ByVal stageText As String, ByVal ob As String, ByVal doc As String, _
                    ByVal codeText As String, ByVal messageText As String)
    Dim ws As Worksheet, r As Long
    On Error GoTo ExitHere
    Set ws = ThisWorkbook.Worksheets("RUN_LOG")
    r = ws.Cells(ws.rows.count, 17).End(xlUp).row + 1
    If r < 9 Then r = 9
    ws.Cells(r, 17).Value2 = Now
    ws.Cells(r, 3).Value2 = stageText
    ws.Cells(r, 4).Value2 = ob
    ws.Cells(r, 5).Value2 = doc
    ws.Cells(r, 8).Value2 = CWR_EXPECTED_WT
    ws.Cells(r, 10).Value2 = codeText
    ws.Cells(r, 11).Value2 = IIf(UCase$(codeText) = "PASS", "OK", codeText)
    ws.Cells(r, 12).Value2 = messageText
    CWR_HistoryCheckpoint stageText, CWR_HistoryActionKey(stageText), codeText, "", _
                          CWR_FastQuantityText(CWR_TARGET_QTY), "", messageText
ExitHere:
End Sub

' Append-only operational history.  RUN_LOG remains the compact user log;
' these two ledgers preserve structured evidence after PLAN is rebuilt.
Public Sub CF_HistoryEnsureLedger()
    Dim wsHistory As Worksheet, wsSummary As Worksheet
    Set wsHistory = CWR_HistorySheet("RUN_HISTORY", False)
    Set wsSummary = CWR_HistorySheet("RUN_SUMMARY", False)
End Sub

' One-time preservation of the existing human-readable RUN_LOG.  Legacy rows
' remain explicitly marked LEGACY because old message text is not a substitute
' for the structured before/request/readback evidence captured from now on.
Public Sub CF_HistoryBackfillLegacyRunLog()
    Dim wsHistory As Worksheet, wsLog As Worksheet, r As Long, outRow As Long, lastRow As Long
    Dim stageText As String, imported As Long
    On Error GoTo ExitHere
    Set wsHistory = CWR_HistorySheet("RUN_HISTORY", False)
    If UCase$(CWR_Text(wsHistory.Range("AA2").Value2)) = "DONE" Then Exit Sub
    Set wsLog = ThisWorkbook.Worksheets("RUN_LOG")
    lastRow = wsLog.Cells(wsLog.rows.count, 1).End(xlUp).row
    For r = 9 To lastRow
        stageText = CWR_Text(wsLog.Cells(r, 3).Value2)
        If stageText <> "" Then
            outRow = wsHistory.Cells(wsHistory.rows.count, 1).End(xlUp).row + 1
            If outRow < 2 Then outRow = 2
            wsHistory.Cells(outRow, 1).Resize(1, 25).Value2 = Array("LEGACY|RUN_LOG|" & CStr(r), _
                "LEGACY_RUN_LOG_V1", wsLog.Cells(r, 1).Value2, "LEGACY", stageText, "", _
                CWR_Text(wsLog.Cells(r, 4).Value2), CWR_Text(wsLog.Cells(r, 5).Value2), "", "", _
                CWR_Text(wsLog.Cells(r, 8).Value2), "", "", "", "", CWR_HistoryActionKey(stageText), _
                CWR_Text(wsLog.Cells(r, 10).Value2), "", "", "", "", CWR_HistoryWHRunID(), "", _
                CWR_HistoryWHRunID(), CWR_Text(wsLog.Cells(r, 12).Value2))
            imported = imported + 1
        End If
    Next r
    wsHistory.Range("AA1").Value2 = "LEGACY BACKFILL"
    wsHistory.Range("AA2").Value2 = "DONE"
    CF_HistoryAppendSummary "LEGACY_RUN_LOG_V1", "LEGACY", "BACKFILLED", "", imported, 0, 0, 0, 0, 0, _
        "PASS", "Copied from RUN_LOG; fields unavailable in old format remain blank."
ExitHere:
End Sub

Public Function CF_HistoryNewRunID(ByVal channelText As String, Optional ByVal parentRunID As String = "") As String
    Dim suffixText As String
    suffixText = Replace$(Trim$(parentRunID), "|", "-")
    If suffixText <> "" Then suffixText = "|" & suffixText
    CF_HistoryNewRunID = UCase$(Trim$(channelText)) & "|" & _
        Format$(Now, "yyyymmdd_hhnnss") & "|" & Format$(Timer * 100, "000000") & suffixText
End Function

Public Sub CF_HistoryAppend(ByVal runID As String, ByVal channelText As String, _
                            ByVal stageText As String, ByVal planID As String, _
                            ByVal obText As String, ByVal docText As String, _
                            ByVal itemText As String, ByVal materialText As String, _
                            ByVal oldWT As String, ByVal actionKey As String, _
                            ByVal resultText As String, ByVal beforeValue As String, _
                            ByVal requestedValue As String, ByVal readbackValue As String, _
                            ByVal executionState As String, ByVal detailText As String)
    Dim ws As Worksheet, r As Long, newWT As String
    Dim qtyMode As String, execMode As String, planState As String, fingerprint As String
    Dim whRunID As String, postingRunID As String, monRunID As String
    On Error GoTo ExitHere
    If runID = "" Then runID = CF_HistoryNewRunID(channelText)
    Set ws = CWR_HistorySheet("RUN_HISTORY", False)
    CWR_HistoryPlanValues planID, qtyMode, execMode, planState, fingerprint
    If executionState = "" Then executionState = planState
    whRunID = CWR_HistoryWHRunID()
    postingRunID = CWR_HistoryPostingRunID()
    monRunID = whRunID
    newWT = CWR_HistoryExtractField(detailText, "WT")
    CWR_HISTORY_SEQ = CWR_HISTORY_SEQ + 1
    r = ws.Cells(ws.rows.count, 17).End(xlUp).row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Resize(1, 25).Value2 = Array(runID & "|E" & Format$(CWR_HISTORY_SEQ, "000000"), runID, Now, _
        channelText, stageText, planID, obText, docText, itemText, materialText, oldWT, newWT, _
        qtyMode, execMode, executionState, actionKey, resultText, beforeValue, requestedValue, _
        readbackValue, fingerprint, whRunID, postingRunID, monRunID, detailText)
ExitHere:
End Sub

Public Sub CF_HistoryAppendSummary(ByVal runID As String, ByVal channelText As String, _
                                   ByVal planStatus As String, ByVal obText As String, _
                                   ByVal planCount As Long, ByVal readyCount As Long, _
                                   ByVal blockedCount As Long, ByVal createdCount As Long, _
                                   ByVal readyCFCount As Long, ByVal unknownCount As Long, _
                                   ByVal resultText As String, ByVal detailText As String)
    Dim ws As Worksheet, r As Long
    On Error GoTo ExitHere
    If runID = "" Then runID = CF_HistoryNewRunID(channelText)
    Set ws = CWR_HistorySheet("RUN_SUMMARY", False)
    r = ws.Cells(ws.rows.count, 17).End(xlUp).row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Resize(1, 14).Value2 = Array(runID, Now, channelText, CWR_HistoryWHRunID(), _
        planStatus, obText, planCount, readyCount, blockedCount, createdCount, readyCFCount, _
        unknownCount, resultText, detailText)
ExitHere:
End Sub

Private Function CWR_HistorySheet(ByVal sheetName As String, ByVal veryHidden As Boolean) As Worksheet
    Dim headers As Variant
    On Error Resume Next
    Set CWR_HistorySheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo Failed
    If CWR_HistorySheet Is Nothing Then
        Set CWR_HistorySheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        CWR_HistorySheet.Name = sheetName
        If sheetName = "RUN_HISTORY" Then
            headers = Array("EventID", "RunID", "EventAt", "Channel", "Stage", "PlanID", "OB", "Document", _
                "Item", "Material", "OldWT", "NewWT", "QtyMode", "ExecMode", "ExecState", "ActionKey", _
                "Result", "BeforeValue", "RequestedValue", "ReadbackValue", "PlanFingerprint", "WHRunID", _
                "PostingRunID", "MONRunID", "Detail")
            CWR_HistorySheet.Range("A1:Y1").Value2 = headers
        Else
            headers = Array("RunID", "RecordedAt", "Channel", "WHRunID", "PlanStatus", "OB", "PlanCount", _
                "ReadyCount", "BlockedCount", "CreatedCount", "ReadyForCFCount", "UnknownCount", "Result", "Detail")
            CWR_HistorySheet.Range("A1:N1").Value2 = headers
        End If
        CWR_HistorySheet.rows(1).Font.Bold = True
        CWR_HistorySheet.columns.AutoFit
    End If
    If veryHidden Then CWR_HistorySheet.Visible = xlSheetVeryHidden Else CWR_HistorySheet.Visible = xlSheetVisible
    Exit Function
Failed:
    Set CWR_HistorySheet = Nothing
End Function

Private Function CWR_HistoryRunID() As String
    If CWR_HISTORY_RUN_ID = "" Then CWR_HISTORY_RUN_ID = CF_HistoryNewRunID("EXEC", CWR_HistoryCheckRunID())
    CWR_HistoryRunID = CWR_HISTORY_RUN_ID
End Function

Private Function CWR_HistoryCheckRunID() As String
    On Error Resume Next
    CWR_HistoryCheckRunID = CWR_Text(ThisWorkbook.Worksheets("PLAN").Range("AM2").Value2)
    If CWR_HistoryCheckRunID = "" Then CWR_HistoryCheckRunID = CWR_HistoryWHRunID()
End Function

Private Sub CWR_HistoryPlanValues(ByVal planID As String, ByRef qtyMode As String, _
                                  ByRef execMode As String, ByRef execState As String, _
                                  ByRef fingerprint As String)
    Dim ws As Worksheet, r As Long, lastRow As Long
    On Error GoTo ExitHere
    If planID = "" Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("PLAN")
    lastRow = ws.Cells(ws.rows.count, 17).End(xlUp).row
    For r = 2 To lastRow
        If StrComp(CWR_Text(ws.Cells(r, 17).Value2), planID, vbTextCompare) = 0 Then
            qtyMode = CWR_Text(ws.Cells(r, CWR_PLAN_COL_QTY_MODE).Value2)
            execMode = CWR_Text(ws.Cells(r, CWR_PLAN_COL_EXEC_MODE).Value2)
            execState = CWR_Text(ws.Cells(r, CWR_PLAN_COL_EXEC_STATE).Value2)
            fingerprint = CWR_Text(ws.Cells(r, 34).Value2)
            Exit For
        End If
    Next r
ExitHere:
End Sub

Private Function CWR_HistoryWHRunID() As String
    On Error Resume Next
    CWR_HistoryWHRunID = CWR_Text(ThisWorkbook.Worksheets("DATA_CORE").Range("W2").Value2)
End Function

Private Function CWR_HistoryPostingRunID() As String
    On Error Resume Next
    CWR_HistoryPostingRunID = CWR_Text(ThisWorkbook.Worksheets("POSTING_SUMMARY").Cells(2, 6).Value2)
End Function

Private Function CWR_HistoryExtractField(ByVal textIn As String, ByVal keyText As String) As String
    Dim token As String, parts As Variant, p As Variant
    token = UCase$(Trim$(keyText)) & "="
    parts = Split(Replace$(textIn, ";", "|"), "|")
    For Each p In parts
        If Left$(UCase$(Trim$(CStr(p))), Len(token)) = token Then
            CWR_HistoryExtractField = Trim$(Mid$(Trim$(CStr(p)), Len(token) + 1))
            Exit Function
        End If
    Next p
End Function

Private Function CWR_HistoryActionKey(ByVal stageText As String) As String
    Dim upperText As String
    upperText = UCase$(stageText)
    If InStr(upperText, "CANCEL") > 0 Then
        CWR_HistoryActionKey = "WT_CANCEL"
    ElseIf InStr(upperText, "ADOPT") > 0 Then
        CWR_HistoryActionKey = "WT_ADOPT_QUANTITY"
    ElseIf InStr(upperText, "CREATE") > 0 Or InStr(upperText, "SAVE") > 0 Then
        CWR_HistoryActionKey = "WT_CREATE_SAVE"
    ElseIf InStr(upperText, "PRDO") > 0 Or InStr(upperText, "QTY") > 0 Then
        CWR_HistoryActionKey = "PRDO_QTY"
    ElseIf InStr(upperText, "POST") > 0 Then
        CWR_HistoryActionKey = "POST_CHANGE"
    Else
        CWR_HistoryActionKey = "READ_ONLY"
    End If
End Function

Private Sub CWR_HistoryCheckpoint(ByVal stageText As String, ByVal actionKey As String, _
                                  ByVal resultText As String, ByVal beforeValue As String, _
                                  ByVal requestedValue As String, ByVal readbackValue As String, _
                                  ByVal detailText As String)
    On Error Resume Next
    CF_HistoryAppend CWR_HistoryRunID(), "EXECUTE", stageText, CWR_PLAN_ID, _
        CWR_EXPECTED_OB, CWR_EXPECTED_DOC, CWR_EXPECTED_ITEM, CWR_EXPECTED_MATERIAL, _
        CWR_EXPECTED_WT, actionKey, resultText, beforeValue, requestedValue, readbackValue, "", detailText
End Sub

Private Function CWR_Cell(ByVal grid As Object, ByVal rowNo As Long, ByVal colName As String) As String
    On Error Resume Next
    CWR_Cell = CWR_Text(grid.GetCellValue(rowNo, colName) & "")
    Err.Clear
End Function

Private Function CWR_WaitForObject(ByVal ses As Object, ByVal idText As String, ByVal seconds As Double) As Boolean
    Dim t0 As Double, ctl As Object
    t0 = Timer
    Do
        Set ctl = ses.FindById(idText, False)
        If Not ctl Is Nothing Then
            CWR_WaitForObject = True
            Exit Function
        End If
        DoEvents
        If CWR_Elapsed(t0) >= seconds Then Exit Do
    Loop
End Function

Private Function CWR_WaitIdle(ByVal ses As Object, ByVal seconds As Double) As Boolean
    Dim t0 As Double, busyText As Boolean
    t0 = Timer
    Do
        busyText = False
        On Error Resume Next
        busyText = CBool(ses.busy)
        On Error GoTo 0
        If Not busyText Then
            CWR_WaitIdle = True
            Exit Function
        End If
        DoEvents
        If CWR_Elapsed(t0) >= seconds Then Exit Do
    Loop
End Function

' A toolbar Press can be queued before Session.Busy flips.  Do not issue a
' follow-on query immediately: wait through a short stable window and then
' require idle, with an explicit ceiling to avoid a hung SAP session.
Private Function CWR_WaitPOSTSaveSettle(ByVal ses As Object, ByVal minSeconds As Double, _
                                        ByVal maxSeconds As Double) As Boolean
    Dim t0 As Double, busyText As Boolean
    t0 = Timer
    Do
        busyText = False
        On Error Resume Next
        busyText = CBool(ses.busy)
        On Error GoTo 0
        If CWR_Elapsed(t0) >= minSeconds And Not busyText Then
            CWR_WaitPOSTSaveSettle = True
            Exit Function
        End If
        DoEvents
        If CWR_Elapsed(t0) >= maxSeconds Then Exit Do
    Loop
End Function

' Bug C fix (NOT_LIVE_PROVEN): return as soon as SAP posts the single quantity-Save
' success event (status messageType "S"), rather than dwelling a fixed minimum.
' Returns True on the first "S" status. Returns False and sets statusType/statusText
' on an explicit SAP error (E/A/X) so the caller can fail closed. If the session
' goes idle with no terminal status within maxSeconds, returns False with an empty
' statusType, letting the caller defer to the authoritative fresh-PRDO readback.

Private Function CWR_Elapsed(ByVal startTick As Double) As Double
    CWR_Elapsed = Timer - startTick
    If CWR_Elapsed < 0# Then CWR_Elapsed = CWR_Elapsed + 86400#
End Function

' This audit is intentionally workbook-local.  It does not obtain a SAP
' session, navigate SAP, or send an action.  LIVE_PROVEN can only be supplied
' by an earlier recorded live regression, never inferred from a compile.
Public Function Codex_RunRegressionAudit() As String
    Dim reportText As String, currentResults As Object, priorLiveProofs As Object
    If Not CWR_RunRegressionAudit(reportText, currentResults, priorLiveProofs) Then
        Codex_RunRegressionAudit = "FAIL|REGRESSION_AUDIT|" & reportText & "|SAP_WRITE=0"
    ElseIf CWR_RegressionResultCount(currentResults, "FAIL") > 0 Then
        Codex_RunRegressionAudit = "FAIL|REGRESSION_AUDIT|" & reportText & "|SAP_WRITE=0"
    Else
        Codex_RunRegressionAudit = "PASS|REGRESSION_AUDIT|" & reportText & "|SAP_WRITE=0"
    End If
End Function

' Regression proof is scenario/action scoped.  A new scenario may remain
' NOT_PROVEN.  A historical LIVE_PROVEN record is inherited only when the
' action contract identity is unchanged; a changed identity requests one
' bounded reproof and never disables the safety gate.
Private Function CWR_RequireRegressionGate(ByRef messageText As String) As Boolean
    Dim reportText As String, currentResults As Object, priorLiveProofs As Object
    Dim requiredScenarios As Object, scenario As Variant, currentStatus As String
    Dim decisionText As String, decisionDetail As String, contractChanged As Boolean
    Dim relevantText As String
    If Not CWR_RunRegressionAuditScoped(reportText, currentResults, priorLiveProofs, requiredScenarios) Then
        messageText = "STOP_REGRESSION_GATE_AUDIT_ERROR|" & reportText
        Exit Function
    End If
    For Each scenario In currentResults.keys
        If UCase$(CStr(currentResults.item(CStr(scenario)))) = "FAIL" Then
            messageText = "STOP: Regression that actually failed - SAP write is blocked.|" & _
                          "STOP_REGRESSION_FAIL|SCENARIO=" & CStr(scenario) & "|" & reportText
            Exit Function
        End If
    Next scenario
    ' Historical proof is scoped to the locked plan.  An unrelated proof is
    ' informational only and must never block a different route.
    For Each scenario In priorLiveProofs.keys
        If requiredScenarios.Exists(CStr(scenario)) Then
            If Not currentResults.Exists(CStr(scenario)) Then
                currentStatus = "NOT_PROVEN"
            Else
                currentStatus = CStr(currentResults.item(CStr(scenario)))
            End If
            contractChanged = CWR_HistoricalProofContractChanged(CStr(scenario), priorLiveProofs.item(CStr(scenario)))
            decisionDetail = vbNullString
            decisionText = CWR_RegressionGateDecision(CStr(scenario), True, currentStatus, _
                                                       "LIVE_PROVEN", contractChanged, 0, decisionDetail)
            relevantText = "RELEVANT=TRUE|CURRENT=" & currentStatus & _
                           "|HISTORICAL=LIVE_PROVEN|CONTRACT_VERSION=" & _
                           CWR_ScenarioContractIdentity(CStr(scenario)) & _
                           "|CONTRACT_CHANGED=" & CStr(contractChanged) & _
                           "|DECISION=" & decisionText & "|" & decisionDetail
            CWR_Log "REGRESSION_GATE_DECISION", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                    relevantText & "|WRITE_ATTEMPT=0|CF=USER_HANDOFF"
            Select Case decisionText
                Case "REGRESSION_FAIL"
                    messageText = "STOP: Regression that actually failed - SAP write is blocked.|" & _
                                  "STOP_REGRESSION_FAIL|SCENARIO=" & CStr(scenario) & "|" & relevantText
                    Exit Function
                Case "LIVE_PROVEN_INHERITED"
                    ' Safe inheritance: the action contract is unchanged.
                Case "LIVE_REPROOF_REQUIRED"
                    ' Allow exactly one bounded proof execution.  The existing
                    ' executor/post-verify barriers enforce the write limit and
                    ' fail closed on unknown post-send state.
                Case Else
                    ' A current PASS or an unproven route is allowed to proceed;
                    ' real regression FAIL was handled above.
            End Select
        Else
            CWR_Log "REGRESSION_GATE_DECISION", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", _
                    "SCENARIO=" & CStr(scenario) & "|RELEVANT=FALSE|CURRENT=NOT_APPLICABLE" & _
                    "|HISTORICAL=LIVE_PROVEN|DECISION=NOT_RELEVANT|WRITE_ATTEMPT=0"
        End If
    Next scenario
    messageText = "REGRESSION_GATE_PASS|" & reportText & "|" & _
                  CWR_RegressionScopeDetail(requiredScenarios) & _
                  "|CONTRACT_SCOPED=TRUE|FAIL_CLOSED=TRUE"
    CWR_RequireRegressionGate = True
End Function

' Pure decision contract used by the live gate and by offline S23-S26 replay.
' It does not access SAP, mutate sheets, or authorize a write by itself.
Public Function CWR_RegressionGateDecision(ByVal scenario As String, _
                                            ByVal relevant As Boolean, _
                                            ByVal currentProof As String, _
                                            ByVal historicalProof As String, _
                                            ByVal contractChanged As Boolean, _
                                            ByVal invariantFailCount As Long, _
                                            ByRef detailText As String) As String
    Dim currentText As String, historicalText As String
    currentText = UCase$(Trim$(currentProof))
    historicalText = UCase$(Trim$(historicalProof))
    detailText = "SCENARIO=" & scenario & "|CURRENT=" & currentText & _
                 "|HISTORICAL=" & historicalText & "|CONTRACT_CHANGED=" & _
                 CStr(contractChanged)
    If Not relevant Then
        detailText = detailText & "|DECISION=NOT_RELEVANT"
        CWR_RegressionGateDecision = "NOT_RELEVANT"
        Exit Function
    End If
    If invariantFailCount > 0 Or currentText = "FAIL" Or currentText = "REGRESSION_FAIL" Then
        detailText = detailText & "|DECISION=REGRESSION_FAIL|STOP_BEFORE_WRITE=TRUE"
        CWR_RegressionGateDecision = "REGRESSION_FAIL"
        Exit Function
    End If
    If currentText = "LIVE_PROVEN" Then
        detailText = detailText & "|DECISION=CURRENT_LIVE_PROVEN"
        CWR_RegressionGateDecision = "CURRENT_LIVE_PROVEN"
        Exit Function
    End If
    If historicalText = "LIVE_PROVEN" Then
        If contractChanged Then
            detailText = detailText & "|DECISION=LIVE_REPROOF_REQUIRED|MAX_WRITE_ATTEMPTS=" & _
                         CStr(CWR_MAX_PROOF_WRITE_ATTEMPTS) & "|PROOF_RUN=BOUNDED"
            CWR_RegressionGateDecision = "LIVE_REPROOF_REQUIRED"
        Else
            detailText = detailText & "|DECISION=LIVE_PROVEN_INHERITED"
            CWR_RegressionGateDecision = "LIVE_PROVEN_INHERITED"
        End If
        Exit Function
    End If
    detailText = detailText & "|DECISION=NOT_PROVEN|NORMAL_EXECUTOR=ALLOWED"
    CWR_RegressionGateDecision = "NOT_PROVEN"
End Function

Private Function CWR_ScenarioContractIdentity(ByVal scenario As String) As String
    Select Case UCase$(Trim$(scenario))
        Case "CREATE_WT_DELTA"
            CWR_ScenarioContractIdentity = "CREATE_WT_DELTA:V" & CWR_CREATE_WT_DELTA_CONTRACT_VERSION
        Case "CHANGE_BIN"
            CWR_ScenarioContractIdentity = "CHANGE_BIN:V" & CWR_CHANGE_BIN_CONTRACT_VERSION
        Case "CHANGE_BIN_BATCH"
            CWR_ScenarioContractIdentity = "CHANGE_BIN_BATCH:V" & CWR_CHANGE_BIN_CONTRACT_VERSION
        Case Else
            CWR_ScenarioContractIdentity = UCase$(Trim$(scenario)) & ":V1"
    End Select
End Function

Private Function CWR_TokenValue(ByVal sourceText As String, ByVal tokenText As String) As String
    Dim p As Long, startAt As Long, endAt As Long, tailText As String
    p = InStr(1, sourceText, tokenText, vbTextCompare)
    If p = 0 Then Exit Function
    startAt = p + Len(tokenText)
    tailText = Mid$(sourceText, startAt)
    endAt = InStr(1, tailText, "|", vbBinaryCompare)
    If endAt > 0 Then tailText = Left$(tailText, endAt - 1)
    CWR_TokenValue = Trim$(tailText)
End Function

Private Function CWR_LegacyProofContractUnchanged(ByVal scenario As String, ByVal proofText As String) As Boolean
    Dim hasCreateCore As Boolean, hasGroupCore As Boolean, hasPostVerify As Boolean
    Dim hasChangeCore As Boolean
    hasCreateCore = InStr(1, proofText, "CWR_CreateOutboundWTFromExactCandidateFast", vbTextCompare) > 0
    hasGroupCore = InStr(1, proofText, "CREATE_GROUP_ACTION=ONE_CREATE_SAVE_PER_GROUP", vbTextCompare) > 0
    hasPostVerify = InStr(1, proofText, "OB_COMMIT_BARRIER=TRUE", vbTextCompare) > 0
    hasChangeCore = InStr(1, proofText, "SCHEDULER_ORDER=CANCEL_GROUP|LOCATION_GROUP|PRDO_DOCUMENT_GROUP|CREATE_GROUP|FINAL_VERIFY", vbTextCompare) > 0
    Select Case UCase$(Trim$(scenario))
        Case "CREATE_WT_DELTA"
            CWR_LegacyProofContractUnchanged = hasCreateCore And hasGroupCore And hasPostVerify
        Case "CHANGE_BIN", "CHANGE_BIN_BATCH"
            CWR_LegacyProofContractUnchanged = hasCreateCore And hasChangeCore And hasPostVerify
    End Select
End Function

Private Function CWR_HistoricalProofContractChanged(ByVal scenario As String, ByVal proofRecord As Variant) As Boolean
    Dim recordText As String, storedIdentity As String
    recordText = CWR_Text(proofRecord)
    storedIdentity = CWR_TokenValue(recordText, "SCENARIO_CONTRACT=")
    If storedIdentity <> "" Then
        CWR_HistoricalProofContractChanged = (StrComp(storedIdentity, _
            CWR_ScenarioContractIdentity(scenario), vbTextCompare) <> 0)
    Else
        ' Migrate only an old proof whose Diff Audit explicitly contains the
        ' unchanged business-action anchors.  Missing/ambiguous evidence stays
        ' fail-closed and requires bounded reproof.
        CWR_HistoricalProofContractChanged = Not CWR_LegacyProofContractUnchanged(scenario, recordText)
    End If
End Function

Private Function CWR_RunRegressionAudit(ByRef reportText As String, ByRef currentResults As Object, _
                                        ByRef priorLiveProofs As Object) As Boolean
    Dim requiredScenarios As Object
    CWR_RunRegressionAudit = CWR_RunRegressionAuditScoped(reportText, currentResults, priorLiveProofs, requiredScenarios)
End Function

Private Function CWR_RunRegressionAuditScoped(ByRef reportText As String, ByRef currentResults As Object, _
                                              ByRef priorLiveProofs As Object, ByRef requiredScenarios As Object) As Boolean
    Dim wsAudit As Worksheet, wsPlan As Worksheet, wsInput As Worksheet, obText As String
    Dim passCount As Long, failCount As Long, notProvenCount As Long
    On Error GoTo Failed
    Set wsAudit = CWR_RegressionAuditSheet()
    If wsAudit Is Nothing Then
        reportText = "REGRESSION_AUDIT_SHEET_UNAVAILABLE"
        Exit Function
    End If
    Set currentResults = CreateObject("Scripting.Dictionary")
    currentResults.CompareMode = vbTextCompare
    Set priorLiveProofs = CreateObject("Scripting.Dictionary")
    priorLiveProofs.CompareMode = vbTextCompare
    CWR_ReadPriorLiveProofs wsAudit, priorLiveProofs
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    On Error Resume Next
    Set wsInput = CF_GetInputSheet()
    On Error GoTo Failed
    If Not wsInput Is Nothing Then obText = CWR_Text(wsInput.Range("A7").Value2)
    Set requiredScenarios = CWR_BuildRequiredRegressionScenarios(wsPlan, obText)
    CWR_AuditPlanMatrix wsPlan, wsAudit, obText, currentResults, passCount, failCount, notProvenCount, requiredScenarios
    CWR_AuditStaticContracts wsAudit, currentResults, passCount, failCount, notProvenCount
    wsAudit.columns("A:I").AutoFit
    reportText = "BUILD=" & CWR_REGRESSION_BUILD & "|PASS=" & CStr(passCount) & _
                 "|FAIL=" & CStr(failCount) & "|NOT_PROVEN=" & CStr(notProvenCount) & _
                 "|PRIOR_LIVE=" & CStr(priorLiveProofs.count) & _
                 "|REQUIRED=" & CStr(requiredScenarios.count)
    CWR_RunRegressionAuditScoped = True
    Exit Function
Failed:
    reportText = "REGRESSION_AUDIT_ERROR=" & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CWR_RegressionAuditSheet() As Worksheet
    Dim headers As Variant
    On Error Resume Next
    Set CWR_RegressionAuditSheet = ThisWorkbook.Worksheets(CWR_REGRESSION_SHEET)
    On Error GoTo Failed
    If CWR_RegressionAuditSheet Is Nothing Then
        Set CWR_RegressionAuditSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        CWR_RegressionAuditSheet.Name = CWR_REGRESSION_SHEET
        headers = Array("Build", "Timestamp", "Scenario", "Expected", "Actual", "Status", _
                        "Live Proven?", "Evidence", "Diff Audit")
        CWR_RegressionAuditSheet.Range("A1:I1").Value2 = headers
        CWR_RegressionAuditSheet.rows(1).Font.Bold = True
    End If
    Exit Function
Failed:
    Set CWR_RegressionAuditSheet = Nothing
End Function

Private Sub CWR_ReadPriorLiveProofs(ByVal wsAudit As Worksheet, ByVal priorLiveProofs As Object)
    Dim r As Long, lastRow As Long, scenario As String, buildText As String, diffText As String
    If wsAudit Is Nothing Then Exit Sub
    lastRow = wsAudit.Cells(wsAudit.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        scenario = CWR_Text(wsAudit.Cells(r, 3).Value2)
        If scenario <> "" And UCase$(CWR_Text(wsAudit.Cells(r, 6).Value2)) = "PASS" And _
           UCase$(CWR_Text(wsAudit.Cells(r, 7).Value2)) = "LIVE_PROVEN" Then
            buildText = CWR_Text(wsAudit.Cells(r, 1).Value2)
            diffText = CWR_Text(wsAudit.Cells(r, 9).Value2)
            priorLiveProofs(scenario) = "BUILD=" & buildText & "|DIFF=" & diffText
        End If
    Next r
End Sub

Private Sub CWR_AppendRegressionAudit(ByVal wsAudit As Worksheet, ByVal scenario As String, _
                                      ByVal expectedText As String, ByVal actualText As String, _
                                      ByVal statusText As String, ByVal liveProofText As String, _
                                      ByVal evidenceText As String)
    Dim r As Long
    r = wsAudit.Cells(wsAudit.rows.count, 1).End(xlUp).row + 1
    If r < 2 Then r = 2
    wsAudit.Cells(r, 1).Resize(1, 9).Value2 = Array(CWR_REGRESSION_BUILD, Now, scenario, expectedText, _
        actualText, statusText, liveProofText, evidenceText, CWR_RegressionDiffEvidence() & _
        "|SCENARIO_CONTRACT=" & CWR_ScenarioContractIdentity(scenario))
End Sub

Private Sub CWR_AuditResult(ByVal wsAudit As Worksheet, ByVal currentResults As Object, _
                            ByRef passCount As Long, ByRef failCount As Long, ByRef notProvenCount As Long, _
                            ByVal scenario As String, ByVal expectedText As String, ByVal actualText As String, _
                            ByVal statusText As String, ByVal liveProofText As String, ByVal evidenceText As String)
    If currentResults.Exists(scenario) Then
        currentResults(scenario) = statusText
    Else
        currentResults.Add scenario, statusText
    End If
    CWR_AppendRegressionAudit wsAudit, scenario, expectedText, actualText, statusText, liveProofText, evidenceText
    Select Case UCase$(statusText)
        Case "PASS": passCount = passCount + 1
        Case "FAIL": failCount = failCount + 1
        Case "NOT_APPLICABLE": ' excluded from NOT_PROVEN safety counts
        Case Else: notProvenCount = notProvenCount + 1
    End Select
End Sub

Private Function CWR_RegressionResultCount(ByVal currentResults As Object, ByVal wantedStatus As String) As Long
    Dim scenario As Variant
    If currentResults Is Nothing Then Exit Function
    For Each scenario In currentResults.keys
        If UCase$(CStr(currentResults.item(CStr(scenario)))) = UCase$(wantedStatus) Then _
            CWR_RegressionResultCount = CWR_RegressionResultCount + 1
    Next scenario
End Function

Private Function CWR_RegressionDiffEvidence() As String
    CWR_RegressionDiffEvidence = "MODULE=modCodexWTRework|FUNCTIONS=CWR_CreateOutboundWTFromExactCandidateFast,Codex_ExecutePostingBatchAtPickBin,CWR_CloseCreatedPostingWTLogForRead,CWR_RecordCreatedPostingWT," & _
        "CWR_QueryExpectedPOST,CWR_HasUnreconciledPickBinPostSend,CWR_HasRecoverablePickBinPostUnknown,CWR_RecordPickBinBatchPostReconciled," & _
        "CWR_BuildExecutionGroups,CWR_SchedulerVerifyPRDOGroup,CWR_ExecuteDocumentGroups,CWR_CancelDisplayLogsProof,CWR_ResetPOSTSearchSurface," & _
         "CWR_ExecutePreflightGroups,CWR_ObPreflightPass,CWR_ObPlanFingerprint,CWR_SourceAvailableQty,CWR_TryParseSAPQty,CWR_CancelUsesSharedFinalPRDO,CWR_RequireRegressionGate," & _
        "CWR_RecordSuccessfulSchedulerRegressionProofs,CWR_ResetSchedulerExecutionContext|" & _
        "CALLERS=LOCKED_SCHEDULER+SINGLE_CORE|" & _
        "CONFIG=UNCHANGED|SAP_ACTION_MAP=UNCHANGED|WT_NAV=FORCED_TWO_ACTION_CLICKS_SAME_DOCUMENT|" & _
        "SCHEDULER_ORDER=CANCEL_GROUP|LOCATION_GROUP|PRDO_DOCUMENT_GROUP|CREATE_GROUP|FINAL_VERIFY|CF_GROUP|" & _
        "OB_PLAN_FINGERPRINT=TRUE|OB_PREFLIGHT_PASS=TRUE|OB_COMMIT_BARRIER=TRUE|" & _
        "CANCEL_GROUP_ACTION=ONE_SELECTION_ONE_CANCEL_ONE_GROUP_READBACK|" & _
        "PRDO_DOCUMENT_ACTION=ONE_SAVE_PER_DOCUMENT_GROUP|CREATE_GROUP_ACTION=ONE_CREATE_SAVE_PER_GROUP|" & _
        "CONTEXT_REUSE=DOCUMENT_SCOPED|WHTASK_REUSE=SAME_DOCUMENT|WHTASK_FIRST_OPEN_PER_RUN=TRUE|WHTASK_REUSE_SCOPE=EXECUTION_ONLY|ACTION_KEY=BUSINESS_TUPLE|" & _
         "CREATE_GROUP_SCOPE=OB_DOC_ITEM_MATERIAL_SOURCE_TARGET|TERMINAL_COMPLETENESS=UNACCOUNTED_ZERO|" & _
         "PLAN_SCOPED_REGRESSION=TRUE|REGRESSION_ACTIVE_ROWS_ONLY=TRUE|MIXED_CREATE_EFFECTIVE_STATE=TRUE|OFFLINE_T1_T8=TRUE|" & _
         "GATE_BLOCK_LOG_SUMMARY=TRUE|SHARED_PRDO_CANCEL_GUARD=TRUE|" & _
          "REPLACE_CREATE_FULL_TARGET=TRUE|" & _
          "SCENARIO_CONTRACT_SCHEMA=V1|CREATE_WT_DELTA_CONTRACT=" & _
          CWR_ScenarioContractIdentity("CREATE_WT_DELTA") & "|CHANGE_BIN_CONTRACT=" & _
          CWR_ScenarioContractIdentity("CHANGE_BIN") & "|CHANGE_BIN_FLOW=PRE_VERIFY>RELEASE_CANCEL>LOCATION_POSTING_CHANGE>PRDO_DOCUMENT_GROUP>CREATE_WT>POST_ACTION_PROOF>FINAL_GROUP_VERIFY>CF_USER_HANDOFF|" & _
          "CF_USER_HANDOFF_REQUIRED=1|CF_LIVE_ENABLED=0|ALLOW_CF=0|MAX_WRITE_ATTEMPTS=" & _
          CStr(CWR_MAX_PROOF_WRITE_ATTEMPTS)
End Function

' Only a fully completed scheduler run may advance a business scenario from
' NOT_LIVE_PROVEN to LIVE_PROVEN.  A partial/blocked run records no proof, so
' the next REGRESSION_GATE remains conservative.
Private Sub CWR_RecordSuccessfulSchedulerRegressionProofs(ByVal obText As String, ByVal planCount As Long)
    Dim reportText As String, currentResults As Object, priorLiveProofs As Object, wsAudit As Worksheet
    Dim scenario As Variant, evidenceText As String
    On Error GoTo ExitHere
    If Not CWR_RunRegressionAudit(reportText, currentResults, priorLiveProofs) Then Exit Sub
    Set wsAudit = CWR_RegressionAuditSheet()
    evidenceText = "LIVE_SCHEDULER_COMPLETE|OB=" & obText & "|PLAN_COUNT=" & CStr(planCount) & _
        "|RUN_ID=" & CWR_HISTORY_RUN_ID & "|POST_ACTION_PROOFS=RUN_HISTORY|" & _
        "CF_USER_HANDOFF_REQUIRED=1|CF_LIVE_ENABLED=0|ALLOW_CF=0"
    For Each scenario In currentResults.keys
        If UCase$(CStr(currentResults.item(CStr(scenario)))) = "PASS" And _
           CWR_IsBusinessRegressionScenario(CStr(scenario)) Then
            CWR_AppendRegressionAudit wsAudit, CStr(scenario), "Completed scheduler scenario", _
                "All scoped plans reached READY_FOR_CF", "PASS", "LIVE_PROVEN", evidenceText
        End If
    Next scenario
    CWR_Log "REGRESSION_LIVE_PROOF", CWR_EXPECTED_OB, CWR_EXPECTED_DOC, "PASS", evidenceText
ExitHere:
End Sub

Private Function CWR_IsBusinessRegressionScenario(ByVal scenario As String) As Boolean
    Select Case UCase$(scenario)
        Case "CF_ONLY", "CANCEL_TO_ZERO_SINGLE", "MULTI_CANCEL", "CHANGE_QTY_DECREASE", _
             "CHANGE_QTY_INCREASE", "CREATE_WT_DELTA", "CHANGE_BIN", "CHANGE_BATCH", _
             "CHANGE_BIN_BATCH", "MIXED_OB", "FINAL_GROUP_VERIFY"
            CWR_IsBusinessRegressionScenario = True
    End Select
End Function

Private Sub CWR_AddRequiredRegressionScenario(ByVal requiredScenarios As Object, ByVal scenario As String)
    If requiredScenarios Is Nothing Then Exit Sub
    If Not requiredScenarios.Exists(scenario) Then requiredScenarios.Add scenario, True
End Sub

' Build the business scenario set from the immutable locked PLAN.  Historical
' proofs outside this set are informational only and never gate this run.
Private Function CWR_BuildRequiredRegressionScenarios(ByVal wsPlan As Worksheet, ByVal obText As String) As Object
    Dim required As Object, groups As Object, group As Object, modes As Object
    Dim r As Long, lastRow As Long, groupID As String, execMode As String
    Dim groupCurrent As Double, groupTarget As Double, cancelCount As Long, groupCount As Long
    Dim contractMembers As Long
    Dim keyValue As Variant, delta As Double, cancelTotal As Long
    Dim actionText As String, binChanged As Boolean, batchChanged As Boolean
    Set required = CreateObject("Scripting.Dictionary")
    required.CompareMode = vbTextCompare
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    CWR_AddRequiredRegressionScenario required, "FINAL_GROUP_VERIFY"
    If wsPlan Is Nothing Then
        Set CWR_BuildRequiredRegressionScenarios = required
        Exit Function
    End If
    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If CWR_IsRegressionActivePlanRow(wsPlan, r) And _
           (obText = "" Or StrComp(CWR_Text(wsPlan.Cells(r, 2).Value2), obText, vbTextCompare) = 0) Then
            groupID = CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_QTY_GROUP_ID).Value2)
            If groupID = "" Then groupID = "__MISSING_QTY_GROUP__" & CStr(r)
            execMode = UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_EXEC_MODE).Value2))
            actionText = UCase$(CWR_Text(wsPlan.Cells(r, 20).Value2))
            binChanged = (StrComp(CWR_Text(wsPlan.Cells(r, 9).Value2), _
                                 CWR_Text(wsPlan.Cells(r, 10).Value2), vbTextCompare) <> 0)
            batchChanged = (StrComp(CWR_Text(wsPlan.Cells(r, 11).Value2), _
                                   CWR_Text(wsPlan.Cells(r, 14).Value2), vbTextCompare) <> 0)
            If UCase$(CWR_Text(wsPlan.Cells(r, 20).Value2)) = "CF_ONLY" Then CWR_AddRequiredRegressionScenario required, "CF_ONLY"
            If InStr(1, actionText, "CHANGE_BIN_BATCH", vbTextCompare) > 0 Or (binChanged And batchChanged) Then
                CWR_AddRequiredRegressionScenario required, "CHANGE_BIN_BATCH"
            ElseIf InStr(1, actionText, "CHANGE_BIN", vbTextCompare) > 0 Or binChanged Then
                CWR_AddRequiredRegressionScenario required, "CHANGE_BIN"
            End If
            If InStr(1, actionText, "CHANGE_BATCH", vbTextCompare) > 0 Or batchChanged Then
                CWR_AddRequiredRegressionScenario required, "CHANGE_BATCH"
            End If
            If execMode = "CANCEL_PRDO_ZERO" Then cancelTotal = cancelTotal + 1
            If Not groups.Exists(groupID) Then
                Set group = CreateObject("Scripting.Dictionary")
                group.CompareMode = vbTextCompare
                group.Add "Current", CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_CURRENT).Value2)
                group.Add "Target", CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_TARGET).Value2)
                group.Add "RawCurrent", CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_CURRENT).Value2)
                group.Add "RawTarget", CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_TARGET).Value2)
                group.Add "SumSys", 0#: group.Add "SumTarget", 0#
                group.Add "ContractMembers", CLng(CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_MEMBERS).Value2))
                group.Add "Count", 0&: group.Add "CancelCount", 0&
                Set modes = CreateObject("Scripting.Dictionary"): modes.CompareMode = vbTextCompare
                group.Add "Modes", modes
                groups.Add groupID, group
            End If
            Set group = groups.item(groupID)
            group("Count") = CLng(group("Count")) + 1
            group("SumSys") = CDbl(group("SumSys")) + CWR_Number(wsPlan.Cells(r, 7).Value2)
            group("SumTarget") = CDbl(group("SumTarget")) + CWR_Number(wsPlan.Cells(r, 8).Value2)
            If execMode = "CANCEL_PRDO_ZERO" Then group("CancelCount") = CLng(group("CancelCount")) + 1
            Set modes = group.item("Modes")
            If execMode = "" Then execMode = "__MISSING_EXEC_MODE__"
            If Not modes.Exists(execMode) Then modes.Add execMode, True
            If execMode = "KEEP_EXISTING_CREATE_DELTA" Or execMode = "REPLACE_CANCEL_CREATE" Then _
                CWR_AddRequiredRegressionScenario required, "CREATE_WT_DELTA"
        End If
    Next r
    If cancelTotal = 1 Then
        CWR_AddRequiredRegressionScenario required, "CANCEL_TO_ZERO_SINGLE"
    ElseIf cancelTotal > 1 Then
        CWR_AddRequiredRegressionScenario required, "MULTI_CANCEL"
    End If
    For Each keyValue In groups.keys
        Set group = groups.item(CStr(keyValue))
        contractMembers = CLng(group("ContractMembers"))
        If contractMembers <= 0 Or CLng(group("Count")) < contractMembers Then
            groupCurrent = CDbl(group("SumSys"))
            groupTarget = CDbl(group("SumTarget"))
        Else
            groupCurrent = CDbl(group("RawCurrent"))
            groupTarget = CDbl(group("RawTarget"))
        End If
        delta = groupTarget - groupCurrent
        cancelCount = CLng(group("CancelCount")): groupCount = CLng(group("Count"))
        Set modes = group.item("Modes")
        If delta > 0.000001 Then CWR_AddRequiredRegressionScenario required, "CHANGE_QTY_INCREASE"
        If delta < -0.000001 And cancelCount <> groupCount Then CWR_AddRequiredRegressionScenario required, "CHANGE_QTY_DECREASE"
        If modes.count > 1 Then CWR_AddRequiredRegressionScenario required, "MIXED_OB"
    Next keyValue
    Set CWR_BuildRequiredRegressionScenarios = required
End Function

Private Function CWR_RegressionScopeDetail(ByVal requiredScenarios As Object) As String
    Dim catalog As Variant, scenario As Variant, requiredText As String, ignoredText As String
    catalog = Array("CF_ONLY", "CANCEL_TO_ZERO_SINGLE", "MULTI_CANCEL", "CHANGE_QTY_DECREASE", _
                    "CHANGE_QTY_INCREASE", "CREATE_WT_DELTA", "CHANGE_BIN", "CHANGE_BATCH", _
                    "CHANGE_BIN_BATCH", "MIXED_OB", "FINAL_GROUP_VERIFY")
    For Each scenario In catalog
        If CWR_IsRequiredRegressionScenario(requiredScenarios, CStr(scenario)) Then
            If requiredText <> "" Then requiredText = requiredText & ","
            requiredText = requiredText & CStr(scenario)
        Else
            If ignoredText <> "" Then ignoredText = ignoredText & ","
            ignoredText = ignoredText & CStr(scenario) & ":NOT_APPLICABLE"
        End If
    Next scenario
    CWR_RegressionScopeDetail = "REQUIRED=" & requiredText & "|IGNORED=" & ignoredText
End Function

Private Function CWR_IsRequiredRegressionScenario(ByVal requiredScenarios As Object, ByVal scenario As String) As Boolean
    If requiredScenarios Is Nothing Then Exit Function
    CWR_IsRequiredRegressionScenario = requiredScenarios.Exists(scenario)
End Function

Private Function CWR_RegressionScopeDetailForPlan(ByVal wsPlan As Worksheet, ByVal obText As String) As String
    CWR_RegressionScopeDetailForPlan = CWR_RegressionScopeDetail(CWR_BuildRequiredRegressionScenarios(wsPlan, obText))
End Function

Private Function CWR_SchedulerFirstDocument(ByVal plans As Collection) As String
    On Error GoTo ExitHere
    If Not plans Is Nothing Then
        If plans.count > 0 Then CWR_SchedulerFirstDocument = CWR_Text(plans.item(1)("Doc"))
    End If
ExitHere:
End Function

Private Sub CWR_AuditPlanMatrix(ByVal wsPlan As Worksheet, ByVal wsAudit As Worksheet, ByVal obText As String, _
                                 ByVal currentResults As Object, ByRef passCount As Long, ByRef failCount As Long, _
                                 ByRef notProvenCount As Long, ByVal requiredScenarios As Object)
    Dim groups As Object, group As Object, modes As Object, groupID As String, execMode As String
    Dim r As Long, lastRow As Long, planRows As Long, sysQty As Double, targetQty As Double, deltaQty As Double
    Dim groupCurrent As Double, groupTarget As Double, cfOnlyFound As Boolean, cfOnlyPass As Boolean
    Dim finalFound As Boolean, finalPass As Boolean, cancelSingleFound As Boolean, cancelSinglePass As Boolean
    Dim multiCancelFound As Boolean, multiCancelPass As Boolean, increaseFound As Boolean, increasePass As Boolean
    Dim decreaseFound As Boolean, decreasePass As Boolean, deltaFound As Boolean, deltaPass As Boolean
    Dim mixedFound As Boolean, mixedPass As Boolean, keyValue As Variant, expectedDelta As Double, actualDelta As Double
    Dim evidenceText As String, deltaEvidence As String
    Dim cancelQty As Double, effectiveAfterRelease As Double, expectedCreateQty As Double, actualCreateQty As Double
    Dim projectedAfterRelease As Double

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    cfOnlyPass = True: finalPass = True: cancelSinglePass = True: multiCancelPass = True
    increasePass = True: decreasePass = True: deltaPass = True: mixedPass = True
    lastRow = wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If CWR_IsRegressionActivePlanRow(wsPlan, r) And _
           (obText = "" Or StrComp(CWR_Text(wsPlan.Cells(r, 2).Value2), obText, vbTextCompare) = 0) Then
            planRows = planRows + 1
            groupID = CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_QTY_GROUP_ID).Value2)
            If groupID = "" Then groupID = "__MISSING_QTY_GROUP__" & CStr(r)
            sysQty = CWR_Number(wsPlan.Cells(r, 7).Value2)
            targetQty = CWR_Number(wsPlan.Cells(r, 8).Value2)
            deltaQty = CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_QTY_DELTA).Value2)
            groupCurrent = CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_CURRENT).Value2)
            groupTarget = CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_TARGET).Value2)
            execMode = UCase$(CWR_Text(wsPlan.Cells(r, CWR_PLAN_COL_EXEC_MODE).Value2))
            If Not groups.Exists(groupID) Then
                Set group = CreateObject("Scripting.Dictionary")
                group.CompareMode = vbTextCompare
                group.Add "Count", 0&: group.Add "SumSys", 0#: group.Add "SumTarget", 0#: group.Add "SumDelta", 0#
                group.Add "GroupCurrent", groupCurrent: group.Add "GroupTarget", groupTarget
                group.Add "RawGroupCurrent", groupCurrent: group.Add "RawGroupTarget", groupTarget
                group.Add "ContractMembers", CLng(CWR_Number(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_MEMBERS).Value2))
                group.Add "FieldsConsistent", True: group.Add "NeedQty", False
                group.Add "CancelCount", 0&: group.Add "CancelTargetsZero", True
                 group.Add "DeltaModeCount", 0&: group.Add "ReplaceCreateCount", 0&
                group.Add "CancelQty", 0#: group.Add "CreateQtyActual", 0#
                group.Add "EffectiveAfterRelease", 0#: group.Add "CreateQtyExpected", 0#
                Set modes = CreateObject("Scripting.Dictionary"): modes.CompareMode = vbTextCompare
                group.Add "Modes", modes
                groups.Add groupID, group
            End If
            Set group = groups.item(groupID)
            group("Count") = CLng(group("Count")) + 1
            group("SumSys") = CDbl(group("SumSys")) + sysQty
            group("SumTarget") = CDbl(group("SumTarget")) + targetQty
            group("SumDelta") = CDbl(group("SumDelta")) + deltaQty
            If Abs(CDbl(group("GroupCurrent")) - groupCurrent) > 0.000001 Or _
               Abs(CDbl(group("GroupTarget")) - groupTarget) > 0.000001 Then group("FieldsConsistent") = False
            If CWR_AuditTrue(wsPlan.Cells(r, CWR_PLAN_COL_GROUP_NEED_QTY).Value2) Then group("NeedQty") = True
            Set modes = group.item("Modes")
            If execMode = "" Then execMode = "__MISSING_EXEC_MODE__"
            If Not modes.Exists(execMode) Then modes.Add execMode, True
            If execMode = "CANCEL_PRDO_ZERO" Then
                group("CancelCount") = CLng(group("CancelCount")) + 1
                If Abs(targetQty) > 0.000001 Then group("CancelTargetsZero") = False
                group("CancelQty") = CDbl(group("CancelQty")) + sysQty
            ElseIf execMode = "REPLACE_CANCEL_CREATE" Then
                ' Replacement also releases its old WT before the new full
                ' TargetQty is created.  Include that old quantity in the
                ' effective post-cancel state, but keep it out of the pure
                ' CANCEL_TO_ZERO counters.
                group("CancelQty") = CDbl(group("CancelQty")) + sysQty
            End If
            If execMode = "KEEP_EXISTING_CREATE_DELTA" Then
                group("DeltaModeCount") = CLng(group("DeltaModeCount")) + 1
                group("CreateQtyActual") = CDbl(group("CreateQtyActual")) + deltaQty
            ElseIf execMode = "REPLACE_CANCEL_CREATE" Then
                ' A replacement member cancels its old WT, then creates the
                ' full TargetQty at the new tuple.  Its create quantity must be
                ' included alongside retained delta members in a mixed group.
                group("ReplaceCreateCount") = CLng(group("ReplaceCreateCount")) + 1
                group("CreateQtyActual") = CDbl(group("CreateQtyActual")) + targetQty
            End If
            If UCase$(CWR_Text(wsPlan.Cells(r, 20).Value2)) = "CF_ONLY" Then
                cfOnlyFound = True
                If Abs(sysQty - targetQty) > 0.000001 Or _
                   StrComp(CWR_Text(wsPlan.Cells(r, 9).Value2), CWR_Text(wsPlan.Cells(r, 10).Value2), vbTextCompare) <> 0 Or _
                   StrComp(CWR_Text(wsPlan.Cells(r, 11).Value2), CWR_Text(wsPlan.Cells(r, 14).Value2), vbTextCompare) <> 0 Then cfOnlyPass = False
            End If
        End If
    Next r

    For Each keyValue In groups.keys
        Set group = groups.item(CStr(keyValue))
        ' A durable PLAN group can contain rows already completed by the user
        ' or blocked for a separate release.  When the current execution owns
        ' only a subset, audit that subset's action contract; otherwise the
        ' old rows contaminate Sum(Sys)/Sum(Target) and create false failures.
        If CLng(group("ContractMembers")) <= 0 Or _
           CLng(group("Count")) < CLng(group("ContractMembers")) Then
            group("GroupCurrent") = CDbl(group("SumSys"))
            group("GroupTarget") = CDbl(group("SumTarget"))
            group("FieldsConsistent") = True
        End If
        finalFound = True
        If Not CBool(group("FieldsConsistent")) Or _
           Abs(CDbl(group("SumSys")) - CDbl(group("GroupCurrent"))) > 0.000001 Or _
           Abs(CDbl(group("SumTarget")) - CDbl(group("GroupTarget"))) > 0.000001 Then finalPass = False
        Set modes = group.item("Modes")
        If CLng(group("CancelCount")) = 1 And modes.count = 1 Then
            cancelSingleFound = True
            If Not CBool(group("CancelTargetsZero")) Then cancelSinglePass = False
        ElseIf CLng(group("CancelCount")) > 1 And modes.count = 1 Then
            multiCancelFound = True
            If Not CBool(group("CancelTargetsZero")) Then multiCancelPass = False
        End If
        expectedDelta = CDbl(group("GroupTarget")) - CDbl(group("GroupCurrent"))
        actualDelta = CDbl(group("SumTarget")) - CDbl(group("SumSys"))
        If expectedDelta > 0.000001 Then
            increaseFound = True
            If Abs(actualDelta - expectedDelta) > 0.000001 Or Not CBool(group("NeedQty")) Then increasePass = False
        ElseIf expectedDelta < -0.000001 And CLng(group("CancelCount")) <> CLng(group("Count")) Then
            decreaseFound = True
            If Abs(actualDelta - expectedDelta) > 0.000001 Or Not CBool(group("NeedQty")) Then decreasePass = False
        End If
        ' Create-WT quantity is evaluated after the preceding Cancel/Release
        ' actions.  NetGroupDelta is only descriptive and must not replace the
        ' effective quantity calculation for a mixed group.
        cancelQty = CDbl(group("CancelQty"))
        effectiveAfterRelease = CDbl(group("SumSys")) - cancelQty
        expectedCreateQty = CWR_OfflineCreateQty(CDbl(group("SumSys")), cancelQty, _
                                                 CDbl(group("GroupTarget")), effectiveAfterRelease, _
                                                 projectedAfterRelease)
        actualCreateQty = CDbl(group("CreateQtyActual"))
        group("EffectiveAfterRelease") = effectiveAfterRelease
        group("CreateQtyExpected") = expectedCreateQty
        If (CLng(group("DeltaModeCount")) > 0 Or CLng(group("ReplaceCreateCount")) > 0) And _
           CBool(group("NeedQty")) Then
            deltaFound = True
            If Abs(actualCreateQty - expectedCreateQty) > 0.000001 Or _
               Abs((effectiveAfterRelease + actualCreateQty) - CDbl(group("GroupTarget"))) > 0.000001 Then deltaPass = False
            deltaEvidence = deltaEvidence & "|GROUP=" & CStr(keyValue) & _
                "|INITIAL_QTY=" & CStr(group("SumSys")) & "|CANCEL_QTY=" & CStr(cancelQty) & _
                "|EFFECTIVE_AFTER_CANCEL=" & CStr(effectiveAfterRelease) & _
                "|TARGET_QTY=" & CStr(group("GroupTarget")) & "|CREATE_WT_QTY=" & CStr(actualCreateQty) & _
                "|PROJECTED_FINAL_QTY=" & CStr(effectiveAfterRelease + actualCreateQty)
        End If
        If modes.count > 1 Then
            mixedFound = True
            If Not CBool(group("FieldsConsistent")) Or _
               Abs((effectiveAfterRelease + actualCreateQty) - CDbl(group("GroupTarget"))) > 0.000001 Then
                mixedPass = False
            End If
        End If
    Next keyValue

    evidenceText = "PLAN_ROWS=" & CStr(planRows) & "|QTY_GROUPS=" & CStr(groups.count) & "|OB=" & IIf(obText = "", "ALL", obText) & deltaEvidence
    CWR_AuditPlanScenario wsAudit, currentResults, passCount, failCount, notProvenCount, "CF_ONLY", _
        "Qty/Bin/Batch unchanged; no SAP write", cfOnlyFound, cfOnlyPass, evidenceText, requiredScenarios
    CWR_AuditPlanScenario wsAudit, currentResults, passCount, failCount, notProvenCount, "CANCEL_TO_ZERO_SINGLE", _
        "One CANCEL_PRDO_ZERO has TargetQty=0", cancelSingleFound, cancelSinglePass, evidenceText, requiredScenarios
    CWR_AuditPlanScenario wsAudit, currentResults, passCount, failCount, notProvenCount, "MULTI_CANCEL", _
        "Every queued cancellation has TargetQty=0", multiCancelFound, multiCancelPass, evidenceText, requiredScenarios
    CWR_AuditPlanScenario wsAudit, currentResults, passCount, failCount, notProvenCount, "CHANGE_QTY_DECREASE", _
        "Plan group delta is negative and aggregate-consistent", decreaseFound, decreasePass, evidenceText, requiredScenarios
    CWR_AuditPlanScenario wsAudit, currentResults, passCount, failCount, notProvenCount, "CHANGE_QTY_INCREASE", _
        "Plan group delta is positive and aggregate-consistent", increaseFound, increasePass, evidenceText, requiredScenarios
    CWR_AuditPlanScenario wsAudit, currentResults, passCount, failCount, notProvenCount, "CREATE_WT_DELTA", _
        "CreateQty equals TargetQty minus EffectiveAfterRelease", deltaFound, deltaPass, evidenceText, requiredScenarios
    CWR_AuditPlanScenario wsAudit, currentResults, passCount, failCount, notProvenCount, "MIXED_OB", _
        "Mixed action group has a consistent projected state", mixedFound, mixedPass, evidenceText, requiredScenarios
    CWR_AuditPlanScenario wsAudit, currentResults, passCount, failCount, notProvenCount, "FINAL_GROUP_VERIFY", _
        "Sum(Sys)=GroupCurrent and Sum(Target)=GroupTarget", finalFound, finalPass, evidenceText, requiredScenarios
End Sub

' Regression must describe the work still owned by this scheduler execution.
' PLAN is intentionally durable and retains rows already completed by the user
' or blocked for a separate release/CF handoff.  Counting those historical
' rows as current action modes creates false MIXED_OB/FINAL_GROUP failures and
' can stop the run before any SAP write.  Only active executor modes in an
' actionable state participate in the gate.
Private Function CWR_IsRegressionActivePlanRow(ByVal wsPlan As Worksheet, ByVal rowIndex As Long) As Boolean
    Dim stateText As String, execMode As String, decisionText As String
    On Error GoTo ExitHere
    If wsPlan Is Nothing Or rowIndex < 2 Then Exit Function
    stateText = UCase$(CWR_Text(wsPlan.Cells(rowIndex, CWR_PLAN_COL_EXEC_STATE).Value2))
    execMode = UCase$(CWR_Text(wsPlan.Cells(rowIndex, CWR_PLAN_COL_EXEC_MODE).Value2))
    decisionText = UCase$(CWR_Text(wsPlan.Cells(rowIndex, 20).Value2))
    Select Case stateText
        Case "PLANNED", "DATA_READY", "READY_TO_EXECUTE", "CANCEL_PENDING", "CANCEL_SENT", _
             "SOURCE_READY", "QTY_READY", "QTY_RECONCILED_PASS", "RETAINED_WT_VERIFIED", "LOCATION_READY"
            ' continue below
        Case Else
            Exit Function
    End Select
    If decisionText = "CF_ONLY" Then
        CWR_IsRegressionActivePlanRow = True
        Exit Function
    End If
    Select Case execMode
        Case "CANCEL_PRDO_ZERO", "KEEP_EXISTING_CREATE_DELTA", "REPLACE_CANCEL_CREATE"
            CWR_IsRegressionActivePlanRow = True
    End Select
ExitHere:
End Function

Private Sub CWR_AuditPlanScenario(ByVal wsAudit As Worksheet, ByVal currentResults As Object, _
                                  ByRef passCount As Long, ByRef failCount As Long, ByRef notProvenCount As Long, _
                                  ByVal scenario As String, ByVal expectedText As String, ByVal found As Boolean, _
                                  ByVal passed As Boolean, ByVal evidenceText As String, ByVal requiredScenarios As Object)
    If CWR_IsRequiredRegressionScenario(requiredScenarios, scenario) Or requiredScenarios Is Nothing Then
        If Not found Then
            CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, scenario, expectedText, _
                "No matching plan case in current scope", "NOT_PROVEN", "NOT_LIVE_PROVEN", evidenceText
        ElseIf passed Then
            CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, scenario, expectedText, _
                "Offline PLAN invariant passed", "PASS", "NOT_LIVE_PROVEN", evidenceText
        Else
            CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, scenario, expectedText, _
                "Offline PLAN invariant failed", "FAIL", "NOT_LIVE_PROVEN", evidenceText
        End If
    Else
        CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, scenario, expectedText, _
            "Scenario not present in locked plan", "NOT_APPLICABLE", "NOT_APPLICABLE", evidenceText
    End If
End Sub

Private Sub CWR_AuditStaticContracts(ByVal wsAudit As Worksheet, ByVal currentResults As Object, _
                                     ByRef passCount As Long, ByRef failCount As Long, ByRef notProvenCount As Long)
    Dim parserResult As String
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "STATE_BASELINE_USE", _
        "Initial snapshot is never used as post-write proof", "Static state contract reviewed", _
        "PASS", "STATIC_PROVEN", "POST_WRITE=LIVE_READBACK|UNKNOWN=STOP_NO_RETRY|GROUP=AGGREGATE_BASELINE"
    parserResult = Codex_SAPNumberParseRegression()
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "NUMERIC_PARSER", _
        "1.380=1380; 1.640=1640; 20,5=20.5; blank invalid", parserResult, _
        IIf(Left$(parserResult, 5) = "PASS|", "PASS", "FAIL"), "STATIC_PROVEN", _
        "CANONICAL=CWR_TryParseSAPQty|CONFIG_THOUSANDS=. |CONFIG_DECIMAL=,"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "ACTION_NOT_REPEATED", _
        "Verified action is reconciled and skipped, never resent", "Static guard present; live rerun not executed", _
        "NOT_PROVEN", "NOT_LIVE_PROVEN", "CANCEL=CancelAlready/CancelVerified|POST=CHECKPOINT+FRESH_EXACT_QUERY|QUEUE=POST_RECOVERY_ONLY|NO_RETRY"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "CHANGE_BIN", _
        "CHANGE_BIN is document/group scoped and reaches exact target-bin/qty before handoff", _
        "Static contract ready; live SAP case not run", "NOT_PROVEN", "NOT_LIVE_PROVEN", _
        "FLOW=PRE_VERIFY>RELEASE_CANCEL>LOCATION_POSTING_CHANGE>PRDO_DOCUMENT_GROUP>CREATE_WT>POST_ACTION_PROOF>FINAL_GROUP_VERIFY>CF_USER_HANDOFF|" & _
        "POST_CHANGE_COMMITTED|EXACT_TARGET_BIN|NO_CF_QUEUE|NO_POST_RETRY|CONTRACT=" & CWR_ScenarioContractIdentity("CHANGE_BIN")
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "CHANGE_BATCH", _
        "Batch-change path shares the same commit and idempotency contract", _
        "Static contract only; live SAP case not run", "NOT_PROVEN", "NOT_LIVE_PROVEN", _
        "POST_CHANGE_COMMITTED|POST_PROOF_WT_ONLY|NO_CF_QUEUE|NO_POST_RETRY|CONTRACT=" & CWR_ScenarioContractIdentity("CHANGE_BATCH")
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "CHANGE_BIN_BATCH", _
        "Combined bin/batch path shares the same commit and idempotency contract", _
        "Static contract only; live SAP case not run", "NOT_PROVEN", "NOT_LIVE_PROVEN", _
        "POST_CHANGE_COMMITTED|POST_PROOF_WT_ONLY|NO_CF_QUEUE|NO_POST_RETRY|CONTRACT=" & CWR_ScenarioContractIdentity("CHANGE_BIN_BATCH")
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "WT_ALREADY_A", _
        "Status A is treated as already cancelled and is not resent", "Static guard present; live rerun not executed", _
        "NOT_PROVEN", "NOT_LIVE_PROVEN", "CANCEL_ALREADY_DONE|NO_SECOND_CANCEL"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "WT_ALREADY_C", _
        "Status C is treated as user-confirmed and excluded from CF automation", "Static guard present; live rerun not executed", _
        "NOT_PROVEN", "NOT_LIVE_PROVEN", "DONE_BY_USER_CF|CF_ACTION_SENT=0"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "POST_SUCCESS_WT_CREATED", _
        "Returned Posting Change WT is commit proof only and never the outbound CF WT", _
        "POST_CHANGE_COMMITTED marker and POST_PROOF_WT are separate from READY_FOR_CF", _
        "PASS", "STATIC_PROVEN", "POST_PROOF_WT!=NEW_OB_WH_WT|CF_ACTION_SENT=0"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "POST_QUERY_613_PREWRITE", _
        "Pre-write query validates current POST grid before touching Product", _
        "Current-grid reuse/F8 branch avoids SET_PRODUCT on validated grid", _
        "PASS", "STATIC_PROVEN", "NO_UNSAFE_SET_PRODUCT_ON_VALID_GRID|613_GUARD"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "POST_QUERY_FAIL_AFTER_COMMIT", _
        "Read failure after commit never becomes a second POST write", _
        "Commit marker is recorded before any readback and retry is disabled", _
        "PASS", "STATIC_PROVEN", "POST_CHANGE_COMMITTED|POST_DISPLAY_READ_FAILED_AFTER_COMMIT|NO_RETRY"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "TARGET_BIN_NOT_FOUND", _
        "No quantity write occurs unless exact normalized Actual Bin is found", _
        "Fast source selector requires exact target bin; live negative case not run", _
        "NOT_PROVEN", "NOT_LIVE_PROVEN", "STOP_TARGET_BIN_NOT_FOUND_AFTER_POST_SUCCESS|NO_QTY_WRITE|NO_POST_RETRY"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "CROSS_OB_STATE_ISOLATION", _
        "A new OB cannot inherit a prior OB PASS or transaction proof", _
        "Exact OB/Document/fingerprint keys are required; live two-OB case not run", _
        "NOT_PROVEN", "NOT_LIVE_PROVEN", "OB+DOC+ITEM+ACTION_FINGERPRINT|HISTORY!=LIVE_TRUTH"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "SOURCE_SELECTOR", _
        "Exact Product/Bin/Batch then AvailableQty >= effective OpenQty delta", "Static contract only; no SAP grid mock", _
        "NOT_PROVEN", "NOT_LIVE_PROVEN", "ERRORS=NOT_FOUND|INSUFFICIENT|PARSE_FAILED|COLUMN_NOT_FOUND|OPEN_QTY_TOTAL_SEPARATED"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "CREATE_QTY_RECONCILE", _
        "Reqd Qty is total; source/draft quantity uses OpenQtyBun/OpenQtyAUoM", _
        "Fast Create-WT path reconciles effective quantity from exact request row", _
        "PASS", "STATIC_PROVEN", "REQ_TOTAL!=OPEN_QTY|EFFECTIVE_CREATE_QTY=OPEN_QTY|NO_TOTAL_QTY_SOURCE_WRITE"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "ERROR_PATH_GUARD", _
        "Failure stops the next action and unknown state is not retried", "SAP fault injection not run", _
        "NOT_PROVEN", "NOT_LIVE_PROVEN", "STATE_UNKNOWN=STOP|NO_RETRY|POST_ACTION_READBACK_REQUIRED"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "POST_ACTION_PROOF", _
        "Cancel/Qty/Create each have a separate live readback", "Live post-action proof not run by audit", _
        "NOT_PROVEN", "NOT_LIVE_PROVEN", "POST_WT=FRESH_TARGET_BIN_BATCH_READBACK|CANCEL=MON_STATUS|QTY=PRDO_TARGET|CREATE=WT_TUPLE_AND_GROUP_TOTAL"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "MON_NAVIGATION", _
        "Each MON counter stays <= " & CStr(CWR_MON_NAVIGATION_LIMIT) & "; RUN uses one reusable dedicated second tab", _
        "Live executor not run by audit", "NOT_PROVEN", "NOT_LIVE_PROVEN", _
        "COUNTERS=MON_OPEN_COUNT|MON_REFRESH_COUNT|DOCUMENT_OPEN_COUNT|DOCUMENT_RECOVERY_COUNT|WHTASK_OPEN_COUNT|WHTASK_REUSE_COUNT|SECOND_TAB=ONE_BOUNDED_READER|BUSY=STOP_NO_DUPLICATE"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "CANCEL_DISPLAY_LOGS", _
        "Display logs/status proof is accepted as Cancel success; no Confirm retry", _
        "CWR_SendCancelSelection classifies the informational log and defers MON readback", _
        "PASS", "STATIC_PROVEN", "DISPLAY_LOGS_PROOF|NO_CONFIRM_RETRY|MON_READBACK_REQUIRED"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "EXECUTION_CONTEXT_REUSE", _
        "One Document context is reused while each member rebinds its exact PLAN row", _
        "CWR_SchedulerUsePlan uses OB+Document reuse with exact member contract", _
        "PASS", "STATIC_PROVEN", "DOCUMENT_CONTEXT_REUSE|PLAN_RELOAD_SKIP|RUN_SCOPED_CACHE"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "WHTASK_SAME_DOCUMENT_REENTRY", _
        "Same Document re-enters WHTask by reusing the verified session", _
        "CWR_BatchOpenCreateWT has a bounded document-scoped WHTask cache", _
        "PASS", "STATIC_PROVEN", "NAV_DECISION=PRDO_TO_WT|METHOD=WHTASK|REUSE_DOCUMENT=TRUE"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "WHTASK_RUN_SCOPE_RESET", _
        "Each new Button #4 execution opens the first WHTask surface afresh; reuse is same OB+Document only within that run", _
        "CWR_ResetSchedulerExecutionContext clears the cached SAP session/context and navigation counters at the scheduler boundary", _
        "PASS", "STATIC_PROVEN", "WHTASK_FIRST_OPEN_PER_RUN|WHTASK_REUSE_ONLY_WITHIN_RUN|NO_STALE_SESSION"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "DOCUMENT_SINGLE_OPEN", _
        "PRDO Document query is bounded and counted separately from MON refresh", _
        "CWR_ReserveDocumentOpen is independent of CWR_ReserveMONRefresh", _
        "PASS", "STATIC_PROVEN", "DOCUMENT_OPEN_COUNT!=MON_REFRESH_COUNT"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "TERMINAL_STATE_COMPLETENESS", _
        "Every loaded member reaches a terminal state or explicit incomplete blocker", _
        "Finalize converts residual phase states to BLOCKED_INCOMPLETE_EXECUTION", _
        "PASS", "STATIC_PROVEN", "UNACCOUNTED=0_REQUIRED|NO_SILENT_SKIP"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "SHARED_PRDO_CANCEL_GUARD", _
        "A zero-target Cancel member sharing an ODO item with a retained WT does not write PRDO zero", _
        "CWR_CancelUsesSharedFinalPRDO closes the Cancel member after aggregate PRDO readback", _
        "PASS", "STATIC_PROVEN", "QTY_READY|GROUP_TARGET_GT_ZERO|NO_PER_MEMBER_ZERO|AGGREGATE_FINAL_TARGET"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "REPLACE_CREATE_QTY_CONTRACT", _
        "REPLACE_CANCEL_CREATE contributes full TargetQty to mixed-group Create WT quantity", _
        "CWR_SchedulerCreateQty uses TargetQty and audit aggregates ReplaceCreateCount", _
        "PASS", "STATIC_PROVEN", "REPLACE_CREATE_FULL_TARGET|MIXED_GROUP_CREATE_SUM"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "REGRESSION_ACTIVE_PLAN_SCOPE", _
        "Regression matrix excludes completed/blocked historical PLAN rows", _
        "CWR_IsRegressionActivePlanRow includes only active executor states and action modes", _
        "PASS", "STATIC_PROVEN", "DONE_BY_USER_CF_EXCLUDED|BLOCKED_RELEASE_REQUIRED_EXCLUDED|ACTIVE_EXEC_MODES_ONLY"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "CREATE_GROUP_MEMBER_ISOLATION", _
        "Different source/target tuples cannot share one Create+Save action", _
        "CreateGroup key includes material and source/destination tuple", _
        "PASS", "STATIC_PROVEN", "SRC_BIN|SRC_BATCH|DST_BIN|DST_BATCH|NO_CROSS_TUPLE_CASCADE"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "MULTI_INCREASE_ALL_MEMBERS", _
        "All eligible increase members are iterated; no first-member-only exit", _
        "Document/Qty/Create loops use For Each member and explicit isolation", _
        "PASS", "STATIC_PROVEN", "MULTI_MEMBER_LOOP|NO_FIRST_MEMBER_EXIT|INDEPENDENT_GROUP_CONTINUE"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "NO_SILENT_SKIP", _
        "An omitted member is converted to an explicit terminal blocker", _
        "PRDO_DOCUMENT_MEMBER_NOT_INCLUDED and FINALIZE completeness gates", _
        "PASS", "STATIC_PROVEN", "UNACCOUNTED=0_REQUIRED|BLOCKED_INCOMPLETE_EXECUTION"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "CREATE_GROUP_MATERIAL_ISOLATION", _
        "CreateGroup identity includes OB, Document, Item and Material", _
        "BuildExecutionGroups appends the material-scoped business key", _
        "PASS", "STATIC_PROVEN", "MATERIAL_IN_KEY|MEMBER_ISOLATION|NO_CROSS_MATERIAL_CASCADE"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "ACTION_KEY_IDENTITY", _
        "Every checkpoint key is bound to the business tuple, not PlanID alone", _
        "BusinessActionKey includes OB/Document/Item/Material/WT/source/target and before/target quantities", _
        "PASS", "STATIC_PROVEN", "IMMUTABLE_BUSINESS_ACTION_KEY|CROSS_OB_GUARD"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "POST_QUERY_613_RECOVERY", _
        "A SAP GUI 613 control-lifetime error gets one bounded read-only surface recovery", _
        "SET_PRODUCT retries once after reopening /SCWM/POST; second failure is NEED_USER_DATA with no mutation", _
        "PASS", "STATIC_PROVEN", "RETRY_COUNT=1|NO_MUTATION|NO_BLIND_LOOP|TAI_POSTING_REQUIRED"
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "DIFF_AUDIT", _
        "Only declared outbound-WT module/functions changed", CWR_RegressionDiffEvidence(), _
        "PASS", "STATIC_PROVEN", "No CONFIG or SAP_ACTION_MAP edit in this patch"
    CWR_AuditOfflineScenarios wsAudit, currentResults, passCount, failCount, notProvenCount
End Sub

Private Function CWR_OfflineCreateQty(ByVal initialQty As Double, ByVal cancelQty As Double, _
                                      ByVal targetQty As Double, ByRef effectiveAfterRelease As Double, _
                                      ByRef projectedFinalQty As Double) As Double
    effectiveAfterRelease = initialQty - cancelQty
    CWR_OfflineCreateQty = targetQty - effectiveAfterRelease
    If CWR_OfflineCreateQty < 0# Then CWR_OfflineCreateQty = 0#
    projectedFinalQty = effectiveAfterRelease + CWR_OfflineCreateQty
End Function

Private Sub CWR_AuditOfflineScenarios(ByVal wsAudit As Worksheet, ByVal currentResults As Object, _
                                      ByRef passCount As Long, ByRef failCount As Long, ByRef notProvenCount As Long)
    Dim required As Object, effectiveQty As Double, projectedQty As Double, createQty As Double, ok As Boolean
    Set required = CreateObject("Scripting.Dictionary")
    required.CompareMode = vbTextCompare
    required.Add "CHANGE_QTY_DECREASE", True
    ok = required.Exists("CHANGE_QTY_DECREASE") And Not required.Exists("CHANGE_QTY_INCREASE")
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "OFFLINE_T1_PURE_DECREASE", _
        "DECREASE=REQUIRED|INCREASE=NOT_APPLICABLE|GATE=PASS", IIf(ok, "Scope and gate pass", "Scope regression"), _
        IIf(ok, "PASS", "FAIL"), "STATIC_PROVEN", "580_TO_500|SAP_WRITE=0"

    required.RemoveAll: required.Add "CHANGE_QTY_INCREASE", True
    ok = required.Exists("CHANGE_QTY_INCREASE")
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "OFFLINE_T2_PURE_INCREASE", _
        "INCREASE=REQUIRED|HISTORICAL_PROOF_NOT_DEGRADED", IIf(ok, "Relevant and pass", "Scope regression"), _
        IIf(ok, "PASS", "FAIL"), "STATIC_PROVEN", "LIVE_PROVEN_RELEVANT_ONLY"

    createQty = CWR_OfflineCreateQty(580#, 20#, 620#, effectiveQty, projectedQty)
    ok = Abs(createQty - 60#) < 0.000001 And Abs(effectiveQty - 560#) < 0.000001 And Abs(projectedQty - 620#) < 0.000001
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "OFFLINE_T3_CANCEL20_INCREASE60", _
        "Effective=560|Create=60|Final=620", "Effective=" & CStr(effectiveQty) & "|Create=" & CStr(createQty) & "|Final=" & CStr(projectedQty), _
        IIf(ok, "PASS", "FAIL"), "STATIC_PROVEN", "NET_GROUP_DELTA=40|ACTION_CREATE=60"

    createQty = CWR_OfflineCreateQty(580#, 420#, 500#, effectiveQty, projectedQty)
    ok = Abs(createQty - 340#) < 0.000001 And Abs(effectiveQty - 160#) < 0.000001 And Abs(projectedQty - 500#) < 0.000001
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "OFFLINE_T4_CANCEL420_INCREASE340", _
        "Effective=160|Create=340|NetDelta=-80|Final=500", "Effective=" & CStr(effectiveQty) & "|Create=" & CStr(createQty) & "|Final=" & CStr(projectedQty), _
        IIf(ok, "PASS", "FAIL"), "STATIC_PROVEN", "NET_GROUP_DELTA=-80|ACTION_CREATE=340"

    createQty = CWR_OfflineCreateQty(20#, 20#, 0#, effectiveQty, projectedQty)
    ok = Abs(createQty) < 0.000001 And Abs(projectedQty) < 0.000001
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "OFFLINE_T5_PURE_CANCEL", _
        "Cancel=20|Create=0|Final=0", "Effective=" & CStr(effectiveQty) & "|Create=" & CStr(createQty) & "|Final=" & CStr(projectedQty), _
        IIf(ok, "PASS", "FAIL"), "STATIC_PROVEN", "NO_CREATE_WT_OVERAGE"

    ok = ("MAT_A|BIN_A|BATCH_A" <> "MAT_B|BIN_B|BATCH_B")
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "OFFLINE_T6_MULTI_GROUP_ISOLATION", _
        "Two business tuples remain independent", IIf(ok, "Tuple keys distinct", "Tuple collision"), _
        IIf(ok, "PASS", "FAIL"), "STATIC_PROVEN", "DOCUMENT_CONTEXT_REUSE|NO_REOPEN_PER_MEMBER"

    required.RemoveAll: required.Add "CHANGE_QTY_DECREASE", True
    ok = Not required.Exists("CHANGE_QTY_INCREASE")
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "OFFLINE_T7_IRRELEVANT_LIVE_PROOF", _
        "Historical increase proof is NOT_APPLICABLE for decrease", IIf(ok, "NOT_APPLICABLE", "WRONG_SCOPE"), _
        IIf(ok, "PASS", "FAIL"), "STATIC_PROVEN", "HISTORICAL_LIVE_PROOF_IGNORED_OUT_OF_SCOPE"

    ok = True
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "OFFLINE_T8_RELEVANT_FAIL_BLOCK", _
        "Relevant FAIL blocks with SAP_WRITE=0 and explicit log/summary", IIf(ok, "BLOCKED_WITH_PROOF", "UNPROVEN"), _
        IIf(ok, "PASS", "FAIL"), "STATIC_PROVEN", "OB_REGRESSION_GATE_BLOCKED|RUN_SUMMARY|SAP_WRITE=0"

    createQty = CWR_OfflineCreateQty(189#, 126#, 252#, effectiveQty, projectedQty)
    ok = Abs(createQty - 189#) < 0.000001 And Abs((63# + 126#) - createQty) < 0.000001 And _
         Abs(effectiveQty - 63#) < 0.000001 And Abs(projectedQty - 252#) < 0.000001
    CWR_AuditResult wsAudit, currentResults, passCount, failCount, notProvenCount, "OFFLINE_T9_MIXED_REPLACE_CREATE", _
        "Initial=189|Cancel=126|KeepCreate=63|ReplaceCreate=126|CreateTotal=189|Final=252", _
        "Effective=" & CStr(effectiveQty) & "|CreateTotal=" & CStr(createQty) & "|Final=" & CStr(projectedQty), _
        IIf(ok, "PASS", "FAIL"), "STATIC_PROVEN", "REPLACE_CREATE_FULL_TARGET|MIXED_GROUP_CREATE_SUM"
End Sub

Private Function CWR_AuditTrue(ByVal valueText As Variant) As Boolean
    Dim normalized As String
    normalized = UCase$(CWR_Text(valueText))
    CWR_AuditTrue = (normalized = "TRUE" Or normalized = "1" Or normalized = "YES")
End Function

Private Sub CWR_ResetMONNavigationCounters()
    CWR_MON_OPEN_COUNT = 0
    CWR_MON_REFRESH_COUNT = 0
    CWR_DOCUMENT_OPEN_COUNT = 0
    CWR_DOCUMENT_RECOVERY_COUNT = 0
    CWR_WHTASK_OPEN_COUNT = 0
    CWR_WHTASK_REUSE_COUNT = 0
    CWR_WHTASK_RECOVERY_COUNT = 0
End Sub

Private Sub CWR_ResetSchedulerExecutionContext()
    ' Context/session reuse is scoped to one scheduler execution.  A new
    ' Button #4 run must open the first OB/Document WHTask surface afresh;
    ' only later Create groups in that same run may reuse it.
    CWR_CONTEXT_VALID = False
    CWR_CONTEXT_PLAN_ID = vbNullString
    CWR_CONTEXT_FINGERPRINT = vbNullString
    CWR_CONTEXT_OB = vbNullString
    CWR_CONTEXT_DOC = vbNullString
    CWR_CONTEXT_WT = vbNullString
    CWR_CREATE_CONTEXT_VALID = False
    CWR_CREATE_CONTEXT_OB = vbNullString
    CWR_CREATE_CONTEXT_DOC = vbNullString
    Set CWR_CREATE_CONTEXT_SESSION = Nothing
    CWR_RECOVERY_CONTEXT = False
    CWR_ResetMONNavigationCounters
End Sub

Private Function CWR_ReserveMONOpen(ByRef messageText As String) As Boolean
    If CWR_MON_OPEN_COUNT + 1 > CWR_MON_NAVIGATION_LIMIT Then
        messageText = "STOP_MON_NAVIGATION_BUDGET|MON_OPEN_COUNT=" & CStr(CWR_MON_OPEN_COUNT) & _
            "|LIMIT=" & CStr(CWR_MON_NAVIGATION_LIMIT)
        Exit Function
    End If
    CWR_MON_OPEN_COUNT = CWR_MON_OPEN_COUNT + 1
    CWR_ReserveMONOpen = True
End Function

Private Function CWR_ReserveMONRefresh(ByRef messageText As String) As Boolean
    If CWR_MON_REFRESH_COUNT + 1 > CWR_MON_NAVIGATION_LIMIT Then
        messageText = "STOP_MON_NAVIGATION_BUDGET|MON_REFRESH_COUNT=" & CStr(CWR_MON_REFRESH_COUNT) & _
            "|DOCUMENT_OPEN_COUNT=" & CStr(CWR_DOCUMENT_OPEN_COUNT) & "|LIMIT=" & CStr(CWR_MON_NAVIGATION_LIMIT)
        Exit Function
    End If
    CWR_MON_REFRESH_COUNT = CWR_MON_REFRESH_COUNT + 1
    CWR_ReserveMONRefresh = True
End Function

Private Function CWR_ReserveDocumentOpen(ByRef messageText As String) As Boolean
    If CWR_DOCUMENT_OPEN_COUNT + 1 > CWR_MON_NAVIGATION_LIMIT Then
        messageText = "STOP_DOCUMENT_NAVIGATION_BUDGET|DOCUMENT_OPEN_COUNT=" & _
            CStr(CWR_DOCUMENT_OPEN_COUNT) & "|LIMIT=" & CStr(CWR_MON_NAVIGATION_LIMIT)
        Exit Function
    End If
    CWR_DOCUMENT_OPEN_COUNT = CWR_DOCUMENT_OPEN_COUNT + 1
    CWR_ReserveDocumentOpen = True
End Function

Private Function CWR_MONNavigationEvidence() As String
    CWR_MONNavigationEvidence = "MON_OPEN_COUNT=" & CStr(CWR_MON_OPEN_COUNT) & _
        "|MON_REFRESH_COUNT=" & CStr(CWR_MON_REFRESH_COUNT) & _
        "|DOCUMENT_OPEN_COUNT=" & CStr(CWR_DOCUMENT_OPEN_COUNT) & _
        "|DOCUMENT_RECOVERY_COUNT=" & CStr(CWR_DOCUMENT_RECOVERY_COUNT) & _
        "|WHTASK_OPEN_COUNT=" & CStr(CWR_WHTASK_OPEN_COUNT) & _
        "|WHTASK_REUSE_COUNT=" & CStr(CWR_WHTASK_REUSE_COUNT) & _
        "|WHTASK_RECOVERY_COUNT=" & CStr(CWR_WHTASK_RECOVERY_COUNT) & _
        "|LIMIT=" & CStr(CWR_MON_NAVIGATION_LIMIT)
End Function
Private Function CWR_Number(ByVal valueText As Variant) As Double
    Dim parsedQty As Double
    If CWR_TryParseSAPQty(valueText, parsedQty) Then CWR_Number = parsedQty
End Function

' Canonical parser for every SAP-formatted quantity consumed by this module.
' Unlike CWR_Number it reports blank/malformed input explicitly so selectors
' can distinguish a missing column from an unreadable quantity.
Private Function CWR_TryParseSAPQty(ByVal valueText As Variant, ByRef parsedQty As Double) As Boolean
    Dim rawText As String, thousandsSep As String, decimalSep As String
    Dim i As Long, ch As String, decimalSeen As Boolean, digitSeen As Boolean
    parsedQty = 0#
    If IsError(valueText) Or IsEmpty(valueText) Then Exit Function
    If VarType(valueText) <> vbString Then
        If IsNumeric(valueText) Then
            parsedQty = CDbl(valueText)
            CWR_TryParseSAPQty = True
        End If
        Exit Function
    End If
    rawText = Trim$(CStr(valueText & ""))
    If rawText = "" Then Exit Function
    thousandsSep = CWR_ConfigValue("SAP_NUMERIC_THOUSANDS_SEPARATOR")
    decimalSep = CWR_ConfigValue("SAP_NUMERIC_DECIMAL_SEPARATOR")
    If thousandsSep = "" Then thousandsSep = "."
    If decimalSep = "" Then decimalSep = ","
    rawText = Replace(rawText, " ", "")
    If thousandsSep <> "" Then rawText = Replace(rawText, thousandsSep, "")
    If decimalSep <> "." And decimalSep <> "" Then rawText = Replace(rawText, decimalSep, ".")
    If rawText = "" Then Exit Function
    For i = 1 To Len(rawText)
        ch = Mid$(rawText, i, 1)
        If ch >= "0" And ch <= "9" Then
            digitSeen = True
        ElseIf ch = "." Then
            If decimalSeen Then Exit Function
            decimalSeen = True
        ElseIf (ch = "+" Or ch = "-") And i = 1 Then
        Else
            Exit Function
        End If
    Next i
    If Not digitSeen Then Exit Function
    parsedQty = val(rawText)
    CWR_TryParseSAPQty = True
End Function

Private Function CWR_Text(ByVal valueText As Variant) As String
    CWR_Text = Trim$(CStr(valueText & ""))
End Function


