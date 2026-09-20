Option Explicit
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Private mCfgBusyTimeoutSeconds As Double
Private mCfgPostPollIntervalMs As Long
Private mCfgDeliveryGraceSeconds As Double
Private mCfgMigoOpenTimeoutSeconds As Double
Private mCfgMigoUiPollMs As Long
Private mCfgInputIdlePollMs As Long
Private mCfgInputIdleStableSamples As Long
Private mCfgInputIdleTimeoutSeconds As Double
Private mCfgStopAfterError As Boolean
Private mCfgMaxProductsPerOB As Long
Private mCfgMaxRowsPerOB As Long
Private mCfgAutoResetAfterError As Boolean
Private mCfgMaxOBPerRun As Long
Private mCfgSaveWorkbookEachOB As Boolean
Private mCfgAutoAcceptKnownPopups As Boolean
Private mCfgStopOnUnknownPopup As Boolean
Private mCfgFailIfMultipleMigo As Boolean
Private mCfgLayoutCache As Boolean
Private mCfgLayoutRecheckEvery As Long
Private mCfgScrollSettleMs As Long
Private mCachedMigoTable As Object
Private mCachedMigoTableId As String
Private mCachedSessionKey As String
Private mCachedBatchCol As Long
Private mCachedTableCols As Long
Private mBoundSessionKey As String
Private mMigoSessionMatchCount As Long
Private mObSinceLayoutCheck As Long
Private mOneClickMigo As Boolean
Public LastMigoRunSucceeded As Boolean
Public LastMigoProcessedCount As Long
Public LastMigoPostedCount As Long
Public LastMigoRunMessage As String
Public LastMigoTotalSeconds As Double
Private mRunUncheckRedOK As Boolean
Private mLastExcludedRowsCsv As String
Private mRunInProgress As Boolean
Private mToolActionRunning As Boolean
Private mToolActionName As String

Public Sub MIGO311_CompileProbe()
    Dim mapReady As Boolean
    mapReady = ModuleHelper.ToolConfigBool("MIGO_EXACT_MAP_READY", False)
    If SAPRun_IsMapMode() And mapReady Then Err.Raise vbObjectError + 9971, , "MIGO_MAP_MODE_CONTRACT_PROBE_FAIL"
End Sub
Public Function ToolTryBegin(ByVal actionName As String) As Boolean
If mToolActionRunning Or mRunInProgress Then
MsgBox "Tool dang xu ly: " & IIf(mToolActionName <> "", mToolActionName, "mot tac vu khac") & "." & vbCrLf & _
"Hay cho tac vu hien tai ket thuc, khong bam chong nut.", vbExclamation, "FIFO + MIGO"
Exit Function
End If
modSupports.CancelDeferredSave
ResetTransientState
mToolActionRunning = True
mToolActionName = actionName
ToolTryBegin = True
End Function
Public Sub ToolEnd()
mToolActionRunning = False
mToolActionName = ""
End Sub
Public Function ToolRunIsActive() As Boolean
ToolRunIsActive = (mToolActionRunning Or mRunInProgress)
End Function
Public Sub GetExistingOBState(ByRef readyCount As Long, ByRef postedCount As Long, ByRef staleCount As Long)
Dim ws As Worksheet, lr As Long, r As Long, ob As String, statusText As String
Dim dReady As Object, dPosted As Object, dStale As Object
readyCount = 0: postedCount = 0: staleCount = 0
On Error Resume Next
Set ws = ThisWorkbook.Worksheets("OB_GROUPS")
On Error GoTo 0
If ws Is Nothing Then Exit Sub
lr = ws.Cells(ws.rows.count, 1).End(xlUp).row
If lr < 2 Then Exit Sub
Set dReady = CreateObject("Scripting.Dictionary")
Set dPosted = CreateObject("Scripting.Dictionary")
Set dStale = CreateObject("Scripting.Dictionary")
For r = 2 To lr
ob = Trim$(CStr(ws.Cells(r, 1).Value & ""))
If ob <> "" Then
statusText = UCase$(Trim$(CStr(ws.Cells(r, 9).Value & "")))
If Trim$(CStr(ws.Cells(r, 11).Value & "")) <> "" Or InStr(1, statusText, "POSTED", vbTextCompare) > 0 Then
If Not dPosted.Exists(ob) Then dPosted.Add ob, True
ElseIf statusText = "READY" Then
If Not dReady.Exists(ob) Then dReady.Add ob, True
ElseIf InStr(1, statusText, "STALE", vbTextCompare) > 0 Then
If Not dStale.Exists(ob) Then dStale.Add ob, True
End If
End If
Next r
readyCount = dReady.count: postedCount = dPosted.count: staleCount = dStale.count
End Sub
Public Function IsRunInProgress() As Boolean
IsRunInProgress = mRunInProgress
End Function
Public Sub ResetTransientState()
If mRunInProgress Then Exit Sub
mOneClickMigo = False
mRunUncheckRedOK = False
mLastExcludedRowsCsv = ""
LastMigoRunSucceeded = False
LastMigoProcessedCount = 0
LastMigoPostedCount = 0
LastMigoRunMessage = ""
LastMigoTotalSeconds = 0#
mBoundSessionKey = ""
mMigoSessionMatchCount = 0
RUN_ClearMigoCache
End Sub
Private Sub RUN_ManualEntry(ByVal onlyFirst As Boolean, ByVal uncheckRed As Boolean, ByVal actionName As String)
If Not ToolTryBegin(actionName) Then Exit Sub
On Error GoTo CleanExit
mOneClickMigo = False
mRunUncheckRedOK = uncheckRed
RUN_RunMIGO311 onlyFirst
CleanExit:
mRunUncheckRedOK = False
mOneClickMigo = False
ToolEnd
End Sub
Public Sub Run_MIGO311_Post_FirstReadyOB()
RUN_ManualEntry True, False, "MIGO 311 - 1 OB"
End Sub
Public Sub Run_MIGO311_Post_AllReadyOB()
RUN_ManualEntry False, False, "MIGO 311 - TAT CA OB"
End Sub
Public Sub Run_MIGO311_Post_FirstReadyOB_AllowYellow()
Run_MIGO311_Post_FirstReadyOB
End Sub
Public Sub Run_MIGO311_Post_AllReadyOB_AllowYellow()
Run_MIGO311_Post_AllReadyOB
End Sub
Public Sub Post_MIGO_311_FromOBGroups_V3()
Run_MIGO311_Post_FirstReadyOB
End Sub
Public Sub Run_MIGO311_OneClick()
Dim scopeText As String, onlyFirst As Boolean
On Error GoTo CleanExit
mOneClickMigo = True
mRunUncheckRedOK = ModuleHelper.ToolConfigBool("MIGO_UNCHECK_RED_MODE", False)
scopeText = UCase$(Trim$(ModuleHelper.ToolConfig("ONE_CLICK_POST_SCOPE", "ALL")))
onlyFirst = (scopeText = "FIRST")
RUN_RunMIGO311 onlyFirst
CleanExit:
mRunUncheckRedOK = False
mOneClickMigo = False
End Sub
Private Sub RUN_LoadConfig()
ModuleHelper.ResetToolConfigCache
mCfgBusyTimeoutSeconds = ModuleHelper.ToolConfigDouble("MIGO_BUSY_TIMEOUT_SEC", 60#, 10#, 300#)
mCfgPostPollIntervalMs = ModuleHelper.ToolConfigLong("MIGO_POST_POLL_MS", 60, 30, 1000)
mCfgDeliveryGraceSeconds = ModuleHelper.ToolConfigDouble("MIGO_DELIVERY_GRACE_SEC", 1#, 0#, 30#)
mCfgMigoOpenTimeoutSeconds = ModuleHelper.ToolConfigDouble("MIGO_OPEN_TIMEOUT_SEC", 20#, 5#, 120#)
mCfgMigoUiPollMs = ModuleHelper.ToolConfigLong("MIGO_UI_POLL_MS", 60, 30, 1000)
mCfgInputIdlePollMs = ModuleHelper.ToolConfigLong("MIGO_INPUT_IDLE_POLL_MS", 60, 20, 1000)
mCfgInputIdleStableSamples = ModuleHelper.ToolConfigLong("MIGO_INPUT_IDLE_STABLE_SAMPLES", 1, 1, 10)
mCfgInputIdleTimeoutSeconds = ModuleHelper.ToolConfigDouble("MIGO_INPUT_IDLE_TIMEOUT_SEC", 4#, 1#, 30#)
mCfgStopAfterError = ModuleHelper.ToolConfigBool("MIGO_STOP_AFTER_ERROR", False)
mCfgMaxProductsPerOB = ModuleHelper.ToolConfigLong("MAX_PRODUCTS_PER_OB", 4, 1, 20)
mCfgMaxRowsPerOB = ModuleHelper.ToolConfigLong("MAX_ROWS_PER_OB", 20, 1, 50)
mCfgAutoResetAfterError = ModuleHelper.ToolConfigBool("MIGO_AUTO_RESET_AFTER_ERROR", True)
mCfgMaxOBPerRun = ModuleHelper.ToolConfigLong("MIGO_MAX_OB_PER_RUN", 0, 0, 1000)
mCfgSaveWorkbookEachOB = ModuleHelper.ToolConfigBool("MIGO_SAVE_WORKBOOK_EACH_OB", False)
mCfgAutoAcceptKnownPopups = ModuleHelper.ToolConfigBool("SAP_AUTO_ACCEPT_KNOWN_POPUPS", True)
mCfgStopOnUnknownPopup = ModuleHelper.ToolConfigBool("SAP_STOP_ON_UNKNOWN_POPUP", True)
mCfgFailIfMultipleMigo = ModuleHelper.ToolConfigBool("SAP_FAIL_IF_MULTIPLE_MIGO", True)
mCfgLayoutCache = ModuleHelper.ToolConfigBool("MIGO_LAYOUT_CACHE", True)
mCfgLayoutRecheckEvery = ModuleHelper.ToolConfigLong("MIGO_LAYOUT_RECHECK_EVERY_N_OB", 5, 1, 50)
mCfgScrollSettleMs = ModuleHelper.ToolConfigLong("MIGO_SCROLL_SETTLE_MS", 30, 0, 500)
End Sub
Public Sub Run_MIGO311_Post_FirstReadyOB_UncheckRedOK()
RUN_ManualEntry True, True, "MIGO 311 - 1 OB BO DONG DO"
End Sub
Public Sub Run_MIGO311_Post_AllReadyOB_UncheckRedOK()
RUN_ManualEntry False, True, "MIGO 311 - TAT CA, BO DONG DO"
End Sub
Public Sub Run_MIGO311_Post_AllReadyOB_PartialRescue()
Run_MIGO311_Post_AllReadyOB_UncheckRedOK
End Sub
Private Sub RUN_RunMIGO311(ByVal onlyFirstReadyOB As Boolean)
On Error GoTo ErrHandler
Dim guardMessage As String
If SAPRun_IsMapMode() Then
LastMigoRunMessage = "STOP_RUN_MODE_REQUIRED: MIGO 311 is not a mapping operation."
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbExclamation, "MIGO 311"
Exit Sub
End If
If Not SAPRun_RequireReadOnly("WRITE", guardMessage) Then
LastMigoRunMessage = guardMessage
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbExclamation, "MIGO 311"
Exit Sub
End If
If Not SAPRun_RequireRunMode("MIGO_POST", guardMessage) Then
LastMigoRunMessage = guardMessage
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbExclamation, "MIGO 311"
Exit Sub
End If
If Not ModuleHelper.ToolConfigBool("MIGO_EXACT_MAP_READY", False) Then
LastMigoRunMessage = "STOP_RUN_EXACT_MAP_REQUIRED: MIGO 311 still contains legacy screen/fallback paths; exact MIGO map is not enabled."
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbExclamation, "MIGO 311"
Exit Sub
End If
RUN_LoadConfig
LastMigoRunSucceeded = False
LastMigoProcessedCount = 0
LastMigoPostedCount = 0
LastMigoRunMessage = "MIGO dung truoc khi tao Delivery."
LastMigoTotalSeconds = 0#
If mRunInProgress Then
LastMigoRunMessage = "Tool MIGO dang chay. Khong duoc bam chay lan thu hai."
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, _
vbExclamation, "MIGO dang xu ly"
Exit Sub
End If
Dim runUiLocked As Boolean
Dim oldEnableEvents As Boolean
Dim oldScreenUpdating As Boolean
Dim oldDisplayAlerts As Boolean
Dim oldCancelKey As XlEnableCancelKey
Dim wsOB As Worksheet
On Error Resume Next
Set wsOB = ThisWorkbook.Worksheets("OB_GROUPS")
On Error GoTo ErrHandler
If wsOB Is Nothing Then
LastMigoRunMessage = "Khong tim thay sheet OB_GROUPS. Hay chay GOM OB truoc."
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbExclamation
Exit Sub
End If
If Not RUN_CheckOBTemplate(wsOB) Then Exit Sub
Dim lastRow As Long
lastRow = RUN_LastRow(wsOB, 1)
If lastRow < 2 Then
LastMigoRunMessage = "OB_GROUPS chua co du lieu."
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbExclamation
Exit Sub
End If
RUN_SortOBGroups wsOB, lastRow
Dim data As Variant
data = wsOB.Range(wsOB.Cells(2, 1), wsOB.Cells(lastRow, 17)).Value2
If Not RUN_PrecheckOBGroupsArray(data, wsOB) Then
LastMigoRunMessage = "OB_GROUPS dang co loi logic. Kiem tra cot Status/SAP Message truoc khi Run SAP."
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbExclamation
Exit Sub
End If
If RUN_CountReadyOBsArray(data) = 0 Then
LastMigoRunSucceeded = True
LastMigoRunMessage = "Khong con OB Ready. Lich su Posted/Error duoc giu nguyen; khong goi SAP."
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbInformation, "MIGO 311"
Exit Sub
End If
Dim session As Object
Application.StatusBar = "Buoc 3/3 - 1/4: Dang ket noi SAP GUI..."
DoEvents
Set session = modSupports.Get_SAPSession()
If session Is Nothing Then Set session = modSupports.session
If session Is Nothing Then
LastMigoRunMessage = "Khong lay duoc SAP session. Hay mo SAP, dang nhap va cap quyen SAP GUI Scripting neu duoc hoi."
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbExclamation, "Khong ket noi duoc SAP"
Exit Sub
End If
Dim foundMigoSession As Object
Dim openMigoError As String
Application.StatusBar = "Buoc 3/3 - 2/4: Dang tim tab MIGO dang mo..."
DoEvents
Set foundMigoSession = RUN_FindExistingMIGOSession(session)
If mCfgFailIfMultipleMigo And mMigoSessionMatchCount > 1 Then
LastMigoRunMessage = "Phat hien " & CStr(mMigoSessionMatchCount) & " tab MIGO cung System/Client/User. Hay dong bot, chi de 1 tab MIGO roi chay lai."
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbExclamation, "Nhieu tab MIGO"
Exit Sub
End If
If Not foundMigoSession Is Nothing Then Set session = foundMigoSession
RUN_ClearMigoCache
mBoundSessionKey = RUN_SessionKey(session)
Application.StatusBar = "Buoc 3/3 - 3/4: Dang mo/cau hinh MIGO 311..."
DoEvents
If Not RUN_EnsureMIGO311Ready(session, openMigoError, False) Then
LastMigoRunMessage = "Khong tu mo/cau hinh duoc MIGO 311: " & openMigoError
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbExclamation, "Auto Open MIGO 311"
Exit Sub
End If
If Not RUN_PrepareReusedMIGO(session, openMigoError) Then
LastMigoRunMessage = "MIGO dang co du lieu/popup tu luot truoc va khong reset an toan duoc: " & openMigoError
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbExclamation, "MIGO chua sach"
Exit Sub
End If
Application.StatusBar = "Buoc 3/3 - 4/4: MIGO san sang. Dang nap danh sach OB..."
DoEvents
RUN_ActivateSAPWindow session
Dim firstReadyUsesPreparedMigo As Boolean
firstReadyUsesPreparedMigo = True
Dim msgConfirm As String
msgConfirm = "Tool da tu tim/mo MIGO -> Transfer Posting -> Other -> 311." & vbCrLf & _
"Khong can mo san MIGO truoc khi chay." & vbCrLf & _
IIf(mRunUncheckRedOK, _
"Co che: Check de bo tick dong RED, sau do Post.", _
"Co che: Post truc tiep, khong bam Check.") & vbCrLf & _
IIf(onlyFirstReadyOB, "Pham vi: chi 1 OB Ready dau tien.", "Pham vi: tat ca OB Ready.") & vbCrLf & _
"Tiep tuc?"
If Not mOneClickMigo Then
If MsgBox(msgConfirm, vbQuestion + vbYesNo, "Run MIGO 311") <> vbYes Then Exit Sub
End If
oldEnableEvents = Application.EnableEvents
oldScreenUpdating = Application.ScreenUpdating
oldDisplayAlerts = Application.DisplayAlerts
oldCancelKey = Application.EnableCancelKey
mRunInProgress = True
runUiLocked = True
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableCancelKey = xlErrorHandler
Application.Cursor = xlWait
Dim totalRunStart As Double
Dim totalOBSeconds As Double
Dim obStart As Double
Dim obSeconds As Double
totalRunStart = Timer
wsOB.Range("A2:A" & lastRow).NumberFormat = "@"
wsOB.Range("K2:K" & lastRow).NumberFormat = "@"
' OB_GROUPS da duoc nap vao Variant array truoc khi ket noi SAP.
Dim i As Long, n As Long, startRow As Long, endRow As Long
Dim obName As String
Dim statusOut As String, deliveryOut As String, msgOut As String
Dim processedCount As Long, postedCount As Long
Dim prepareStart As Double, prepareSeconds As Double
n = UBound(data, 1)
i = 1
Do While i <= n
obName = RUN_Clean(data(i, 1))
If obName = "" Then Exit Do
startRow = i
endRow = RUN_GetGroupEndInArray(data, i, n, obName)
If Not RUN_ShouldSkipOB(data, startRow, endRow) Then
Application.StatusBar = "Dang xu ly " & obName & "..."
obStart = Timer
prepareSeconds = 0#
If firstReadyUsesPreparedMigo Then
    firstReadyUsesPreparedMigo = False
Else
    prepareStart = Timer
    If Not RUN_PrepareReusedMIGO(session, openMigoError) Then
        statusOut = "MIGO Reset Error"
        deliveryOut = ""
        msgOut = openMigoError
        RUN_WriteResult wsOB, startRow, endRow, statusOut, msgOut, deliveryOut
        processedCount = processedCount + 1
        Exit Do
    End If
    prepareSeconds = RUN_ElapsedSeconds(prepareStart)
End If
RUN_PostOneOB_MIGO311 session, data, startRow, endRow, _
statusOut, deliveryOut, msgOut
If prepareSeconds > 0# Then msgOut = RUN_AppendMessage(msgOut, "Prepare=" & Format$(prepareSeconds, "0.00") & "s")
obSeconds = RUN_ElapsedSeconds(obStart)
totalOBSeconds = totalOBSeconds + obSeconds
msgOut = RUN_AppendMessage(msgOut, _
"OB Time=" & Format$(obSeconds, "0.00") & "s")
RUN_WriteResult wsOB, startRow, endRow, statusOut, msgOut, deliveryOut
If mCfgSaveWorkbookEachOB Then Application.StatusBar = "Da cap nhat OB; se luu file moi khi ket thuc."
processedCount = processedCount + 1
If statusOut = "Posted" Or statusOut = "Posted Partial" Then
postedCount = postedCount + 1
End If
Application.StatusBar = _
"Da xu ly " & processedCount & " OB | OK " & postedCount & _
" | OB vua roi " & Format$(obSeconds, "0.00") & "s" & _
IIf(processedCount > 0, _
" | TB " & Format$(totalOBSeconds / processedCount, "0.00") & "s/OB", "")
If onlyFirstReadyOB Then Exit Do
If mCfgMaxOBPerRun > 0 And processedCount >= mCfgMaxOBPerRun Then Exit Do
If statusOut <> "Posted" And statusOut <> "Posted Partial" Then
If statusOut = "System Error" Or statusOut = "MIGO Table Error" Or _
statusOut = "SAP Session Lost" Or _
statusOut = "Check Popup Error" Or statusOut = "Check Warning Loop" Or _
statusOut = "Post Popup Error" Or statusOut = "Post Popup Blocked" Or _
statusOut = "Post Warning Loop" Then Exit Do
If mCfgStopAfterError Then Exit Do
If mCfgAutoResetAfterError Then RUN_ResetMIGO311Screen session
End If
End If
i = endRow + 1
Loop
Dim totalRunSeconds As Double
Dim averageSeconds As Double
totalRunSeconds = RUN_ElapsedSeconds(totalRunStart)
If processedCount > 0 Then averageSeconds = totalOBSeconds / processedCount
LastMigoProcessedCount = processedCount
LastMigoPostedCount = postedCount
LastMigoTotalSeconds = totalRunSeconds
LastMigoRunSucceeded = (processedCount > 0 And postedCount = processedCount)
LastMigoRunMessage = "Da xu ly " & processedCount & " OB; Posted " & postedCount & _
"; Loi/khong Posted " & (processedCount - postedCount) & "."
If Not mOneClickMigo Then
MsgBox "Da xu ly " & processedCount & " OB, posted OK " & postedCount & " OB." & _
vbCrLf & "Tong thoi gian: " & Format$(totalRunSeconds, "0.00") & " giay." & _
vbCrLf & "Trung binh: " & Format$(averageSeconds, "0.00") & " giay/OB." & _
vbCrLf & "Chi tiet tung OB nam trong cot SAP Message.", _
vbInformation, "MIGO 311 hoan tat"
End If
GoTo CleanRun
ErrHandler:
LastMigoRunMessage = "System Error: " & Err.Description
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbCritical
CleanRun:
Application.StatusBar = False
If runUiLocked Then
Application.EnableEvents = oldEnableEvents
Application.ScreenUpdating = oldScreenUpdating
Application.DisplayAlerts = oldDisplayAlerts
Application.EnableCancelKey = oldCancelKey
Application.Cursor = xlDefault
End If
mRunInProgress = False
End Sub
Private Sub RUN_PostOneOB_MIGO311(ByVal session As Object, ByVal data As Variant, _
ByVal startRow As Long, ByVal endRow As Long, _
ByRef statusOut As String, ByRef deliveryOut As String, ByRef msgOut As String)
On Error GoTo ErrHandler
statusOut = ""
deliveryOut = ""
msgOut = ""
mLastExcludedRowsCsv = ""
' FAST: khong xoa cache Batch theo chu ky. RUN_GetBatchColumnForRow tu kiem tra lai title + Columns.Count moi lan dung.
If Not RUN_IsSessionAlive(session) Then
statusOut = "SAP Session Lost"
msgOut = "SAP session khong con hoat dong truoc khi dien OB."
GoTo CleanExit
End If
Dim tbl As Object
Dim colMapPre As Object
Set tbl = RUN_FindMigoItemTable(session)
If tbl Is Nothing Then
statusOut = "MIGO Table Error"
msgOut = "Mat bang MIGO sau khi Auto Open. Tool da dung truoc khi dien OB."
GoTo CleanExit
End If
Dim rowCount As Long
rowCount = endRow - startRow + 1
Dim hasAnyBatch As Boolean
Dim rescueRowsCsv As String, rescueDetail As String
Dim rescueCount As Long, postRescueTried As Boolean
hasAnyBatch = RUN_OBHasAnyBatch(data, startRow, endRow)
postRescueTried = False
Dim fillMainStart As Double, fillMainSeconds As Double, fillBatchStart As Double, fillBatchSeconds As Double
Dim enter1Start As Double, enter1Seconds As Double, enter2Start As Double, enter2Seconds As Double
Dim fastPostAfterBatch As Boolean, directPostTried As Boolean, directPostFallbackUsed As Boolean
fastPostAfterBatch = (UCase$(Trim$(ModuleHelper.ToolConfig("MIGO_FAST_POST_AFTER_BATCH", "TRUE"))) <> "FALSE")
fillMainStart = Timer
Set colMapPre = RUN_GetPreBatchColumnMap()
RUN_FillPreBatch_ByGrid tbl, colMapPre, data, startRow, endRow
fillMainSeconds = RUN_ElapsedSeconds(fillMainStart)
enter1Start = Timer
session.FindById("wnd[0]").sendVKey 0
If hasAnyBatch Then
    Set tbl = Nothing
    RUN_WaitForBatchEntryReady session, mCfgInputIdleTimeoutSeconds, tbl
Else
    RUN_WaitForStableIdle session, mCfgInputIdleTimeoutSeconds
End If
enter1Seconds = RUN_ElapsedSeconds(enter1Start)
Dim preType As String, preText As String
preType = RUN_GetSbarType(session)
preText = RUN_GetSbarText(session)
If RUN_IsFatalSbarType(preType) Or InStr(1, preText, "Enter ", vbTextCompare) > 0 Then
If Not (hasAnyBatch And RUN_IsExpectedBatchPrompt(preText)) Then
statusOut = "PreBatch Error"
msgOut = "Sau khi dien Product/Qty/Text/Plant/SLoc/SLoc Transfer va Enter, SAP bao: SBar(" & preType & "): " & preText
GoTo CleanExit
End If
End If
If hasAnyBatch Then
If tbl Is Nothing Then Err.Raise vbObjectError + 311, , "Mat MIGO table sau Enter 1."
fillBatchStart = Timer
RUN_FillBatchDynamic_ByGrid tbl, data, startRow, endRow
fillBatchSeconds = RUN_ElapsedSeconds(fillBatchStart)
Dim batchType As String, batchText As String
If Not fastPostAfterBatch Then
    enter2Start = Timer
    session.FindById("wnd[0]").sendVKey 0
    RUN_WaitForStableIdle session, mCfgInputIdleTimeoutSeconds
    enter2Seconds = RUN_ElapsedSeconds(enter2Start)
    batchType = RUN_GetSbarType(session)
    batchText = RUN_GetSbarText(session)
    If RUN_IsFatalSbarType(batchType) Or InStr(1, batchText, "Enter ", vbTextCompare) > 0 Then
        rescueRowsCsv = "": rescueDetail = ""
        rescueCount = RUN_TryRescueRedLines(session, rowCount, rescueRowsCsv, rescueDetail)
        If rescueCount > 0 And rescueCount < rowCount Then
            mLastExcludedRowsCsv = rescueRowsCsv
        Else
            statusOut = "Batch Error"
            msgOut = "Sau khi dien Batch, SAP bao: SBar(" & batchType & "): " & batchText & IIf(rescueDetail <> "", " | " & rescueDetail, "")
            GoTo CleanExit
        End If
    End If
End If
End If
If mRunUncheckRedOK Then
If Not RUN_PressCheck(session) Then
statusOut = "Check Error"
msgOut = "Khong bam duoc nut Check. Hay kiem tra du lieu dong va nut Check co active khong."
GoTo CleanExit
End If
RUN_Wait session, 700
Dim popupText As String
Dim checkWarningTrace As String
checkWarningTrace = ""
popupText = RUN_ReadPopupTextIfAny(session)
If popupText <> "" Then
checkWarningTrace = "Popup=" & popupText & "; "
If Not RUN_AcceptPopupIfAny(session) Then
statusOut = "Check Popup Error"
msgOut = "Khong tu dong xac nhan duoc popup khi Check: " & popupText
GoTo CleanExit
End If
RUN_Wait session, 300
End If
Dim hasRed As Boolean, hasYellow As Boolean, logText As String
Set tbl = RUN_FindMigoItemTable(session)
RUN_ReadItemLogIcons session, tbl, rowCount, hasRed, hasYellow, logText
Dim sbarType As String, sbarText As String
sbarType = RUN_GetSbarType(session)
sbarText = RUN_GetSbarText(session)
If sbarType = "W" Then hasYellow = True
If InStr(1, sbarText, "Function code cannot be selected", vbTextCompare) > 0 Then
statusOut = "Check Error"
msgOut = "Check chua active: SBar(" & sbarType & "): " & sbarText
GoTo CleanExit
End If
If RUN_IsFatalSbarType(sbarType) Or hasRed Then
If hasRed And mRunUncheckRedOK Then
Dim uncheckedRed As Long, redRowsCsv As String
redRowsCsv = RUN_GetRedRowsCsv(session, tbl, rowCount)
uncheckedRed = RUN_UncheckRedLinesAndReCheck(session, tbl, rowCount, redRowsCsv)
If uncheckedRed > 0 And uncheckedRed < rowCount Then
mLastExcludedRowsCsv = redRowsCsv
Set tbl = RUN_FindMigoItemTable(session)
popupText = RUN_ReadPopupTextIfAny(session)
If popupText <> "" Then RUN_DismissPopupIfAny session
RUN_ReadItemLogIconsExcludingCsv session, tbl, rowCount, mLastExcludedRowsCsv, hasRed, hasYellow, logText
sbarType = RUN_GetSbarType(session)
sbarText = RUN_GetSbarText(session)
If hasRed Then
statusOut = "Check Error"
msgOut = "Da bo tick OK dong RED nhung cac dong con lai van loi: SBar(" & sbarType & "): " & sbarText & IIf(logText <> "", " | " & logText, "")
GoTo CleanExit
End If
Else
statusOut = "Check Error"
msgOut = "Co dong RED nhung khong the bo tick OK hoac tat ca dong deu loi. SBar(" & sbarType & "): " & sbarText & IIf(logText <> "", " | " & logText, "") & IIf(popupText <> "", " | Popup=" & popupText, "")
GoTo CleanExit
End If
Else
statusOut = "Check Error"
msgOut = "SBar(" & sbarType & "): " & sbarText & IIf(logText <> "", " | " & logText, "") & IIf(popupText <> "", " | Popup=" & popupText, "")
GoTo CleanExit
End If
End If
Dim checkAutoContinueCount As Long
checkAutoContinueCount = 0
Do While sbarType = "W" Or RUN_HasPopup(session)
checkAutoContinueCount = checkAutoContinueCount + 1
If checkAutoContinueCount > 5 Then
statusOut = "Check Warning Loop"
msgOut = "SAP lap popup/canh bao VANG qua 5 lan sau Check. Da dung de tranh loop."
GoTo CleanExit
End If
popupText = RUN_ReadPopupTextIfAny(session)
If popupText <> "" Then
checkWarningTrace = checkWarningTrace & "Popup=" & popupText & "; "
If Not RUN_AcceptPopupIfAny(session) Then
statusOut = "Check Popup Error"
msgOut = "Khong tu dong xac nhan duoc popup sau Check: " & popupText
GoTo CleanExit
End If
Else
checkWarningTrace = checkWarningTrace & "W=" & sbarText & "; "
session.FindById("wnd[0]").sendVKey 0
End If
RUN_Wait session, 500
Set tbl = RUN_FindMigoItemTable(session)
RUN_ReadItemLogIcons session, tbl, rowCount, hasRed, hasYellow, logText
sbarType = RUN_GetSbarType(session)
sbarText = RUN_GetSbarText(session)
If RUN_IsFatalSbarType(sbarType) Or hasRed Then
statusOut = "Check Error"
msgOut = "Sau khi tu dong xac nhan YELLOW, SAP bao loi: SBar(" & sbarType & "): " & sbarText & IIf(logText <> "", " | " & logText, "")
GoTo CleanExit
End If
Loop
End If
Dim prePostSbarType As String
Dim prePostSbarText As String
prePostSbarType = RUN_GetSbarType(session)
prePostSbarText = RUN_GetSbarText(session)
Dim prePostDelivery As String
prePostDelivery = RUN_ExtractDeliveryNumber(prePostSbarText)
If Not RUN_IsSessionAlive(session) Then
statusOut = "SAP Session Lost"
msgOut = "SAP session bi mat ngay truoc khi bam Post."
GoTo CleanExit
End If
directPostTried = (hasAnyBatch And fastPostAfterBatch)
If Not RUN_PressPost(session) Then
    If directPostTried Then
        directPostFallbackUsed = True
        enter2Start = Timer
        session.FindById("wnd[0]").sendVKey 0
        RUN_WaitForStableIdle session, mCfgInputIdleTimeoutSeconds
        enter2Seconds = enter2Seconds + RUN_ElapsedSeconds(enter2Start)
        batchType = RUN_GetSbarType(session)
        batchText = RUN_GetSbarText(session)
        If RUN_IsFatalSbarType(batchType) Or InStr(1, batchText, "Enter ", vbTextCompare) > 0 Then
            statusOut = "Batch Error"
            msgOut = "FAST Post fallback: SAP van yeu cau Batch/Enter. SBar(" & batchType & "): " & batchText
            GoTo CleanExit
        End If
        If Not RUN_PressPost(session) Then
            statusOut = "Post Error"
            msgOut = "FAST Post fallback: da Enter nhung van khong bam duoc Post/Save."
            GoTo CleanExit
        End If
        directPostTried = False
    Else
        statusOut = "Post Error"
        msgOut = "Khong bam duoc Post/Save."
        GoTo CleanExit
    End If
End If
Dim postStartTime As Double
Dim successNoDeliveryStart As Double
Dim autoContinueCount As Long
Dim warningTrace As String
Dim postPopupText As String
Dim postSbarType As String
Dim postSbarText As String
Dim deliveryCandidate As String
Dim postStatusChanged As Boolean
Dim wasBusy As Boolean
postStartTime = Timer
successNoDeliveryStart = 0#
autoContinueCount = 0
warningTrace = IIf(mLastExcludedRowsCsv <> "", rescueDetail, "")
If mLastExcludedRowsCsv <> "" Then postRescueTried = True
postStatusChanged = False
wasBusy = False
PostPollLoop:
Do
If Not RUN_IsSessionAlive(session) Then
statusOut = "SAP Session Lost"
msgOut = "Mat SAP session trong luc cho ket qua Post."
deliveryOut = ""
GoTo CleanExit
End If
Do While session.busy
wasBusy = True
DoEvents
Sleep mCfgPostPollIntervalMs
If RUN_ElapsedSeconds(postStartTime) >= mCfgBusyTimeoutSeconds Then
statusOut = "Post Timeout"
msgOut = "SAP Busy qua " & CStr(mCfgBusyTimeoutSeconds) & _
" giay khi dang sinh Delivery."
deliveryOut = ""
GoTo CleanExit
End If
Loop
postPopupText = RUN_ReadPopupTextIfAny(session)
postSbarType = RUN_GetSbarType(session)
postSbarText = RUN_GetSbarText(session)
If postSbarType <> prePostSbarType Or postSbarText <> prePostSbarText Then
postStatusChanged = True
End If
If postSbarType = "S" And postPopupText = "" Then
deliveryCandidate = RUN_ExtractDeliveryNumber( _
postSbarText & " | " & postPopupText & " | " & warningTrace)
If deliveryCandidate <> "" And deliveryCandidate <> prePostDelivery Then
deliveryOut = deliveryCandidate
If mLastExcludedRowsCsv <> "" Then
statusOut = "Posted Partial"
Else
statusOut = "Posted"
End If
msgOut = postSbarText & " | Delivery OK." & _
IIf(checkWarningTrace <> "", " | Check warning: " & checkWarningTrace, "") & _
IIf(warningTrace <> "", " | Post warning: " & warningTrace, "")
GoTo CleanExit
End If
If postStatusChanged Or wasBusy Then
If mLastExcludedRowsCsv <> "" Then
statusOut = "Posted Partial"
Else
statusOut = "Posted"
End If
deliveryOut = ""
msgOut = postSbarText & " | Post OK."
GoTo CleanExit
End If
If successNoDeliveryStart = 0# Then successNoDeliveryStart = Timer
If RUN_ElapsedSeconds(successNoDeliveryStart) >= mCfgDeliveryGraceSeconds Then
If mLastExcludedRowsCsv <> "" Then
statusOut = "Posted Partial"
Else
statusOut = "Posted"
End If
deliveryOut = ""
msgOut = postSbarText & " | SAP=S; grace=" & Format$(mCfgDeliveryGraceSeconds, "0.0") & "s."
GoTo CleanExit
End If
Else
successNoDeliveryStart = 0#
End If
If directPostTried And Not directPostFallbackUsed And Not wasBusy And _
   postSbarText = prePostSbarText And _
   (InStr(1, postSbarText, "Enter ", vbTextCompare) > 0 Or InStr(1, postSbarText, "BATCH", vbTextCompare) > 0 Or _
    InStr(1, postSbarText, "Function code cannot be selected", vbTextCompare) > 0) Then
    directPostFallbackUsed = True
    directPostTried = False
    enter2Start = Timer
    session.FindById("wnd[0]").sendVKey 0
    RUN_WaitForStableIdle session, mCfgInputIdleTimeoutSeconds
    enter2Seconds = enter2Seconds + RUN_ElapsedSeconds(enter2Start)
    batchType = RUN_GetSbarType(session)
    batchText = RUN_GetSbarText(session)
    If RUN_IsFatalSbarType(batchType) Or InStr(1, batchText, "Enter ", vbTextCompare) > 0 Then
        statusOut = "Batch Error"
        msgOut = "FAST Post fallback: " & batchText
        GoTo CleanExit
    End If
    prePostSbarType = batchType
    prePostSbarText = batchText
    prePostDelivery = RUN_ExtractDeliveryNumber(prePostSbarText)
    If Not RUN_PressPost(session) Then
        statusOut = "Post Error"
        msgOut = "FAST Post fallback: khong bam lai duoc Post."
        GoTo CleanExit
    End If
    postStartTime = Timer
    successNoDeliveryStart = 0#
    autoContinueCount = 0
    postStatusChanged = False
    wasBusy = False
    GoTo PostPollLoop
End If
If Not postRescueTried And RUN_IsPostLineDataWarning(postSbarType, postSbarText, postPopupText) Then
postRescueTried = True
warningTrace = RUN_AppendMessage(warningTrace, "SAP data warning: " & postSbarText & IIf(postPopupText <> "", " | Popup=" & postPopupText, ""))
rescueRowsCsv = "": rescueDetail = ""
rescueCount = RUN_TryRescueRedLines(session, rowCount, rescueRowsCsv, rescueDetail)
If rescueCount > 0 And rescueCount < rowCount Then
mLastExcludedRowsCsv = rescueRowsCsv
warningTrace = RUN_AppendMessage(warningTrace, rescueDetail)
prePostSbarType = RUN_GetSbarType(session)
prePostSbarText = RUN_GetSbarText(session)
prePostDelivery = RUN_ExtractDeliveryNumber(prePostSbarText)
If Not RUN_PressPost(session) Then
statusOut = "Post Error"
msgOut = "Da bo tick dong RED nhung khong bam lai duoc Post."
GoTo CleanExit
End If
postStartTime = Timer
successNoDeliveryStart = 0#
autoContinueCount = 0
postStatusChanged = False
wasBusy = False
GoTo PostPollLoop
ElseIf rescueCount >= rowCount Then
statusOut = "Post Error - All Lines Invalid"
deliveryOut = ""
msgOut = "Tat ca dong cua OB deu RED; khong con dong hop le de Post. " & rescueDetail
GoTo CleanExit
Else
statusOut = "Post Data Error"
deliveryOut = ""
msgOut = "SAP bao loi du lieu theo dong; tool khong tu Enter lap. " & postSbarText & IIf(postPopupText <> "", " | Popup=" & postPopupText, "") & IIf(rescueDetail <> "", " | " & rescueDetail, "")
GoTo CleanExit
End If
End If
If RUN_IsFatalSbarType(postSbarType) Then
Dim postHasRed As Boolean
Dim postHasYellow As Boolean
Dim postLogText As String
Set tbl = RUN_FindMigoItemTable(session)
If Not tbl Is Nothing Then
RUN_ReadItemLogIcons session, tbl, rowCount, _
postHasRed, postHasYellow, postLogText
End If
statusOut = "Post Error"
deliveryOut = ""
msgOut = "Post bi SAP tu choi: SBar(" & postSbarType & "): " & _
postSbarText & _
IIf(postLogText <> "", _
" | Item log: " & postLogText, "") & _
IIf(checkWarningTrace <> "", _
" | Check warning: " & checkWarningTrace, "") & _
IIf(warningTrace <> "", _
" | Post warning: " & warningTrace, "")
GoTo CleanExit
End If
If postPopupText <> "" Or postSbarType = "W" Then
autoContinueCount = autoContinueCount + 1
If autoContinueCount > 5 Then
statusOut = "Post Warning Loop"
deliveryOut = ""
msgOut = "SAP lap popup/canh bao VANG qua 5 lan khi Post." & _
IIf(warningTrace <> "", _
" | " & warningTrace, "")
GoTo CleanExit
End If
If postPopupText <> "" Then
warningTrace = warningTrace & _
"Popup=" & postPopupText & "; "
If Not RUN_AcceptPopupIfAny(session) Then
statusOut = "Post Popup Blocked"
deliveryOut = ""
msgOut = "Popup SAP la/khong duoc phep tu dong xac nhan. " & _
"Hay xu ly thu cong roi chay lai OB nay: " & postPopupText
GoTo CleanExit
End If
Else
warningTrace = warningTrace & _
"W=" & postSbarText & "; "
session.FindById("wnd[0]").sendVKey 0
End If
RUN_Wait session, mCfgPostPollIntervalMs
Else
DoEvents
Sleep mCfgPostPollIntervalMs
End If
If RUN_ElapsedSeconds(postStartTime) >= mCfgBusyTimeoutSeconds Then
statusOut = "Post Timeout"
deliveryOut = ""
msgOut = "Qua " & CStr(mCfgBusyTimeoutSeconds) & _
" giay khong nhan duoc ket qua cuoi sau Post." & _
IIf(postSbarText <> "", _
" | SBar(" & postSbarType & "): " & postSbarText, "") & _
IIf(postPopupText <> "", _
" | Popup=" & postPopupText, "") & _
IIf(warningTrace <> "", _
" | Post warning: " & warningTrace, "")
GoTo CleanExit
End If
Loop
CleanExit:
Dim postWaitSeconds As Double
If postStartTime > 0# Then postWaitSeconds = RUN_ElapsedSeconds(postStartTime)
If fillMainSeconds > 0# Or fillBatchSeconds > 0# Or postWaitSeconds > 0# Then
    msgOut = RUN_AppendMessage(msgOut, "MIGO Stage Main=" & Format$(fillMainSeconds, "0.00") & "s; Enter1=" & Format$(enter1Seconds, "0.00") & "s; Batch=" & Format$(fillBatchSeconds, "0.00") & "s; Enter2=" & Format$(enter2Seconds, "0.00") & "s; Post=" & Format$(postWaitSeconds, "0.00") & "s" & IIf(directPostFallbackUsed, "; FastFallback=1", "; FastFallback=0"))
End If
Set colMapPre = Nothing
Set tbl = Nothing
Exit Sub
ErrHandler:
statusOut = "System Error"
msgOut = Err.Description
deliveryOut = ""
Resume CleanExit
End Sub
Private Function RUN_IsPostLineDataWarning(ByVal msgType As String, ByVal msgText As String, ByVal popupText As String) As Boolean
Dim t As String
t = UCase$(Trim$(msgText & " " & popupText))
If InStr(1, t, "DISPLAY LOG", vbTextCompare) > 0 Then RUN_IsPostLineDataWarning = True: Exit Function
If UCase$(Trim$(msgType)) <> "W" And Not RUN_IsFatalSbarType(msgType) Then Exit Function
Dim keys As Variant, k As Variant
keys = Array("BATCH", "DOES NOT EXIST", "NOT EXIST", "MATERIAL", "PRODUCT", "QUANTITY", "STOCK", "STORAGE LOCATION", "SLOC", "PLANT", "DEFICIT", "NOT AVAILABLE", "INSUFFICIENT", "INVALID", "NOT DEFINED", "NOT MAINTAINED")
For Each k In keys
If InStr(1, t, CStr(k), vbTextCompare) > 0 Then RUN_IsPostLineDataWarning = True: Exit Function
Next k
End Function
Private Function RUN_TryRescueRedLines(ByVal session As Object, ByVal rowCount As Long, ByRef redRowsCsv As String, ByRef detail As String) As Long
On Error GoTo RescueFail
Dim tbl As Object, hasRed As Boolean, hasYellow As Boolean, logText As String
Dim uncheckedCount As Long, sbarType As String, sbarText As String
If RUN_HasPopup(session) Then RUN_DismissPopupIfAny session
RUN_Wait session, mCfgPostPollIntervalMs
Set tbl = RUN_FindMigoItemTable(session)
If tbl Is Nothing Then detail = "Khong tim thay bang MIGO de cuu phan.": Exit Function
RUN_ReadItemLogIcons session, tbl, rowCount, hasRed, hasYellow, logText
redRowsCsv = RUN_GetRedRowsCsv(session, tbl, rowCount)
If redRowsCsv = "" Then
If RUN_PressCheck(session) Then
RUN_Wait session, 300
If RUN_HasPopup(session) Then RUN_DismissPopupIfAny session
Set tbl = RUN_FindMigoItemTable(session)
If Not tbl Is Nothing Then redRowsCsv = RUN_GetRedRowsCsv(session, tbl, rowCount)
End If
End If
If redRowsCsv = "" Then detail = "Khong xac dinh duoc dong RED de bo tick.": Exit Function
uncheckedCount = RUN_UncheckRedLinesAndReCheck(session, tbl, rowCount, redRowsCsv)
If uncheckedCount <= 0 Then detail = "Tim thay RED dong " & redRowsCsv & " nhung khong bo duoc tick OK.": Exit Function
If uncheckedCount >= rowCount Then detail = "RED dong " & redRowsCsv & "; tat ca dong bi loai.": RUN_TryRescueRedLines = uncheckedCount: Exit Function
Set tbl = RUN_FindMigoItemTable(session)
If tbl Is Nothing Then detail = "Mat bang MIGO sau khi bo tick RED.": Exit Function
RUN_ReadItemLogIconsExcludingCsv session, tbl, rowCount, redRowsCsv, hasRed, hasYellow, logText
sbarType = RUN_GetSbarType(session): sbarText = RUN_GetSbarText(session)
If hasRed Then detail = "Sau khi loai RED " & redRowsCsv & ", dong con lai van RED. " & logText: Exit Function
detail = "Da bo tick OK dong RED " & redRowsCsv & "; tiep tuc Post cac dong hop le." & IIf(sbarText <> "", " SBar=" & sbarText, "")
RUN_TryRescueRedLines = uncheckedCount
Exit Function
RescueFail:
detail = "Loi cuu phan " & CStr(Err.Number) & ": " & Err.Description
End Function
Private Function RUN_GetPreBatchColumnMap() As Object
Dim d As Object
Set d = CreateObject("Scripting.Dictionary")
d("Product") = 3
d("Qty") = 4
d("Text") = 5
d("Plant") = 6
d("SLoc") = 7
d("SLocTo") = 8
Set RUN_GetPreBatchColumnMap = d
End Function
Private Sub RUN_FillPreBatch_ByGrid(ByVal tbl As Object, ByVal colMap As Object, ByVal data As Variant, ByVal startRow As Long, ByVal endRow As Long)
Dim rowCount As Long, firstRow As Long, visibleRows As Long
Dim nextLogical As Long, lastLogical As Long
Dim logicalRow As Long, tableRow As Long, r As Long
Dim cleanTram As String, useGetCell As Boolean
rowCount = endRow - startRow + 1
If rowCount <= 0 Then Exit Sub
If Not RUN_ReadGridViewport(tbl, firstRow, visibleRows) Then
RUN_FillPreBatch_ByGrid_Legacy tbl, colMap, data, startRow, endRow
Exit Sub
End If
If firstRow <> 0 Then
If Not RUN_ScrollGridToLogicalRow(tbl, 0, firstRow, visibleRows) Then
Err.Raise vbObjectError + 605, , "Khong dua duoc bang MIGO ve dong dau truoc khi dien OB."
End If
End If
useGetCell = RUN_ProbeGetCell(tbl, 0, CLng(colMap("Product")))
nextLogical = 0
Do While nextLogical < rowCount
If nextLogical < firstRow Or nextLogical >= firstRow + visibleRows Then
If Not RUN_ScrollGridToLogicalRow(tbl, nextLogical, firstRow, visibleRows) Then
Err.Raise vbObjectError + 606, , "Khong cuon duoc bang MIGO toi dong " & CStr(nextLogical + 1) & "."
End If
End If
lastLogical = firstRow + visibleRows - 1
If lastLogical > rowCount - 1 Then lastLogical = rowCount - 1
If lastLogical < nextLogical Then
Err.Raise vbObjectError + 607, , "Khung hien thi MIGO khong chua dong can dien " & CStr(nextLogical + 1) & "."
End If
For logicalRow = nextLogical To lastLogical
r = startRow + logicalRow
tableRow = logicalRow - firstRow
RUN_SetGridCellMode tbl, tableRow, CLng(colMap("Product")), RUN_Clean(data(r, 3)), "Product", useGetCell
RUN_SetGridCellMode tbl, tableRow, CLng(colMap("Qty")), RUN_Clean(data(r, 4)), "Qty", useGetCell
cleanTram = RUN_CleanTextForSap(RUN_Clean(data(r, 2)))
If cleanTram = "" Then Err.Raise vbObjectError + 602, , "Text/Tram rong tai dong OB_GROUPS " & (r + 1)
RUN_SetGridCellMode tbl, tableRow, CLng(colMap("Text")), cleanTram, "Text/Tram", useGetCell
RUN_SetGridCellMode tbl, tableRow, CLng(colMap("Plant")), RUN_Clean(data(r, 5)), "Plant", useGetCell
RUN_SetGridCellMode tbl, tableRow, CLng(colMap("SLoc")), RUN_Clean(data(r, 6)), "SLoc", useGetCell
RUN_SetGridCellMode tbl, tableRow, CLng(colMap("SLocTo")), RUN_Clean(data(r, 7)), "SLoc Transfer", useGetCell
Next logicalRow
nextLogical = lastLogical + 1
Loop
End Sub
Private Sub RUN_FillPreBatch_ByGrid_Legacy(ByVal tbl As Object, ByVal colMap As Object, ByVal data As Variant, ByVal startRow As Long, ByVal endRow As Long)
Dim r As Long, logicalRow As Long, tableRow As Long
Dim cleanTram As String, useGetCell As Boolean
useGetCell = RUN_ProbeGetCell(tbl, 0, CLng(colMap("Product")))
For r = startRow To endRow
logicalRow = r - startRow
tableRow = RUN_EnsureVisibleRow(tbl, logicalRow)
RUN_SetGridCellMode tbl, tableRow, CLng(colMap("Product")), RUN_Clean(data(r, 3)), "Product", useGetCell
RUN_SetGridCellMode tbl, tableRow, CLng(colMap("Qty")), RUN_Clean(data(r, 4)), "Qty", useGetCell
cleanTram = RUN_CleanTextForSap(RUN_Clean(data(r, 2)))
If cleanTram = "" Then Err.Raise vbObjectError + 602, , "Text/Tram rong tai dong OB_GROUPS " & (r + 1)
RUN_SetGridCellMode tbl, tableRow, CLng(colMap("Text")), cleanTram, "Text/Tram", useGetCell
RUN_SetGridCellMode tbl, tableRow, CLng(colMap("Plant")), RUN_Clean(data(r, 5)), "Plant", useGetCell
RUN_SetGridCellMode tbl, tableRow, CLng(colMap("SLoc")), RUN_Clean(data(r, 6)), "SLoc", useGetCell
RUN_SetGridCellMode tbl, tableRow, CLng(colMap("SLocTo")), RUN_Clean(data(r, 7)), "SLoc Transfer", useGetCell
Next r
End Sub
Private Function RUN_CleanTextForSap(ByVal rawText As String) As String
RUN_CleanTextForSap = Trim$(rawText)
End Function
Private Sub RUN_FillBatchDynamic_ByGrid(ByVal tbl As Object, ByVal data As Variant, ByVal startRow As Long, ByVal endRow As Long)
Dim rowCount As Long, firstRow As Long, visibleRows As Long
Dim nextLogical As Long, lastLogical As Long
Dim logicalRow As Long, tableRow As Long, r As Long
Dim realBatchCol As Long, batchValue As String, useGetCell As Boolean
rowCount = endRow - startRow + 1
If rowCount <= 0 Then Exit Sub
realBatchCol = RUN_GetBatchColumnForRow(tbl, 0)
If realBatchCol < 0 Then
Err.Raise vbObjectError + 604, , _
"Khong xac dinh duoc cot Batch that tren SAP. Tool dung de tranh ghi nham Batch Trfr Pstg."
End If
If Not RUN_ReadGridViewport(tbl, firstRow, visibleRows) Then
RUN_FillBatchDynamic_ByGrid_Legacy tbl, data, startRow, endRow, realBatchCol
Exit Sub
End If
If firstRow <> 0 Then
If Not RUN_ScrollGridToLogicalRow(tbl, 0, firstRow, visibleRows) Then
Err.Raise vbObjectError + 608, , "Khong dua duoc bang MIGO ve dong dau truoc khi dien Batch."
End If
End If
useGetCell = RUN_ProbeGetCell(tbl, 0, realBatchCol)
nextLogical = 0
Do While nextLogical < rowCount
If nextLogical < firstRow Or nextLogical >= firstRow + visibleRows Then
If Not RUN_ScrollGridToLogicalRow(tbl, nextLogical, firstRow, visibleRows) Then
Err.Raise vbObjectError + 609, , "Khong cuon duoc bang MIGO toi dong Batch " & CStr(nextLogical + 1) & "."
End If
End If
lastLogical = firstRow + visibleRows - 1
If lastLogical > rowCount - 1 Then lastLogical = rowCount - 1
If lastLogical < nextLogical Then
Err.Raise vbObjectError + 610, , "Khung hien thi MIGO khong chua dong Batch " & CStr(nextLogical + 1) & "."
End If
For logicalRow = nextLogical To lastLogical
r = startRow + logicalRow
batchValue = RUN_Clean(data(r, 8))
If batchValue <> "" Then
tableRow = logicalRow - firstRow
RUN_SetGridCellMode tbl, tableRow, realBatchCol, batchValue, "Batch", useGetCell
End If
Next logicalRow
nextLogical = lastLogical + 1
Loop
End Sub
Private Sub RUN_FillBatchDynamic_ByGrid_Legacy(ByVal tbl As Object, ByVal data As Variant, ByVal startRow As Long, ByVal endRow As Long, ByVal realBatchCol As Long)
Dim r As Long, logicalRow As Long, tableRow As Long
Dim batchValue As String, useGetCell As Boolean
useGetCell = RUN_ProbeGetCell(tbl, 0, realBatchCol)
For r = startRow To endRow
batchValue = RUN_Clean(data(r, 8))
If batchValue <> "" Then
logicalRow = r - startRow
tableRow = RUN_EnsureVisibleRow(tbl, logicalRow)
RUN_SetGridCellMode tbl, tableRow, realBatchCol, batchValue, "Batch", useGetCell
End If
Next r
End Sub
Private Function RUN_OBHasAnyBatch(ByVal data As Variant, ByVal startRow As Long, ByVal endRow As Long) As Boolean
Dim r As Long
For r = startRow To endRow
If RUN_Clean(data(r, 8)) <> "" Then
RUN_OBHasAnyBatch = True
Exit Function
End If
Next r
End Function
Private Function RUN_IsExpectedBatchPrompt(ByVal statusText As String) As Boolean
statusText = UCase$(Trim$(statusText))
RUN_IsExpectedBatchPrompt = (InStr(1, statusText, "BATCH", vbTextCompare) > 0)
End Function
Private Function RUN_ReadGridViewport(ByVal tbl As Object, ByRef firstRow As Long, ByRef visibleRows As Long) As Boolean
Dim readFirstOK As Boolean, readVisibleOK As Boolean
firstRow = 0
visibleRows = 0
On Error Resume Next
Err.Clear
firstRow = CLng(tbl.FirstVisibleRow)
readFirstOK = (Err.Number = 0)
On Error GoTo 0
If Not readFirstOK Then
On Error Resume Next
Err.Clear
firstRow = CLng(tbl.VerticalScrollbar.Position)
readFirstOK = (Err.Number = 0)
On Error GoTo 0
End If
On Error Resume Next
Err.Clear
visibleRows = CLng(tbl.VisibleRowCount)
readVisibleOK = (Err.Number = 0)
On Error GoTo 0
If visibleRows <= 0 Then readVisibleOK = False
RUN_ReadGridViewport = (readFirstOK And readVisibleOK)
End Function
Private Function RUN_ScrollGridToLogicalRow(ByVal tbl As Object, ByVal logicalRow As Long, ByRef firstRow As Long, ByRef visibleRows As Long) As Boolean
Dim setOK As Boolean
On Error Resume Next
Err.Clear
tbl.FirstVisibleRow = logicalRow
setOK = (Err.Number = 0)
On Error GoTo 0
If Not setOK Then
On Error Resume Next
Err.Clear
tbl.VerticalScrollbar.Position = logicalRow
setOK = (Err.Number = 0)
On Error GoTo 0
End If
If Not setOK Then Exit Function
DoEvents
If mCfgScrollSettleMs > 0 Then Sleep mCfgScrollSettleMs
If Not RUN_ReadGridViewport(tbl, firstRow, visibleRows) Then Exit Function
RUN_ScrollGridToLogicalRow = (logicalRow >= firstRow And logicalRow < firstRow + visibleRows)
End Function
Private Function RUN_EnsureVisibleRow(ByVal tbl As Object, ByVal logicalRow As Long) As Long
Dim firstRow As Long, visibleRows As Long
If Not RUN_ReadGridViewport(tbl, firstRow, visibleRows) Then
Err.Raise vbObjectError + 611, , "Khong doc duoc khung hien thi cua bang MIGO."
End If
If logicalRow < firstRow Or logicalRow >= firstRow + visibleRows Then
If Not RUN_ScrollGridToLogicalRow(tbl, logicalRow, firstRow, visibleRows) Then
Err.Raise vbObjectError + 612, , "Khong cuon duoc bang MIGO toi dong " & CStr(logicalRow + 1) & "."
End If
End If
RUN_EnsureVisibleRow = logicalRow - firstRow
If RUN_EnsureVisibleRow < 0 Or RUN_EnsureVisibleRow >= visibleRows Then
Err.Raise vbObjectError + 613, , "Dong MIGO " & CStr(logicalRow + 1) & " nam ngoai khung hien thi sau khi cuon."
End If
End Function
Private Function RUN_GetBatchColumnForRow(ByVal tbl As Object, ByVal tableRow As Long) As Long
Dim i As Long, titleText As String, colCount As Long
Dim readOK As Boolean
RUN_GetBatchColumnForRow = -1
On Error Resume Next
Err.Clear
colCount = CLng(tbl.columns.count)
readOK = (Err.Number = 0)
On Error GoTo 0
If Not readOK Or colCount <= 0 Then Exit Function
If mCfgLayoutCache And mCachedBatchCol >= 0 And mCachedBatchCol < colCount And mCachedTableCols = colCount Then
On Error Resume Next
Err.Clear
titleText = UCase$(Trim$(CStr(tbl.columns.ElementAt(mCachedBatchCol).title)))
readOK = (Err.Number = 0)
On Error GoTo 0
If readOK And RUN_IsRealBatchTitle(titleText) Then
RUN_GetBatchColumnForRow = mCachedBatchCol
Exit Function
End If
mCachedBatchCol = -1
mCachedTableCols = 0
End If
For i = 0 To colCount - 1
On Error Resume Next
Err.Clear
titleText = UCase$(Trim$(CStr(tbl.columns.ElementAt(i).title)))
readOK = (Err.Number = 0)
On Error GoTo 0
If readOK Then
If RUN_IsRealBatchTitle(titleText) Then
RUN_GetBatchColumnForRow = i
If mCfgLayoutCache Then mCachedBatchCol = i: mCachedTableCols = colCount
Exit Function
End If
End If
Next i
End Function
Private Function RUN_IsRealBatchTitle(ByVal titleText As String) As Boolean
titleText = UCase$(Trim$(titleText))
RUN_IsRealBatchTitle = (InStr(1, titleText, "BATCH", vbTextCompare) > 0 _
And InStr(1, titleText, "TRFR", vbTextCompare) = 0 _
And InStr(1, titleText, "TRANSFER", vbTextCompare) = 0 _
And InStr(1, titleText, "PSTG", vbTextCompare) = 0)
End Function
Private Function RUN_TryGetCellControl(ByVal tbl As Object, ByVal rowNo As Long, ByVal colNo As Long, ByRef ctl As Object) As Boolean
Dim callOK As Boolean
Set ctl = Nothing
On Error Resume Next
Err.Clear
Set ctl = tbl.GetCell(rowNo, colNo)
callOK = (Err.Number = 0)
On Error GoTo 0
RUN_TryGetCellControl = (callOK And Not ctl Is Nothing)
End Function
Private Function RUN_ProbeGetCell(ByVal tbl As Object, ByVal rowNo As Long, ByVal colNo As Long) As Boolean
Dim ctl As Object
RUN_ProbeGetCell = RUN_TryGetCellControl(tbl, rowNo, colNo, ctl)
End Function
Private Function RUN_TrySetControlText(ByVal ctl As Object, ByVal valueText As String) As Boolean
Dim callOK As Boolean
If ctl Is Nothing Then Exit Function
On Error Resume Next
Err.Clear
ctl.text = valueText
callOK = (Err.Number = 0)
On Error GoTo 0
RUN_TrySetControlText = callOK
End Function
Private Function RUN_TryReadControlText(ByVal ctl As Object, ByRef valueText As String) As Boolean
Dim callOK As Boolean
valueText = ""
If ctl Is Nothing Then Exit Function
On Error Resume Next
Err.Clear
valueText = Trim$(CStr(ctl.text))
callOK = (Err.Number = 0)
On Error GoTo 0
RUN_TryReadControlText = callOK
End Function
Private Sub RUN_SetGridCellMode(ByVal tbl As Object, ByVal rowNo As Long, ByVal colNo As Long, ByVal valueText As String, ByVal label As String, ByRef useGetCell As Boolean)
Dim ctl As Object
If useGetCell Then
If RUN_TryGetCellControl(tbl, rowNo, colNo, ctl) Then
If RUN_TrySetControlText(ctl, valueText) Then Exit Sub
End If
useGetCell = False
End If
Set ctl = RUN_FindControlByCoord(tbl, colNo, rowNo)
If Not ctl Is Nothing Then
If RUN_TrySetControlText(ctl, valueText) Then Exit Sub
End If
Err.Raise vbObjectError + 601, , "Khong dien duoc " & label & " tai row=" & rowNo & ", col=" & colNo & ", value=" & valueText
End Sub
Private Sub RUN_SetGridCell(ByVal tbl As Object, ByVal rowNo As Long, ByVal colNo As Long, ByVal valueText As String, ByVal label As String)
Dim useGetCell As Boolean
useGetCell = True
RUN_SetGridCellMode tbl, rowNo, colNo, valueText, label, useGetCell
End Sub
Private Sub RUN_WriteResult(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, _
ByVal statusText As String, ByVal sapMessage As String, ByVal deliveryNo As String)
Dim r As Long, rowCount As Long
Dim outData() As Variant
On Error GoTo CleanExit
rowCount = endRow - startRow + 1
ReDim outData(1 To rowCount, 1 To 4)
For r = 1 To rowCount
If statusText = "Posted Partial" And RUN_CsvHasNumber(mLastExcludedRowsCsv, r) Then
outData(r, 1) = "Excluded Red Line"
outData(r, 2) = "Dong bi RED trong che do cuu phan, da BO TICK OK nen khong vao Delivery. " & sapMessage
Else
outData(r, 1) = IIf(statusText = "Posted Partial", "Posted", statusText)
outData(r, 2) = sapMessage
If deliveryNo <> "" Then outData(r, 3) = CStr(deliveryNo)
End If
outData(r, 4) = Now
Next r
With ws.Range(ws.Cells(startRow + 1, 9), ws.Cells(endRow + 1, 12))
.columns(3).NumberFormat = "@"
.Value = outData
End With
CleanExit:
mLastExcludedRowsCsv = ""
End Sub
Private Function RUN_CsvHasNumber(ByVal csv As String, ByVal valueNo As Long) As Boolean
RUN_CsvHasNumber = (InStr(1, "," & csv & ",", "," & CStr(valueNo) & ",", vbTextCompare) > 0)
End Function
Private Function RUN_PrecheckOBGroupsArray(ByRef data As Variant, ByVal ws As Worksheet) As Boolean
Dim ok As Boolean, n As Long, i As Long, startIdx As Long, endIdx As Long, r As Long
Dim groupName As String, errMsg As String, groupByTL As Boolean
Dim statusOut() As Variant
ok = True
n = UBound(data, 1)
groupByTL = ModuleHelper.ToolConfigBool("GROUP_BY_TL", True)
i = 1
Do While i <= n
    groupName = RUN_Clean(data(i, 1))
    If groupName = "" Then Exit Do
    startIdx = i
    endIdx = RUN_GetGroupEndInArray(data, i, n, groupName)
    errMsg = RUN_ValidateOneOBArray(data, startIdx, endIdx, groupByTL)
    If errMsg <> "" Then
        ok = False
        For r = startIdx To endIdx
            data(r, 9) = "Logic Error"
            data(r, 10) = errMsg
            ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + 1, 17)).Interior.Color = RGB(255, 199, 206)
        Next r
    End If
    i = endIdx + 1
Loop
If Not ok Then
    ReDim statusOut(1 To n, 1 To 2)
    For r = 1 To n
        statusOut(r, 1) = data(r, 9)
        statusOut(r, 2) = data(r, 10)
    Next r
    ws.Range("I2:J" & CStr(n + 1)).Value2 = statusOut
End If
RUN_PrecheckOBGroupsArray = ok
End Function
Private Function RUN_ValidateOneOBArray(ByRef data As Variant, ByVal startIdx As Long, ByVal endIdx As Long, ByVal groupByTL As Boolean) As String
Dim hasReady As Boolean, hasDelivery As Boolean, r As Long, errMsg As String
For r = startIdx To endIdx
    If UCase$(RUN_Clean(data(r, 9))) = "READY" Then hasReady = True
    If RUN_Clean(data(r, 11)) <> "" Then hasDelivery = True
Next r
If hasDelivery Or Not hasReady Then Exit Function
Dim tram0 As String, tramKey0 As String, tramKey As String, plant0 As String, sloc0 As String, slocTo0 As String
tram0 = RUN_CleanTextForSap(RUN_Clean(data(startIdx, 2)))
If groupByTL Then tramKey0 = ModuleHelper.TramTLKey(tram0) Else tramKey0 = UCase$(tram0)
plant0 = RUN_Clean(data(startIdx, 5))
sloc0 = RUN_Clean(data(startIdx, 6))
slocTo0 = RUN_Clean(data(startIdx, 7))
Dim dictProd As Object, dictProdBatchState As Object
Set dictProd = CreateObject("Scripting.Dictionary")
Set dictProdBatchState = CreateObject("Scripting.Dictionary")
Dim p As String, qtyText As String, batchState As Long, oldBatchState As Long
For r = startIdx To endIdx
    If UCase$(RUN_Clean(data(r, 9))) <> "READY" Then errMsg = errMsg & "Co dong trong cung don khong phai Ready; "
    If groupByTL Then
        tramKey = ModuleHelper.TramTLKey(RUN_CleanTextForSap(RUN_Clean(data(r, 2))))
    Else
        tramKey = UCase$(RUN_CleanTextForSap(RUN_Clean(data(r, 2))))
    End If
    If tramKey <> tramKey0 Then errMsg = errMsg & "Khac TL/Tram trong cung don; "
    If RUN_CleanTextForSap(RUN_Clean(data(r, 2))) = "" Then errMsg = errMsg & "Text/Tram rong; "
    If RUN_Clean(data(r, 5)) <> plant0 Then errMsg = errMsg & "Khac Plant trong cung don; "
    If RUN_Clean(data(r, 6)) <> sloc0 Then errMsg = errMsg & "Khac SLoc trong cung don; "
    If RUN_Clean(data(r, 7)) <> slocTo0 Then errMsg = errMsg & "Khac SLoc Transfer trong cung don; "
    If RUN_Clean(data(r, 3)) = "" Then errMsg = errMsg & "Thieu Product; "
    qtyText = RUN_Clean(data(r, 4))
    If qtyText = "" Or Not IsNumeric(qtyText) Or CDbl(qtyText) <= 0 Then errMsg = errMsg & "Qty loi; "
    If RUN_Clean(data(r, 5)) = "" Then errMsg = errMsg & "Thieu Plant; "
    If RUN_Clean(data(r, 6)) = "" Then errMsg = errMsg & "Thieu SLoc; "
    If RUN_Clean(data(r, 7)) = "" Then errMsg = errMsg & "Thieu SLoc Transfer; "
    If RUN_Clean(data(r, 6)) <> "" And RUN_Clean(data(r, 6)) = RUN_Clean(data(r, 7)) Then errMsg = errMsg & "SLoc = SLoc Transfer; "
    p = UCase$(RUN_Clean(data(r, 3)))
    If p <> "" Then
        If Not dictProd.Exists(p) Then dictProd.Add p, True
        If RUN_Clean(data(r, 8)) = "" Then batchState = 1 Else batchState = 2
        If Not dictProdBatchState.Exists(p) Then
            dictProdBatchState.Add p, batchState
        Else
            oldBatchState = CLng(dictProdBatchState(p))
            If oldBatchState <> batchState Then dictProdBatchState(p) = 3
        End If
    End If
Next r
Dim prodKey As Variant
For Each prodKey In dictProdBatchState.keys
    If CLng(dictProdBatchState(CStr(prodKey))) = 3 Then errMsg = errMsg & "Product " & CStr(prodKey) & " vua co Batch vua NO_BATCH; "
Next prodKey
If dictProd.count > mCfgMaxProductsPerOB Then errMsg = errMsg & "Qua " & mCfgMaxProductsPerOB & " Product khac nhau trong 1 don; "
If (endIdx - startIdx + 1) > mCfgMaxRowsPerOB Then errMsg = errMsg & "Qua " & mCfgMaxRowsPerOB & " dong trong 1 don; "
Set dictProd = Nothing
Set dictProdBatchState = Nothing
RUN_ValidateOneOBArray = errMsg
End Function
Private Function RUN_CountReadyOBsArray(ByRef data As Variant) As Long
Dim d As Object, r As Long, ob As String, n As Long
Set d = CreateObject("Scripting.Dictionary")
n = UBound(data, 1)
For r = 1 To n
    ob = RUN_Clean(data(r, 1))
    If ob <> "" And UCase$(RUN_Clean(data(r, 9))) = "READY" And RUN_Clean(data(r, 11)) = "" Then
        If Not d.Exists(ob) Then d.Add ob, True
    End If
Next r
RUN_CountReadyOBsArray = d.count
End Function
Private Function RUN_PrecheckOBGroups(ByVal ws As Worksheet, ByVal lastRow As Long) As Boolean
Dim ok As Boolean: ok = True
Dim i As Long, startRow As Long, endRow As Long, groupName As String
i = 2
Do While i <= lastRow
groupName = RUN_Clean(ws.Cells(i, 1).Value)
If groupName = "" Then Exit Do
startRow = i
endRow = i
Do While endRow < lastRow
If RUN_Clean(ws.Cells(endRow + 1, 1).Value) = groupName Then
endRow = endRow + 1
Else
Exit Do
End If
Loop
If Not RUN_ValidateOneOBOnSheet(ws, startRow, endRow) Then ok = False
i = endRow + 1
Loop
RUN_PrecheckOBGroups = ok
End Function
Private Function RUN_ValidateOneOBOnSheet(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long) As Boolean
Dim hasReady As Boolean, hasDelivery As Boolean
Dim r As Long, errMsg As String
For r = startRow To endRow
If UCase$(RUN_Clean(ws.Cells(r, 9).Value)) = "READY" Then hasReady = True
If RUN_Clean(ws.Cells(r, 11).Value) <> "" Then hasDelivery = True
Next r
If hasDelivery Or Not hasReady Then
RUN_ValidateOneOBOnSheet = True
Exit Function
End If
Dim tram0 As String, tramKey0 As String, tramKey As String, plant0 As String, sloc0 As String, slocTo0 As String
Dim groupByTL As Boolean
groupByTL = ModuleHelper.ToolConfigBool("GROUP_BY_TL", True)
tram0 = RUN_CleanTextForSap(RUN_Clean(ws.Cells(startRow, 2).Value))
If groupByTL Then tramKey0 = ModuleHelper.TramTLKey(tram0) Else tramKey0 = UCase$(tram0)
plant0 = RUN_Clean(ws.Cells(startRow, 5).Value)
sloc0 = RUN_Clean(ws.Cells(startRow, 6).Value)
slocTo0 = RUN_Clean(ws.Cells(startRow, 7).Value)
Dim dictProd As Object, dictProdBatchState As Object
Set dictProd = CreateObject("Scripting.Dictionary")
Set dictProdBatchState = CreateObject("Scripting.Dictionary")
Dim p As String, qtyText As String, batchState As Long, oldBatchState As Long
For r = startRow To endRow
If UCase$(RUN_Clean(ws.Cells(r, 9).Value)) <> "READY" Then
errMsg = errMsg & "Co dong trong cung don khong phai Ready; "
End If
If groupByTL Then
    tramKey = ModuleHelper.TramTLKey(RUN_CleanTextForSap(RUN_Clean(ws.Cells(r, 2).Value)))
Else
    tramKey = UCase$(RUN_CleanTextForSap(RUN_Clean(ws.Cells(r, 2).Value)))
End If
If tramKey <> tramKey0 Then errMsg = errMsg & "Khac TL/Tram trong cung don; "
If RUN_CleanTextForSap(RUN_Clean(ws.Cells(r, 2).Value)) = "" Then errMsg = errMsg & "Text/Tram rong; "
If RUN_Clean(ws.Cells(r, 5).Value) <> plant0 Then errMsg = errMsg & "Khac Plant trong cung don; "
If RUN_Clean(ws.Cells(r, 6).Value) <> sloc0 Then errMsg = errMsg & "Khac SLoc trong cung don; "
If RUN_Clean(ws.Cells(r, 7).Value) <> slocTo0 Then errMsg = errMsg & "Khac SLoc Transfer trong cung don; "
If RUN_Clean(ws.Cells(r, 3).Value) = "" Then errMsg = errMsg & "Thieu Product; "
qtyText = RUN_Clean(ws.Cells(r, 4).Value)
If qtyText = "" Or Not IsNumeric(qtyText) Or CDbl(qtyText) <= 0 Then errMsg = errMsg & "Qty loi; "
If RUN_Clean(ws.Cells(r, 5).Value) = "" Then errMsg = errMsg & "Thieu Plant; "
If RUN_Clean(ws.Cells(r, 6).Value) = "" Then errMsg = errMsg & "Thieu SLoc; "
If RUN_Clean(ws.Cells(r, 7).Value) = "" Then errMsg = errMsg & "Thieu SLoc Transfer; "
If RUN_Clean(ws.Cells(r, 6).Value) <> "" And RUN_Clean(ws.Cells(r, 6).Value) = RUN_Clean(ws.Cells(r, 7).Value) Then errMsg = errMsg & "SLoc = SLoc Transfer; "
p = UCase$(RUN_Clean(ws.Cells(r, 3).Value))
If p <> "" Then
If Not dictProd.Exists(p) Then dictProd.Add p, True
If RUN_Clean(ws.Cells(r, 8).Value) = "" Then batchState = 1 Else batchState = 2
If Not dictProdBatchState.Exists(p) Then
dictProdBatchState.Add p, batchState
Else
oldBatchState = CLng(dictProdBatchState(p))
If oldBatchState <> batchState Then dictProdBatchState(p) = 3
End If
End If
Next r
Dim prodKey As Variant
For Each prodKey In dictProdBatchState.keys
If CLng(dictProdBatchState(CStr(prodKey))) = 3 Then errMsg = errMsg & "Product " & CStr(prodKey) & " vua co Batch vua NO_BATCH; "
Next prodKey
If dictProd.count > mCfgMaxProductsPerOB Then errMsg = errMsg & "Qua " & mCfgMaxProductsPerOB & " Product khac nhau trong 1 don; "
If (endRow - startRow + 1) > mCfgMaxRowsPerOB Then errMsg = errMsg & "Qua " & mCfgMaxRowsPerOB & " dong trong 1 don; "
If errMsg <> "" Then
For r = startRow To endRow
ws.Cells(r, 9).Value = "Logic Error"
ws.Cells(r, 10).Value = errMsg
ws.Range(ws.Cells(r, 1), ws.Cells(r, 17)).Interior.Color = RGB(255, 199, 206)
Next r
RUN_ValidateOneOBOnSheet = False
Else
RUN_ValidateOneOBOnSheet = True
End If
Set dictProd = Nothing
Set dictProdBatchState = Nothing
End Function
Private Function RUN_CountReadyOBs(ByVal ws As Worksheet, ByVal lastRow As Long) As Long
Dim d As Object, r As Long, ob As String
Set d = CreateObject("Scripting.Dictionary")
For r = 2 To lastRow
ob = RUN_Clean(ws.Cells(r, 1).Value)
If ob <> "" And UCase$(RUN_Clean(ws.Cells(r, 9).Value)) = "READY" And RUN_Clean(ws.Cells(r, 11).Value) = "" Then
If Not d.Exists(ob) Then d.Add ob, True
End If
Next r
RUN_CountReadyOBs = d.count
End Function
Private Function RUN_PrepareReusedMIGO(ByRef session As Object, ByRef errorOut As String) As Boolean
On Error GoTo PrepError
errorOut = ""
If Not RUN_IsSessionAlive(session) Then errorOut = "SAP session khong con hoat dong.": Exit Function
If mBoundSessionKey <> "" And RUN_SessionKey(session) <> mBoundSessionKey Then errorOut = "SAP session da thay doi/bi chon nham tab; tool dung an toan.": Exit Function
If RUN_HasPopup(session) Then
If Not RUN_AcceptPopupIfAny(session) Then errorOut = "Popup chua duoc xu ly: " & RUN_ReadPopupTextIfAny(session): Exit Function
End If
RUN_WaitForStableIdle session, mCfgInputIdleTimeoutSeconds
If RUN_MigoHasResidualInput(session) Then
Application.StatusBar = "MIGO con du lieu tu luot truoc; dang reset an toan trong cung SAP session..."
If Not RUN_EnsureMIGO311Ready(session, errorOut, True) Then Exit Function
End If
RUN_PrepareReusedMIGO = True
Exit Function
PrepError:
errorOut = Err.Description
End Function
Private Function RUN_MigoHasResidualInput(ByVal session As Object) As Boolean
Dim tbl As Object, ctl As Object
Dim r As Long, txt As String, maxRows As Long
Dim firstRow As Long, visibleRows As Long
Set tbl = RUN_FindMigoItemTable(session)
If tbl Is Nothing Then Exit Function
' FAST: MIGO input lien tuc tu dong dau. Neu Product dong dau doc duoc va rong thi man hinh da sach.
Set ctl = Nothing
If Not RUN_TryGetCellControl(tbl, 0, 3, ctl) Then Set ctl = RUN_FindControlByCoord(tbl, 3, 0)
If RUN_TryReadControlText(ctl, txt) Then
    RUN_MigoHasResidualInput = (txt <> "")
    Exit Function
End If
' Fallback chi khi control dong dau khong doc duoc.
If RUN_ReadGridViewport(tbl, firstRow, visibleRows) Then
    maxRows = visibleRows
    If maxRows > 5 Then maxRows = 5
Else
    maxRows = 1
End If
If maxRows < 1 Then maxRows = 1
For r = 0 To maxRows - 1
    Set ctl = Nothing
    If Not RUN_TryGetCellControl(tbl, r, 3, ctl) Then Set ctl = RUN_FindControlByCoord(tbl, 3, r)
    If RUN_TryReadControlText(ctl, txt) Then
        If txt <> "" Then RUN_MigoHasResidualInput = True: Exit Function
    End If
Next r
End Function
Private Function RUN_FindExistingMIGOSession(ByVal fallbackSession As Object) As Object
On Error Resume Next
Dim sapApp As Object, conn As Object, sess As Object, firstMatch As Object, preferred As Object
Dim i As Long, j As Long
mMigoSessionMatchCount = 0
Set sapApp = modSupports.Apps
If sapApp Is Nothing Then
Dim sapGuiAutoLocal As Object
Set sapGuiAutoLocal = GetObject("SAPGUI")
If Not sapGuiAutoLocal Is Nothing Then Set sapApp = sapGuiAutoLocal.GetScriptingEngine
End If
If sapApp Is Nothing Then Exit Function
For i = 0 To sapApp.children.count - 1
Set conn = sapApp.children(i)
If Not conn Is Nothing Then
For j = 0 To conn.children.count - 1
Set sess = conn.children(j)
If RUN_SameSAPContext(sess, fallbackSession) And RUN_IsMIGOSession(sess) Then
mMigoSessionMatchCount = mMigoSessionMatchCount + 1
If firstMatch Is Nothing Then Set firstMatch = sess
If RUN_SessionKey(sess) = RUN_SessionKey(fallbackSession) Then Set preferred = sess
End If
Next j
End If
Next i
If Not preferred Is Nothing Then Set RUN_FindExistingMIGOSession = preferred Else Set RUN_FindExistingMIGOSession = firstMatch
On Error GoTo 0
End Function
Private Function RUN_SessionKey(ByVal sess As Object) As String
On Error Resume Next
RUN_SessionKey = CStr(sess.id) & "|" & CStr(sess.Info.SystemName) & "|" & CStr(sess.Info.Client) & "|" & CStr(sess.Info.User)
On Error GoTo 0
End Function
Private Function RUN_SameSAPContext(ByVal a As Object, ByVal b As Object) As Boolean
On Error GoTo Nope
RUN_SameSAPContext = (UCase$(CStr(a.Info.SystemName)) = UCase$(CStr(b.Info.SystemName)) And CStr(a.Info.Client) = CStr(b.Info.Client) And UCase$(CStr(a.Info.User)) = UCase$(CStr(b.Info.User)))
Exit Function
Nope:
RUN_SameSAPContext = False
End Function
Private Function RUN_IsMIGOSession(ByVal sess As Object) As Boolean
On Error Resume Next
If sess Is Nothing Then Exit Function
Dim trans As String
Dim titleText As String
trans = UCase$(Trim$(CStr(sess.Info.Transaction)))
titleText = UCase$(Trim$(CStr(sess.FindById("wnd[0]").text)))
RUN_IsMIGOSession = _
(trans = "MIGO") Or _
(InStr(1, titleText, "MIGO", vbTextCompare) > 0) Or _
RUN_SessionHasMigoTable(sess)
On Error GoTo 0
End Function
Private Function RUN_EnsureMIGO311Ready(ByRef session As Object, _
ByRef errorOut As String, _
Optional ByVal forceReset As Boolean = False) As Boolean
On Error GoTo OpenError
errorOut = ""
If session Is Nothing Then
errorOut = "SAP session = Nothing."
Exit Function
End If
If forceReset Then RUN_ClearMigoCache
If Not forceReset Then
If RUN_IsMIGO311ModeReady(session) Then
RUN_EnsureMIGO311Ready = True
Exit Function
End If
End If
RUN_ActivateSAPWindow session
Dim openStart As Double
If forceReset Or Not RUN_IsMIGOSession(session) Then
RUN_ClearMigoCache
session.FindById("wnd[0]/tbar[0]/okcd").text = "/NMIGO"
session.FindById("wnd[0]").sendVKey 0
openStart = Timer
Do
Do While session.busy
DoEvents
Sleep mCfgMigoUiPollMs
If RUN_ElapsedSeconds(openStart) >= mCfgMigoOpenTimeoutSeconds Then
errorOut = "SAP Busy qua " & CStr(mCfgMigoOpenTimeoutSeconds) & " giay khi mo /NMIGO."
Exit Function
End If
Loop
If RUN_HasPopup(session) Then
If Not RUN_AcceptPopupIfAny(session) Then
errorOut = "Co popup khi mo MIGO nhung khong dong/xac nhan duoc: " & RUN_ReadPopupTextIfAny(session)
Exit Function
End If
End If
If RUN_IsMIGOSession(session) Then Exit Do
DoEvents
Sleep mCfgMigoUiPollMs
If RUN_ElapsedSeconds(openStart) >= mCfgMigoOpenTimeoutSeconds Then
errorOut = "Da gui /NMIGO nhung SAP khong chuyen sang transaction MIGO."
Exit Function
End If
Loop
End If
Dim rootWindow As Object
Dim actionCombo As Object
Dim refCombo As Object
Dim movementField As Object
Set rootWindow = session.FindById("wnd[0]")
openStart = Timer
Do
Set actionCombo = RUN_FindChildByNameExact(rootWindow, "GODYNPRO-ACTION")
If Not actionCombo Is Nothing Then Exit Do
DoEvents
Sleep mCfgMigoUiPollMs
If RUN_ElapsedSeconds(openStart) >= mCfgMigoOpenTimeoutSeconds Then
errorOut = "Da mo MIGO nhung khong tim thay combobox GODYNPRO-ACTION."
Exit Function
End If
Loop
If Not RUN_SetComboByCaption(actionCombo, "TRANSFER POSTING", "TRANSFER POST") Then
errorOut = "Khong chon duoc ACTION = Transfer Posting. Entries: " & _
RUN_ListComboEntries(actionCombo)
Exit Function
End If
RUN_Wait session, 60
Set rootWindow = session.FindById("wnd[0]")
Set refCombo = RUN_FindChildByNameExact(rootWindow, "GODYNPRO-REFDOC")
If refCombo Is Nothing Then
errorOut = "Khong tim thay combobox GODYNPRO-REFDOC sau khi chon Transfer Posting."
Exit Function
End If
If Not RUN_SetComboByCaption(refCombo, "OTHER", "") Then
errorOut = "Khong chon duoc REFERENCE = Other. Entries: " & _
RUN_ListComboEntries(refCombo)
Exit Function
End If
RUN_Wait session, 60
Set rootWindow = session.FindById("wnd[0]")
Set movementField = RUN_FindChildByNameExact(rootWindow, "GODEFAULT_TV-BWART")
If movementField Is Nothing Then
errorOut = "Khong tim thay Movement Type field GODEFAULT_TV-BWART."
Exit Function
End If
movementField.text = "311"
session.FindById("wnd[0]").sendVKey 0
RUN_Wait session, 60
Dim sbarType As String
Dim sbarText As String
sbarType = RUN_GetSbarType(session)
sbarText = RUN_GetSbarText(session)
If RUN_IsFatalSbarType(sbarType) Then
errorOut = "SAP khong chap nhan MIGO 311: SBar(" & sbarType & "): " & sbarText
Exit Function
End If
If Not RUN_SessionHasMigoTable(session) Then
On Error Resume Next
session.SendCommand "MIGO_OK_HEADER_CLOSE"
Err.Clear
session.SendCommand "MIGO_OK_DETAIL_CLOSE"
On Error GoTo OpenError
RUN_Wait session, 60
End If
openStart = Timer
Do
If RUN_IsMIGO311ModeReady(session) Then
RUN_EnsureMIGO311Ready = True
Exit Function
End If
If RUN_HasPopup(session) Then
If Not RUN_AcceptPopupIfAny(session) Then
errorOut = "Popup khi cau hinh MIGO 311 khong xu ly duoc: " & _
RUN_ReadPopupTextIfAny(session)
Exit Function
End If
End If
DoEvents
Sleep mCfgMigoUiPollMs
If RUN_ElapsedSeconds(openStart) >= mCfgMigoOpenTimeoutSeconds Then
errorOut = "MIGO da mo nhung chua san sang dung mode " & _
"Transfer Posting -> Other -> 311 hoac khong tim thay bang GOITEM." & _
IIf(sbarText <> "", " SBar: " & sbarText, "")
Exit Function
End If
Loop
OpenError:
errorOut = "Auto Open MIGO 311 loi: " & Err.Description
RUN_EnsureMIGO311Ready = False
End Function
Private Function RUN_IsMIGO311ModeReady(ByVal session As Object) As Boolean
On Error Resume Next
If session Is Nothing Then Exit Function
If Not RUN_SessionHasMigoTable(session) Then Exit Function
Dim rootWindow As Object
Dim actionCombo As Object
Dim refCombo As Object
Dim movementField As Object
Dim actionText As String
Dim refText As String
Dim movementText As String
Set rootWindow = session.FindById("wnd[0]")
Set actionCombo = RUN_FindChildByNameExact(rootWindow, "GODYNPRO-ACTION")
Set refCombo = RUN_FindChildByNameExact(rootWindow, "GODYNPRO-REFDOC")
Set movementField = RUN_FindChildByNameExact(rootWindow, "GODEFAULT_TV-BWART")
actionText = UCase$(RUN_GetComboSelectedText(actionCombo))
refText = UCase$(RUN_GetComboSelectedText(refCombo))
If Not movementField Is Nothing Then movementText = Trim$(CStr(movementField.text))
RUN_IsMIGO311ModeReady = _
(InStr(1, actionText, "TRANSFER", vbTextCompare) > 0) And _
(InStr(1, refText, "OTHER", vbTextCompare) > 0) And _
(movementText = "311")
On Error GoTo 0
End Function
Private Function RUN_FindChildByNameExact(ByVal obj As Object, _
ByVal targetName As String) As Object
On Error Resume Next
If obj Is Nothing Then Exit Function
If StrComp(Trim$(CStr(obj.Name)), targetName, vbTextCompare) = 0 Then
Set RUN_FindChildByNameExact = obj
Exit Function
End If
Dim child As Object
Dim found As Object
For Each child In obj.children
Set found = RUN_FindChildByNameExact(child, targetName)
If Not found Is Nothing Then
Set RUN_FindChildByNameExact = found
Exit Function
End If
Next child
On Error GoTo 0
End Function
Private Function RUN_SetComboByCaption(ByVal combo As Object, _
ByVal caption1 As String, _
Optional ByVal caption2 As String = "") As Boolean
On Error Resume Next
If combo Is Nothing Then Exit Function
Dim entry As Object
Dim entryText As String
Dim entryKey As String
Dim normalizedText As String
Dim needle1 As String
Dim needle2 As String
needle1 = UCase$(Trim$(caption1))
needle2 = UCase$(Trim$(caption2))
For Each entry In combo.Entries
Err.Clear
entryText = CStr(entry.Value)
If entryText = "" Then entryText = CStr(entry.text)
entryKey = CStr(entry.key)
normalizedText = UCase$(Trim$(entryText))
If (needle1 <> "" And InStr(1, normalizedText, needle1, vbTextCompare) > 0) Or _
(needle2 <> "" And InStr(1, normalizedText, needle2, vbTextCompare) > 0) Then
combo.key = entryKey
If Err.Number = 0 Then
RUN_SetComboByCaption = True
Exit Function
End If
End If
Next entry
On Error GoTo 0
End Function
Private Function RUN_GetComboSelectedText(ByVal combo As Object) As String
On Error Resume Next
If combo Is Nothing Then Exit Function
Dim txt As String
Dim selectedKey As String
Dim entry As Object
txt = Trim$(CStr(combo.text))
selectedKey = CStr(combo.key)
If txt = "" Then
For Each entry In combo.Entries
If CStr(entry.key) = selectedKey Then
txt = CStr(entry.Value)
If txt = "" Then txt = CStr(entry.text)
Exit For
End If
Next entry
End If
RUN_GetComboSelectedText = Trim$(txt)
On Error GoTo 0
End Function
Private Function RUN_ListComboEntries(ByVal combo As Object) As String
On Error Resume Next
If combo Is Nothing Then Exit Function
Dim entry As Object
Dim txt As String
Dim itemText As String
For Each entry In combo.Entries
itemText = CStr(entry.Value)
If itemText = "" Then itemText = CStr(entry.text)
If itemText <> "" Then
If Len(txt) > 250 Then Exit For
txt = txt & IIf(txt <> "", " | ", "") & _
CStr(entry.key) & "=" & itemText
End If
Next entry
RUN_ListComboEntries = txt
On Error GoTo 0
End Function
Private Sub RUN_ActivateSAPWindow(ByVal session As Object)
On Error Resume Next
session.FindById("wnd[0]").maximize
session.FindById("wnd[0]").SetFocus
DoEvents
On Error GoTo 0
End Sub
Private Function RUN_SessionHasMigoTable(ByVal sess As Object) As Boolean
On Error Resume Next
If sess Is Nothing Then Exit Function
Dim tbl As Object
Set tbl = RUN_FindMigoItemTable(sess)
RUN_SessionHasMigoTable = Not tbl Is Nothing
On Error GoTo 0
End Function
Private Sub RUN_ClearMigoCache()
Set mCachedMigoTable = Nothing
mCachedMigoTableId = ""
mCachedSessionKey = ""
mCachedBatchCol = -1
mCachedTableCols = 0
mObSinceLayoutCheck = 0
End Sub
Private Function RUN_FindMigoItemTable(ByVal session As Object) As Object
On Error Resume Next
Dim rootWindow As Object, tbl As Object, sid As String, tid As String
sid = RUN_SessionKey(session)
If mCfgLayoutCache And Not mCachedMigoTable Is Nothing And sid = mCachedSessionKey Then
tid = "": Err.Clear: tid = CStr(mCachedMigoTable.id)
If Err.Number = 0 And tid = mCachedMigoTableId Then Set RUN_FindMigoItemTable = mCachedMigoTable: Exit Function
RUN_ClearMigoCache
End If
Set tbl = session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM", False)
If tbl Is Nothing Then Set tbl = session.FindById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM", False)
Set rootWindow = session.FindById("wnd[0]", False)
If rootWindow Is Nothing Then Set rootWindow = session.FindById("wnd[0]")
If tbl Is Nothing And Not rootWindow Is Nothing Then Set tbl = RUN_FindChildByIdContains(rootWindow, "tblSAPLMIGOTV_GOITEM")
If tbl Is Nothing And Not rootWindow Is Nothing Then Set tbl = RUN_FindChildByIdContains(rootWindow, "GOITEM")
If tbl Is Nothing And Not rootWindow Is Nothing Then Set tbl = RUN_FindMigoTableByType(rootWindow)
If Not tbl Is Nothing Then
Set RUN_FindMigoItemTable = tbl
If mCfgLayoutCache Then
Set mCachedMigoTable = tbl
mCachedMigoTableId = CStr(tbl.id)
mCachedSessionKey = sid
mCachedTableCols = CLng(tbl.columns.count)
End If
End If
End Function
Private Function RUN_FindMigoTableByType(ByVal obj As Object) As Object
On Error Resume Next
Dim objectType As String
Dim objectId As String
objectType = UCase$(CStr(obj.Type))
objectId = UCase$(CStr(obj.id))
If (objectType = "GUITABLECONTROL" Or objectType = "GUIGRIDVIEW") And _
(InStr(1, objectId, "MIGO", vbTextCompare) > 0 Or _
InStr(1, objectId, "GOITEM", vbTextCompare) > 0) Then
Set RUN_FindMigoTableByType = obj
Exit Function
End If
Dim child As Object
Dim found As Object
For Each child In obj.children
Set found = RUN_FindMigoTableByType(child)
If Not found Is Nothing Then
Set RUN_FindMigoTableByType = found
Exit Function
End If
Next child
End Function
Private Function RUN_FindChildByIdContains(ByVal obj As Object, ByVal needle As String) As Object
On Error Resume Next
If InStr(1, CStr(obj.id), needle, vbTextCompare) > 0 Then
Set RUN_FindChildByIdContains = obj
Exit Function
End If
Dim child As Object, found As Object
For Each child In obj.children
Set found = RUN_FindChildByIdContains(child, needle)
If Not found Is Nothing Then
Set RUN_FindChildByIdContains = found
Exit Function
End If
Next child
End Function
Private Function RUN_FindControlByCoord(ByVal obj As Object, ByVal colNo As Long, ByVal rowNo As Long) As Object
On Error Resume Next
Dim cid As String
cid = CStr(obj.id)
If Right$(cid, Len("[" & colNo & "," & rowNo & "]")) = "[" & colNo & "," & rowNo & "]" Then
Set RUN_FindControlByCoord = obj
Exit Function
End If
Dim child As Object, found As Object
For Each child In obj.children
Set found = RUN_FindControlByCoord(child, colNo, rowNo)
If Not found Is Nothing Then
Set RUN_FindControlByCoord = found
Exit Function
End If
Next child
End Function
Private Function RUN_PressCheck(ByVal session As Object) As Boolean
On Error Resume Next
Err.Clear
session.FindById("wnd[0]/tbar[1]/btn[7]").Press
If Err.Number = 0 Then
RUN_PressCheck = True
Exit Function
End If
Err.Clear
session.SendCommand "MIGO_CHECK"
RUN_PressCheck = (Err.Number = 0)
End Function
Private Function RUN_PressPost(ByVal session As Object) As Boolean
On Error Resume Next
Err.Clear
session.FindById("wnd[0]/tbar[1]/btn[23]").Press
If Err.Number = 0 Then
RUN_PressPost = True
Exit Function
End If
Err.Clear
session.FindById("wnd[0]/tbar[0]/btn[11]").Press
If Err.Number = 0 Then
RUN_PressPost = True
Exit Function
End If
Err.Clear
session.SendCommand "MIGO_POST"
RUN_PressPost = (Err.Number = 0)
End Function
Private Sub RUN_ReadItemLogIcons(ByVal session As Object, ByVal tbl As Object, ByVal rowCount As Long, ByRef hasRed As Boolean, ByRef hasYellow As Boolean, ByRef logText As String)
Dim r As Long, iconName As String, visibleRow As Long
hasRed = False
hasYellow = False
logText = ""
If tbl Is Nothing Then Exit Sub
For r = 0 To rowCount - 1
visibleRow = RUN_EnsureVisibleRow(tbl, r)
iconName = RUN_GetLogIconName(session, tbl, visibleRow)
If iconName = "S_TL_R" Or iconName = "S_LED_RED" Then
hasRed = True
logText = logText & "D" & (r + 1) & ":RED; "
ElseIf iconName = "S_TL_Y" Or iconName = "S_LED_YELLOW" Then
hasYellow = True
logText = logText & "D" & (r + 1) & ":YELLOW; "
End If
Next r
End Sub
Private Function RUN_GetLogIconName(ByVal session As Object, ByVal tbl As Object, ByVal visibleRow As Long) As String
On Error Resume Next
RUN_GetLogIconName = CStr(session.FindById(tbl.id & "/btnGOITEM-LOGLEVEL[1," & visibleRow & "]").iconName)
End Function
Private Function RUN_GetRedRowsCsv(ByVal session As Object, ByVal tbl As Object, ByVal rowCount As Long) As String
Dim r As Long, visibleRow As Long, iconName As String, csv As String
If tbl Is Nothing Then Exit Function
For r = 0 To rowCount - 1
visibleRow = RUN_EnsureVisibleRow(tbl, r)
iconName = RUN_GetLogIconName(session, tbl, visibleRow)
If iconName = "S_TL_R" Or iconName = "S_LED_RED" Then
csv = csv & IIf(csv <> "", ",", "") & CStr(r + 1)
End If
Next r
RUN_GetRedRowsCsv = csv
End Function
Private Function RUN_UncheckRedLinesAndReCheck(ByVal session As Object, ByVal tbl As Object, ByVal rowCount As Long, ByVal redRowsCsv As String) As Long
On Error Resume Next
Dim r As Long, visibleRow As Long, iconName As String, uncheckedCount As Long
uncheckedCount = 0
If tbl Is Nothing Then Exit Function
For r = 0 To rowCount - 1
visibleRow = RUN_EnsureVisibleRow(tbl, r)
iconName = RUN_GetLogIconName(session, tbl, visibleRow)
If iconName = "S_TL_R" Or iconName = "S_LED_RED" Then
If RUN_SetItemOKSelected(tbl, visibleRow, False) Then
uncheckedCount = uncheckedCount + 1
End If
End If
Next r
If uncheckedCount > 0 Then
session.FindById("wnd[0]").sendVKey 0
RUN_Wait session, 400
RUN_PressCheck session
RUN_Wait session, 700
If RUN_ReadPopupTextIfAny(session) <> "" Then RUN_DismissPopupIfAny session
End If
RUN_UncheckRedLinesAndReCheck = uncheckedCount
On Error GoTo 0
End Function
Private Function RUN_SetItemOKSelected(ByVal tbl As Object, ByVal visibleRow As Long, ByVal selectedValue As Boolean) As Boolean
On Error Resume Next
Dim candidates As Variant, c As Variant, okCol As Long
Dim ctl As Object
okCol = RUN_FindOKColumn(tbl)
candidates = Array(okCol, 14, 13, 5, 4)
For Each c In candidates
If CLng(c) >= 0 Then
Err.Clear
Set ctl = tbl.GetCell(visibleRow, CLng(c))
If Err.Number = 0 And Not ctl Is Nothing Then
Err.Clear
ctl.selected = selectedValue
If Err.Number = 0 Then
RUN_SetItemOKSelected = True
Exit Function
End If
End If
End If
Next c
Set ctl = RUN_FindCheckboxByRow(tbl, visibleRow)
If Not ctl Is Nothing Then
Err.Clear
ctl.selected = selectedValue
If Err.Number = 0 Then RUN_SetItemOKSelected = True
End If
On Error GoTo 0
End Function
Private Function RUN_FindOKColumn(ByVal tbl As Object) As Long
On Error Resume Next
Dim i As Long, titleText As String
RUN_FindOKColumn = -1
For i = 0 To tbl.columns.count - 1
titleText = UCase$(Trim$(CStr(tbl.columns.ElementAt(i).title)))
If titleText = "OK" Then
RUN_FindOKColumn = i
Exit Function
End If
Next i
On Error GoTo 0
End Function
Private Function RUN_FindCheckboxByRow(ByVal obj As Object, ByVal rowNo As Long) As Object
On Error Resume Next
Dim child As Object, found As Object
Dim cid As String, typ As String
cid = CStr(obj.id)
typ = UCase$(CStr(obj.Type))
If InStr(1, typ, "CHECK", vbTextCompare) > 0 Or InStr(1, typ, "GuiCheckBox", vbTextCompare) > 0 Then
If InStr(1, cid, "," & CStr(rowNo) & "]", vbTextCompare) > 0 Then
Set RUN_FindCheckboxByRow = obj
Exit Function
End If
End If
For Each child In obj.children
Set found = RUN_FindCheckboxByRow(child, rowNo)
If Not found Is Nothing Then
Set RUN_FindCheckboxByRow = found
Exit Function
End If
Next child
End Function
Private Sub RUN_ReadItemLogIconsExcludingCsv(ByVal session As Object, ByVal tbl As Object, ByVal rowCount As Long, ByVal excludeCsv As String, ByRef hasRed As Boolean, ByRef hasYellow As Boolean, ByRef logText As String)
Dim r As Long, visibleRow As Long, iconName As String, logicalIdx As Long
hasRed = False
hasYellow = False
logText = ""
If tbl Is Nothing Then Exit Sub
For r = 0 To rowCount - 1
logicalIdx = r + 1
visibleRow = RUN_EnsureVisibleRow(tbl, r)
iconName = RUN_GetLogIconName(session, tbl, visibleRow)
If iconName = "S_TL_R" Or iconName = "S_LED_RED" Then
If RUN_CsvHasNumber(excludeCsv, logicalIdx) Then
logText = logText & "D" & logicalIdx & ":RED/OK_OFF; "
Else
hasRed = True
logText = logText & "D" & logicalIdx & ":RED; "
End If
ElseIf iconName = "S_TL_Y" Or iconName = "S_LED_YELLOW" Then
If Not RUN_CsvHasNumber(excludeCsv, logicalIdx) Then hasYellow = True
logText = logText & "D" & logicalIdx & ":YELLOW; "
End If
Next r
End Sub
Private Function RUN_HasPopup(ByVal session As Object) As Boolean
On Error Resume Next
RUN_HasPopup = (session.children.count > 1 Or session.ActiveWindow.Name = "wnd[1]")
On Error GoTo 0
End Function
Private Function RUN_ReadPopupTextIfAny(ByVal session As Object) As String
On Error GoTo SafeExit
Dim wnd As Object
Dim txt As String
Dim bodyText As String
If RUN_HasPopup(session) Then
Set wnd = session.FindById("wnd[1]")
txt = Trim$(CStr(wnd.text))
bodyText = RUN_CollectPopupText(wnd, 0)
If bodyText <> "" Then txt = RUN_AppendMessage(txt, bodyText)
End If
RUN_ReadPopupTextIfAny = Trim$(txt)
Exit Function
SafeExit:
RUN_ReadPopupTextIfAny = Trim$(txt)
End Function
Private Function RUN_CollectPopupText(ByVal parent As Object, ByVal depth As Long) As String
On Error Resume Next
If parent Is Nothing Then Exit Function
If depth > 5 Then Exit Function
Dim i As Long
Dim child As Object
Dim txt As String
Dim nestedText As String
Dim resultText As String
For i = 0 To parent.children.count - 1
Set child = parent.children(i)
txt = Trim$(CStr(child.text))
If txt <> "" Then resultText = RUN_AppendMessage(resultText, txt)
If child.children.count > 0 Then
nestedText = RUN_CollectPopupText(child, depth + 1)
If nestedText <> "" Then resultText = RUN_AppendMessage(resultText, nestedText)
End If
Next i
RUN_CollectPopupText = resultText
On Error GoTo 0
End Function
Private Function RUN_AcceptPopupIfAny(ByVal session As Object) As Boolean
On Error GoTo SafeExit
If Not RUN_HasPopup(session) Then
RUN_AcceptPopupIfAny = True
Exit Function
End If
Dim popupText As String
Dim isKnownSafe As Boolean
popupText = RUN_ReadPopupTextIfAny(session)
If RUN_IsSecurityPopupText(popupText) Then
RUN_AcceptPopupIfAny = False
Exit Function
End If
If RUN_IsBlockingBusinessPopupText(popupText) Then
RUN_AcceptPopupIfAny = False
Exit Function
End If
If Not mCfgAutoAcceptKnownPopups Then
RUN_AcceptPopupIfAny = False
Exit Function
End If
isKnownSafe = RUN_IsKnownSafePopupText(popupText)
If Not isKnownSafe And mOneClickMigo Then
isKnownSafe = RUN_HasAffirmativePopupButton(session) And Trim$(popupText) <> ""
End If
If Not isKnownSafe And mCfgStopOnUnknownPopup Then
RUN_AcceptPopupIfAny = False
Exit Function
End If
On Error Resume Next
Err.Clear
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
If Err.Number <> 0 Then
Err.Clear
session.FindById("wnd[1]/usr/btnSPOP-OPTION1").Press
End If
If Err.Number <> 0 Then
Err.Clear
session.FindById("wnd[1]").sendVKey 0
End If
On Error GoTo SafeExit
RUN_Wait session, mCfgPostPollIntervalMs
RUN_AcceptPopupIfAny = Not RUN_HasPopup(session)
Exit Function
SafeExit:
RUN_AcceptPopupIfAny = False
End Function
Private Function RUN_IsSecurityPopupText(ByVal popupText As String) As Boolean
Dim t As String
t = UCase$(Trim$(popupText))
RUN_IsSecurityPopupText = _
(InStr(1, t, "AUTHORIZATION", vbTextCompare) > 0) Or _
(InStr(1, t, "AUTHORISATION", vbTextCompare) > 0) Or _
(InStr(1, t, "NOT AUTHORIZED", vbTextCompare) > 0) Or _
(InStr(1, t, "SECURITY", vbTextCompare) > 0) Or _
(InStr(1, t, "SCRIPT", vbTextCompare) > 0) Or _
(InStr(1, t, "PASSWORD", vbTextCompare) > 0) Or _
(InStr(1, t, "LOGON", vbTextCompare) > 0) Or _
(InStr(1, t, "LOGIN", vbTextCompare) > 0) Or _
(InStr(1, t, "PERMISSION", vbTextCompare) > 0) Or _
(InStr(1, t, "CREDENTIAL", vbTextCompare) > 0) Or _
(InStr(1, t, "ACCESS DENIED", vbTextCompare) > 0) Or _
(InStr(1, t, "KHONG CO QUYEN", vbTextCompare) > 0) Or _
(InStr(1, t, "KHONG DUOC PHEP", vbTextCompare) > 0) Or _
(InStr(1, t, "UY QUYEN", vbTextCompare) > 0)
End Function
Private Function RUN_IsBlockingBusinessPopupText(ByVal popupText As String) As Boolean
Dim t As String
t = UCase$(Trim$(popupText))
RUN_IsBlockingBusinessPopupText = _
(InStr(1, t, "NOT POSSIBLE", vbTextCompare) > 0) Or _
(InStr(1, t, "ERROR", vbTextCompare) > 0) Or _
(InStr(1, t, "INVALID", vbTextCompare) > 0) Or _
(InStr(1, t, "MISSING", vbTextCompare) > 0) Or _
(InStr(1, t, "DEFICIT", vbTextCompare) > 0) Or _
(InStr(1, t, "SHORTAGE", vbTextCompare) > 0) Or _
(InStr(1, t, "BLOCKED", vbTextCompare) > 0) Or _
(InStr(1, t, "FAILED", vbTextCompare) > 0) Or _
(InStr(1, t, "EXCEEDS", vbTextCompare) > 0) Or _
(InStr(1, t, "NO AUTHORITY", vbTextCompare) > 0) Or _
(InStr(1, t, "LOI", vbTextCompare) > 0) Or _
(InStr(1, t, "KHONG THE", vbTextCompare) > 0) Or _
(InStr(1, t, "THIEU", vbTextCompare) > 0) Or _
(InStr(1, t, "KHONG DU", vbTextCompare) > 0) Or _
(InStr(1, t, "BI CHAN", vbTextCompare) > 0)
End Function
Private Function RUN_IsKnownSafePopupText(ByVal popupText As String) As Boolean
Dim t As String
t = UCase$(Trim$(popupText))
If t = "" Then
RUN_IsKnownSafePopupText = False
Exit Function
End If
RUN_IsKnownSafePopupText = _
(InStr(1, t, "WARNING", vbTextCompare) > 0) Or _
(InStr(1, t, "CONTINUE", vbTextCompare) > 0) Or _
(InStr(1, t, "QUESTION", vbTextCompare) > 0) Or _
(InStr(1, t, "CONFIRM", vbTextCompare) > 0) Or _
(InStr(1, t, "INFORMATION", vbTextCompare) > 0) Or _
(InStr(1, t, "MESSAGE", vbTextCompare) > 0) Or _
(InStr(1, t, "DISPLAY LOG", vbTextCompare) > 0) Or _
(InStr(1, t, "DOCUMENT IS COMPLETE", vbTextCompare) > 0) Or _
(InStr(1, t, "TRANSFER POSTING", vbTextCompare) > 0) Or _
(InStr(1, t, "DELIVERY", vbTextCompare) > 0) Or _
(InStr(1, t, "POST", vbTextCompare) > 0) Or _
(InStr(1, t, "CHECK", vbTextCompare) > 0) Or _
(InStr(1, t, "CANH BAO", vbTextCompare) > 0) Or _
(InStr(1, t, "TIEP TUC", vbTextCompare) > 0) Or _
(InStr(1, t, "XAC NHAN", vbTextCompare) > 0) Or _
(InStr(1, t, "THONG TIN", vbTextCompare) > 0) Or _
(InStr(1, t, "CHUNG TU", vbTextCompare) > 0)
End Function
Private Function RUN_HasAffirmativePopupButton(ByVal session As Object) As Boolean
On Error Resume Next
Dim ctl As Object
Set ctl = session.FindById("wnd[1]/usr/btnSPOP-OPTION1")
If Not ctl Is Nothing Then
RUN_HasAffirmativePopupButton = True
Exit Function
End If
Set ctl = Nothing
Set ctl = session.FindById("wnd[1]/tbar[0]/btn[0]")
If Not ctl Is Nothing Then RUN_HasAffirmativePopupButton = True
On Error GoTo 0
End Function
Private Sub RUN_DismissPopupIfAny(ByVal session As Object)
On Error Resume Next
If Not RUN_HasPopup(session) Then Exit Sub
Err.Clear
session.FindById("wnd[1]/tbar[0]/btn[12]").Press
If Err.Number <> 0 Then
Err.Clear
session.FindById("wnd[1]/usr/btnSPOP-OPTION2").Press
End If
If Err.Number <> 0 Then
Err.Clear
session.FindById("wnd[1]").sendVKey 12
End If
If RUN_HasPopup(session) Then
Err.Clear
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
End If
RUN_Wait session, 300
On Error GoTo 0
End Sub
Private Function RUN_IsFatalSbarType(ByVal messageType As String) As Boolean
Select Case UCase$(Trim$(messageType))
Case "E", "A", "X"
RUN_IsFatalSbarType = True
End Select
End Function
Private Function RUN_GetSbarType(ByVal session As Object) As String
On Error Resume Next
RUN_GetSbarType = CStr(session.FindById("wnd[0]/sbar").messageType)
End Function
Private Function RUN_GetSbarText(ByVal session As Object) As String
On Error Resume Next
RUN_GetSbarText = CStr(session.FindById("wnd[0]/sbar").text)
End Function
Private Function RUN_IsSessionAlive(ByVal session As Object) As Boolean
On Error Resume Next
If session Is Nothing Then Exit Function
Dim wnd0 As Object
Err.Clear
Set wnd0 = session.FindById("wnd[0]", False)
RUN_IsSessionAlive = (Err.Number = 0 And Not wnd0 Is Nothing)
On Error GoTo 0
End Function
Private Sub RUN_WaitForStableIdle(ByVal session As Object, _
ByVal timeoutSeconds As Double)
On Error GoTo WaitError
Dim waitStart As Double
Dim stableSamples As Long
waitStart = Timer
stableSamples = 0
Do
If Not RUN_IsSessionAlive(session) Then
Err.Raise vbObjectError + 621, , _
"SAP session bi mat khi cho Grid on dinh."
End If
If session.busy Then
stableSamples = 0
Else
stableSamples = stableSamples + 1
If stableSamples >= mCfgInputIdleStableSamples Then Exit Sub
End If
DoEvents
Sleep mCfgInputIdlePollMs
If RUN_ElapsedSeconds(waitStart) >= timeoutSeconds Then
Err.Raise vbObjectError + 622, , _
"Grid MIGO khong ve trang thai idle on dinh sau " & _
CStr(timeoutSeconds) & " giay."
End If
Loop
WaitError:
Err.Raise Err.Number, "RUN_WaitForStableIdle", Err.Description
End Sub
Private Sub RUN_WaitForBatchEntryReady(ByVal session As Object, ByVal timeoutSeconds As Double, ByRef readyTable As Object)
On Error GoTo WaitError
Dim waitStart As Double, tbl As Object, batchCol As Long
Set readyTable = Nothing
waitStart = Timer
Do
If Not RUN_IsSessionAlive(session) Then Err.Raise vbObjectError + 623, , "SAP session bi mat khi cho cot Batch."
If Not session.busy Then
    Set tbl = RUN_FindMigoItemTable(session)
    If Not tbl Is Nothing Then
        batchCol = RUN_GetBatchColumnForRow(tbl, 0)
        If batchCol >= 0 Then
            Set readyTable = tbl
            Exit Sub
        End If
    End If
End If
DoEvents
Sleep mCfgInputIdlePollMs
If RUN_ElapsedSeconds(waitStart) >= timeoutSeconds Then
    Err.Raise vbObjectError + 624, , "Cot Batch chua san sang sau " & CStr(timeoutSeconds) & " giay."
End If
Loop
WaitError:
Err.Raise Err.Number, "RUN_WaitForBatchEntryReady", Err.Description
End Sub
Private Function RUN_ExtractDeliveryNumber(ByVal statusText As String) As String
On Error GoTo SafeExit
Dim re As Object
Dim matches As Object
Set re = CreateObject("VBScript.RegExp")
re.Global = False
re.IgnoreCase = True
re.pattern = "(OUTBOUND[ ]+DELIVERY|DELIVERY|DELIV[.]?)[^0-9]{0,40}([0-9]{8,12})"
Set matches = re.Execute(statusText)
If matches.count > 0 Then
RUN_ExtractDeliveryNumber = CStr(matches(0).SubMatches(1))
Else
RUN_ExtractDeliveryNumber = ""
End If
Exit Function
SafeExit:
RUN_ExtractDeliveryNumber = ""
End Function
Private Function RUN_AppendMessage(ByVal baseText As String, _
ByVal extraText As String) As String
baseText = Trim$(baseText)
extraText = Trim$(extraText)
If baseText = "" Then
RUN_AppendMessage = extraText
ElseIf extraText = "" Then
RUN_AppendMessage = baseText
Else
RUN_AppendMessage = baseText & " | " & extraText
End If
End Function
Private Function RUN_ExtractLongestNumber(ByVal s As String) As String
Dim i As Long, ch As String, cur As String, best As String
cur = ""
best = ""
For i = 1 To Len(s)
ch = Mid$(s, i, 1)
If ch >= "0" And ch <= "9" Then
cur = cur & ch
Else
If Len(cur) > Len(best) Then best = cur
cur = ""
End If
Next i
If Len(cur) > Len(best) Then best = cur
If Len(best) >= 8 Then RUN_ExtractLongestNumber = best
End Function
Private Function RUN_ShouldSkipOB(ByVal data As Variant, ByVal startRow As Long, ByVal endRow As Long) As Boolean
Dim r As Long
For r = startRow To endRow
If RUN_Clean(data(r, 11)) <> "" Then
RUN_ShouldSkipOB = True
Exit Function
End If
If UCase$(RUN_Clean(data(r, 9))) <> "READY" Then
RUN_ShouldSkipOB = True
Exit Function
End If
Next r
End Function
Private Function RUN_CheckOBTemplate(ByVal ws As Worksheet) As Boolean
Dim expected As Variant
expected = Array("OB", "Tram", "Product", "Qty", "Plant", "SLoc", "SLoc Transfer", "Batch", _
"Status", "SAP Message", "Delivery", "Post Time", "Run ID", "Product Count", _
"Row Count", "Group Key", "Warning")
Dim i As Long, headerText As String
For i = LBound(expected) To UBound(expected)
headerText = UCase$(RUN_Clean(ws.Cells(1, i + 1).Value))
If i = 0 Then
If headerText <> "OB" And headerText <> "DON" Then
LastMigoRunMessage = "Sai layout OB_GROUPS cot A. Can la OB hoac Don."
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbCritical
RUN_CheckOBTemplate = False
Exit Function
End If
ElseIf headerText <> UCase$(CStr(expected(i))) Then
LastMigoRunMessage = "Sai layout OB_GROUPS cot " & RUN_ColumnLetter(i + 1) & ". Can la: " & expected(i)
If Not mOneClickMigo Then MsgBox LastMigoRunMessage, vbCritical
RUN_CheckOBTemplate = False
Exit Function
End If
Next i
RUN_CheckOBTemplate = True
End Function
Private Sub RUN_SortOBGroups(ByVal ws As Worksheet, ByVal lastRow As Long)
On Error Resume Next
Dim seen As Object, r As Long, ob As String, lastOb As String, needSort As Boolean
Dim obArr As Variant
Set seen = CreateObject("Scripting.Dictionary")
obArr = ws.Range("A2:A" & CStr(lastRow)).Value2
For r = 1 To UBound(obArr, 1)
ob = RUN_Clean(obArr(r, 1))
If ob <> "" And ob <> lastOb Then
If seen.Exists(ob) Then needSort = True: Exit For
seen(ob) = True
lastOb = ob
End If
Next r
If Not needSort Then Exit Sub
With ws.Sort
.SortFields.Clear
.SortFields.Add key:=ws.Range("A2:A" & lastRow), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortNormal
.SetRange ws.Range("A1:Q" & lastRow)
.Header = xlYes
.Apply
End With
End Sub
Private Sub RUN_ResetMIGO311Screen(ByVal session As Object)
On Error Resume Next
Dim resetError As String
Call RUN_EnsureMIGO311Ready(session, resetError, True)
RUN_ClearMigoCache
On Error GoTo 0
End Sub
Private Function RUN_GetGroupEndInArray(ByVal data As Variant, ByVal startIdx As Long, ByVal n As Long, ByVal obName As String) As Long
Dim endIdx As Long
endIdx = startIdx
Do While endIdx < n
If RUN_Clean(data(endIdx + 1, 1)) = obName Then
endIdx = endIdx + 1
Else
Exit Do
End If
Loop
RUN_GetGroupEndInArray = endIdx
End Function
Private Function RUN_LastRow(ByVal ws As Worksheet, ByVal colNo As Long) As Long
RUN_LastRow = ws.Cells(ws.rows.count, colNo).End(xlUp).row
End Function
Private Function RUN_Clean(ByVal v As Variant) As String
If IsError(v) Or IsNull(v) Then
RUN_Clean = ""
Else
RUN_Clean = Trim$(CStr(v))
End If
End Function
Private Function RUN_ColumnLetter(ByVal colNo As Long) As String
RUN_ColumnLetter = Split(Cells(1, colNo).Address(False, False), "1")(0)
End Function
Private Sub RUN_Wait(ByVal session As Object, ByVal ms As Long)
Dim waitStart As Double
Dim idleSleepMs As Long
waitStart = Timer
If mCfgMigoUiPollMs > 0 Then
idleSleepMs = mCfgMigoUiPollMs
Else
idleSleepMs = 50
End If
On Error GoTo WaitError
Do While session.busy
DoEvents
Sleep idleSleepMs
If RUN_ElapsedSeconds(waitStart) >= mCfgBusyTimeoutSeconds Then
Err.Raise vbObjectError + 605, , _
"SAP Busy qua " & CStr(mCfgBusyTimeoutSeconds) & " giay. Tool da dung de tranh treo Excel."
End If
Loop
If ms > 0 Then Sleep ms
waitStart = Timer
Do While session.busy
DoEvents
Sleep idleSleepMs
If RUN_ElapsedSeconds(waitStart) >= mCfgBusyTimeoutSeconds Then
Err.Raise vbObjectError + 606, , _
"SAP van Busy sau khi cho. Tool da dung de tranh treo Excel."
End If
Loop
Exit Sub
WaitError:
Err.Raise Err.Number, "RUN_Wait", Err.Description
End Sub
Private Function RUN_ElapsedSeconds(ByVal startTime As Double) As Double
Dim currentTime As Double
currentTime = Timer
If currentTime < startTime Then currentTime = currentTime + 86400#
RUN_ElapsedSeconds = currentTime - startTime
End Function

