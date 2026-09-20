Attribute VB_Name = "ModuleMain"
Option Explicit
Public gOneClickMode As Boolean
Private mOneClickRunning As Boolean
Public Sub CreateAllocationTable()
If Not MIGO311_RunPost_V24_ErrorGuard.ToolTryBegin("CHI TAO BANG FIFO") Then Exit Sub
Dim startTime As Double
startTime = Timer

Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean
Dim oldCalculation As XlCalculation
oldScreenUpdating = Application.ScreenUpdating
oldEnableEvents = Application.EnableEvents
oldCalculation = Application.Calculation

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

On Error GoTo Fail

ModuleLog.ClearLog
ModuleFIFO.ProcessAllocation

Application.ScreenUpdating = oldScreenUpdating
Application.Calculation = oldCalculation
Application.EnableEvents = oldEnableEvents

Dim elapsed As String
elapsed = Format$(Timer - startTime, "0.00")

Dim summary As String
summary = BuildRunSummary(elapsed)

If Not ModuleFIFO.LastAllocationSucceeded Then
MsgBox "KHONG THE HOAN THANH CAP HANG." & vbCrLf & vbCrLf & _
"Mo sheet LOG de xem loi cau hinh/du lieu can sua." & vbCrLf & vbCrLf & summary, _
vbCritical, "Cap hang that bai"
ElseIf ModuleFIFO.LastAllocationHasBlockingIssues Then
MsgBox "DA XU LY XONG NHUNG CON LOI NANG CHAN GOM OB / POST SAP." & vbCrLf & vbCrLf & _
"Kiem tra cac dong LOG loai LOI/ERROR va sua du lieu truoc khi Gom OB." & vbCrLf & vbCrLf & summary, _
vbExclamation, "Hoan thanh co loi"
ElseIf ModuleFIFO.LastShortageCount > 0 Then
MsgBox "HOAN THANH CO CANH BAO THIEU TON." & vbCrLf & vbCrLf & _
"Van co the Gom OB tu cac dong RESULT da cap hop le." & vbCrLf & _
"Phan con thieu chi duoc ghi trong LOG, khong dua vao OB/MIGO." & vbCrLf & vbCrLf & _
IIf(ModuleFIFO.LastShortageSummary <> "", ModuleFIFO.LastShortageSummary & vbCrLf & vbCrLf, "") & summary, _
vbExclamation, "Hoan thanh co thieu ton"
Else
MsgBox "HOAN THANH!" & vbCrLf & vbCrLf & _
"Kiem tra RESULT truoc khi GOM OB." & vbCrLf & vbCrLf & summary, _
vbInformation, "Hoan thanh"
End If
MIGO311_RunPost_V24_ErrorGuard.ToolEnd
Exit Sub

Fail:
Application.ScreenUpdating = oldScreenUpdating
Application.Calculation = oldCalculation
Application.EnableEvents = oldEnableEvents
MIGO311_RunPost_V24_ErrorGuard.ToolEnd
MsgBox "Co loi xay ra: " & Err.Description, vbCritical, "Loi"
End Sub

Private Function BuildRunSummary(ByVal elapsed As String) As String
Dim resRows As Long
Dim nLoi As Long, nCanhBao As Long, nThieu As Long

resRows = CountDataRows("RESULT", 1)
CountLogByType nLoi, nCanhBao, nThieu

BuildRunSummary = _
"- Dong RESULT hop le: " & resRows & vbCrLf & _
"- LOG: " & nLoi & " loi, " & nCanhBao & " canh bao, " & nThieu & " thieu hang" & vbCrLf & _
"- Thoi gian xu ly: " & elapsed & " giay"
End Function

Private Function CountDataRows(ByVal sheetName As String, ByVal keyCol As Long) As Long
Dim ws As Worksheet
On Error Resume Next
Set ws = ThisWorkbook.Worksheets(sheetName)
On Error GoTo 0
If ws Is Nothing Then Exit Function

Dim lr As Long
lr = ws.Cells(ws.rows.Count, keyCol).End(xlUp).Row
If lr >= 2 Then CountDataRows = lr - 1
End Function

Private Sub CountLogByType(ByRef nLoi As Long, ByRef nCanhBao As Long, ByRef nThieu As Long)
nLoi = 0: nCanhBao = 0: nThieu = 0

Dim ws As Worksheet
On Error Resume Next
Set ws = ThisWorkbook.Worksheets("LOG")
On Error GoTo 0
If ws Is Nothing Then Exit Sub

Dim lr As Long
lr = ws.Cells(ws.rows.Count, 2).End(xlUp).Row
If lr < 2 Then Exit Sub

Dim data As Variant
data = ws.Range(ws.Cells(2, 2), ws.Cells(lr, 2)).value

Dim i As Long, t As String
If IsArray(data) Then
For i = 1 To UBound(data, 1)
t = UCase$(Trim$(CStr(data(i, 1) & "")))
If t = "LOI" Then
nLoi = nLoi + 1
ElseIf t = "CANH BAO" Then
nCanhBao = nCanhBao + 1
ElseIf t = "THIEU HANG" Or t = "THIEU TON" Then
nThieu = nThieu + 1
End If
Next i
Else
t = UCase$(Trim$(CStr(ws.Cells(2, 2).value & "")))
If t = "LOI" Then
nLoi = 1
ElseIf t = "CANH BAO" Then
nCanhBao = 1
ElseIf t = "THIEU HANG" Or t = "THIEU TON" Then
nThieu = 1
End If
End If
End Sub

Public Sub Run_FIFO_MIGO311_OneClick()
Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean
Dim oldCalculation As XlCalculation, oldDisplayAlerts As Boolean
Dim configError As String, scopeText As String, confirmText As String
Dim allowShortage As Boolean, showSummary As Boolean, confirmBeforeSap As Boolean
Dim startedAt As Double, elapsedSeconds As Double
Dim summaryText As String
Dim migoFailed As Long
Dim summaryStyle As VbMsgBoxStyle
Dim errorText As String
Dim saveError As String
Dim saveScheduled As Boolean
Dim runStage As String
Dim resumeExistingOB As Boolean
Dim readyOBCount As Long, postedOBCount As Long, staleOBCount As Long
Dim noPendingOB As Boolean
Dim runMode As Long, existingResultRows As Long, existingOBRows As Long

If Not MIGO311_RunPost_V24_ErrorGuard.ToolTryBegin("CHAY TOAN BO") Then Exit Sub

ModuleHelper.ResetToolConfigCache
If Not ModuleHelper.ValidateToolConfig(configError) Then
MsgBox "CONFIG chua hop le:" & vbCrLf & vbCrLf & configError, vbCritical, "CONFIG"
MIGO311_RunPost_V24_ErrorGuard.ToolEnd
Exit Sub
End If

scopeText = UCase$(Trim$(ModuleHelper.ToolConfig("ONE_CLICK_POST_SCOPE", "ALL")))
allowShortage = ModuleHelper.ToolConfigBool("ALLOW_SHORTAGE_TO_MIGO", True)
showSummary = ModuleHelper.ToolConfigBool("ONE_CLICK_SHOW_FINAL_SUMMARY", True)
confirmBeforeSap = ModuleHelper.ToolConfigBool("ONE_CLICK_CONFIRM_BEFORE_SAP", True)
MIGO311_RunPost_V24_ErrorGuard.GetExistingOBState readyOBCount, postedOBCount, staleOBCount
existingResultRows = CountDataRows("RESULT", 1): existingOBRows = CountDataRows("OB_GROUPS", 1)
runMode = ModuleFIFO.OneClickSelectMode(scopeText, (existingResultRows > 0 Or existingOBRows > 0), readyOBCount, postedOBCount, staleOBCount, confirmBeforeSap)
If runMode = 0 Then MIGO311_RunPost_V24_ErrorGuard.ToolEnd: Exit Sub
resumeExistingOB = (runMode = 2)
noPendingOB = (resumeExistingOB And readyOBCount = 0)

oldScreenUpdating = Application.ScreenUpdating
oldEnableEvents = Application.EnableEvents
oldCalculation = Application.Calculation
oldDisplayAlerts = Application.DisplayAlerts

On Error GoTo Failed
runStage = "Khoi tao luot chay"
startedAt = Timer
mOneClickRunning = True
gOneClickMode = True

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
If Not resumeExistingOB Then
    If ModuleHelper.ToolConfigBool("AUTO_DOWNLOAD_POSTING_NEW_RUN", True) Then
        ModuleFIFO.OneClickClearPostingData
        Application.StatusBar = "Buoc 1/4: Dang tai ton /SCWM/POST theo ORDER..."
        runStage = "Tai ton POSTING"
        If Not ModuleSAP.DownloadPostingForOrder(False) Then _
            Err.Raise vbObjectError + 2652, , "Khong tai duoc POSTING moi. RESULT va OB_GROUPS cu duoc giu nguyen."
    End If
    ModuleFIFO.OneClickClearRunData False
End If
If resumeExistingOB Then
Application.StatusBar = "Tiep tuc OB Ready; bo qua tai POSTING/FIFO/Gom OB..."
runStage = "Tiep tuc OB thu cong"
Else
Application.StatusBar = "Buoc 2/4: Dang kiem tra va phan bo FIFO/FEFO..."
runStage = "Xoa LOG"
ModuleLog.ClearLog
runStage = "Phan bo FIFO/FEFO"
ModuleFIFO.ProcessAllocation
If Not ModuleFIFO.LastAllocationSucceeded Then _
Err.Raise vbObjectError + 2653, , "FIFO/FEFO khong hoan thanh. Mo LOG de xem loi."
If ModuleFIFO.LastAllocationHasBlockingIssues Then _
Err.Raise vbObjectError + 2654, , "LOG con loi nghiem trong, dung truoc khi gom OB/MIGO."
If ModuleFIFO.LastShortageCount > 0 And Not allowShortage Then _
Err.Raise vbObjectError + 2655, , "Co thieu ton va ALLOW_SHORTAGE_TO_MIGO=FALSE."

Application.StatusBar = "Buoc 3/4: Dang gom OB_GROUPS..."
runStage = "Gom OB_GROUPS"
ModuleOBGroup_V8_ErrorGuard.GroupResultIntoOB_ProductBased
If Not ModuleOBGroup_V8_ErrorGuard.LastOBGroupSucceeded Then _
Err.Raise vbObjectError + 2656, , ModuleOBGroup_V8_ErrorGuard.LastOBGroupMessage
Application.StatusBar = "Da gom OB; chuyen sang SAP..."
End If

If scopeText <> "BUILD_ONLY" Then
If noPendingOB Then
Application.StatusBar = "Khong con OB Ready."
Else
Application.StatusBar = "Buoc 4/4: Ket noi SAP/MIGO..."
runStage = "Ket noi SAP / MIGO 311"
DoEvents
MIGO311_RunPost_V24_ErrorGuard.Run_MIGO311_OneClick
If MIGO311_RunPost_V24_ErrorGuard.LastMigoProcessedCount = 0 And Not noPendingOB Then _
Err.Raise vbObjectError + 2658, , MIGO311_RunPost_V24_ErrorGuard.LastMigoRunMessage
migoFailed = MIGO311_RunPost_V24_ErrorGuard.LastMigoProcessedCount - _
MIGO311_RunPost_V24_ErrorGuard.LastMigoPostedCount
If migoFailed < 0 Then migoFailed = 0
End If
End If

Application.StatusBar = "Dang hoan tat..."
Application.ScreenUpdating = oldScreenUpdating
Application.EnableEvents = oldEnableEvents
Application.Calculation = oldCalculation
Application.DisplayAlerts = oldDisplayAlerts
On Error Resume Next
ThisWorkbook.Worksheets("ORDER").Range("E15,G15,I15,E18,G18,I18").Calculate
On Error GoTo Failed

If ModuleHelper.ToolConfigBool("SAVE_WORKBOOK_AT_END", True) Then
runStage = "Len lich luu"
saveError = ""
saveScheduled = modSupports.ScheduleDeferredSave(saveError)
End If

If Timer >= startedAt Then
elapsedSeconds = Timer - startedAt
Else
elapsedSeconds = 86400# - startedAt + Timer
End If

Application.StatusBar = False
gOneClickMode = False
mOneClickRunning = False
MIGO311_RunPost_V24_ErrorGuard.ToolEnd

If showSummary Then
summaryText = "FIFO + MIGO ket thuc." & vbCrLf & _
IIf(Not resumeExistingOB And ModuleHelper.ToolConfigBool("AUTO_DOWNLOAD_POSTING_NEW_RUN", True), "- POSTING da tai moi tu /SCWM/POST." & vbCrLf, "") & _
IIf(resumeExistingOB, "- Tiep tuc OB thu cong." & vbCrLf, "") & _
IIf(noPendingOB, "- Khong con OB Ready." & vbCrLf, "") & _
"- RESULT hop le: " & CountDataRows("RESULT", 1) & vbCrLf & _
"- OB: " & IIf(resumeExistingOB, readyOBCount + postedOBCount, ModuleOBGroup_V8_ErrorGuard.LastOBGroupCount) & vbCrLf & _
"- Thieu ton: " & IIf(resumeExistingOB, "khong tinh lai", CStr(ModuleFIFO.LastShortageCount))
If scopeText <> "BUILD_ONLY" Then
summaryText = summaryText & vbCrLf & _
"- MIGO da xu ly: " & MIGO311_RunPost_V24_ErrorGuard.LastMigoProcessedCount & vbCrLf & _
"- Posted: " & MIGO311_RunPost_V24_ErrorGuard.LastMigoPostedCount & vbCrLf & _
"- Loi/khong Posted: " & migoFailed
If migoFailed > 0 Then
summaryStyle = vbExclamation
Else
summaryStyle = vbInformation
End If
Else
summaryText = summaryText & vbCrLf & "- Che do BUILD_ONLY: khong chay SAP"
summaryStyle = vbInformation
End If
summaryText = summaryText & vbCrLf & _
"- Tong thoi gian: " & Format$(elapsedSeconds, "0.00") & " giay"
If saveScheduled Then
summaryText = summaryText & vbCrLf & _
"- Tu dong luu: sau khi dong thong bao nay."
ElseIf saveError <> "" Then
summaryText = summaryText & vbCrLf & "- Khong the len lich luu: " & saveError
summaryStyle = vbExclamation
End If
MsgBox summaryText, summaryStyle, "FIFO + MIGO 311"
End If
Exit Sub

Failed:
errorText = "Buoc: " & runStage & vbCrLf & _
"Loi " & CStr(Err.Number) & ": " & Err.Description
Application.StatusBar = False
Application.ScreenUpdating = oldScreenUpdating
Application.EnableEvents = oldEnableEvents
Application.Calculation = oldCalculation
Application.DisplayAlerts = oldDisplayAlerts
gOneClickMode = False
mOneClickRunning = False
MIGO311_RunPost_V24_ErrorGuard.ToolEnd
MsgBox "Tool da dung:" & vbCrLf & vbCrLf & errorText, _
vbCritical, "FIFO + MIGO 311"
End Sub

