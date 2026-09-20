$targetDir = "g:\SAP\v184_patch"
if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir | Out-Null
}

$targetFile = Join-Path $targetDir "modSupports.bas"

$content = @'
Attribute VB_Name = "modSupports"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub SUPPORT_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub SUPPORT_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

Public Apps
Public session
Public WScript
Public sapGuiAuto
Public connection

Private Const SUPPORT_MAX_RETRY As Long = 2
Private Const SUPPORT_WAIT_SECONDS As Long = 45
Private mDeferredSaveTime As Date
Private mDeferredSaveScheduled As Boolean
Private mDeferredSaveAttempt As Long
Private mDeferredSaveLastError As String

Public Function Get_SAPSession() As Object
Set Get_SAPSession = SUPPORT_GetSAPSessionInternal(0)
End Function

Private Function SUPPORT_GetSAPSessionInternal(ByVal retryCount As Long) As Object
On Error GoTo SafeExit

Dim sapApp As Object, SAPcon As Object, ses As Object
Dim i As Long, j As Long

Set sapGuiAuto = GetObject("SAPGUI")
If sapGuiAuto Is Nothing Then Exit Function
Set sapApp = sapGuiAuto.GetScriptingEngine
If sapApp Is Nothing Then Exit Function
Set Apps = sapApp

' Uu tien session dang active (ho tro nhieu user SAP cung luc)
On Error Resume Next
Set ses = sapApp.ActiveSession
On Error GoTo SafeExit
If Not ses Is Nothing Then
    Set session = ses
    Set connection = ses.parent
    Set SUPPORT_GetSAPSessionInternal = ses
    Exit Function
End If

' Fallback: duyet tat ca connection va session
For i = 0 To sapApp.Children.Count - 1
    Set SAPcon = Nothing
    On Error Resume Next
    Set SAPcon = sapApp.Children(i)
    On Error GoTo SafeExit
    If Not SAPcon Is Nothing Then
        For j = 0 To SAPcon.Children.Count - 1
            Set ses = Nothing
            On Error Resume Next
            Set ses = SAPcon.Children(j)
            On Error GoTo SafeExit
            If Not ses Is Nothing Then
                Set session = ses
                Set connection = SAPcon
                Set SUPPORT_GetSAPSessionInternal = ses
                Exit Function
            End If
        Next j
    End If
Next i
Exit Function

SafeExit:
Set SUPPORT_GetSAPSessionInternal = Nothing
End Function
Public Sub Logon_SAP()
Dim sapLogonPath As String

On Error GoTo LogonError

sapLogonPath = SUPPORT_FindSAPLogonPath()
If sapLogonPath = "" Then
MsgBox "Khong tim thay SAP Logon. Hay mo SAP thu cong.", vbExclamation, "SAP"
Exit Sub
End If

Shell Chr$(34) & sapLogonPath & Chr$(34), vbNormalFocus
MsgBox "SAP Logon da duoc mo." & vbCrLf & _
"Hay dang nhap thu cong, mo dung he thong SAP va chay lai tool." & vbCrLf & _
"Tool khong luu ten dang nhap hoac mat khau.", _
vbInformation, "SAP Logon"
Exit Sub

LogonError:
MsgBox "Khong the mo SAP Logon: " & Err.Description, vbExclamation, "SAP"
End Sub
Private Function SUPPORT_FindSAPLogonPath() As String
Dim candidate As String

candidate = Environ$("ProgramFiles(x86)") & _
"\SAP\FrontEnd\SAPgui\saplogon.exe"
If Dir$(candidate) <> "" Then
SUPPORT_FindSAPLogonPath = candidate
Exit Function
End If

candidate = Environ$("ProgramFiles") & _
"\SAP\FrontEnd\SAPgui\saplogon.exe"
If Dir$(candidate) <> "" Then
SUPPORT_FindSAPLogonPath = candidate
Exit Function
End If

candidate = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
If Dir$(candidate) <> "" Then SUPPORT_FindSAPLogonPath = candidate
End Function


Public Sub CancelDeferredSave()
Dim procName As String
On Error Resume Next
If mDeferredSaveScheduled Then
procName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!modSupports.DeferredSaveCurrentWorkbook"
Application.OnTime mDeferredSaveTime, procName, , False
End If
mDeferredSaveScheduled = False
mDeferredSaveAttempt = 0
mDeferredSaveLastError = ""
Application.Cursor = xlDefault
Application.EnableCancelKey = xlInterrupt
On Error GoTo 0
End Sub

Public Function ScheduleDeferredSave(ByRef msg As String) As Boolean
msg = "": ScheduleDeferredSave = False
CancelDeferredSave
If ThisWorkbook.path = "" Then msg = "File chua co duong dan.": Exit Function
If LCase$(Right$(ThisWorkbook.Name, 5)) <> ".xlsm" Then msg = "Can dung file XLSM.": Exit Function
If ThisWorkbook.ReadOnly Then msg = "File dang Read-only.": Exit Function
mDeferredSaveAttempt = 0
mDeferredSaveLastError = ""
ScheduleDeferredSaveNext 2
ScheduleDeferredSave = True
End Function

Private Sub ScheduleDeferredSaveNext(ByVal delaySeconds As Long)
Dim procName As String
procName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!modSupports.DeferredSaveCurrentWorkbook"
mDeferredSaveTime = Now + TimeSerial(0, 0, delaySeconds)
mDeferredSaveScheduled = True
Application.OnTime mDeferredSaveTime, procName, , True
End Sub

Public Sub DeferredSaveCurrentWorkbook()
Dim oldEvents As Boolean, oldAlerts As Boolean
Dim n As Long, d As String
mDeferredSaveScheduled = False
mDeferredSaveAttempt = mDeferredSaveAttempt + 1
If MIGO311_RunPost_V24_ErrorGuard.ToolRunIsActive Or Not Application.Ready Or Application.CalculationState <> xlDone Then
If mDeferredSaveAttempt < 6 Then ScheduleDeferredSaveNext 2: Exit Sub
mDeferredSaveLastError = "Bo qua tu dong luu vi Excel/tool chua san sang."
mDeferredSaveScheduled = False
mDeferredSaveAttempt = 0
Application.Cursor = xlDefault
Application.EnableCancelKey = xlInterrupt
Exit Sub
End If
oldEvents = Application.EnableEvents
oldAlerts = Application.DisplayAlerts
On Error GoTo SaveFailed
Application.EnableEvents = False
Application.DisplayAlerts = True
ThisWorkbook.CheckCompatibility = False
ThisWorkbook.Save
Application.EnableEvents = oldEvents
Application.DisplayAlerts = oldAlerts
If ThisWorkbook.Saved Then
Application.StatusBar = "Da tu dong luu ket qua: " & ThisWorkbook.Name
Exit Sub
End If
Err.Raise vbObjectError + 2681, , "Workbook chua ve trang thai Saved."
SaveFailed:
mDeferredSaveLastError = "Loi " & CStr(Err.Number) & ": " & Err.Description
On Error Resume Next
Application.EnableEvents = oldEvents
Application.DisplayAlerts = oldAlerts
mDeferredSaveScheduled = False
mDeferredSaveAttempt = 0
Application.Cursor = xlDefault
Application.EnableCancelKey = xlInterrupt
Application.StatusBar = False
On Error GoTo 0
End Sub
'@

# Write file without BOM
[System.IO.File]::WriteAllText($targetFile, $content, (New-Object System.Text.UTF8Encoding $false))

# Verify BOM
$bytes = [System.IO.File]::ReadAllBytes($targetFile)
if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    Write-Host "Error: BOM found in $targetFile" -ForegroundColor Red
} else {
    Write-Host "Success: $targetFile created without BOM." -ForegroundColor Green
}
