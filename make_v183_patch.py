import os

base_file = r"g:\SAP\v182_extracted\vba_exported\MIGO311_RunPost_V24_ErrorGuard.bas"
dest_file = r"g:\SAP\v183_patch\MIGO311_RunPost_V24_ErrorGuard.bas"

with open(base_file, "r", encoding="utf-8", errors="replace") as f:
    text = f.read()

# Replace version string in header
text = text.replace("V1.7 SMART GOS ATT", "V1.8.3 SMART GOS ATT REBUILT")

gos_adapter_code = """
' =============================================================================
'  GOS ADAPTER LAYER - V1.8.3 SMART GOS ATT REBUILT
'  Bao gom day du 7 ham GOS Adapter, Re-bind Session, Auto-Discovery & Dynamic Map
' =============================================================================

Public Function ATT_WaitDocReady(ByVal ses As Object, ByVal timeoutSec As Long) As Boolean
    Dim startTime As Double
    startTime = Timer
    On Error Resume Next
    Do While (Timer - startTime) < timeoutSec
        If Not ses Is Nothing Then
            If Not ses.Busy Then
                Dim headCtrl As Object
                Set headCtrl = Nothing
                Set headCtrl = ses.findById("wnd[0]/usr/subTAXI_HEADER:SAPLMIGO:0115/subHEADER_SYSTEM:SAPLMIGO:0100/txtGOHEAD-HEADER_TXT", False)
                If Not headCtrl Is Nothing Then
                    ATT_WaitDocReady = True
                    Exit Function
                End If
            End If
        End If
        DoEvents
        ATT_Sleep 200
    Loop
    ATT_WaitDocReady = False
End Function

Public Function ATT_FindGosShell(ByVal ses As Object, ByRef outGosId As String) As Boolean
    On Error Resume Next
    Dim primaryId As String, savedMap As String
    primaryId = "wnd[0]/titl/shellcont/shell"
    
    ' 0. Kiem tra map da luu trong CONFIG
    Dim wsCfg As Worksheet
    Set wsCfg = Nothing
    Set wsCfg = ThisWorkbook.Sheets("CONFIG")
    If Not wsCfg Is Nothing Then
        savedMap = Trim$(wsCfg.Range("B16").Value)
        If savedMap <> "" Then
            Dim testObj As Object
            Set testObj = ses.findById(savedMap, False)
            If Not testObj Is Nothing Then
                If testObj.Enabled Then
                    outGosId = savedMap
                    ATT_FindGosShell = True
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' 1. Primary Technical ID
    Dim gosObj As Object
    Set gosObj = ses.findById(primaryId, False)
    If Not gosObj Is Nothing Then
        If gosObj.Enabled Then
            outGosId = primaryId
            ATT_FindGosShell = True
            Exit Function
        End If
    End If
    
    ' 2. Auto-Discovery qua Control Tree
    Dim titlContainer As Object
    Set titlContainer = ses.findById("wnd[0]/titl/shellcont", False)
    If Not titlContainer Is Nothing Then
        Dim i As Long
        For i = 0 To titlContainer.Children.Count - 1
            Dim child As Object
            Set child = titlContainer.Children(CInt(i))
            If child.Type = "GuiShell" Then
                outGosId = child.Id
                ATT_FindGosShell = True
                Exit Function
            End If
        Next i
    End If
    
    outGosId = ""
    ATT_FindGosShell = False
End Function

Public Function ATT_DetectSecurity() As Boolean
    #If Win64 Then
        Dim hwndSec As LongPtr
    #Else
        Dim hwndSec As Long
    #End If
    hwndSec = modWinFileDialog.WinFindWindow("#32770", "SAP GUI Security")
    If hwndSec = 0 Then hwndSec = modWinFileDialog.WinFindWindow("#32770", "SAP GUI Security Prompt")
    ATT_DetectSecurity = (hwndSec <> 0)
End Function

Public Function ATT_WaitSecurityClosed(ByVal timeoutSec As Long, ByRef logStage As String) As Boolean
    Dim startTime As Double
    startTime = Timer
    logStage = "SAP_SECURITY_DETECTED"
    
    #If Win64 Then
        Dim hwndSec As LongPtr
    #Else
        Dim hwndSec As Long
    #End If
    
    hwndSec = modWinFileDialog.WinFindWindow("#32770", "SAP GUI Security")
    If hwndSec = 0 Then hwndSec = modWinFileDialog.WinFindWindow("#32770", "SAP GUI Security Prompt")
    If hwndSec <> 0 Then modWinFileDialog.WinBringToFront hwndSec
    
    logStage = "SAP_SECURITY_WAIT_USER"
    
    Do While (Timer - startTime) < timeoutSec
        If Not ATT_DetectSecurity() Then
            logStage = "SAP_SECURITY_ALLOWED"
            ATT_WaitSecurityClosed = True
            Exit Function
        End If
        DoEvents
        ATT_Sleep 300
    Loop
    
    logStage = "SAP_SECURITY_TIMEOUT"
    ATT_WaitSecurityClosed = False
End Function

Public Function ATT_RebindSessionAfterSecurity(ByRef ses As Object) As Boolean
    On Error Resume Next
    Dim guiApp As Object, connection As Object
    Set guiApp = GetObject("SAPGUI").GetScriptingEngine
    Set connection = guiApp.Children(0)
    Set ses = connection.Children(0)
    ATT_RebindSessionAfterSecurity = (Not ses Is Nothing)
End Function

Public Sub ATT_SaveSuccessfulGosMap(ByVal ses As Object, ByVal workingGosId As String)
    On Error Resume Next
    Dim sysName As String, clientNo As String, userName As String
    sysName = ses.Info.SystemName
    clientNo = ses.Info.Client
    userName = ses.Info.User
    
    Dim mapKey As String
    mapKey = sysName & "_" & clientNo & "_" & userName & "_MIGO"
    
    Dim wsCfg As Worksheet
    Set wsCfg = ThisWorkbook.Sheets("CONFIG")
    If Not wsCfg Is Nothing Then
        wsCfg.Range("B15").Value = mapKey
        wsCfg.Range("B16").Value = workingGosId
    End If
End Sub

Public Function ATT_OpenGosWithRetry(ByVal ses As Object, ByRef logStage As String) As Boolean
    Dim retryCount As Long
    Dim gosId As String
    Dim gosShell As Object
    
    For retryCount = 1 To 3
        On Error Resume Next
        Select Case retryCount
            Case 1 ' Luot 1: Direct Call
                If ATT_FindGosShell(ses, gosId) Then
                    Set gosShell = ses.findById(gosId, False)
                    gosShell.pressContextButton "%GOS_TOOLBOX"
                    gosShell.selectContextMenuItem "%GOS_PCATTA_CREA"
                    If Err.Number = 0 Then
                        logStage = "ACTION_SENT"
                        ATT_OpenGosWithRetry = True
                        Exit Function
                    End If
                End If
                
            Case 2 ' Luot 2: Active Window + Wait + Re-bind + Retry
                AppActivate "SAP GUI"
                ATT_Sleep 500
                If ATT_RebindSessionAfterSecurity(ses) Then
                    If ATT_FindGosShell(ses, gosId) Then
                        Set gosShell = ses.findById(gosId, False)
                        gosShell.pressContextButton "%GOS_TOOLBOX"
                        gosShell.selectContextMenuItem "%GOS_PCATTA_CREA"
                        If Err.Number = 0 Then
                            logStage = "ACTION_SENT"
                            ATT_OpenGosWithRetry = True
                            Exit Function
                        End If
                    End If
                End If
                
            Case 3 ' Luot 3: Auto-Discovery & Save Map
                If ATT_FindGosShell(ses, gosId) Then
                    Set gosShell = ses.findById(gosId, False)
                    gosShell.pressContextButton "%GOS_TOOLBOX"
                    gosShell.selectContextMenuItem "%GOS_PCATTA_CREA"
                    If Err.Number = 0 Then
                        Call ATT_SaveSuccessfulGosMap(ses, gosId)
                        logStage = "ACTION_SENT"
                        ATT_OpenGosWithRetry = True
                        Exit Function
                    End If
                End If
        End Select
        On Error GoTo 0
        ATT_Sleep 300
    Next retryCount
    
    logStage = "GOS_ATTACH_ITEM_NOT_FOUND"
    ATT_OpenGosWithRetry = False
End Function
"""

text += "\n" + gos_adapter_code

os.makedirs(r"g:\SAP\v183_patch", exist_ok=True)
with open(dest_file, "w", encoding="utf-8") as f:
    f.write(text)

print("Created g:\\SAP\\v183_patch\\MIGO311_RunPost_V24_ErrorGuard.bas successfully")
