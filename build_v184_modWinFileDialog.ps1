$path = "g:\SAP\v184_patch\modWinFileDialog.bas"
$dir = Split-Path $path
if (!(Test-Path $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$content = @'
Attribute VB_Name = "modWinFileDialog"
Option Explicit

' ======================================================================
'  modWinFileDialog  - V1.8.4 MERGED (FIX29 robust + V1.8.3 Unicode)
'  Combines: EnumWindows security auto-click + GetFileAttributesW Unicode
' ======================================================================

#If VBA7 Then
    ' --- FIX29 APIs: Window enumeration & security auto-click ---
    Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetDlgItem Lib "user32" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long) As LongPtr
    Private Declare PtrSafe Function GetClassNameW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpClassName As LongPtr, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function GetWindowTextW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpString As LongPtr, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    ' --- V1.8.3 APIs: Unicode file check + FindWindow + BringWindowToTop ---
    Private Declare PtrSafe Function WFD_FindWindowW Lib "user32" Alias "FindWindowW" _
        (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
    Private Declare PtrSafe Function WFD_FindWindowExW Lib "user32" Alias "FindWindowExW" _
        (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
         ByVal lpszClass As LongPtr, ByVal lpszWindow As LongPtr) As LongPtr
    Private Declare PtrSafe Function WFD_BringWindowToTop Lib "user32" Alias "BringWindowToTop" _
        (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function WFD_GetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" _
        (ByVal lpFileName As LongPtr) As Long
#Else
    Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
    Private Declare Function GetClassNameW Lib "user32" (ByVal hwnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
    Private Declare Function GetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function WFD_FindWindowW Lib "user32" Alias "FindWindowW" _
        (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
    Private Declare Function WFD_FindWindowExW Lib "user32" Alias "FindWindowExW" _
        (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, _
         ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
    Private Declare Function WFD_BringWindowToTop Lib "user32" Alias "BringWindowToTop" _
        (ByVal hWnd As Long) As Long
    Private Declare Function WFD_GetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" _
        (ByVal lpFileName As Long) As Long
#End If

' --- Module-level state (FIX29) ---
#If VBA7 Then
    Private mFileDlg As LongPtr, mSecurityDlg As LongPtr, mChildEdit As LongPtr, mChildButton As LongPtr
#Else
    Private mFileDlg As Long, mSecurityDlg As Long, mChildEdit As Long, mChildButton As Long
#End If

Private Const BM_CLICK As Long = &HF5
Private Const WM_SETTEXT As Long = &HC
Private mFindChildMode As Long
Private mLastMethod As String

' ======================================================================
'  FIX29: WinLastMethod / WinElapsed
' ======================================================================
Public Function WinLastMethod() As String
    WinLastMethod = mLastMethod
End Function

Private Function WinElapsed(ByVal t0 As Double) As Double
    If Timer >= t0 Then WinElapsed = Timer - t0 Else WinElapsed = 86400# - t0 + Timer
End Function

' ======================================================================
'  FIX29: EnumWindows callbacks
' ======================================================================
#If VBA7 Then
Public Function WinEnumTopProc(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Public Function WinEnumTopProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
#End If
    WinEnumTopProc = 1
    If IsWindowVisible(hwnd) = 0 Then Exit Function
    If WinClass(hwnd) <> "#32770" Then Exit Function
    Dim ttl As String: ttl = UCase$(WinText(hwnd))
    If InStr(ttl, "SAP") > 0 And (InStr(ttl, "SECUR") > 0 Or InStr(ttl, "GUI") > 0) Then mSecurityDlg = hwnd
    If WinFileEdit(hwnd) <> 0 Then mFileDlg = hwnd
End Function

#If VBA7 Then
Public Function WinEnumChildProc(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Public Function WinEnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
#End If
    WinEnumChildProc = 1
    Dim cls As String: cls = UCase$(WinClass(hwnd))
    If mFindChildMode = 1 Then
        If cls = "EDIT" And IsWindowVisible(hwnd) <> 0 Then mChildEdit = hwnd
    ElseIf mFindChildMode = 2 Then
        If cls = "BUTTON" Then
            Dim txt As String: txt = UCase$(WinText(hwnd))
            If InStr(txt, "ALLOW") > 0 Or InStr(txt, "PERMIT") > 0 Or InStr(txt, "YES") > 0 _
               Or InStr(txt, "CONTINUE") > 0 Or InStr(txt, "CHO PHEP") > 0 Or InStr(txt, "OK") > 0 Then
                mChildButton = hwnd: WinEnumChildProc = 0
            End If
        End If
    End If
End Function

#If VBA7 Then
Private Function WinFileEdit(ByVal hDlg As LongPtr) As LongPtr
#Else
Private Function WinFileEdit(ByVal hDlg As Long) As Long
#End If
    WinFileEdit = GetDlgItem(hDlg, 1148)
    If WinFileEdit = 0 Then
        mChildEdit = 0: mFindChildMode = 1
        EnumChildWindows hDlg, AddressOf WinEnumChildProc, 0
        WinFileEdit = mChildEdit
    End If
End Function

Private Function WinClass(ByVal hwnd As LongPtr) As String
    Dim buf As String: buf = String$(260, vbNullChar)
    Dim n As Long: n = GetClassNameW(hwnd, StrPtr(buf), 260)
    If n > 0 Then WinClass = Left$(buf, n)
End Function

Private Function WinText(ByVal hwnd As LongPtr) As String
    Dim buf As String: buf = String$(512, vbNullChar)
    Dim n As Long: n = GetWindowTextW(hwnd, StrPtr(buf), 512)
    If n > 0 Then WinText = Left$(buf, n)
End Function

' ======================================================================
'  FIX29: WinHandleSapSecurity (auto-click Allow/Permit/Yes)
' ======================================================================
Public Function WinHandleSapSecurity(Optional ByVal timeoutSeconds As Long = 0) As Boolean
    Dim t0 As Double: t0 = Timer
    Do
        mSecurityDlg = 0: mChildButton = 0
        EnumWindows AddressOf WinEnumTopProc, 0
        If mSecurityDlg <> 0 Then
            mFindChildMode = 2: mChildButton = 0
            EnumChildWindows mSecurityDlg, AddressOf WinEnumChildProc, 0
            If mChildButton <> 0 Then
                SetForegroundWindow mSecurityDlg
                DoEvents
                SendMessageW mChildButton, BM_CLICK, 0, 0
                mLastMethod = "WINDOWS_SECURITY_BUTTON"
                WinHandleSapSecurity = True
                Exit Function
            End If
        End If
        If timeoutSeconds <= 0 Then Exit Do
        DoEvents
        If WinElapsed(t0) > CDbl(timeoutSeconds) Then Exit Do
    Loop
End Function

' ======================================================================
'  V1.8.3: Unicode File/Folder Existence (GetFileAttributesW)
' ======================================================================
Public Function ATT_FolderExistsW(ByVal pathStr As String) As Boolean
    Dim attr As Long
    pathStr = Trim$(pathStr)
    If pathStr = "" Then Exit Function
    attr = WFD_GetFileAttributesW(StrPtr(pathStr))
    If attr <> -1 Then ATT_FolderExistsW = ((attr And 16) <> 0)
End Function

Public Function ATT_FileExistsW(ByVal pathStr As String) As Boolean
    Dim attr As Long
    pathStr = Trim$(pathStr)
    If pathStr = "" Then Exit Function
    attr = WFD_GetFileAttributesW(StrPtr(pathStr))
    If attr <> -1 Then ATT_FileExistsW = ((attr And 16) = 0)
End Function

' ======================================================================
'  V1.8.3: WinFindWindow / WinBringToFront (Public wrappers)
' ======================================================================
#If VBA7 Then
Public Function WinFindWindow(ByVal className As String, ByVal windowName As String) As LongPtr
    Dim pClass As LongPtr, pName As LongPtr
    If className <> "" Then pClass = StrPtr(className)
    If windowName <> "" Then pName = StrPtr(windowName)
    WinFindWindow = WFD_FindWindowW(pClass, pName)
End Function

Public Sub WinBringToFront(ByVal hWnd As LongPtr)
    If hWnd <> 0 Then
        WFD_BringWindowToTop hWnd
        SetForegroundWindow hWnd
    End If
End Sub
#Else
Public Function WinFindWindow(ByVal className As String, ByVal windowName As String) As Long
    Dim pClass As Long, pName As Long
    If className <> "" Then pClass = StrPtr(className)
    If windowName <> "" Then pName = StrPtr(windowName)
    WinFindWindow = WFD_FindWindowW(pClass, pName)
End Function

Public Sub WinBringToFront(ByVal hWnd As Long)
    If hWnd <> 0 Then
        WFD_BringWindowToTop hWnd
        SetForegroundWindow hWnd
    End If
End Sub
#End If

' ======================================================================
'  V1.8.3: Security Dialog Detection (passive)
' ======================================================================
Public Function WinBringSecurityDialogToFront() As Boolean
#If VBA7 Then
    Dim hSec As LongPtr
#Else
    Dim hSec As Long
#End If
    hSec = WFD_FindSapSecurityWindow()
    If hSec <> 0 Then
        WFD_BringWindowToTop hSec
        SetForegroundWindow hSec
        WinBringSecurityDialogToFront = True
    End If
End Function

Public Function WinIsSapSecurityVisible() As Boolean
#If VBA7 Then
    Dim hSec As LongPtr
#Else
    Dim hSec As Long
#End If
    hSec = WFD_FindSapSecurityWindow()
    WinIsSapSecurityVisible = (hSec <> 0 And IsWindow(hSec) <> 0)
End Function

#If VBA7 Then
Private Function WFD_FindSapSecurityWindow() As LongPtr
    Dim hWnd As LongPtr
#Else
Private Function WFD_FindSapSecurityWindow() As Long
    Dim hWnd As Long
#End If
    Dim titles(3) As String
    titles(0) = "SAP GUI Security"
    titles(1) = "SAP Security"
    titles(2) = "Security Warning"
    titles(3) = "SAP GUI-Sicherheit"
    Dim i As Integer
    For i = 0 To 3
        hWnd = WFD_FindWindowW(0, StrPtr(titles(i)))
        If hWnd <> 0 Then WFD_FindSapSecurityWindow = hWnd: Exit Function
    Next i
End Function

' ======================================================================
'  MERGED: WinImportFile (FIX29 robustness + V1.8.3 Unicode)
' ======================================================================
Public Function WinImportFile(ByVal fullPath As String, _
                              Optional ByVal timeoutSeconds As Long = 20) As Boolean
#If VBA7 Then
    Dim hDlg As LongPtr, hFile As LongPtr, hOpen As LongPtr, hFore As LongPtr
#Else
    Dim hDlg As Long, hFile As Long, hOpen As Long, hFore As Long
#End If
    Dim t0 As Double
    mLastMethod = ""

    ' V1.8.3 Unicode file check
    If Not ATT_FileExistsW(fullPath) Then Exit Function

    t0 = Timer
    Do
        ' FIX29: auto-dismiss SAP Security popups
        WinHandleSapSecurity 0

        ' FIX29: check foreground window first
        hFore = GetForegroundWindow()
        If hFore <> 0 And IsWindowVisible(hFore) <> 0 Then
            If WinClass(hFore) = "#32770" Then
                If WinFileEdit(hFore) <> 0 Then hDlg = hFore
            End If
        End If

        ' FIX29: EnumWindows fallback
        If hDlg = 0 Then
            mFileDlg = 0
            EnumWindows AddressOf WinEnumTopProc, 0
            hDlg = mFileDlg
        End If

        If hDlg <> 0 Then Exit Do
        DoEvents
        If WinElapsed(t0) > CDbl(timeoutSeconds) Then Exit Function
    Loop

    ' Bring dialog to front
    SetForegroundWindow hDlg
    DoEvents

    ' Find edit control: try ID 1148, then V1.8.3 ComboBox hierarchy, then EnumChildWindows
    hFile = GetDlgItem(hDlg, 1148)
    If hFile = 0 Then hFile = WFD_FindEditControl(hDlg)
    If hFile = 0 Then
        mChildEdit = 0: mFindChildMode = 1
        EnumChildWindows hDlg, AddressOf WinEnumChildProc, 0
        hFile = mChildEdit
    End If

    hOpen = GetDlgItem(hDlg, 1)
    If hFile = 0 Or hOpen = 0 Then Exit Function

    ' Set file path and click Open
    SendMessageW hFile, WM_SETTEXT, 0, StrPtr(fullPath)
    DoEvents
    SendMessageW hOpen, BM_CLICK, 0, 0
    mLastMethod = "WINDOWS_UNICODE_DIALOG"

    ' Wait for dialog to close, handling security popups
    t0 = Timer
    Do While IsWindow(hDlg) <> 0
        WinHandleSapSecurity 0
        DoEvents
        If WinElapsed(t0) > CDbl(timeoutSeconds) Then Exit Function
    Loop
    WinImportFile = True
End Function

' ======================================================================
'  V1.8.3: WFD_FindEditControl (ComboBoxEx32 hierarchy)
' ======================================================================
#If VBA7 Then
Private Function WFD_FindEditControl(ByVal hDlg As LongPtr) As LongPtr
    Dim h As LongPtr
    h = WFD_FindWindowExW(hDlg, 0, StrPtr("Edit"), 0)
    If h = 0 Then
        Dim hCombo As LongPtr
        hCombo = WFD_FindWindowExW(hDlg, 0, StrPtr("ComboBoxEx32"), 0)
        If hCombo <> 0 Then
            h = WFD_FindWindowExW(hCombo, 0, StrPtr("ComboBox"), 0)
            If h <> 0 Then h = WFD_FindWindowExW(h, 0, StrPtr("Edit"), 0)
        End If
    End If
    WFD_FindEditControl = h
End Function
#Else
Private Function WFD_FindEditControl(ByVal hDlg As Long) As Long
    Dim h As Long
    h = WFD_FindWindowExW(hDlg, 0, StrPtr("Edit"), 0)
    If h = 0 Then
        Dim hCombo As Long
        hCombo = WFD_FindWindowExW(hDlg, 0, StrPtr("ComboBoxEx32"), 0)
        If hCombo <> 0 Then
            h = WFD_FindWindowExW(hCombo, 0, StrPtr("ComboBox"), 0)
            If h <> 0 Then h = WFD_FindWindowExW(h, 0, StrPtr("Edit"), 0)
        End If
    End If
    WFD_FindEditControl = h
End Function
#End If
'@

[System.IO.File]::WriteAllText($path, $content, (New-Object System.Text.UTF8Encoding $false))
Write-Host "File created successfully without BOM: $path"
