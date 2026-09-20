Attribute VB_Name = "modWinFileDialog"
Option Explicit

#If VBA7 Then
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
    Private mFileDlg As LongPtr, mSecurityDlg As LongPtr, mChildEdit As LongPtr, mChildButton As LongPtr
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
    Private mFileDlg As Long, mSecurityDlg As Long, mChildEdit As Long, mChildButton As Long
#End If

Private Const BM_CLICK As Long = &HF5
Private Const WM_SETTEXT As Long = &HC
Private mFindChildMode As Long
Private mLastMethod As String

Public Function WinLastMethod() As String
    WinLastMethod = mLastMethod
End Function

Private Function WinElapsed(ByVal t0 As Double) As Double
    If Timer >= t0 Then WinElapsed = Timer - t0 Else WinElapsed = 86400# - t0 + Timer
End Function

#If VBA7 Then
Public Function WinEnumTopProc(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Public Function WinEnumTopProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
#End If
    Dim cls As String, ttl As String
    WinEnumTopProc = 1
    If IsWindowVisible(hwnd) = 0 Then Exit Function
    cls = WinClass(hwnd)
    If cls <> "#32770" Then Exit Function
    ttl = UCase$(WinText(hwnd))
    If InStr(ttl, "SAP") > 0 And (InStr(ttl, "SECUR") > 0 Or InStr(ttl, "GUI") > 0) Then mSecurityDlg = hwnd
    If WinFileEdit(hwnd) <> 0 Then mFileDlg = hwnd
End Function

#If VBA7 Then
Public Function WinEnumChildProc(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Public Function WinEnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
#End If
    Dim cls As String, ttl As String
    WinEnumChildProc = 1
    cls = UCase$(WinClass(hwnd)): ttl = UCase$(Trim$(WinText(hwnd)))
    If mFindChildMode = 1 Then
        If cls = "EDIT" And IsWindowVisible(hwnd) <> 0 Then mChildEdit = hwnd
    ElseIf mFindChildMode = 2 Then
        If cls = "BUTTON" Then
            If InStr(ttl, "ALLOW") > 0 Or InStr(ttl, "PERMIT") > 0 Or InStr(ttl, "YES") > 0 Or _
               InStr(ttl, "CONTINUE") > 0 Or InStr(ttl, "CHO PHEP") > 0 Or InStr(ttl, "OK") > 0 Then
                mChildButton = hwnd: WinEnumChildProc = 0
            End If
        End If
    End If
End Function

#If VBA7 Then
Private Function WinFileEdit(ByVal hDlg As LongPtr) As LongPtr
    Dim h As LongPtr
#Else
Private Function WinFileEdit(ByVal hDlg As Long) As Long
    Dim h As Long
#End If
    h = GetDlgItem(hDlg, 1148)
    If h = 0 Then
        mChildEdit = 0: mFindChildMode = 1
        EnumChildWindows hDlg, AddressOf WinEnumChildProc, 0
        h = mChildEdit
    End If
    WinFileEdit = h
End Function

#If VBA7 Then
Private Function WinClass(ByVal hwnd As LongPtr) As String
#Else
Private Function WinClass(ByVal hwnd As Long) As String
#End If
    Dim b As String, n As Long
    b = String$(128, vbNullChar)
    n = GetClassNameW(hwnd, StrPtr(b), Len(b))
    If n > 0 Then WinClass = Left$(b, n)
End Function

#If VBA7 Then
Private Function WinText(ByVal hwnd As LongPtr) As String
#Else
Private Function WinText(ByVal hwnd As Long) As String
#End If
    Dim b As String, n As Long
    b = String$(512, vbNullChar)
    n = GetWindowTextW(hwnd, StrPtr(b), Len(b))
    If n > 0 Then WinText = Left$(b, n)
End Function

Public Function WinHandleSapSecurity(Optional ByVal timeoutSeconds As Long = 0) As Boolean
#If VBA7 Then
    Dim h As LongPtr
#Else
    Dim h As Long
#End If
    Dim t0 As Double
    t0 = Timer
    Do
        mSecurityDlg = 0: EnumWindows AddressOf WinEnumTopProc, 0
        h = mSecurityDlg
        If h <> 0 Then
            mChildButton = 0: mFindChildMode = 2
            EnumChildWindows h, AddressOf WinEnumChildProc, 0
            If mChildButton <> 0 Then
                SetForegroundWindow h
                SendMessageW mChildButton, BM_CLICK, 0, 0
                mLastMethod = "WINDOWS_SECURITY_BUTTON"
                WinHandleSapSecurity = True
                Exit Function
            End If
        End If
        DoEvents
    Loop While WinElapsed(t0) < timeoutSeconds
End Function

Public Function WinImportFile(ByVal fullPath As String, Optional ByVal timeoutSeconds As Long = 15) As Boolean
#If VBA7 Then
    Dim hDlg As LongPtr, hFile As LongPtr, hOpen As LongPtr, hFore As LongPtr
#Else
    Dim hDlg As Long, hFile As Long, hOpen As Long, hFore As Long
#End If
    Dim t0 As Double
    If Len(Dir$(fullPath)) = 0 Then Exit Function
    t0 = Timer
    Do
        WinHandleSapSecurity 0
        hDlg = 0: hFile = 0: hOpen = 0
        hFore = GetForegroundWindow()
        If hFore <> 0 Then
            If WinClass(hFore) = "#32770" Then
                hFile = WinFileEdit(hFore)
                If hFile <> 0 Then hDlg = hFore
            End If
        End If
        If hDlg = 0 Then
            mFileDlg = 0: EnumWindows AddressOf WinEnumTopProc, 0
            hDlg = mFileDlg
            If hDlg <> 0 Then hFile = WinFileEdit(hDlg)
        End If
        If hDlg <> 0 And hFile <> 0 Then
            hOpen = GetDlgItem(hDlg, 1)
            If hOpen <> 0 Then
                SetForegroundWindow hDlg
                SendMessageW hFile, WM_SETTEXT, 0, StrPtr(fullPath)
                SendMessageW hOpen, BM_CLICK, 0, 0
                mLastMethod = "WINDOWS_UNICODE_DIALOG"
                Do While IsWindow(hDlg) <> 0 And WinElapsed(t0) < timeoutSeconds
                    WinHandleSapSecurity 0
                    DoEvents
                Loop
                WinImportFile = (IsWindow(hDlg) = 0)
                Exit Function
            End If
        End If
        DoEvents
    Loop While WinElapsed(t0) < timeoutSeconds
End Function
