Attribute VB_Name = "modWinFileDialog"
Option Explicit

' ======================================================================
'  modWinFileDialog  - V1.8.3 SMART GOS ATT REBUILT
'  Win32 Unicode API  (FindWindowW / SendMessageW / SetForegroundWindow / GetFileAttributesW)
'  Ho tro ten file + duong dan tieng Viet co dau & UNC Network Paths
' ======================================================================

#If VBA7 Then
    Private Declare PtrSafe Function WFD_FindWindowW Lib "user32" Alias "FindWindowW" _
        (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
    Private Declare PtrSafe Function WFD_FindWindowExW Lib "user32" Alias "FindWindowExW" _
        (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
         ByVal lpszClass As LongPtr, ByVal lpszWindow As LongPtr) As LongPtr
    Private Declare PtrSafe Function WFD_GetDlgItem Lib "user32" Alias "GetDlgItem" _
        (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long) As LongPtr
    Private Declare PtrSafe Function WFD_IsWindow Lib "user32" Alias "IsWindow" _
        (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function WFD_SendMessageW Lib "user32" Alias "SendMessageW" _
        (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function WFD_SetForegroundWindow Lib "user32" Alias "SetForegroundWindow" _
        (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function WFD_BringWindowToTop Lib "user32" Alias "BringWindowToTop" _
        (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function WFD_GetWindowTextW Lib "user32" Alias "GetWindowTextW" _
        (ByVal hwnd As LongPtr, ByVal lpString As LongPtr, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function WFD_GetWindowTextLengthW Lib "user32" Alias "GetWindowTextLengthW" _
        (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function WFD_EnumChildWindows Lib "user32" Alias "EnumChildWindows" _
        (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function WFD_GetClassNameW Lib "user32" Alias "GetClassNameW" _
        (ByVal hwnd As LongPtr, ByVal lpClassName As LongPtr, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function WFD_GetWindowLongPtrW Lib "user32" Alias "GetWindowLongPtrW" _
        (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function WFD_GetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" _
        (ByVal lpFileName As LongPtr) As Long
#Else
    Private Declare Function WFD_FindWindowW Lib "user32" Alias "FindWindowW" _
        (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
    Private Declare Function WFD_FindWindowExW Lib "user32" Alias "FindWindowExW" _
        (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, _
         ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
    Private Declare Function WFD_GetDlgItem Lib "user32" Alias "GetDlgItem" _
        (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
    Private Declare Function WFD_IsWindow Lib "user32" Alias "IsWindow" _
        (ByVal hwnd As Long) As Long
    Private Declare Function WFD_SendMessageW Lib "user32" Alias "SendMessageW" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function WFD_SetForegroundWindow Lib "user32" Alias "SetForegroundWindow" _
        (ByVal hwnd As Long) As Long
    Private Declare Function WFD_BringWindowToTop Lib "user32" Alias "BringWindowToTop" _
        (ByVal hwnd As Long) As Long
    Private Declare Function WFD_GetWindowTextW Lib "user32" Alias "GetWindowTextW" _
        (ByVal hwnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
    Private Declare Function WFD_GetWindowTextLengthW Lib "user32" Alias "GetWindowTextLengthW" _
        (ByVal hwnd As Long) As Long
    Private Declare Function WFD_GetClassNameW Lib "user32" Alias "GetClassNameW" _
        (ByVal hwnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
    Private Declare Function WFD_GetWindowLongPtrW Lib "user32" Alias "GetWindowLongW" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function WFD_GetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" _
        (ByVal lpFileName As Long) As Long
#End If

Private Const WM_SETTEXT  As Long = &HC
Private Const BM_CLICK    As Long = &HF5
Private Const GWL_STYLE   As Long = -16
Private Const ES_READONLY As Long = &H800

' ======================================================================
'  Public Unicode Win32 File/Folder Existence Helpers
' ======================================================================
Public Function ATT_FolderExistsW(ByVal pathStr As String) As Boolean
    Dim attr As Long
    pathStr = Trim`(pathStr)
    If pathStr = "" Then Exit Function
    attr = WFD_GetFileAttributesW(StrPtr(pathStr))
    If attr <> -1 Then
        ATT_FolderExistsW = ((attr And 16) <> 0)
    End If
End Function

Public Function ATT_FileExistsW(ByVal pathStr As String) As Boolean
    Dim attr As Long
    pathStr = Trim`(pathStr)
    If pathStr = "" Then Exit Function
    attr = WFD_GetFileAttributesW(StrPtr(pathStr))
    If attr <> -1 Then
        ATT_FileExistsW = ((attr And 16) = 0)
    End If
End Function

' ======================================================================
'  Public Window Helper Wrappers
' ======================================================================
#If VBA7 Then
Public Function WinFindWindow(ByVal className As String, ByVal windowName As String) As LongPtr
    Dim pClass As LongPtr, pName As LongPtr
    If className <> "" Then pClass = StrPtr(className)
    If windowName <> "" Then pName = StrPtr(windowName)
    WinFindWindow = WFD_FindWindowW(pClass, pName)
End Function

Public Sub WinBringToFront(ByVal hwnd As LongPtr)
    If hwnd <> 0 Then
        WFD_BringWindowToTop hwnd
        WFD_SetForegroundWindow hwnd
    End If
End Sub
#Else
Public Function WinFindWindow(ByVal className As String, ByVal windowName As String) As Long
    Dim pClass As Long, pName As Long
    If className <> "" Then pClass = StrPtr(className)
    If windowName <> "" Then pName = StrPtr(windowName)
    WinFindWindow = WFD_FindWindowW(pClass, pName)
End Function

Public Sub WinBringToFront(ByVal hwnd As Long)
    If hwnd <> 0 Then
        WFD_BringWindowToTop hwnd
        WFD_SetForegroundWindow hwnd
    End If
End Sub
#End If

' ======================================================================
'  WinImportFile
'  Dien duong dan vao Windows Native File Dialog (Unicode, hot tro tieng Viet)
' ======================================================================
Public Function WinImportFile(ByVal fullPath As String, _
                              Optional ByVal timeoutSeconds As Long = 20) As Boolean
#If VBA7 Then
    Dim hDlg  As LongPtr
    Dim hFile As LongPtr
    Dim hOpen As LongPtr
#Else
    Dim hDlg  As Long
    Dim hFile As Long
    Dim hOpen As Long
#End If
    Dim t0 As Single

    If Not ATT_FileExistsW(fullPath) Then
        WinImportFile = False
        Exit Function
    End If

    t0 = Timer
    Do
        hDlg = WFD_FindDialog()
        If hDlg <> 0 Then Exit Do
        DoEvents
        If Timer - t0 > timeoutSeconds Then
            WinImportFile = False
            Exit Function
        End If
    Loop

    If hDlg = 0 Then Exit Function

    WFD_BringWindowToTop hDlg
    WFD_SetForegroundWindow hDlg

    hFile = WFD_GetDlgItem(hDlg, 1148)
    If hFile = 0 Then
        hFile = WFD_FindEditControl(hDlg)
    End If

    hOpen = WFD_GetDlgItem(hDlg, 1)

    If hFile = 0 Or hOpen = 0 Then
        WinImportFile = False
        Exit Function
    End If

    WFD_SendMessageW hFile, WM_SETTEXT, 0, StrPtr(fullPath)
    DoEvents

    WFD_SendMessageW hOpen, BM_CLICK, 0, 0

    t0 = Timer
    Do While WFD_IsWindow(hDlg) <> 0 And Timer - t0 < timeoutSeconds
        DoEvents
    Loop

    WinImportFile = (WFD_IsWindow(hDlg) = 0)
End Function

' ======================================================================
'  WinBringSecurityDialogToFront
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
        WFD_SetForegroundWindow hSec
        WinBringSecurityDialogToFront = True
    End If
End Function

' ======================================================================
'  WinIsSapSecurityVisible
' ======================================================================
Public Function WinIsSapSecurityVisible() As Boolean
#If VBA7 Then
    Dim hSec As LongPtr
#Else
    Dim hSec As Long
#End If
    hSec = WFD_FindSapSecurityWindow()
    WinIsSapSecurityVisible = (hSec <> 0 And WFD_IsWindow(hSec) <> 0)
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
        If hWnd <> 0 Then
            WFD_FindSapSecurityWindow = hWnd
            Exit Function
        End If
    Next i
End Function

#If VBA7 Then
Private Function WFD_FindDialog() As LongPtr
    Dim className As String
    Dim hWnd As LongPtr
    className = "#32770"
    hWnd = WFD_FindWindowW(StrPtr(className), 0)
    WFD_FindDialog = hWnd
End Function
#Else
Private Function WFD_FindDialog() As Long
    Dim className As String
    Dim hWnd As Long
    className = "#32770"
    hWnd = WFD_FindWindowW(StrPtr(className), 0)
    WFD_FindDialog = hWnd
End Function
#End If

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
