Attribute VB_Name = "modWinFileDialog"
Option Explicit

' ======================================================================
'  modWinFileDialog  - V1.7 SMART GOS ATT
'  Win32 Unicode API  (FindWindowW / SendMessageW / SetForegroundWindowW)
'  Ho tro ten file + duong dan tieng Viet co dau
'  Ho tro Local Drive (C:\), Network Drive (V:\), UNC (\\server\folder)
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
#End If

Private Const WM_SETTEXT  As Long = &HC
Private Const BM_CLICK    As Long = &HF5
Private Const GWL_STYLE   As Long = -16
Private Const ES_READONLY As Long = &H800

' ======================================================================
'  WinImportFile
'  Dien duong dan vao Windows Native File Dialog (Unicode, hot tro tieng Viet)
'  Tra ve True neu dialog da dong thanh cong sau khi nhan OK
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

    ' Kiem tra file ton tai truoc khi dien vao dialog
    If Len(Dir$(fullPath)) = 0 Then
        WinImportFile = False
        Exit Function
    End If

    ' Cho dialog xuat hien (toi da timeoutSeconds giay)
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

    ' Dua dialog len foreground truoc khi dien
    WFD_BringWindowToTop hDlg
    WFD_SetForegroundWindow hDlg

    ' Tim Edit control de dien duong dan (uu tien control ID 1148, fallback Edit class)
    hFile = WFD_GetDlgItem(hDlg, 1148)
    If hFile = 0 Then
        hFile = WFD_FindEditControl(hDlg)
    End If

    ' Tim nut Open/OK (ID=1)
    hOpen = WFD_GetDlgItem(hDlg, 1)

    If hFile = 0 Or hOpen = 0 Then
        WinImportFile = False
        Exit Function
    End If

    ' Dien duong dan bang WM_SETTEXT (Unicode, dung StrPtr)
    WFD_SendMessageW hFile, WM_SETTEXT, 0, StrPtr(fullPath)
    DoEvents

    ' Nhan nut Open
    WFD_SendMessageW hOpen, BM_CLICK, 0, 0

    ' Cho dialog dong (toi da timeoutSeconds giay)
    t0 = Timer
    Do While WFD_IsWindow(hDlg) <> 0 And Timer - t0 < timeoutSeconds
        DoEvents
    Loop

    WinImportFile = (WFD_IsWindow(hDlg) = 0)
End Function

' ======================================================================
'  WinBringSecurityDialogToFront
'  Tim cua so SAP GUI Security va dua len foreground
'  Tra ve True neu tim thay
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
'  Kiem tra cua so SAP GUI Security co dang mo khong
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

' ======================================================================
'  Private: Tim cua so SAP GUI Security
' ======================================================================
#If VBA7 Then
Private Function WFD_FindSapSecurityWindow() As LongPtr
    Dim hWnd As LongPtr
#Else
Private Function WFD_FindSapSecurityWindow() As Long
    Dim hWnd As Long
#End If
    Dim title As String
    ' Tim theo tieu de cua so SAP Security (tieng Anh va tieng Viet)
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

' ======================================================================
'  Private: Tim Windows File Dialog (class #32770)
' ======================================================================
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

' ======================================================================
'  Private: Tim Edit control trong dialog de dien duong dan
' ======================================================================
#If VBA7 Then
Private Function WFD_FindEditControl(ByVal hDlg As LongPtr) As LongPtr
    Dim h As LongPtr
    h = WFD_FindWindowExW(hDlg, 0, StrPtr("Edit"), 0)
    ' Uu tien tim ComboBoxEx32 > Edit de ho tro Vista+ file dialog
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
