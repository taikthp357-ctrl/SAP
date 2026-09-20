Attribute VB_Name = "modWinFileDialog"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
    Private Declare PtrSafe Function GetDlgItem Lib "user32" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String) As Long
    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
    Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Private Const BM_CLICK As Long = &HF5

Public Function WinImportFile(ByVal fullPath As String, Optional ByVal timeoutSeconds As Long = 15) As Boolean
#If VBA7 Then
    Dim hDlg As LongPtr, hFile As LongPtr, hOpen As LongPtr
#Else
    Dim hDlg As Long, hFile As Long, hOpen As Long
#End If
    Dim t0 As Single
    If Len(Dir$(fullPath)) = 0 Then Exit Function
    t0 = Timer
    Do
        hDlg = FindWindow("#32770", vbNullString)
        If hDlg <> 0 Then Exit Do
        DoEvents
    Loop While Timer - t0 < timeoutSeconds
    If hDlg = 0 Then Exit Function
    hFile = GetDlgItem(hDlg, 1148)
    If hFile = 0 Then hFile = FindWindowEx(hDlg, 0, "Edit", vbNullString)
    hOpen = GetDlgItem(hDlg, 1)
    If hFile = 0 Or hOpen = 0 Then Exit Function
    If SetWindowText(hFile, fullPath) = 0 Then Exit Function
    Call SendMessage(hOpen, BM_CLICK, 0, 0)
    t0 = Timer
    Do While IsWindow(hDlg) <> 0 And Timer - t0 < timeoutSeconds
        DoEvents
    Loop
    WinImportFile = (IsWindow(hDlg) = 0)
End Function
