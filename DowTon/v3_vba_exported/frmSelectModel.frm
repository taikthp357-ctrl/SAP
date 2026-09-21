VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectModel 
   Caption         =   "Ch?n Model (Multi-Select)"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmSelectModel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim wsCheck As Worksheet
    Dim lastRow As Long
    Dim dicModels As Object
    Dim r As Long, m As String
    Dim k As Variant
    Dim curVal As String
    Dim isAll As Boolean
    Dim i As Long
    Dim parts() As String, p As Variant
    Dim selDic As Object
    
    Me.Caption = "Ch" & ChrW(7885) & "n Model (Multi-Select)"
    Me.lblInfo.Caption = "Ch" & ChrW(7885) & "n m" & ChrW(7897) & "t ho" & ChrW(7863) & "c nhi" & ChrW(7873) & "u Model c" & ChrW(7847) & "n theo d" & ChrW(245) & "i:"
    Me.btnSelectAll.Caption = "Ch" & ChrW(7885) & "n T" & ChrW(7845) & "t C" & ChrW(7843)
    Me.btnClearAll.Caption = "B" & ChrW(7887) & " Ch" & ChrW(7885) & "n"
    Me.btnApply.Caption = ChrW(193) & "p D" & ChrW(7909) & "ng"
    
    Set wsCheck = GetCheckSheet(ThisWorkbook)
    lastRow = wsCheck.Cells(wsCheck.Rows.Count, "A").End(xlUp).Row
    
    Set dicModels = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        m = Trim(CStr(wsCheck.Cells(r, 10).Value))
        If Len(m) > 0 Then dicModels(m) = True
    Next r
    
    Me.lstModels.Clear
    For Each k In dicModels.Keys
        Me.lstModels.AddItem CStr(k)
    Next k
    
    ' Pre-select based on C3
    curVal = Trim(CStr(ThisWorkbook.Sheets("DASHBOARD").Range("C3").Value))
    isAll = (Len(curVal) = 0 Or UCase(curVal) = "ALL" Or curVal = UniText("Tat_ca") Or curVal = "T" & ChrW(7845) & "t c" & ChrW(7843))
    
    If isAll Then
        For i = 0 To Me.lstModels.ListCount - 1
            Me.lstModels.Selected(i) = True
        Next i
    Else
        parts = Split(curVal, ",")
        Set selDic = CreateObject("Scripting.Dictionary")
        For Each p In parts
            If Len(Trim(CStr(p))) > 0 Then selDic(UCase(Trim(CStr(p)))) = True
        Next p
        
        For i = 0 To Me.lstModels.ListCount - 1
            If selDic.Exists(UCase(Me.lstModels.List(i))) Then
                Me.lstModels.Selected(i) = True
            Else
                Me.lstModels.Selected(i) = False
            End If
        Next i
    End If
End Sub

Private Sub btnSelectAll_Click()
    Dim i As Long
    For i = 0 To Me.lstModels.ListCount - 1
        Me.lstModels.Selected(i) = True
    Next i
End Sub

Private Sub btnClearAll_Click()
    Dim i As Long
    For i = 0 To Me.lstModels.ListCount - 1
        Me.lstModels.Selected(i) = False
    Next i
End Sub

Private Sub btnApply_Click()
    Dim selCount As Long, i As Long
    Dim res As String
    selCount = 0: res = ""
    
    For i = 0 To Me.lstModels.ListCount - 1
        If Me.lstModels.Selected(i) Then
            selCount = selCount + 1
            If Len(res) > 0 Then res = res & ", "
            res = res & Me.lstModels.List(i)
        End If
    Next i
    
    application.EnableEvents = False
    If selCount = 0 Or selCount = Me.lstModels.ListCount Then
        ThisWorkbook.Sheets("DASHBOARD").Range("C3").Value = "ALL"
    Else
        ThisWorkbook.Sheets("DASHBOARD").Range("C3").Value = res
    End If
    
    ' Ð?ng b? Pending theo Model dã ch?n (unique list)
    Call Sync_Pending_From_Model(ThisWorkbook.Sheets("DASHBOARD").Range("C3").Value)
    
    ' Ð?i Model KHÔNG gi? l?p Ðã t?i MB52
    ThisWorkbook.Sheets("DASHBOARD").Range("C7").Value = "Chua t?i MB52"
    
    Call Calculate_All_Stock(False, False)
    application.EnableEvents = True
    
    Unload Me
End Sub

