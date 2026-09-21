VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit
    If Not Intersect(Target, Me.Range("C3:C4")) Is Nothing Then
        application.EnableEvents = False
        If Not Intersect(Target, Me.Range("C3")) Is Nothing Then
            ' Khi ngu?i dùng d?i ô C3 -> d?ng b? l?i Pending và không gi? l?p t?i MB52
            Call Sync_Pending_From_Model(Me.Range("C3").Value)
            Me.Range("C7").Value = "Chua t?i MB52"
        End If
        Call Calculate_All_Stock(False, False)
        application.EnableEvents = True
    End If
CleanExit:
    application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, Me.Range("C3")) Is Nothing Then
        Cancel = True
        Call Show_Select_Models
    End If
End Sub

