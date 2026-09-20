VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next

    ' Initialize status cells (MON/SAP/LOAD state display)
    RT_InitializeUI

    ' Navigate to CF_CONTROL and apply window settings
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")

    If Not ws Is Nothing Then
        ws.Activate

        With ActiveWindow
            .Zoom = 95
            .DisplayGridlines = False
            .DisplayHeadings = False

            ' Freeze pane after worklist header (row 8)
            .FreezePanes = False
            ws.Range("A9").Select
            .FreezePanes = True

            ' Scroll to top-left
            ws.Range("A1").Select
        End With

        ' Update button state
        modCF_InputLayout.CF_UpdateCheckExecuteButton

        ' Refresh worklist if DATA_CORE has data
        On Error Resume Next
        Dim wsCore As Worksheet
        Set wsCore = ThisWorkbook.Worksheets("DATA_CORE")
        If Not wsCore Is Nothing Then
            If wsCore.Cells(5, 1).Value2 <> "" Then
                modCF_InputWorklist.CF_RefreshInputWorklist
            End If
        End If
        On Error GoTo 0

        modCF_InputLayout.CF_UpdateStatusBar "Ready"
    End If

    On Error GoTo 0
End Sub
