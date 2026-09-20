Attribute VB_Name = "ThisWorkbook"

' ═══════════════════════════════════════════════════════════════════════════════
' ThisWorkbook — Workbook_Open hook for CF_CONTROL
' ═══════════════════════════════════════════════════════════════════════════════

Private Sub Workbook_Open()
    On Error Resume Next

    ' Navigate to CF_CONTROL sheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")

    If Not ws Is Nothing Then
        ws.Activate

        ' Set zoom
        ActiveWindow.Zoom = 95

        ' Scroll to top
        ws.Range("A1").Select

        ' Set freeze pane (freeze row 11, keep headers visible)
        ActiveWindow.FreezePanes = False
        ws.Range("A11").Select
        ActiveWindow.FreezePanes = True

        ' Return to top-left
        ws.Range("A1").Select

        ' Update button state and refresh data if DATA_CORE has rows
        Dim wsCore As Worksheet
        Set wsCore = ThisWorkbook.Worksheets("DATA_CORE")
        If Not wsCore Is Nothing Then
            Dim lastRow As Long
            lastRow = wsCore.Cells(wsCore.Rows.Count, 1).End(xlUp).Row
            If lastRow >= 5 Then
                ' Has data, refresh table
                modCF_ControlData.CF_RefreshControlTable
                modCF_ControlLayout.CF_UpdateCheckExecuteButton
            End If
        End If

        modCF_ControlLayout.CF_UpdateStatusBar "Ready"
    End If

    On Error GoTo 0
End Sub
