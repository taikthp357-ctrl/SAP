Attribute VB_Name = "Sheet_CF_CONTROL"

' ═══════════════════════════════════════════════════════════════════════════════
' CF_CONTROL Sheet — Worksheet event handlers (28-point spec)
' ═══════════════════════════════════════════════════════════════════════════════

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next

    ' OB input area (A7:A40) — validate for numeric OBs and update count
    If Not Intersect(Target, Me.Range("A7:A40")) Is Nothing Then
        CF_UpdateOBCount Me
    End If

    On Error GoTo 0
End Sub

Private Sub Worksheet_Deactivate()
    On Error Resume Next

    ' Auto-sync user edits in ACTUAL cols (J:L) back to DATA_CORE
    modCF_InputWorklist.CF_SyncControlToCore

    On Error GoTo 0
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next

    ' Update CHECK/XỬ LÝ button state when returning to the sheet
    modCF_InputLayout.CF_UpdateCheckExecuteButton

    On Error GoTo 0
End Sub

' ─────────────────────────────────────────────────────────────────────────────
' Update OB count display in B41:C41
' ─────────────────────────────────────────────────────────────────────────────
Private Sub CF_UpdateOBCount(ws As Worksheet)
    Dim r As Long, count As Long, dups As Long
    Dim seen As Object
    Dim obVal As String

    Set seen = CreateObject("Scripting.Dictionary")
    count = 0
    dups = 0

    For r = 7 To 40
        obVal = Trim(CStr(ws.Cells(r, 1).Value))
        If obVal = "" Then GoTo NextRow
        If Not IsNumeric(obVal) Then GoTo NextRow

        count = count + 1
        If seen.Exists(obVal) Then
            dups = dups + 1
            ws.Cells(r, 1).Interior.Color = RGB(255, 182, 182) ' Light red = duplicate
        Else
            seen.Add obVal, 1
            ws.Cells(r, 1).Interior.Color = RGB(255, 255, 192) ' Restore yellow
        End If
NextRow:
    Next r

    ws.Cells(41, 2).Value = "Count:"
    ws.Cells(41, 3).Value = count
    If dups > 0 Then
        ws.Cells(41, 3).Font.Color = RGB(192, 0, 0)
        ws.Cells(41, 3).Value = count & " (" & dups & " dup)"
    Else
        ws.Cells(41, 3).Font.Color = RGB(0, 0, 0)
    End If
End Sub
