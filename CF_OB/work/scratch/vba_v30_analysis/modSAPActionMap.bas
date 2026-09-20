Attribute VB_Name = "modSAPActionMap"
Option Explicit

' Action map is a discovery/review layer. It is deliberately separate from
' CONFIG lock so a discovered command can never become a RUN permission by
' accident.

Public Sub SAPActionMap_CompileProbe()
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("SYSTEM")
    If Trim$(CStr(ws.Cells(1, 27).Value2)) <> "ActionKey" Then
        Err.Raise vbObjectError + 9601, , "SAP_ACTION_MAP header missing."
    End If
    Exit Sub
Failed:
    Err.Raise Err.Number, "SAPActionMap_CompileProbe", Err.Description
End Sub

Public Function SAPActionMap_RequiredKeys() As String
    SAPActionMap_RequiredKeys = _
        "MON_WT_GRID_ID,MON_WT_ACTION_ID,MON_CANCEL_COMMAND_ID," & _
        "MON_CONFIRM_FG_COMMAND_ID,MON_CONFIRM_BG_COMMAND_ID," & _
        "WT_REQUEST_TOOLBAR_ID,WT_SAVE_COMMAND_ID,WT_CREATE_AND_SAVE_COMMAND_ID," & _
        "WT_CREATE_COMMAND_ID,WT_DELETE_COMMAND_ID,WT_SOURCE_BIN_ID,WT_QUANTITY_ID," & _
        "OB_EDIT_COMMAND_ID,OB_QUANTITY_COL," & _
        "WT_REQUEST_GRID,WT_TAB_STK_CAN_BE_REMOVED,WT_STOCK_GRID," & _
        "WT_STOCK_TOOLBAR,WT_STOCK_SET_QTY,WT_STOCK_SELECT_ROW," & _
        "WT_ADOPT_QUANTITY,WT_TAB_WAREHOUSE_TASK,WT_WAREHOUSE_TASK_GRID," & _
        "WT_CREATE_FOLLOWON,WT_CREATE_QUERY,WT_CREATE_COMMAND_ID,WT_SAVE_COMMAND_ID,WT_CREATE_SAVE"
End Function

Public Function SAPActionMap_Review(ByRef reportText As String) As Boolean
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim keyText As String, idText As String, statusText As String
    Dim seen As Object, errors As Long

    reportText = ""
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("SYSTEM")
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    lastRow = ws.Cells(ws.rows.count, 27).End(xlUp).row

    For r = 2 To lastRow
        keyText = Trim$(CStr(ws.Cells(r, 27).Value2))
        idText = Trim$(CStr(ws.Cells(r, 31).Value2))
        statusText = UCase$(Trim$(CStr(ws.Cells(r, 32).Value2)))
        If keyText <> "" Then
            If seen.Exists(keyText) Then
                errors = errors + 1
                reportText = reportText & "DUPLICATE_KEY=" & keyText & ";"
            Else
                seen.Add keyText, True
            End If
            If statusText = "DISCOVERED" And idText = "" Then
                errors = errors + 1
                reportText = reportText & "DISCOVERED_WITHOUT_ID=" & keyText & ";"
            End If
        End If
    Next r

    If errors = 0 Then
        SAPActionMap_Review = True
        reportText = "PASS; rows=" & CStr(lastRow - 1)
    Else
        reportText = "FAIL; errors=" & CStr(errors) & ";" & reportText
    End If
    Exit Function
Failed:
    reportText = "FAIL; " & Err.Description
End Function

