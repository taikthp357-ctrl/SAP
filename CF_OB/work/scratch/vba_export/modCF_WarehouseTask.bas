Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub CFWT_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub CFWT_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

Private Const CFWT_GRID_ID As String = "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell"
Private Const CFWT_UPPER_GRID_ID As String = "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell"
Public CFWT_LastDiagnostic As String
Public CFWT_WT_ACTION_CLICK_COUNT As Long
Private mCFWT_NodeBudget As Long

' Fail-safe: touching SAP GUI controls while the session is mid-repaint
' ("Busy") is the most common trigger of an untrappable "Automation error:
' Catastrophic failure". This version re-checks ses.Busy on every poll and
' only calls FindById once the session is confirmed idle.
Public Sub CFWT_ResetNavigationCounters()
    CFWT_WT_ACTION_CLICK_COUNT = 0
End Sub

Public Function CFWT_SelectWarehouseTask(ByVal ses As Object, ByVal timeoutSeconds As Double, _
                                          Optional ByVal forceActionClick As Boolean = False) As Boolean
    Dim started As Double, grid As Object, isBusy As Boolean, activated As Boolean, activationAttempted As Boolean
    Dim gridReady As Boolean
    CFWT_LastDiagnostic = ""
    mCFWT_NodeBudget = 4000
    If ses Is Nothing Then CFWT_LastDiagnostic = "SAP session is Nothing.": Exit Function
    started = Timer
    Do
        DoEvents

        isBusy = True
        On Error Resume Next
        isBusy = CBool(ses.busy)
        On Error GoTo 0

        If Not isBusy Then
            Set grid = Nothing
            On Error Resume Next
            Set grid = ses.FindById(CFWT_GridID(), False)
            On Error GoTo 0

            gridReady = False
            If Not grid Is Nothing Then
                gridReady = CFWT_HasColumn(grid, CFWT_ColumnID("MON_WT_DOC_COL", "DOCNO")) And _
                            CFWT_HasColumn(grid, CFWT_ColumnID("MON_WT_WHO_COL", "WHO")) And _
                            CFWT_HasColumn(grid, CFWT_ColumnID("MON_WT_COL", "TANUM")) And _
                            CFWT_HasColumn(grid, CFWT_ColumnID("MON_WT_BIN_COL", "VLPLA"))

                ' Normal mode may reuse an already materialized WT grid.  In
                ' forced mode the caller explicitly requires one fresh action
                ' click, but the function must still succeed once that click
                ' has produced a valid grid.  The previous implementation
                ' skipped this return path whenever forceActionClick=True and
                ' therefore timed out on every forced navigation attempt.
                If gridReady And ((Not forceActionClick) Or (activationAttempted And activated)) Then
                    CFWT_SelectWarehouseTask = True
                    Set grid = Nothing
                    Exit Function
                End If
            End If
            Set grid = Nothing

            If SAPRun_IsRunMode() Then
                If Not activationAttempted Then
                    activationAttempted = True
                    activated = CFWT_ActivateWarehouseTaskMapped(ses)
                If Not activated Then CFWT_LastDiagnostic = "STOP_CONTROL_NOT_FOUND: MON_WT_ACTION_ID"
                End If
            ElseIf Not activationAttempted Then
                activationAttempted = True
                activated = CFWT_ActivateWarehouseTaskToolbar(ses)
                If Not activated Then activated = CFWT_ActivateWarehouseTaskContextMenu(ses)
                If Not activated Then activated = CFWT_ActivateWarehouseTaskTab(ses)
                If activated Then CFWT_LastDiagnostic = "Da activate tab Warehouse Task; dang cho grid."
            End If
            If activated Then CFWT_WT_ACTION_CLICK_COUNT = CFWT_WT_ACTION_CLICK_COUNT + 1
        End If

        If CFWT_Elapsed(started) >= timeoutSeconds Then Exit Do
        CFWT_Pause 100
    Loop
    If CFWT_LastDiagnostic = "" Then CFWT_LastDiagnostic = "Khong tim thay grid Warehouse Task hoac thieu cot DOCNO/WHO/TANUM/VLPLA."
    mCFWT_NodeBudget = 0
End Function

Private Function CFWT_GridID() As String
    If SAPRun_IsRunMode() Then
        CFWT_GridID = Trim$(ModuleHelper.ToolConfig("MON_WT_GRID_ID", ""))
    Else
        CFWT_GridID = CFWT_GRID_ID
    End If
End Function

Private Function CFWT_ColumnID(ByVal keyName As String, ByVal fallbackText As String) As String
    If SAPRun_IsRunMode() Then
        CFWT_ColumnID = Trim$(ModuleHelper.ToolConfig(keyName, ""))
    Else
        CFWT_ColumnID = fallbackText
    End If
End Function

Private Function CFWT_ActivateWarehouseTaskMapped(ByVal ses As Object) As Boolean
    Dim grid As Object, actionID As String
    If ses Is Nothing Then Exit Function
    On Error Resume Next
    Set grid = ses.FindById(ModuleHelper.ToolConfig("MON_RESULT_GRID_ID", ""), False)
    actionID = Trim$(ModuleHelper.ToolConfig("MON_WT_ACTION_ID", ""))
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If grid Is Nothing Or actionID = "" Then Exit Function
    On Error Resume Next
    Err.Clear
    CallByName grid, "PressToolbarButton", VbMethod, actionID
    CFWT_ActivateWarehouseTaskMapped = (Err.Number = 0)
    If Err.Number <> 0 Then CFWT_LastDiagnostic = "STOP_CONTROL_NOT_FOUND: MON_WT_ACTION_ID=" & actionID & "; Err " & CStr(Err.Number) & ": " & Err.Description
    Err.Clear
    On Error GoTo 0
End Function

Private Function CFWT_ActivateWarehouseTaskToolbar(ByVal ses As Object) As Boolean
    Dim grid As Object, buttonCount As Long, pos As Long
    Dim buttonId As String, buttonText As String, buttonTip As String
    Dim errText As String
    If ses Is Nothing Then Exit Function
    On Error Resume Next
    Set grid = ses.FindById(CFWT_UPPER_GRID_ID, False)
    If grid Is Nothing Then
        Err.Clear
        Exit Function
    End If
    Err.Clear
    buttonCount = CLng(CallByName(grid, "ToolbarButtonCount", VbGet))
    If Err.Number <> 0 Then
        CFWT_LastDiagnostic = "Upper grid has no readable toolbar. Err " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    For pos = 0 To buttonCount - 1
        buttonId = "": buttonText = "": buttonTip = ""
        On Error Resume Next
        Err.Clear
        buttonId = CStr(CallByName(grid, "GetToolbarButtonId", VbMethod, pos))
        buttonText = CStr(CallByName(grid, "GetToolbarButtonText", VbMethod, pos))
        buttonTip = CStr(CallByName(grid, "GetToolbarButtonTooltip", VbMethod, pos))
        Err.Clear
        On Error GoTo 0
        If InStr(1, UCase$(buttonId & "|" & buttonText & "|" & buttonTip), "WAREHOUSE TASK", vbTextCompare) > 0 Then
            If buttonId = "" Then
                CFWT_LastDiagnostic = "Warehouse Task toolbar button found without technical ID at position " & CStr(pos)
                Exit Function
            End If
            On Error Resume Next
            Err.Clear
            CallByName grid, "PressToolbarButton", VbMethod, buttonId
            If Err.Number = 0 Then CFWT_ActivateWarehouseTaskToolbar = True
            If Err.Number <> 0 Then errText = " Err " & CStr(Err.Number) & ": " & Err.Description
            Err.Clear
            On Error GoTo 0
            If Not CFWT_ActivateWarehouseTaskToolbar Then CFWT_LastDiagnostic = "Warehouse Task toolbar press failed; ID=" & buttonId & errText
            Exit Function
        End If
    Next pos
    CFWT_LastDiagnostic = "Warehouse Task toolbar button not found; Count=" & CStr(buttonCount)
End Function

Private Function CFWT_ActivateWarehouseTaskContextMenu(ByVal ses As Object) As Boolean
    Dim grid As Object, root As Object, ctl As Object, actionText As String
    If ses Is Nothing Then Exit Function
    On Error Resume Next
    Set grid = ses.FindById(CFWT_UPPER_GRID_ID, False)
    If grid Is Nothing Then
        Err.Clear
        Exit Function
    End If
    Err.Clear
    grid.ContextMenu
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    CFWT_Pause 150
    On Error Resume Next
    Set root = ses.FindById("wnd[1]", False)
    On Error GoTo 0
    If root Is Nothing Then Exit Function

    Set ctl = CFWT_FindTextControl(root, "WAREHOUSE TASK", 0)
    If ctl Is Nothing Then Set ctl = CFWT_FindTextControl(root, "WAREHOUSE TASKS", 0)
    If ctl Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    CallByName ctl, "Select", VbMethod
    If Err.Number = 0 Then CFWT_ActivateWarehouseTaskContextMenu = True
    If Not CFWT_ActivateWarehouseTaskContextMenu Then
        Err.Clear
        CallByName ctl, "Press", VbMethod
        If Err.Number = 0 Then CFWT_ActivateWarehouseTaskContextMenu = True
    End If
    If Not CFWT_ActivateWarehouseTaskContextMenu Then
        Err.Clear
        CallByName ctl, "Click", VbMethod
        If Err.Number = 0 Then CFWT_ActivateWarehouseTaskContextMenu = True
    End If
    If Err.Number <> 0 Then actionText = " Err " & CStr(Err.Number) & ": " & Err.Description
    Err.Clear
    On Error GoTo 0
    If Not CFWT_ActivateWarehouseTaskContextMenu Then CFWT_LastDiagnostic = "Context menu found but Warehouse Task action failed." & actionText
End Function

Private Function CFWT_ActivateWarehouseTaskTab(ByVal ses As Object) As Boolean
    Dim root As Object, ctl As Object, actionText As String, ignored As Variant
    If ses Is Nothing Then Exit Function
    On Error Resume Next
    Set root = ses.FindById("wnd[0]", False)
    On Error GoTo 0
    If root Is Nothing Then
        CFWT_LastDiagnostic = "Khong tim thay wnd[0] de tim tab Warehouse Task."
        Exit Function
    End If
    Set ctl = CFWT_FindTextControl(root, "WAREHOUSE TASK", 0)
    If ctl Is Nothing Then
        Set ctl = CFWT_FindTextControl(root, "WAREHOUSE TASKS", 0)
    End If
    If ctl Is Nothing Then
        CFWT_LastDiagnostic = "Khong tim thay control/tab co text Warehouse Task tren MON."
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    ignored = CallByName(ctl, "Select", VbMethod)
    If Err.Number = 0 Then CFWT_ActivateWarehouseTaskTab = True
    If Not CFWT_ActivateWarehouseTaskTab Then
        Err.Clear
        ignored = CallByName(ctl, "Press", VbMethod)
        If Err.Number = 0 Then CFWT_ActivateWarehouseTaskTab = True
    End If
    If Not CFWT_ActivateWarehouseTaskTab Then
        Err.Clear
        ignored = CallByName(ctl, "Click", VbMethod)
        If Err.Number = 0 Then CFWT_ActivateWarehouseTaskTab = True
    End If
    If Err.Number <> 0 Then actionText = " Err " & CStr(Err.Number) & ": " & Err.Description
    Err.Clear
    On Error GoTo 0
    If Not CFWT_ActivateWarehouseTaskTab Then CFWT_LastDiagnostic = "Tim thay tab nhung khong activate duoc." & actionText
End Function

Private Function CFWT_FindTextControl(ByVal obj As Object, ByVal wantedText As String, ByVal depth As Long) As Object
    Dim children As Object, child As Object, i As Long, count As Long, ignored As Variant
    Dim textValue As String, tipValue As String, nameValue As String, typeValue As String
    If obj Is Nothing Or depth > 12 Then Exit Function
    If mCFWT_NodeBudget <= 0 Then Exit Function
    mCFWT_NodeBudget = mCFWT_NodeBudget - 1
    On Error Resume Next
    textValue = UCase$(Trim$(CStr(CallByName(obj, "Text", VbGet) & "")))
    tipValue = UCase$(Trim$(CStr(CallByName(obj, "Tooltip", VbGet) & "")))
    nameValue = UCase$(Trim$(CStr(CallByName(obj, "Name", VbGet) & "")))
    typeValue = UCase$(Trim$(CStr(CallByName(obj, "Type", VbGet) & "")))
    On Error GoTo 0
    If InStr(1, textValue, wantedText, vbTextCompare) > 0 Or _
       InStr(1, tipValue, wantedText, vbTextCompare) > 0 Or _
       InStr(1, nameValue, wantedText, vbTextCompare) > 0 Then
        If InStr(typeValue, "TAB") > 0 Or InStr(typeValue, "BUTTON") > 0 Or _
           InStr(typeValue, "SHELL") > 0 Or textValue <> "" Then
            Set CFWT_FindTextControl = obj
            Exit Function
        End If
    End If

    On Error Resume Next
    Set children = obj.children
    count = 0
    If Not children Is Nothing Then count = CLng(children.count)
    On Error GoTo 0
    For i = 0 To count - 1
        Set child = CFWT_ChildAt(children, i)
        If Not child Is Nothing Then
            Set CFWT_FindTextControl = CFWT_FindTextControl(child, wantedText, depth + 1)
            If Not CFWT_FindTextControl Is Nothing Then Exit Function
        End If
    Next i
End Function

Private Function CFWT_ChildAt(ByVal items As Object, ByVal indexNo As Long) As Object
    If items Is Nothing Then Exit Function
    On Error Resume Next
    Set CFWT_ChildAt = items.ElementAt(indexNo)
    If CFWT_ChildAt Is Nothing Then Set CFWT_ChildAt = items.item(indexNo)
    If CFWT_ChildAt Is Nothing Then Set CFWT_ChildAt = items(indexNo)
    On Error GoTo 0
End Function

Private Function CFWT_HasColumn(ByVal grid As Object, ByVal columnKey As String) As Boolean
    Dim order As Object, i As Long, n As Long, valueText As String
    If grid Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear
    Set order = grid.columnOrder
    If Err.Number <> 0 Or order Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    n = CLng(order.count)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    For i = 0 To n - 1
        valueText = CFWT_CollectionValue(order, i)
        If StrComp(Trim$(valueText), Trim$(columnKey), vbTextCompare) = 0 Then
            CFWT_HasColumn = True
            Exit Function
        End If
    Next i
End Function

Private Function CFWT_CollectionValue(ByVal items As Object, ByVal indexNo As Long) As String
    If items Is Nothing Then Exit Function
    On Error Resume Next
    Err.Clear
    CFWT_CollectionValue = CStr(items.ElementAt(indexNo) & "")
    If Err.Number <> 0 Then
        Err.Clear
        CFWT_CollectionValue = CStr(items.item(indexNo) & "")
    End If
    If Err.Number <> 0 Then
        Err.Clear
        CFWT_CollectionValue = CStr(items(indexNo) & "")
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function CFWT_Elapsed(ByVal started As Double) As Double
    If Timer >= started Then
        CFWT_Elapsed = Timer - started
    Else
        CFWT_Elapsed = 86400# - started + Timer
    End If
End Function

Private Sub CFWT_Pause(ByVal milliseconds As Long)
    If milliseconds <= 0 Then Exit Sub
    CFWT_Sleep milliseconds
    DoEvents
End Sub

