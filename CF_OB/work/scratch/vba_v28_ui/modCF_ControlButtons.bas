Attribute VB_Name = "modCF_ControlButtons"
Option Explicit

' ═════════════════════════════════════════════════════════════════════════════
' modCF_ControlButtons — 3-button handlers for CF_CONTROL (28-point spec)
' ═════════════════════════════════════════════════════════════════════════════
' Button contract:
'   1. TẢI WH (blue) → CF_ButtonLoadWH → calls existing CF_LoadWarehouseData
'   2. CHECK+XỬ LÝ (state-aware: green CHECK / orange-blue XỬ LÝ) → CF_ButtonCheckExecute
'   3. LÀM MỚI (gray) → CF_ButtonReset → calls CF_ResetUI
' ═════════════════════════════════════════════════════════════════════════════

' ═════════════════════════════════════════════════════════════════════════════
' BUTTON 1: TẢI WH — Load warehouse data from SAP
' ═════════════════════════════════════════════════════════════════════════════
Public Sub CF_ButtonLoadWH()
    On Error GoTo Failed

    ' Get OB from A7 (engine anchor)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")
    Dim obValue As String
    obValue = Trim(CStr(ws.Range("A7").Value2))

    If obValue = "" Then
        Call modCF_InputLayout.CF_UpdateStatusBar("ERROR: No OB entered in A7")
        MsgBox "Vui lòng nhập OB tại ô A7", vbExclamation, "TẢI WH"
        Exit Sub
    End If

    ' Call existing warehouse loader (from modCF_Entry or modCF_CheckPlanner)
    Call modCF_InputLayout.CF_UpdateStatusBar("Loading WH data for OB: " & obValue & "...")

    ' Check if multi-OB mode (A8:A40 has numeric OBs)
    Dim hasMultiOB As Boolean
    hasMultiOB = CF_HasMultipleOBs(ws)

    If hasMultiOB Then
        ' Multi-OB mode: call modCF_MultiOB.CF_RunMultiOB
        On Error Resume Next
        CF_RunMultiOB
        If Err.Number <> 0 Then
            Call modCF_InputLayout.CF_UpdateStatusBar("Multi-OB load failed: " & Err.Description)
        Else
            Call modCF_InputLayout.CF_UpdateStatusBar("Multi-OB load complete")
            Call modCF_InputWorklist.CF_RefreshInputWorklist
        End If
        On Error GoTo 0
    Else
        ' Single-OB mode: call standard CF_Check or CF_LoadWarehouseData
        On Error Resume Next
        ' Try to find the entry point (varies by CP version)
        If Not CF_TryCallLoadWH() Then
            Call modCF_InputLayout.CF_UpdateStatusBar("Load WH entry point not found")
            MsgBox "Load WH function not available (check modCF_Entry or modCF_CheckPlanner)", vbExclamation, "TẢI WH"
        Else
            Call modCF_InputLayout.CF_UpdateStatusBar("WH data loaded for OB: " & obValue)
            Call modCF_InputWorklist.CF_RefreshInputWorklist
        End If
        On Error GoTo 0
    End If

    Exit Sub

Failed:
    Call modCF_InputLayout.CF_UpdateStatusBar("TẢI WH failed: " & Err.Description)
    MsgBox "TẢI WH gặp lỗi: " & Err.Description, vbCritical, "TẢI WH"
End Sub

' Helper: Check if multi-OB list exists in A8:A40
Private Function CF_HasMultipleOBs(ws As Worksheet) As Boolean
    Dim r As Long
    Dim count As Long
    count = 0

    For r = 8 To 40
        If IsNumeric(ws.Cells(r, 1).Value) Then
            count = count + 1
            If count >= 2 Then
                CF_HasMultipleOBs = True
                Exit Function
            End If
        End If
    Next r

    CF_HasMultipleOBs = False
End Function

' Helper: Try to call the Load WH entry point (CP version-agnostic)
Private Function CF_TryCallLoadWH() As Boolean
    On Error Resume Next

    ' Try V2.8 entry point (modCF_CheckPlanner)
    CF_Check
    If Err.Number = 0 Then
        CF_TryCallLoadWH = True
        Exit Function
    End If
    Err.Clear

    ' Try V2.7 entry point (modCF_Entry)
    CF_LoadWarehouseData
    If Err.Number = 0 Then
        CF_TryCallLoadWH = True
        Exit Function
    End If
    Err.Clear

    CF_TryCallLoadWH = False
    On Error GoTo 0
End Function

' ═════════════════════════════════════════════════════════════════════════════
' BUTTON 2: CHECK+XỬ LÝ — State-aware button (CHECK or EXECUTE)
' ═════════════════════════════════════════════════════════════════════════════
Public Sub CF_ButtonCheckExecute()
    Dim wsPlan As Worksheet
    Dim hasPlan As Boolean

    On Error GoTo Failed

    ' Check PLAN sheet state (PLAN!AJ2 determines mode)
    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    If wsPlan Is Nothing Then
        MsgBox "PLAN sheet not found", vbCritical, "CHECK/XỬ LÝ"
        Exit Sub
    End If

    hasPlan = (Trim(CStr(wsPlan.Range("AJ2").Value2)) = "LOCKED")

    If hasPlan Then
        ' XỬ LÝ mode: call EXECUTE
        Call CF_ExecuteMode
    Else
        ' CHECK mode: call CHECK
        Call CF_CheckMode
    End If

    Exit Sub

Failed:
    Call modCF_InputLayout.CF_UpdateStatusBar("CHECK/XỬ LÝ failed: " & Err.Description)
    MsgBox "CHECK/XỬ LÝ gặp lỗi: " & Err.Description, vbCritical, "CHECK/XỬ LÝ"
End Sub

' CHECK mode: run plan checker
Private Sub CF_CheckMode()
    On Error Resume Next

    Call modCF_InputLayout.CF_UpdateStatusBar("Running CHECK...")

    ' Call CHECK entry point (modCF_CheckPlanner or modCF_Phase3Planner)
    CF_Check
    If Err.Number <> 0 Then
        Call modCF_InputLayout.CF_UpdateStatusBar("CHECK failed: " & Err.Description)
        MsgBox "CHECK gặp lỗi: " & Err.Description, vbExclamation, "CHECK"
    Else
        Call modCF_InputLayout.CF_UpdateStatusBar("CHECK complete — Review PLAN sheet")
        ' Refresh worklist with new Result data
        Call modCF_InputWorklist.CF_RefreshInputWorklist
        ' Update button to XỬ LÝ mode
        Call modCF_InputLayout.CF_UpdateCheckExecuteButton
    End If

    On Error GoTo 0
End Sub

' EXECUTE mode: run locked plan
Private Sub CF_ExecuteMode()
    On Error Resume Next

    Call modCF_InputLayout.CF_UpdateStatusBar("Running XỬ LÝ...")

    ' Call EXECUTE entry point (modCF_Executor or modCFX_Entry)
    ' Try V2.8 CFX executor
    CFX_Execute
    If Err.Number <> 0 Then
        Err.Clear
        ' Try V2.7 executor
        CF_Execute
        If Err.Number <> 0 Then
            Call modCF_InputLayout.CF_UpdateStatusBar("EXECUTE failed: " & Err.Description)
            MsgBox "XỬ LÝ gặp lỗi: " & Err.Description, vbExclamation, "XỬ LÝ"
            Exit Sub
        End If
    End If

    Call modCF_InputLayout.CF_UpdateStatusBar("XỬ LÝ complete — Check RUN_LOG")
    ' Refresh worklist (Status=C rows will show DONE)
    Call modCF_InputWorklist.CF_RefreshInputWorklist
    ' Update button back to CHECK mode
    Call modCF_InputLayout.CF_UpdateCheckExecuteButton

    On Error GoTo 0
End Sub

' ═════════════════════════════════════════════════════════════════════════════
' BUTTON 3: LÀM MỚI — Reset UI to clean state
' ═════════════════════════════════════════════════════════════════════════════
Public Sub CF_ButtonReset()
    On Error GoTo Failed

    ' Confirm reset
    Dim response As VbMsgBoxResult
    response = MsgBox("Làm mới sẽ xóa PLAN lock và status. Tiếp tục?", _
                      vbQuestion + vbYesNo, "LÀM MỚI")

    If response <> vbYes Then
        Call modCF_InputLayout.CF_UpdateStatusBar("Reset cancelled")
        Exit Sub
    End If

    ' Call existing CF_ResetUI from modCFResetUI
    CF_ResetUI

    ' Refresh worklist
    Call modCF_InputWorklist.CF_RefreshInputWorklist

    ' Update button state
    Call modCF_InputLayout.CF_UpdateCheckExecuteButton

    Call modCF_InputLayout.CF_UpdateStatusBar("Đã làm mới xong. Ready.")

    Exit Sub

Failed:
    Call modCF_InputLayout.CF_UpdateStatusBar("Reset failed: " & Err.Description)
    MsgBox "LÀM MỚI gặp lỗi: " & Err.Description, vbCritical, "LÀM MỚI"
End Sub
