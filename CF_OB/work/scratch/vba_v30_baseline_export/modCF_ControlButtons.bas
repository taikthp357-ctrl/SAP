Attribute VB_Name = "modCF_ControlButtons"
Option Explicit

' V3.0: three-button operator flow.  CHECK first writes yellow user input back
' by immutable Source Key, then runs the multi-material Posting loader + CHECK.

Public Sub CF_ButtonLoadWH()
    On Error GoTo Failed
    If CF_InputOBCount() = 0 Then
        modCF_InputLayout.CF_UpdateStatusBar "Nhap it nhat mot OB tai A7:A40."
        MsgBox "Nhap it nhat mot OB tai A7:A40.", vbExclamation, CF_UIText("LOAD_WH")
        Exit Sub
    End If
    CF_InvalidatePlanLock "RECHECK_REQUIRED_AFTER_LOAD_WH_V3"
    modCF_InputLayout.CF_UpdateStatusBar "Dang TAI WH cho danh sach OB..."
    CF_LoadWH
    CF_RefreshInputWorklist
    modCF_InputLayout.CF_UpdateCheckExecuteButton
    Exit Sub
Failed:
    modCF_InputLayout.CF_UpdateStatusBar "TAI WH loi: " & Err.Description
    MsgBox "TAI WH loi: " & Err.Description, vbCritical, CF_UIText("LOAD_WH")
End Sub

Public Sub CF_ButtonCheckExecute()
    On Error GoTo Failed
    If CF_CurrentPlanIsExecutable() Then
        modCF_InputLayout.CF_UpdateStatusBar "Dang XU LY PLAN da khoa..."
        CF_ExecutePlanToCFHandoff
    Else
        If CF_InputOBCount() = 0 Then Err.Raise vbObjectError + 2891, , _
            "Nhap it nhat mot OB tai A7:A40."
        CF_InvalidatePlanLock "RECHECK_REQUIRED_AFTER_USER_EDIT_V3"
        modCF_InputLayout.CF_UpdateStatusBar "Dang dong bo du lieu va CHECK multi-OB..."
        CF_SyncControlToCore
        CF_LoadPosting
    End If
    CF_RefreshInputWorklist
    modCF_InputLayout.CF_UpdateCheckExecuteButton
    Exit Sub
Failed:
    modCF_InputLayout.CF_UpdateStatusBar "CHECK/XU LY loi: " & Err.Description
    MsgBox "CHECK/XU LY loi: " & Err.Description, vbCritical, "CHECK/XU LY"
End Sub

Public Sub CF_ButtonReset()
    Dim response As VbMsgBoxResult
    On Error GoTo Failed
    response = MsgBox("LAM MOI se huy PLAN lock va trang thai hien tai. Tiep tuc?", _
                      vbQuestion + vbYesNo, CF_UIText("RESET"))
    If response <> vbYes Then
        modCF_InputLayout.CF_UpdateStatusBar "Da huy thao tac LAM MOI."
        Exit Sub
    End If
    CF_ResetUI
    CF_RefreshInputWorklist
    modCF_InputLayout.CF_UpdateCheckExecuteButton
    modCF_InputLayout.CF_UpdateStatusBar "Da LAM MOI. San sang."
    Exit Sub
Failed:
    modCF_InputLayout.CF_UpdateStatusBar "LAM MOI loi: " & Err.Description
    MsgBox "LAM MOI loi: " & Err.Description, vbCritical, CF_UIText("RESET")
End Sub
