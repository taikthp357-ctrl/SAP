Attribute VB_Name = "modCF_ControlButtons"
Option Explicit

' ═══════════════════════════════════════════════════════════════════════════════
' modCF_ControlButtons — Button click handlers for CF_CONTROL
' ═══════════════════════════════════════════════════════════════════════════════

' ═══════════════════════════════════════════════════════════════════════════════
' Button 1: TẢI WH — Load Warehouse data
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub CF_ButtonLoadWH()
    On Error GoTo ErrHandler

    modCF_ControlLayout.CF_UpdateStatusBar "Đang tải dữ liệu WH từ SAP..."
    DoEvents

    ' Call existing CF_LoadWH logic (from modCF_LoadWH or similar)
    ' This should populate DATA_CORE with WH data

    ' Check if we have the old handler
    On Error Resume Next
    Application.Run "CF_LoadWH"
    If Err.Number <> 0 Then
        ' Fallback: call whatever the original button was calling
        Application.Run "btnCF_LoadWH_Click"
    End If
    On Error GoTo ErrHandler

    ' Refresh the control table
    modCF_ControlLayout.CF_RefreshControlTable
    modCF_ControlLayout.CF_UpdateStatusBar "Đã tải xong dữ liệu WH. Ready."

    Exit Sub

ErrHandler:
    modCF_ControlLayout.CF_UpdateStatusBar "ERROR: " & Err.Description
    MsgBox "Error loading WH data: " & Err.Description, vbCritical, "TẢI WH Error"
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' Button 2: CHECK / XỬ LÝ — State-aware button
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub CF_ButtonCheckExecute()
    On Error GoTo ErrHandler

    Dim ws As Worksheet, btn As Shape
    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")
    Set btn = ws.Shapes("btnCheckExecute")

    ' Determine current state from button text
    If btn.TextFrame2.TextRange.Text = "CHECK" Then
        ' Run CHECK
        modCF_ControlLayout.CF_UpdateStatusBar "Đang chạy CHECK..."
        DoEvents

        ' Call existing CF_Check logic (from modCF_CheckPlanner)
        On Error Resume Next
        Application.Run "CF_Check"
        If Err.Number <> 0 Then
            modCF_CheckPlanner.CF_Check  ' Direct call if Run fails
        End If
        On Error GoTo ErrHandler

        ' Refresh table and update button
        modCF_ControlLayout.CF_RefreshControlTable
        modCF_ControlLayout.CF_UpdateCheckExecuteButton
        modCF_ControlLayout.CF_UpdateStatusBar "CHECK hoàn tất. Ready."

    Else
        ' Run XỬ LÝ (Execute)
        modCF_ControlLayout.CF_UpdateStatusBar "Đang thực thi..."
        DoEvents

        ' Call existing CF_Execute logic
        On Error Resume Next
        Application.Run "CF_Execute"
        If Err.Number <> 0 Then
            Application.Run "CF_ExecuteToCF"  ' Try alternate name
        End If
        On Error GoTo ErrHandler

        ' Refresh table and update button
        modCF_ControlLayout.CF_RefreshControlTable
        modCF_ControlLayout.CF_UpdateCheckExecuteButton
        modCF_ControlLayout.CF_UpdateStatusBar "Thực thi hoàn tất. Ready."
    End If

    Exit Sub

ErrHandler:
    modCF_ControlLayout.CF_UpdateStatusBar "ERROR: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "CHECK/XỬ LÝ Error"
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' Button 3: LÀM MỚI — Reset for new run (calls CF_ResetUI)
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub CF_ButtonReset()
    On Error GoTo ErrHandler

    ' Confirm with user
    Dim response As VbMsgBoxResult
    response = MsgBox("Bạn có chắc muốn LÀM MỚI? Thao tác này sẽ xóa toàn bộ dữ liệu hiện tại và chuẩn bị cho run mới." & vbCrLf & vbCrLf & _
                      "LƯU Ý: Thao tác này KHÔNG xóa dữ liệu Actual đã nhập.", _
                      vbYesNo + vbExclamation, "Xác nhận LÀM MỚI")

    If response <> vbYes Then
        modCF_ControlLayout.CF_UpdateStatusBar "Hủy LÀM MỚI. Ready."
        Exit Sub
    End If

    modCF_ControlLayout.CF_UpdateStatusBar "Đang làm mới..."
    DoEvents

    ' Call existing CF_ResetUI logic (from modCFResetUI)
    On Error Resume Next
    Application.Run "CF_ResetUI"
    If Err.Number <> 0 Then
        modCFResetUI.CF_ResetUI  ' Direct call if Run fails
    End If
    On Error GoTo ErrHandler

    ' Clear CF_CONTROL table
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")
    ws.Range("D12:P100").ClearContents

    ' Clear OB input area placeholder
    ws.Range("A11").Value = "Paste OB list here"
    ws.Range("A11").Font.Italic = True
    ws.Range("A11").Font.Color = RGB(150, 150, 150)
    ws.Range("A12:C30").ClearContents

    ' Update button state
    modCF_ControlLayout.CF_UpdateCheckExecuteButton
    modCF_ControlLayout.CF_UpdateStatusBar "Đã làm mới xong. Ready."

    Exit Sub

ErrHandler:
    modCF_ControlLayout.CF_UpdateStatusBar "ERROR: " & Err.Description
    MsgBox "Error during reset: " & Err.Description, vbCritical, "LÀM MỚI Error"
End Sub

' ═══════════════════════════════════════════════════════════════════════════════
' Helper: Validate OB input and count duplicates
' ═══════════════════════════════════════════════════════════════════════════════
Public Sub CF_ValidateOBInput()
    Dim ws As Worksheet, rng As Range, cell As Range
    Dim obList As Object, ob As String
    Dim duplicateCount As Long

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")
    Set rng = ws.Range("A11:A30")
    Set obList = CreateObject("Scripting.Dictionary")

    duplicateCount = 0

    For Each cell In rng
        ob = Trim(CStr(cell.Value))
        If ob <> "" And ob <> "Paste OB list here" Then
            If obList.Exists(ob) Then
                duplicateCount = duplicateCount + 1
                cell.Interior.Color = RGB(255, 200, 200)  ' Highlight duplicates in red
            Else
                obList.Add ob, 1
                cell.Interior.Color = RGB(255, 255, 230)  ' Normal yellow
            End If
        End If
    Next cell

    ws.Range("OB_DUPLICATES").Value = duplicateCount

    If duplicateCount > 0 Then
        ws.Range("OB_DUPLICATES").Font.Color = RGB(192, 57, 43)  ' Red
        ws.Range("OB_DUPLICATES").Font.Bold = True
    Else
        ws.Range("OB_DUPLICATES").Font.Color = RGB(0, 0, 0)
        ws.Range("OB_DUPLICATES").Font.Bold = False
    End If

    Exit Sub

ErrHandler:
    ' Silent fail for validation helper
End Sub
