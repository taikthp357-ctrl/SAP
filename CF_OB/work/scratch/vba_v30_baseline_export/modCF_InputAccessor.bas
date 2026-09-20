Attribute VB_Name = "modCF_InputAccessor"
Option Explicit

' One source of truth for the operator-input sheet.
Public Function CF_GetInputSheet() As Worksheet
    On Error GoTo Failed
    Set CF_GetInputSheet = ThisWorkbook.Worksheets("CF_CONTROL")
    Exit Function
Failed:
    Err.Raise vbObjectError + 2840, "CF_GetInputSheet", _
              "CF_CONTROL input sheet is unavailable: " & Err.Description
End Function

' Unicode-safe visible UI labels. ChrW avoids VBA code-page corruption.
Public Function CF_UIText(ByVal keyName As String) As String
    Select Case UCase$(Trim$(keyName))
        Case "LOAD_WH"
            CF_UIText = "T" & ChrW(&H1EA2) & "I WH"
        Case "EXECUTE"
            CF_UIText = "X" & ChrW(&H1EEC) & " L" & ChrW(&HDD)
        Case "RESET"
            CF_UIText = "L" & ChrW(&HC0) & "M M" & ChrW(&H1EDA) & "I"
        Case "INPUT_OB"
            CF_UIText = "NH" & ChrW(&H1EAC) & "P OB"
        Case "SYSTEM"
            CF_UIText = "H" & ChrW(&H1EC6) & " TH" & ChrW(&H1ED0) & "NG"
        Case "PLAN"
            CF_UIText = "K" & ChrW(&H1EBE) & " HO" & ChrW(&H1EA0) & "CH"
        Case "ACTUAL"
            CF_UIText = "TH" & ChrW(&H1EF0) & "C T" & ChrW(&H1EBE)
        Case "RESULT"
            CF_UIText = "K" & ChrW(&H1EBF) & "t qu" & ChrW(&H1EA3)
        Case "DETAIL"
            CF_UIText = "Chi ti" & ChrW(&H1EBF) & "t"
        Case Else
            CF_UIText = keyName
    End Select
End Function

' Static, SAP-free compile/accessor probe.
Public Function CF_Offline_CompileProbe() As String
    Dim ws As Worksheet
    Set ws = CF_GetInputSheet()
    If ws Is Nothing Then
        CF_Offline_CompileProbe = "FAIL|CF_CONTROL_NOT_FOUND"
    Else
        CF_Offline_CompileProbe = "PASS|INPUT_SHEET=" & ws.Name
    End If
End Function

