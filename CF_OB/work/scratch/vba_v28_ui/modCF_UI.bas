Attribute VB_Name = "modCF_UI"
Option Explicit

Public Function RT_GetSingleInputOB(ByRef ob As String, ByRef messageText As String) As Boolean
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")
    ob = Trim$(CStr(ws.Range("A7").Value2 & ""))
    If ob = "" Then
        messageText = "INPUT_OB_EMPTY: nhap OB tai CF_CONTROL!A7."
    Else
        RT_GetSingleInputOB = True
    End If
    Exit Function
Failed:
    messageText = "CF_CONTROL_OB_READ_FAILED: " & Err.Description
End Function

Public Function RT_CellText(ByVal cell As Range) As String
    On Error Resume Next
    RT_CellText = Trim$(CStr(cell.Value2 & ""))
    Err.Clear
End Function

Public Function RT_Number(ByVal rawValue As Variant) As Double
    Dim valueText As String, normalizedText As String
    On Error GoTo Failed
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    If VarType(rawValue) <> vbString Then
        If IsNumeric(rawValue) Then
            RT_Number = CDbl(rawValue)
            Exit Function
        End If
    End If
    valueText = Trim$(CStr(rawValue & ""))
    If valueText = "" Then Exit Function
    normalizedText = RT_NormalizeSAPNumber(valueText)
    If normalizedText <> "" Then RT_Number = val(normalizedText)
    Exit Function
Failed:
    RT_Number = 0#
End Function

Private Function RT_NormalizeSAPNumber(ByVal valueText As String) As String
    Dim commaCount As Long, dotCount As Long
    Dim commaPos As Long, dotPos As Long, splitPos As Long
    Dim wholeText As String, fractionText As String
    Dim separatorText As String, signText As String

    valueText = Replace(valueText, " ", "")
    valueText = Replace(valueText, ChrW(160), "")
    valueText = Replace(valueText, "'", "")
    If valueText = "" Then Exit Function

    If Left$(valueText, 1) = "+" Or Left$(valueText, 1) = "-" Then
        signText = Left$(valueText, 1)
        valueText = Mid$(valueText, 2)
    End If
    If valueText = "" Then Exit Function

    commaCount = RT_CountChar(valueText, ",")
    dotCount = RT_CountChar(valueText, ".")
    commaPos = InStrRev(valueText, ",")
    dotPos = InStrRev(valueText, ".")

    If commaCount > 0 And dotCount > 0 Then
        If commaPos > dotPos Then
            splitPos = commaPos
            separatorText = ","
        Else
            splitPos = dotPos
            separatorText = "."
        End If
        wholeText = Left$(valueText, splitPos - 1)
        fractionText = Mid$(valueText, splitPos + 1)
        If separatorText = "," Then
            wholeText = Replace(wholeText, ".", "")
        Else
            wholeText = Replace(wholeText, ",", "")
        End If
    ElseIf commaCount > 0 Then
        If RT_LooksLikeSAPThousands(valueText, ",") Then
            wholeText = Replace(valueText, ",", "")
            fractionText = ""
        Else
            splitPos = commaPos
            wholeText = Replace(Left$(valueText, splitPos - 1), ",", "")
            fractionText = Mid$(valueText, splitPos + 1)
        End If
    ElseIf dotCount > 0 Then
        If RT_LooksLikeSAPThousands(valueText, ".") Then
            wholeText = Replace(valueText, ".", "")
            fractionText = ""
        Else
            splitPos = dotPos
            wholeText = Replace(Left$(valueText, splitPos - 1), ".", "")
            fractionText = Mid$(valueText, splitPos + 1)
        End If
    Else
        wholeText = valueText
        fractionText = ""
    End If

    If Not RT_AllDigits(wholeText) Then Exit Function
    If fractionText <> "" And Not RT_AllDigits(fractionText) Then Exit Function
    If wholeText = "" Then wholeText = "0"
    RT_NormalizeSAPNumber = signText & wholeText
    If fractionText <> "" Then RT_NormalizeSAPNumber = RT_NormalizeSAPNumber & "." & fractionText
End Function

Private Function RT_LooksLikeSAPThousands(ByVal valueText As String, ByVal separatorText As String) As Boolean
    Dim parts() As String, i As Long
    parts = Split(valueText, separatorText)
    If UBound(parts) < 1 Then Exit Function
    If Len(parts(0)) < 1 Or Len(parts(0)) > 3 Then Exit Function
    If Not RT_AllDigits(parts(0)) Then Exit Function
    For i = 1 To UBound(parts)
        If Len(parts(i)) <> 3 Then Exit Function
        If Not RT_AllDigits(parts(i)) Then Exit Function
    Next i
    RT_LooksLikeSAPThousands = True
End Function

Private Function RT_AllDigits(ByVal valueText As String) As Boolean
    Dim i As Long, ch As String
    If valueText = "" Then Exit Function
    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    RT_AllDigits = True
End Function

Private Function RT_CountChar(ByVal valueText As String, ByVal charText As String) As Long
    RT_CountChar = Len(valueText) - Len(Replace(valueText, charText, ""))
End Function

Public Function RT_StartTimer() As Double
    RT_StartTimer = Timer
End Function

Public Function RT_Elapsed(ByVal started As Double) As Double
    If Timer >= started Then
        RT_Elapsed = Timer - started
    Else
        RT_Elapsed = 86400# - started + Timer
    End If
End Function

Public Sub RT_SetStatus(ByVal areaName As String, ByVal valueText As String)
    Dim ws As Worksheet, addr As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CF_CONTROL")
    Select Case UCase$(areaName)
        Case "SAP": addr = "D6"
        Case "MON": addr = "F6"
        Case "LOAD": addr = "D10"
        Case "PLAN": addr = "F10"
        Case "STAGE": addr = "H6"
        Case "TIME": addr = "H10"
        Case Else: Exit Sub
    End Select
    ws.Range(addr).Value2 = valueText
    On Error GoTo 0
End Sub

Public Sub RT_Log(ByVal stageText As String, ByVal ob As String, ByVal doc As String, ByVal codeText As String, ByVal messageText As String)
    Dim ws As Worksheet, NextRow As Long, statusText As String, logRunID As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("RUN_LOG")
    If ws Is Nothing Then Exit Sub
    NextRow = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    If NextRow < 9 Then NextRow = 9
    If UCase$(codeText) = "PASS" Then
        statusText = "OK"
    ElseIf UCase$(codeText) = "FAIL" Then
        statusText = "ERROR"
    Else
        statusText = codeText
    End If
    If Trim$(gRT_ActiveBatchRunID) <> "" Then
        logRunID = gRT_ActiveBatchRunID
    Else
        logRunID = gRT_RunID
    End If
    ws.Cells(NextRow, 1).Value = Now
    ws.Cells(NextRow, 2).Value2 = logRunID
    ws.Cells(NextRow, 3).Value2 = stageText
    ws.Cells(NextRow, 4).Value2 = ob
    ws.Cells(NextRow, 5).Value2 = doc
    ws.Cells(NextRow, 10).Value2 = codeText
    ws.Cells(NextRow, 11).Value2 = statusText
    ws.Cells(NextRow, 12).Value2 = messageText
    On Error GoTo 0
End Sub
