Option Explicit

Public Sub CF_SetVerifiedActualForOB1011764693()
    Const EXPECTED_OB As String = "1011764693"
    Const EXPECTED_DOC As String = "1012022347"
    Const EXPECTED_ITEM As String = "10"
    Const EXPECTED_MAT As String = "BIW32026259AA"
    Const EXPECTED_BATCH As String = "260623-01"
    Const EXPECTED_BIN As String = "APFL-ZF11-02-18"
    Const EXPECTED_QTY As Double = 10#
    Dim ws As Worksheet, r As Long, hits As Long
    Dim wt As String, detail As String

    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("WH_ACTUAL")
    For r = 5 To 504
        If CStr(ws.Cells(r, 1).Value2) = EXPECTED_OB _
           And CStr(ws.Cells(r, 2).Value2) = EXPECTED_DOC _
           And CStr(ws.Cells(r, 3).Value2) = EXPECTED_ITEM _
           And CStr(ws.Cells(r, 6).Value2) = EXPECTED_MAT _
           And CStr(ws.Cells(r, 8).Value2) = EXPECTED_BIN _
           And CStr(ws.Cells(r, 9).Value2) = EXPECTED_BATCH Then
            wt = CStr(ws.Cells(r, 5).Value2)
            If wt <> "5042661772" And wt <> "5042662681" Then
                Err.Raise vbObjectError + 9951, , "UNEXPECTED_WT=" & wt
            End If
            If CDbl(ws.Cells(r, 7).Value2) <> EXPECTED_QTY Then
                Err.Raise vbObjectError + 9952, , "SYSTEM_QTY_NOT_10 WT=" & wt
            End If
            ws.Cells(r, 11).Value2 = "ZF"
            ws.Cells(r, 12).Value2 = EXPECTED_QTY
            ws.Cells(r, 13).Value2 = "11-02-18"
            ws.Cells(r, 14).Value2 = EXPECTED_BIN
            hits = hits + 1
            detail = detail & IIf(detail = "", "", ";") & wt & "=10"
        End If
    Next r

    If hits <> 2 Then Err.Raise vbObjectError + 9953, , _
        "EXPECTED_TWO_WT_ROWS; found=" & CStr(hits)
    RT_Log "ACTUAL_VERIFIED", EXPECTED_OB, EXPECTED_DOC, "PASS", _
            "WT_QTY=10; ROWS=2; " & detail & _
            "; BIN=" & EXPECTED_BIN & "; BATCH=" & EXPECTED_BATCH & _
            "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
    Exit Sub
Failed:
    RT_Log "ACTUAL_VERIFIED", EXPECTED_OB, EXPECTED_DOC, "FAIL", _
            "Err=" & CStr(Err.Number) & ": " & Err.Description
    Err.Raise Err.Number, , Err.Description
End Sub
