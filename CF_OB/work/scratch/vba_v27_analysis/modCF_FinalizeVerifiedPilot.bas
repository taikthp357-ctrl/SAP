Attribute VB_Name = "modCF_FinalizeVerifiedPilot"
Option Explicit

Public Sub CF_FinalizeVerifiedPilot()
    Const ob As String = "1011764693"
    Const doc As String = "1012022347"
    Const wt1 As String = "5042662681"
    Const wt2 As String = "5042661772"
    Dim wsPlan As Worksheet, wsWH As Worksheet, wsCache As Worksheet
    Dim r As Long, wt As String, planDone As Long, cacheDone As Long

    On Error GoTo Failed
    Set wsWH = ThisWorkbook.Worksheets("WH_ACTUAL")
    For r = 5 To 504
        If CStr(wsWH.Cells(r, 1).Value2) = ob And CStr(wsWH.Cells(r, 2).Value2) = doc Then
            wt = CStr(wsWH.Cells(r, 5).Value2)
            If wt = wt1 Or wt = wt2 Then
                If UCase$(Trim$(CStr(wsWH.Cells(r, 10).Value2))) <> "C" Then _
                    Err.Raise vbObjectError + 9971, , "WH_STATUS_NOT_C WT=" & wt
                wsWH.Cells(r, 15).Value2 = "DONE"
                wsWH.Cells(r, 16).Value2 = "POST_VERIFY"
                wsWH.Cells(r, 18).Value2 = "CF"
                wsWH.Cells(r, 19).Value2 = "Status=C; exact Background confirmation"
                wsWH.Cells(r, 26).Value2 = "WT_DONE"
                wsWH.Cells(r, 29).Value2 = "DONE"
                wsWH.Cells(r, 30).Value2 = "WT status C: WT is already confirmed."
            End If
        End If
    Next r

    Set wsPlan = ThisWorkbook.Worksheets("PLAN")
    For r = 2 To wsPlan.Cells(wsPlan.rows.count, 1).End(xlUp).row
        If CStr(wsPlan.Cells(r, 2).Value2) = ob And CStr(wsPlan.Cells(r, 3).Value2) = doc Then
            wt = CStr(wsPlan.Cells(r, 5).Value2)
            If wt = wt1 Or wt = wt2 Then
                wsPlan.Cells(r, 20).Value2 = "DONE"
                wsPlan.Cells(r, 24).Value2 = "POST_VERIFY"
                wsPlan.Cells(r, 25).Value2 = "DONE"
                wsPlan.Cells(r, 26).Value2 = "Status=C; exact Background confirmation"
                wsPlan.Cells(r, 29).Value2 = "DONE"
                planDone = planDone + 1
            End If
        End If
    Next r
    If planDone <> 2 Then Err.Raise vbObjectError + 9972, , "PLAN_DONE_ROWS=" & CStr(planDone)

    Set wsCache = ThisWorkbook.Worksheets("OB_WH_CACHE")
    For r = 2 To wsCache.Cells(wsCache.rows.count, 1).End(xlUp).row
        If CStr(wsCache.Cells(r, 2).Value2) = ob And CStr(wsCache.Cells(r, 4).Value2) = doc Then
            wt = CStr(wsCache.Cells(r, 7).Value2)
            If wt = wt1 Or wt = wt2 Then
                wsCache.Cells(r, 12).Value2 = "C"
                wsCache.Cells(r, 22).Value2 = "WT_DONE"
                wsCache.Cells(r, 24).Value2 = "FALSE"
                wsCache.Cells(r, 25).Value2 = "DONE"
                wsCache.Cells(r, 26).Value2 = "WT status C: WT is already confirmed."
                cacheDone = cacheDone + 1
            End If
        End If
    Next r
    If cacheDone <> 2 Then Err.Raise vbObjectError + 9973, , "CACHE_DONE_ROWS=" & CStr(cacheDone)

    RT_Log "CF_FINALIZE", ob, doc, "PASS", _
            "PLAN=2; CACHE=2; WH=2; LOCAL_STATUS=C; SAP_WRITE=0"
    Exit Sub
Failed:
    RT_Log "CF_FINALIZE", ob, doc, "FAIL", "Err=" & CStr(Err.Number) & ": " & Err.Description
    Err.Raise Err.Number, , Err.Description
End Sub
