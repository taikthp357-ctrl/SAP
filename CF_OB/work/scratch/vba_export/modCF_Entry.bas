Option Explicit

Public Sub CF_LoadWH()
    CF_RunAllMultiOB
End Sub

Public Sub CF_LoadPosting()
    MPOST_LoadPosting
    If gMPOST_LastRunOK Then CF_Check
End Sub

' Normal post-input flow: require a current WH snapshot for the exact input OB
' before reading Posting or building PLAN.  This blocks a newly entered OB from
' inheriting the previous OB's WH rows.
Public Sub CF_CheckAndLoadPosting()
    Dim oldSilent As Boolean, finalResult As String
    Dim inputOB As String, snapshotMessage As String
    On Error GoTo Failed

    oldSilent = gRT_Silent
    If Not CF_RequireInputOBSnapshot(inputOB, snapshotMessage) Then
        finalResult = "STOP|" & snapshotMessage
        RT_Log "CHECK_POSTING", inputOB, "", "STOP", finalResult & _
               "; SAP_WRITE=0; CANCEL=0; SAVE=0; CREATE=0; CF=0"
        If Not oldSilent Then MsgBox finalResult, vbExclamation, "CHECK + POSTING"
        Exit Sub
    End If

    gRT_Silent = True
    MPOST_LoadPosting
    If gMPOST_LastRunOK Then
        CF_Check
        finalResult = gCP_LastCheckResult
    Else
        finalResult = "STOP|TAI_POSTING_FAILED|" & gMPOST_LastRunMessage
    End If

    gRT_Silent = oldSilent
    If Left$(finalResult, 5) = "STOP|" Or Left$(finalResult, 5) = "FAIL|" Then
        RT_Log "CHECK_POSTING", inputOB, "", "STOP", finalResult & _
               "; SAP_WRITE=0; CANCEL=0; SAVE=0; CREATE=0; CF=0"
    Else
        RT_Log "CHECK_POSTING", inputOB, "", "PASS", finalResult & _
               "; SAP_WRITE=0; CANCEL=0; SAVE=0; CREATE=0; CF=0"
    End If
    If Not oldSilent Then
        If Left$(finalResult, 5) = "STOP|" Or Left$(finalResult, 5) = "FAIL|" Then
            MsgBox finalResult, vbExclamation, "CHECK + POSTING"
        Else
            MsgBox finalResult & vbCrLf & "SAP_WRITE=0; CF=0", vbInformation, "CHECK + POSTING"
        End If
    End If
    Exit Sub

Failed:
    gRT_Silent = oldSilent
    RT_Log "CHECK_POSTING", inputOB, "", "FAIL", "Err " & CStr(Err.Number) & ": " & Err.Description & _
           "; SAP_WRITE=0; CANCEL=0; SAVE=0; CREATE=0; CF=0"
    If Not oldSilent Then MsgBox "CHECK + POSTING failed" & vbCrLf & Err.Description, vbExclamation, "CHECK + POSTING"
End Sub

' Requires the last read-only TAI WH snapshot to contain the selected OB.
' All paths that return False end before any SAP call.
Private Function CF_RequireInputOBSnapshot(ByRef inputOB As String, _
                                            ByRef messageText As String) As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim runID As String, gateText As String

    messageText = ""
    If Not RT_GetSingleInputOB(inputOB, messageText) Then Exit Function
    Set ws = ThisWorkbook.Worksheets("WH_ACTUAL")
    runID = Trim$(CStr(ws.Range("W2").Value2 & ""))
    gateText = Trim$(CStr(ws.Range("X2").Value2 & ""))
    If runID = "" Or UCase$(Left$(gateText, 5)) = "STOP_" Then
        messageText = "STOP_WH_SNAPSHOT_UNPROVEN_FOR_INPUT_OB|OB=" & inputOB & _
                      "|Run 1. TAI WH for this OB first; no SAP read/write was run."
        Exit Function
    End If

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow > 504 Then lastRow = 504
    For r = 5 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value2 & "")), inputOB, vbTextCompare) = 0 Then
            CF_RequireInputOBSnapshot = True
            Exit Function
        End If
    Next r
    messageText = "STOP_WH_SNAPSHOT_MISSING_FOR_INPUT_OB|OB=" & inputOB & _
                  "|Current WH snapshot belongs to another OB. Run 1. TAI WH first; no SAP read/write was run."
End Function

Public Sub CF_Phase2()
    MPOST_LoadPosting
End Sub

Public Sub CF_RunAll()
    CF_RunAllMultiOB
End Sub

Public Sub CF_RunAllSingle()
    Dim ob As String, msg As String
    If Not RT_GetSingleInputOB(ob, msg) Then
        RT_Log "RUN_ALL_SINGLE", ob, "", "FAIL", msg
        MsgBox msg, vbExclamation, "SAP CF OB"
        Exit Sub
    End If
    RT_LoadMON
    CF_LoadPostingFromWH
End Sub

Public Sub CF_CheckSAP()
    RT_Preflight
End Sub

Public Sub CF_OpenLog()
    RT_OpenLog
End Sub
