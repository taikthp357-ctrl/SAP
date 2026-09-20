Attribute VB_Name = "modCF_AutoRun_MultiSilent"
Option Explicit

' Silent wrapper used by the tool-run read-only WH loader.  Any VBA error is
' recorded before propagation so RUN_LOG has a usable error signature.
Public Sub CF_RunAllMultiOB_AutoSilent()
    Dim oldSilent As Boolean
    Dim failureNo As Long, failureSource As String, failureDescription As String
    On Error GoTo Failed
    oldSilent = gRT_Silent
    gRT_Silent = True
    CF_RunAllMultiOB
CleanExit:
    gRT_Silent = oldSilent
    Exit Sub
Failed:
    failureNo = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    RT_Log "MULTI_WH_AUTOSILENT", "", "", "FAIL", _
           "UNHANDLED_VBA_ERROR|ERR_NO=" & CStr(failureNo) & _
           "|ERR_SOURCE=" & failureSource & _
           "|ERR_DESCRIPTION=" & failureDescription & _
           "|SAP_DIAGNOSTIC=" & gRT_LastSessionDiagnostic & _
           "|RUN_ID=" & gRT_RunID & _
           "; SAP_WRITE=0; CANCEL=0; SAVE=0; CF=0"
    gRT_Silent = oldSilent
    If failureNo <> 0 Then Err.Raise failureNo, "CF_RunAllMultiOB_AutoSilent", _
                                      "MULTI_WH_AUTOSILENT_ERROR|" & failureDescription
End Sub
