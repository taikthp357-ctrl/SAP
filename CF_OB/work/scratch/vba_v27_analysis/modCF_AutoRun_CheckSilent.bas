Attribute VB_Name = "modCF_AutoRun_CheckSilent"
Option Explicit

Public Sub CF_Check_AutoSilent()
    Dim oldSilent As Boolean
    On Error GoTo Failed
    oldSilent = gRT_Silent
    gRT_Silent = True
    CF_CheckAndLoadPosting
CleanExit:
    gRT_Silent = oldSilent
    Exit Sub
Failed:
    gRT_Silent = oldSilent
    Err.Raise Err.Number, , Err.Description
End Sub

Public Sub CF_Check_AutoSilent_CompileProbe()
    If False Then CF_Check_AutoSilent
End Sub
