Option Explicit
Public Sub CF_LoadWH_AutoSilent_Exact()
    On Error GoTo Failed
    gRT_Silent = True
    CF_LoadWH
CleanExit:
    gRT_Silent = False
    Exit Sub
Failed:
    gRT_Silent = False
    Err.Raise Err.Number, , Err.Description
End Sub
