Option Explicit

Sub Export_MB52()

    On Error GoTo ErrorHandler

    Dim SapGuiAuto As Object
    Dim SapApp As Object
    Dim SAPConnection As Object
    Dim session As Object

    Dim wsPending As Worksheet
    Dim wsMB52 As Worksheet

    Dim LastRowProduct As Long
    Dim LastRowPlant As Long

    Dim SavePath As String
    Dim FileName As String
    Dim FullFileName As String

    Dim ExportWB As Workbook
    Dim ExportWS As Worksheet

    Dim OpenSuccess As Boolean
    Dim Retry As Long

    application.ScreenUpdating = False
    application.DisplayAlerts = False

    Set wsPending = ThisWorkbook.Sheets("Pending")
    Set wsMB52 = ThisWorkbook.Sheets("MB52")

    SavePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\MB52 Export"

    If Dir(SavePath, vbDirectory) = "" Then
    MkDir SavePath
    End If

    FileName = Format(Now, "yymmdd") & "_" & _
               Format(Now, "hh") & "h" & _
               Format(Now, "nn") & ".xlsx"

    FullFileName = SavePath & "\" & FileName

    LastRowProduct = wsPending.Cells(wsPending.Rows.Count, "A").End(xlUp).Row
    LastRowPlant = wsPending.Cells(wsPending.Rows.Count, "B").End(xlUp).Row

    '=========================================
    ' KI?M TRA D? LI?U
    '=========================================
    If LastRowProduct < 2 Then
        MsgBox "Không có Product trong sheet Pending.", vbExclamation
        GoTo CleanExit
    End If

    If LastRowPlant < 2 Then
        MsgBox "Không có Plant trong sheet Pending.", vbExclamation
        GoTo CleanExit
    End If

    '=========================================
    ' XÓA FILE CU N?U T?N T?I
    '=========================================
    If Dir(FullFileName) <> "" Then
        Kill FullFileName
    End If

    '=========================================
    ' K?T N?I SAP
    '=========================================
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapApp = SapGuiAuto.GetScriptingEngine
    Set SAPConnection = SapApp.Children(0)
    Set session = SAPConnection.Children(0)

    session.findById("wnd[0]").maximize

    '=========================================
    ' THOÁT MÀN HÌNH HI?N T?I
    '=========================================
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[12]").Press
    On Error GoTo ErrorHandler

    application.Wait Now + TimeValue("00:00:01")

    '=========================================
    ' VÀO MB52
    '=========================================
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmb52"
    session.findById("wnd[0]").sendVKey 0

    application.Wait Now + TimeValue("00:00:03")

    '=========================================
    ' PRODUCT
    '=========================================
    wsPending.Range("A2:A" & LastRowProduct).Copy

    session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    application.CutCopyMode = False

    '=========================================
    ' PLANT
    '=========================================
    wsPending.Range("B2:B" & LastRowPlant).Copy

    session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    application.CutCopyMode = False

    '=========================================
    ' SLOC
    '=========================================
    session.findById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    '=========================================
' BATCH - XÓA TOÀN B? GIÁ TR?
'=========================================
session.findById("wnd[0]/usr/btn%_CHARG_%_APP_%-VALU_PUSH").Press
session.findById("wnd[1]/tbar[0]/btn[16]").Press
session.findById("wnd[1]/tbar[0]/btn[8]").Press

    '=========================================
    ' VARIANT
    '=========================================
    session.findById("wnd[0]/usr/ctxtP_VARI").Text = "/TANTK1"

    '=========================================
    ' EXECUTE
    '=========================================
    session.findById("wnd[0]/tbar[1]/btn[8]").Press

    application.Wait Now + TimeValue("00:00:03")

    '=========================================
    ' EXPORT EXCEL
    '=========================================
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").ContextMenu
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"

    session.findById("wnd[1]/tbar[0]/btn[0]").Press

    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = SavePath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = FileName

    session.findById("wnd[1]/tbar[0]/btn[0]").Press

    '=========================================
    ' CH? FILE XU?T HI?N
    '=========================================
    application.Wait Now + TimeValue("00:00:02")

    If Not WaitForFile(FullFileName, 120) Then
        MsgBox "Không tìm th?y file export SAP.", vbCritical
        GoTo CleanExit
    End If

    '=========================================
    ' XÓA D? LI?U CU MB52
    '=========================================
    wsMB52.Cells.ClearContents

    '=========================================
    ' M? FILE EXPORT
    '=========================================
    For Retry = 1 To 30

        On Error Resume Next

        Set ExportWB = Workbooks.Open(FullFileName)

        If Err.Number = 0 Then
            OpenSuccess = True
            On Error GoTo ErrorHandler
            Exit For
        End If

        Err.Clear
        On Error GoTo ErrorHandler

        application.Wait Now + TimeValue("00:00:01")

    Next Retry

    If Not OpenSuccess Then
        MsgBox "Không m? du?c file export SAP.", vbCritical
        GoTo CleanExit
    End If

    Set ExportWS = ExportWB.Sheets(1)

    '=========================================
    ' Ð? D? LI?U VÀO SHEET MB52
    '=========================================
    With ExportWS.UsedRange
        wsMB52.Range("A1").Resize(.Rows.Count, .Columns.Count).Value = .Value
    End With

    ExportWB.Close False

    MsgBox "Ðã c?p nh?t d? li?u MB52 thành công." & vbCrLf & _
           "Product: " & LastRowProduct - 1 & vbCrLf & _
           "Plant: " & LastRowPlant - 1 & vbCrLf & _
           "File: " & FileName, vbInformation

CleanExit:

    application.ScreenUpdating = True
    application.DisplayAlerts = True

    Set ExportWS = Nothing
    Set ExportWB = Nothing

    Set session = Nothing
    Set SAPConnection = Nothing
    Set SapApp = Nothing
    Set SapGuiAuto = Nothing

    Exit Sub

ErrorHandler:

    MsgBox "L?i: " & Err.Description, vbCritical

    Resume CleanExit

End Sub

Function WaitForFile(ByVal FileName As String, _
                     Optional ByVal TimeOutSeconds As Long = 120) As Boolean

    Dim StartTime As Double

    StartTime = Timer

    Do While Timer - StartTime < TimeOutSeconds

        If Dir(FileName) <> "" Then
            WaitForFile = True
            Exit Function
        End If

        DoEvents

    Loop

    WaitForFile = False

End Function




