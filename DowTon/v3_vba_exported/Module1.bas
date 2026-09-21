Attribute VB_Name = "Module1"
Option Explicit

Sub Export_MB52()
    Dim SapGuiAuto As Object
    Dim SapApp As Object
    Dim SAPConnection As Object
    Dim session As Object

    Dim wsPending As Worksheet
    Dim wsMB52 As Worksheet
    Dim wsDash As Worksheet

    Dim LastRowProduct As Long
    Dim LastRowPlant As Long

    Dim SavePath As String
    Dim FileName As String
    Dim FullFileName As String

    Dim ExportWB As Workbook
    Dim ExportWS As Worksheet

    Dim OpenSuccess As Boolean
    Dim Retry As Long

    Set wsPending = ThisWorkbook.Sheets("Pending")
    Set wsMB52 = ThisWorkbook.Sheets("MB52")
    Set wsDash = ThisWorkbook.Sheets("DASHBOARD")

    LastRowProduct = wsPending.Cells(wsPending.Rows.Count, "A").End(xlUp).Row
    LastRowPlant = wsPending.Cells(wsPending.Rows.Count, "B").End(xlUp).Row

    If wsPending.AutoFilterMode Then
        On Error Resume Next
        wsPending.ShowAllData
        On Error GoTo 0
    End If
    wsPending.Rows.Hidden = False

    LastRowProduct = wsPending.Cells(wsPending.Rows.Count, "A").End(xlUp).Row
    LastRowPlant = wsPending.Cells(wsPending.Rows.Count, "B").End(xlUp).Row

    ' Ki?m tra d? li?u d?u vào trong sheet Pending
    If LastRowProduct < 2 Then
        MsgBox "Không có Material trong sheet Pending. Vui lòng ch?n Model tru?c!", vbExclamation, "Thông báo"
        Exit Sub
    End If

    If LastRowPlant < 2 Then
        MsgBox "Không có Plant trong sheet Pending!", vbExclamation, "Thông báo"
        Exit Sub
    End If

    ' Ðu?ng d?n luu file export
    SavePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\MB52 Export"
    If Dir(SavePath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir SavePath
        On Error GoTo 0
    End If

    FileName = Format(Now, "yymmdd") & "_" & Format(Now, "hh") & "h" & Format(Now, "nn") & ".xlsx"
    FullFileName = SavePath & "\" & FileName

    If Dir(FullFileName) <> "" Then
        On Error Resume Next
        Kill FullFileName
        On Error GoTo 0
    End If

    ' Ki?m tra k?t n?i SAP GUI
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Or SapGuiAuto Is Nothing Then
        MsgBox "Không th? k?t n?i SAP GUI!" & vbCrLf & _
               "Vui lòng m? SAP Logon và dang nh?p vào h? th?ng tru?c khi th?c hi?n t?i t?n MB52.", vbCritical, "L?i k?t n?i SAP GUI"
        Err.Clear
        Exit Sub
    End If

    Set SapApp = SapGuiAuto.GetScriptingEngine
    If SapApp Is Nothing Then
        MsgBox "Không th? kh?i d?ng SAP Scripting Engine. Vui lòng b?t tính nang Scripting trong SAP GUI!", vbCritical, "L?i SAP Scripting"
        Exit Sub
    End If

    If SapApp.Children.Count = 0 Then
        MsgBox "SAP GUI dang m? nhung chua có k?t n?i nào du?c thi?t l?p!" & vbCrLf & _
               "Vui lòng dang nh?p SAP tru?c khi ti?p t?c.", vbCritical, "L?i k?t n?i SAP"
        Exit Sub
    End If

    Set SAPConnection = SapApp.Children(0)
    If SAPConnection.Children.Count = 0 Then
        MsgBox "Không tìm th?y c?a s? (session) làm vi?c c?a SAP!", vbCritical, "L?i k?t n?i SAP"
        Exit Sub
    End If

    Set session = SAPConnection.Children(0)
    On Error GoTo ErrorHandler

    application.ScreenUpdating = False
    application.DisplayAlerts = False

    session.findById("wnd[0]").maximize

    ' Thoát màn hình hi?n t?i
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[12]").Press
    On Error GoTo ErrorHandler

    application.Wait Now + TimeValue("00:00:01")

    ' Vào giao d?ch MB52
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmb52"
    session.findById("wnd[0]").sendVKey 0
    application.Wait Now + TimeValue("00:00:02")

    ' Dán danh sách Product t? Pending
    wsPending.Range("A2:A" & LastRowProduct).Copy
    session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    application.CutCopyMode = False

    ' Dán danh sách Plant t? Pending
    wsPending.Range("B2:B" & LastRowPlant).Copy
    session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    application.CutCopyMode = False

    ' Xóa SLoc filter
    session.findById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    ' Xóa Batch filter
    session.findById("wnd[0]/usr/btn%_CHARG_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    ' Variant
    session.findById("wnd[0]/usr/ctxtP_VARI").Text = "/TANTK1"

    ' Execute
    session.findById("wnd[0]/tbar[1]/btn[8]").Press
    application.Wait Now + TimeValue("00:00:03")

    ' Export Excel
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").ContextMenu
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
    session.findById("wnd[1]/tbar[0]/btn[0]").Press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = SavePath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = FileName
    session.findById("wnd[1]/tbar[0]/btn[0]").Press

    ' Ch? file xu?t hi?n
    application.Wait Now + TimeValue("00:00:02")
    If Not WaitForFile(FullFileName, 120) Then
        MsgBox "Không tìm th?y file export t? SAP MB52!", vbCritical, "L?i Export SAP"
        GoTo CleanExit
    End If

    ' Xóa d? li?u cu trong sheet MB52
    wsMB52.Cells.ClearContents

    ' M? file export d? n?p d? li?u
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
        MsgBox "Không th? m? file export t? SAP MB52!", vbCritical, "L?i d?c file"
        GoTo CleanExit
    End If

    Set ExportWS = ExportWB.Sheets(1)
    With ExportWS.UsedRange
        wsMB52.Range("A1").Resize(.Rows.Count, .Columns.Count).Value = .Value
    End With
    ExportWB.Close False

    ' Sau khi n?p MB52 thành công -> C?p nh?t Dashboard và hi?n th? tr?ng thái Ðã t?i
    Call Calculate_All_Stock(False, True)
    wsDash.Range("C5").Value = Format(Now, "dd/mm/yyyy hh:nn:ss")
    wsDash.Range("C7").Value = "Ðã t?i MB52"

    MsgBox "Ðã t?i d? li?u MB52 t? SAP thành công!" & vbCrLf & _
           "- S? Material trong danh sách: " & (LastRowProduct - 1) & vbCrLf & _
           "- Th?i gian: " & Format(Now, "dd/mm/yyyy hh:nn:ss"), vbInformation, "Thành Công"

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
    application.ScreenUpdating = True
    application.DisplayAlerts = True
    MsgBox "L?i trong quá trình t?i MB52 t? SAP: " & Err.Description, vbCritical, "L?i SAP"
    Resume CleanExit
End Sub

Function WaitForFile(ByVal FileName As String, Optional ByVal TimeOutSeconds As Long = 120) As Boolean
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


