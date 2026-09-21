Attribute VB_Name = "Module2"
Option Explicit

Sub Run_VL06O()
    Dim wsI As Worksheet, wsP As Worksheet, wsDO As Worksheet
    Dim LastRowProduct As Long, LastRowInput As Long, r As Long
    Dim SapGuiAuto As Object, application As Object, connection As Object, session As Object
    Dim Grid As Object
    Dim dicSLI As Object, dicDOO As Object, dicSlog As Object, dicBatchFilled As Object
    Dim mat As String, qty As Double, DOO As Double, SlogQty As Double, BatchFilledQty As Double
    Dim Batch As String, Bin As String, STO As String, DONum As String
    Dim ArrReport() As Variant, ReportIdx As Long
    Dim ArrDODetail() As Variant, DetailIdx As Long
    Dim oldLastDO As Long, shipPoint As String, shipTo As String
    Dim statusStr As String, sliVal As Double

    Set wsP = ThisWorkbook.Sheets("Pending")
    Set wsDO = ThisWorkbook.Sheets("KQ_DO")
    On Error Resume Next
    Set wsI = ThisWorkbook.Sheets("INPUT")
    On Error GoTo 0

    If wsP.AutoFilterMode Then
        On Error Resume Next
        wsP.ShowAllData
        On Error GoTo 0
    End If
    wsP.Rows.Hidden = False

    LastRowProduct = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    If LastRowProduct < 2 Then
        MsgBox "Không có Material trong sheet Pending. Vui lòng ch?n Model tru?c!", vbExclamation, "Thông báo"
        Exit Sub
    End If

    ' Shipping Point & Ship To from Pending
    shipPoint = Trim(CStr(wsP.Range("D3").Value))
    If Len(shipPoint) = 0 Then shipPoint = "3093"
    shipTo = Trim(CStr(wsP.Range("E3").Value))
    If Len(shipTo) = 0 Then shipTo = "3001"

    ' L?y SLI t? sheet INPUT (n?u có)
    Set dicSLI = CreateObject("Scripting.Dictionary")
    If Not wsI Is Nothing Then
        LastRowInput = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row
        For r = 3 To LastRowInput
            mat = Trim(CStr(wsI.Cells(r, "A").Value))
            If Len(mat) > 0 Then
                qty = Val(wsI.Cells(r, "D").Value)
                dicSLI(mat) = dicSLI(mat) + qty
            End If
        Next r
    End If

    ' K?t n?i SAP GUI
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Or SapGuiAuto Is Nothing Then
        MsgBox "Không th? k?t n?i SAP GUI!" & vbCrLf & _
               "Vui lòng m? SAP Logon và dang nh?p vào h? th?ng tru?c khi ki?m tra DO.", vbCritical, "L?i k?t n?i SAP GUI"
        Err.Clear
        Exit Sub
    End If

    Set application = SapGuiAuto.GetScriptingEngine
    If application Is Nothing Or application.Children.Count = 0 Then
        MsgBox "SAP GUI dang m? nhung chua có phiên làm vi?c nào!", vbCritical, "L?i k?t n?i SAP GUI"
        Exit Sub
    End If

    Set connection = application.Children(0)
    If connection.Children.Count = 0 Then
        MsgBox "Không tìm th?y session làm vi?c trong SAP!", vbCritical, "L?i k?t n?i SAP GUI"
        Exit Sub
    End If

    Set session = connection.Children(0)
    On Error GoTo ErrorHandler

    session.findById("wnd[0]").maximize

    ' Thoát màn hình hi?n t?i
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[12]").Press
    On Error GoTo ErrorHandler

    ' Vào giao d?ch VL06O
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl06o"
    session.findById("wnd[0]").sendVKey 0

    ' Nút For Outbound Deliveries
    session.findById("wnd[0]/usr/btnBUTTON6").Press

    ' Shipping Point
    session.findById("wnd[0]/usr/btn%_IF_VSTEL_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    wsP.Range("D3").Copy
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    application.CutCopyMode = False

    ' Xóa kho?ng th?i gian Date
    session.findById("wnd[0]/usr/btn%_IT_WADAT_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    ' Material t? sheet Pending
    session.findById("wnd[0]/usr/btn%_IT_MATNR_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    wsP.Range("A2:A" & LastRowProduct).Copy
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    application.CutCopyMode = False

    ' WBSTK = A
    session.findById("wnd[0]/usr/btn%_IT_WBSTK_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "A"
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    ' Execute
    session.findById("wnd[0]/tbar[1]/btn[8]").Press

    ' Item View
    session.findById("wnd[0]/tbar[1]/btn[18]").Press

    ' Filter Ship To
    wsP.Range("E3").Copy
    Set Grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    Grid.SetCurrentCell -1, "KUNNR"
    Grid.SelectColumn "KUNNR"
    Grid.DoubleClickCurrentCell
    session.findById("wnd[0]/tbar[1]/btn[29]").Press
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[2]/tbar[0]/btn[16]").Press
    session.findById("wnd[2]/tbar[0]/btn[16]").Press
    session.findById("wnd[2]/tbar[0]/btn[24]").Press
    session.findById("wnd[2]/tbar[0]/btn[8]").Press
    session.findById("wnd[1]/tbar[0]/btn[0]").Press
    application.CutCopyMode = False

    ' Ð?c d? li?u t? ALV Grid
    Set Grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    Set dicDOO = CreateObject("Scripting.Dictionary")
    Set dicSlog = CreateObject("Scripting.Dictionary")
    Set dicBatchFilled = CreateObject("Scripting.Dictionary")

    If Grid.RowCount > 0 Then
        ReDim ArrDODetail(1 To Grid.RowCount, 1 To 4)
        DetailIdx = 0

        For r = 0 To Grid.RowCount - 1
            If r Mod 20 = 0 Then
                Grid.FirstVisibleRow = r
                DoEvents
            End If

            mat = Trim(Grid.GetCellValue(r, "MATNR"))
            qty = Val(Grid.GetCellValue(r, "S_LGMNG"))
            DONum = Trim(Grid.GetCellValue(r, "VBELN"))
            If Len(DONum) = 0 Then DONum = Trim(Grid.GetCellValue(r, "EBELN"))
            Batch = ""
            Bin = ""
            On Error Resume Next
            Batch = Trim(Grid.GetCellValue(r, "S_CHARG"))
            Bin = Trim(Grid.GetCellValue(r, "LGORT"))
            On Error GoTo ErrorHandler

            If Batch = "" And Bin = "" Then
                dicDOO(mat) = dicDOO(mat) + qty
            ElseIf Batch = "" And Bin <> "" Then
                dicSlog(mat) = dicSlog(mat) + qty
            ElseIf Batch <> "" And Bin <> "" Then
                dicBatchFilled(mat) = dicBatchFilled(mat) + qty
            End If

            If Len(DONum) > 0 And Len(mat) > 0 Then
                DetailIdx = DetailIdx + 1
                ArrDODetail(DetailIdx, 1) = DetailIdx
                ArrDODetail(DetailIdx, 2) = DONum
                ArrDODetail(DetailIdx, 3) = qty
                ArrDODetail(DetailIdx, 4) = mat
            End If
        Next r
    End If

    ' Xu?t k?t qu? ra sheet KQ_DO
    oldLastDO = wsDO.Cells(wsDO.Rows.Count, "A").End(xlUp).Row
    If oldLastDO >= 5 Then wsDO.Range("A5:L" & oldLastDO).ClearContents
    wsDO.Range("C2").Value = Format(Now, "dd/mm/yyyy hh:nn:ss")

    ' B?ng 1: T?ng h?p DO theo danh sách Pending
    If LastRowProduct >= 2 Then
        ReDim ArrReport(1 To (LastRowProduct - 1), 1 To 7)
        For r = 2 To LastRowProduct
            mat = Trim(CStr(wsP.Cells(r, "A").Value))
            ReportIdx = r - 1
            ArrReport(ReportIdx, 1) = ReportIdx
            ArrReport(ReportIdx, 2) = mat

            sliVal = 0
            If dicSLI.Exists(mat) Then sliVal = dicSLI(mat)
            ArrReport(ReportIdx, 3) = sliVal

            DOO = 0: If dicDOO.Exists(mat) Then DOO = dicDOO(mat)
            ArrReport(ReportIdx, 4) = DOO

            SlogQty = 0: If dicSlog.Exists(mat) Then SlogQty = dicSlog(mat)
            ArrReport(ReportIdx, 5) = SlogQty

            BatchFilledQty = 0: If dicBatchFilled.Exists(mat) Then BatchFilledQty = dicBatchFilled(mat)
            ArrReport(ReportIdx, 6) = BatchFilledQty

            If sliVal > 0 Then
                If DOO >= sliVal Then
                    statusStr = "OK TO SHIP"
                Else
                    statusStr = "THI" & ChrW(7870) & "U DO (" & (sliVal - DOO) & ")"
                End If
            Else
                If (DOO + SlogQty + BatchFilledQty) > 0 Then
                    statusStr = "C" & ChrW(211) & " DO"
                Else
                    statusStr = "CH" & ChrW(431) & "A C" & ChrW(211) & " DO"
                End If
            End If
            ArrReport(ReportIdx, 7) = statusStr
        Next r

        wsDO.Range("A5").Resize(LastRowProduct - 1, 7).Value = ArrReport
    End If

    ' B?ng 2: Chi ti?t các mã DO
    If DetailIdx > 0 Then
        Dim subDODetail() As Variant
        ReDim subDODetail(1 To DetailIdx, 1 To 4)
        Dim dRow As Long, dCol As Long
        For dRow = 1 To DetailIdx
            For dCol = 1 To 4: subDODetail(dRow, dCol) = ArrDODetail(dRow, dCol): Next dCol
        Next dRow
        wsDO.Range("I5").Resize(DetailIdx, 4).Value = subDODetail
    End If

    MsgBox "Ki?m tra DO (VL06O) hoàn t?t!" & vbCrLf & _
           "- S? Material ki?m tra: " & (LastRowProduct - 1) & vbCrLf & _
           "- K?t qu? dã du?c xu?t ra sheet KQ_DO.", vbInformation, "Thành Công"

CleanExit:
    Set Grid = Nothing
    Set session = Nothing
    Set connection = Nothing
    Set application = Nothing
    Set SapGuiAuto = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "L?i trong quá trình ki?m tra DO: " & Err.Description, vbCritical, "L?i"
    Resume CleanExit
End Sub


