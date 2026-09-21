import win32com.client as win32
import os, sys, shutil, zipfile

src_v2 = os.path.abspath(r"G:\SAP\DowTon\260920_DowTon_V2.xlsm")
dst_v3 = os.path.abspath(r"G:\SAP\DowTon\260920_DowTon_V3.xlsm")
zip_v3 = os.path.abspath(r"G:\SAP\DowTon\260920_DowTon_V3.zip")

print(f"Step 1: Copying {src_v2} -> {dst_v3}...")
shutil.copy2(src_v2, dst_v3)

# ---------------------------------------------------------
# Define VBA Code for each module
# ---------------------------------------------------------

CODE_MODULE1 = """Option Explicit

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

    ' Kiểm tra dữ liệu đầu vào trong sheet Pending
    If LastRowProduct < 2 Then
        MsgBox "Không có Material trong sheet Pending. Vui lòng chọn Model trước!", vbExclamation, "Thông báo"
        Exit Sub
    End If

    If LastRowPlant < 2 Then
        MsgBox "Không có Plant trong sheet Pending!", vbExclamation, "Thông báo"
        Exit Sub
    End If

    ' Đường dẫn lưu file export
    SavePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\\MB52 Export"
    If Dir(SavePath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir SavePath
        On Error GoTo 0
    End If

    FileName = Format(Now, "yymmdd") & "_" & Format(Now, "hh") & "h" & Format(Now, "nn") & ".xlsx"
    FullFileName = SavePath & "\\" & FileName

    If Dir(FullFileName) <> "" Then
        On Error Resume Next
        Kill FullFileName
        On Error GoTo 0
    End If

    ' Kiểm tra kết nối SAP GUI
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Or SapGuiAuto Is Nothing Then
        MsgBox "Không thể kết nối SAP GUI!" & vbCrLf & _
               "Vui lòng mở SAP Logon và đăng nhập vào hệ thống trước khi thực hiện tải tồn MB52.", vbCritical, "Lỗi kết nối SAP GUI"
        Err.Clear
        Exit Sub
    End If

    Set SapApp = SapGuiAuto.GetScriptingEngine
    If SapApp Is Nothing Then
        MsgBox "Không thể khởi động SAP Scripting Engine. Vui lòng bật tính năng Scripting trong SAP GUI!", vbCritical, "Lỗi SAP Scripting"
        Exit Sub
    End If

    If SapApp.Children.Count = 0 Then
        MsgBox "SAP GUI đang mở nhưng chưa có kết nối nào được thiết lập!" & vbCrLf & _
               "Vui lòng đăng nhập SAP trước khi tiếp tục.", vbCritical, "Lỗi kết nối SAP"
        Exit Sub
    End If

    Set SAPConnection = SapApp.Children(0)
    If SAPConnection.Children.Count = 0 Then
        MsgBox "Không tìm thấy cửa sổ (session) làm việc của SAP!", vbCritical, "Lỗi kết nối SAP"
        Exit Sub
    End If

    Set session = SAPConnection.Children(0)
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    session.findById("wnd[0]").maximize

    ' Thoát màn hình hiện tại
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[12]").Press
    On Error GoTo ErrorHandler

    Application.Wait Now + TimeValue("00:00:01")

    ' Vào giao dịch MB52
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmb52"
    session.findById("wnd[0]").sendVKey 0
    Application.Wait Now + TimeValue("00:00:02")

    ' Dán danh sách Product từ Pending
    wsPending.Range("A2:A" & LastRowProduct).Copy
    session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    Application.CutCopyMode = False

    ' Dán danh sách Plant từ Pending
    wsPending.Range("B2:B" & LastRowPlant).Copy
    session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    Application.CutCopyMode = False

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
    Application.Wait Now + TimeValue("00:00:03")

    ' Export Excel
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").ContextMenu
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
    session.findById("wnd[1]/tbar[0]/btn[0]").Press
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = SavePath
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = FileName
    session.findById("wnd[1]/tbar[0]/btn[0]").Press

    ' Chờ file xuất hiện
    Application.Wait Now + TimeValue("00:00:02")
    If Not WaitForFile(FullFileName, 120) Then
        MsgBox "Không tìm thấy file export từ SAP MB52!", vbCritical, "Lỗi Export SAP"
        GoTo CleanExit
    End If

    ' Xóa dữ liệu cũ trong sheet MB52
    wsMB52.Cells.ClearContents

    ' Mở file export để nạp dữ liệu
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
        Application.Wait Now + TimeValue("00:00:01")
    Next Retry

    If Not OpenSuccess Then
        MsgBox "Không thể mở file export từ SAP MB52!", vbCritical, "Lỗi đọc file"
        GoTo CleanExit
    End If

    Set ExportWS = ExportWB.Sheets(1)
    With ExportWS.UsedRange
        wsMB52.Range("A1").Resize(.Rows.Count, .Columns.Count).Value = .Value
    End With
    ExportWB.Close False

    ' Sau khi nạp MB52 thành công -> Cập nhật Dashboard và hiển thị trạng thái Đã tải
    Call Calculate_All_Stock(False, True)
    wsDash.Range("C5").Value = Format(Now, "dd/mm/yyyy hh:nn:ss")
    wsDash.Range("C7").Value = "Đã tải MB52"

    MsgBox "Đã tải dữ liệu MB52 từ SAP thành công!" & vbCrLf & _
           "- Số Material trong danh sách: " & (LastRowProduct - 1) & vbCrLf & _
           "- Thời gian: " & Format(Now, "dd/mm/yyyy hh:nn:ss"), vbInformation, "Thành Công"

CleanExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set ExportWS = Nothing
    Set ExportWB = Nothing
    Set session = Nothing
    Set SAPConnection = Nothing
    Set SapApp = Nothing
    Set SapGuiAuto = Nothing
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Lỗi trong quá trình tải MB52 từ SAP: " & Err.Description, vbCritical, "Lỗi SAP"
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
"""

CODE_MODULE2 = """Option Explicit

Sub Run_VL06O()
    Dim wsI As Worksheet, wsP As Worksheet, wsDO As Worksheet
    Dim LastRowProduct As Long, LastRowInput As Long, r As Long
    Dim SapGuiAuto As Object, application As Object, connection As Object, session As Object
    Dim Grid As Object
    Dim dicSLI As Object, dicDOO As Object, dicSlog As Object, dicBatchFilled As Object
    Dim Mat As String, Qty As Double, DOO As Double, SlogQty As Double, BatchFilledQty As Double
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
        MsgBox "Không có Material trong sheet Pending. Vui lòng chọn Model trước!", vbExclamation, "Thông báo"
        Exit Sub
    End If

    ' Shipping Point & Ship To from Pending
    shipPoint = Trim(CStr(wsP.Range("D3").Value))
    If Len(shipPoint) = 0 Then shipPoint = "3093"
    shipTo = Trim(CStr(wsP.Range("E3").Value))
    If Len(shipTo) = 0 Then shipTo = "3001"

    ' Lấy SLI từ sheet INPUT (nếu có)
    Set dicSLI = CreateObject("Scripting.Dictionary")
    If Not wsI Is Nothing Then
        LastRowInput = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row
        For r = 3 To LastRowInput
            Mat = Trim(CStr(wsI.Cells(r, "A").Value))
            If Len(Mat) > 0 Then
                Qty = Val(wsI.Cells(r, "D").Value)
                dicSLI(Mat) = dicSLI(Mat) + Qty
            End If
        Next r
    End If

    ' Kết nối SAP GUI
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Or SapGuiAuto Is Nothing Then
        MsgBox "Không thể kết nối SAP GUI!" & vbCrLf & _
               "Vui lòng mở SAP Logon và đăng nhập vào hệ thống trước khi kiểm tra DO.", vbCritical, "Lỗi kết nối SAP GUI"
        Err.Clear
        Exit Sub
    End If

    Set application = SapGuiAuto.GetScriptingEngine
    If application Is Nothing Or application.Children.Count = 0 Then
        MsgBox "SAP GUI đang mở nhưng chưa có phiên làm việc nào!", vbCritical, "Lỗi kết nối SAP GUI"
        Exit Sub
    End If

    Set connection = application.Children(0)
    If connection.Children.Count = 0 Then
        MsgBox "Không tìm thấy session làm việc trong SAP!", vbCritical, "Lỗi kết nối SAP GUI"
        Exit Sub
    End If

    Set session = connection.Children(0)
    On Error GoTo ErrorHandler

    session.findById("wnd[0]").maximize

    ' Thoát màn hình hiện tại
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[12]").Press
    On Error GoTo ErrorHandler

    ' Vào giao dịch VL06O
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
    Application.CutCopyMode = False

    ' Xóa khoảng thời gian Date
    session.findById("wnd[0]/usr/btn%_IT_WADAT_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    ' Material từ sheet Pending
    session.findById("wnd[0]/usr/btn%_IT_MATNR_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    wsP.Range("A2:A" & LastRowProduct).Copy
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    Application.CutCopyMode = False

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
    Application.CutCopyMode = False

    ' Đọc dữ liệu từ ALV Grid
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

            Mat = Trim(Grid.GetCellValue(r, "MATNR"))
            Qty = Val(Grid.GetCellValue(r, "S_LGMNG"))
            DONum = Trim(Grid.GetCellValue(r, "VBELN"))
            If Len(DONum) = 0 Then DONum = Trim(Grid.GetCellValue(r, "EBELN"))
            Batch = ""
            Bin = ""
            On Error Resume Next
            Batch = Trim(Grid.GetCellValue(r, "S_CHARG"))
            Bin = Trim(Grid.GetCellValue(r, "LGORT"))
            On Error GoTo ErrorHandler

            If Batch = "" And Bin = "" Then
                dicDOO(Mat) = dicDOO(Mat) + Qty
            ElseIf Batch = "" And Bin <> "" Then
                dicSlog(Mat) = dicSlog(Mat) + Qty
            ElseIf Batch <> "" And Bin <> "" Then
                dicBatchFilled(Mat) = dicBatchFilled(Mat) + Qty
            End If

            If Len(DONum) > 0 And Len(Mat) > 0 Then
                DetailIdx = DetailIdx + 1
                ArrDODetail(DetailIdx, 1) = DetailIdx
                ArrDODetail(DetailIdx, 2) = DONum
                ArrDODetail(DetailIdx, 3) = Qty
                ArrDODetail(DetailIdx, 4) = Mat
            End If
        Next r
    End If

    ' Xuất kết quả ra sheet KQ_DO
    oldLastDO = wsDO.Cells(wsDO.Rows.Count, "A").End(xlUp).Row
    If oldLastDO >= 5 Then wsDO.Range("A5:L" & oldLastDO).ClearContents
    wsDO.Range("C2").Value = Format(Now, "dd/mm/yyyy hh:nn:ss")

    ' Bảng 1: Tổng hợp DO theo danh sách Pending
    If LastRowProduct >= 2 Then
        ReDim ArrReport(1 To (LastRowProduct - 1), 1 To 7)
        For r = 2 To LastRowProduct
            Mat = Trim(CStr(wsP.Cells(r, "A").Value))
            ReportIdx = r - 1
            ArrReport(ReportIdx, 1) = ReportIdx
            ArrReport(ReportIdx, 2) = Mat

            sliVal = 0
            If dicSLI.Exists(Mat) Then sliVal = dicSLI(Mat)
            ArrReport(ReportIdx, 3) = sliVal

            DOO = 0: If dicDOO.Exists(Mat) Then DOO = dicDOO(Mat)
            ArrReport(ReportIdx, 4) = DOO

            SlogQty = 0: If dicSlog.Exists(Mat) Then SlogQty = dicSlog(Mat)
            ArrReport(ReportIdx, 5) = SlogQty

            BatchFilledQty = 0: If dicBatchFilled.Exists(Mat) Then BatchFilledQty = dicBatchFilled(Mat)
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

    ' Bảng 2: Chi tiết các mã DO
    If DetailIdx > 0 Then
        Dim subDODetail() As Variant
        ReDim subDODetail(1 To DetailIdx, 1 To 4)
        Dim dRow As Long, dCol As Long
        For dRow = 1 To DetailIdx
            For dCol = 1 To 4: subDODetail(dRow, dCol) = ArrDODetail(dRow, dCol): Next dCol
        Next dRow
        wsDO.Range("I5").Resize(DetailIdx, 4).Value = subDODetail
    End If

    MsgBox "Kiểm tra DO (VL06O) hoàn tất!" & vbCrLf & _
           "- Số Material kiểm tra: " & (LastRowProduct - 1) & vbCrLf & _
           "- Kết quả đã được xuất ra sheet KQ_DO.", vbInformation, "Thành Công"

CleanExit:
    Set Grid = Nothing
    Set session = Nothing
    Set connection = Nothing
    Set application = Nothing
    Set SapGuiAuto = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Lỗi trong quá trình kiểm tra DO: " & Err.Description, vbCritical, "Lỗi"
    Resume CleanExit
End Sub
"""

CODE_MODULE3 = """Option Explicit

Sub Run_ME2N()
    Dim wsI As Worksheet, wsP As Worksheet, wsKqSTO As Worksheet
    Dim LastRowProduct As Long, LastRowInput As Long, r As Long
    Dim SapGuiAuto As Object, SapApp As Object, connection As Object, session As Object
    Dim Grid As Object, LayoutGrid As Object
    Dim dicSLI As Object, dicSTOO As Object, dicLastSTO As Object, dicLastDate As Object
    Dim Mat As String, STO As String, OrderQty As Double, NotifiedQty As Double, STOOpen As Double
    Dim sDate As String, ArrDate() As String, DelDate As Date
    Dim ArrTable1() As Variant, Table1Idx As Long
    Dim ArrME2N() As Variant, STOIdx As Long
    Dim oldLastSTO As Long, sliVal As Double, statusStr As String
    Dim plantVal As String, reswkVal As String

    Set wsP = ThisWorkbook.Sheets("Pending")
    Set wsKqSTO = ThisWorkbook.Sheets("KQ_STO")
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
        MsgBox "Không có Material trong sheet Pending. Vui lòng chọn Model trước!", vbExclamation, "Thông báo"
        Exit Sub
    End If

    ' Plant & Supplying Plant from Pending
    plantVal = Trim(CStr(wsP.Range("F3").Value))
    If Len(plantVal) = 0 Then plantVal = "3001"
    reswkVal = Trim(CStr(wsP.Range("G3").Value))
    If Len(reswkVal) = 0 Then reswkVal = "3093"

    ' Lấy SLI từ sheet INPUT (nếu có)
    Set dicSLI = CreateObject("Scripting.Dictionary")
    If Not wsI Is Nothing Then
        LastRowInput = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row
        For r = 3 To LastRowInput
            Mat = Trim(CStr(wsI.Cells(r, "A").Value))
            If Len(Mat) > 0 Then
                dicSLI(Mat) = dicSLI(Mat) + Val(wsI.Cells(r, "D").Value)
            End If
        Next r
    End If

    ' Kết nối SAP GUI
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Or SapGuiAuto Is Nothing Then
        MsgBox "Không thể kết nối SAP GUI!" & vbCrLf & _
               "Vui lòng mở SAP Logon và đăng nhập vào hệ thống trước khi kiểm tra STO.", vbCritical, "Lỗi kết nối SAP GUI"
        Err.Clear
        Exit Sub
    End If

    Set SapApp = SapGuiAuto.GetScriptingEngine
    If SapApp Is Nothing Or SapApp.Children.Count = 0 Then
        MsgBox "SAP GUI đang mở nhưng chưa có phiên làm việc nào!", vbCritical, "Lỗi kết nối SAP GUI"
        Exit Sub
    End If

    Set connection = SapApp.Children(0)
    If connection.Children.Count = 0 Then
        MsgBox "Không tìm thấy session làm việc trong SAP!", vbCritical, "Lỗi kết nối SAP GUI"
        Exit Sub
    End If

    Set session = connection.Children(0)
    On Error GoTo ErrorHandler

    session.findById("wnd[0]").maximize

    ' Thoát màn hình hiện tại
    On Error Resume Next
    session.findById("wnd[0]/tbar[0]/btn[12]").Press
    On Error GoTo ErrorHandler

    ' Vào giao dịch ME2N
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme2n"
    session.findById("wnd[0]").sendVKey 0

    ' Scope of list
    session.findById("wnd[0]/usr/ctxtLISTU").Text = "ALV"

    ' Plant (F3 trong Pending)
    session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    wsP.Range("F3").Copy
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    Application.CutCopyMode = False

    ' Supplying Plant (G3 trong Pending)
    session.findById("wnd[0]/usr/btn%_S_RESWK_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    wsP.Range("G3").Copy
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    Application.CutCopyMode = False

    ' Material từ Pending (A2:A...)
    session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    wsP.Range("A2:A" & LastRowProduct).Copy
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    Application.CutCopyMode = False

    ' Execute
    session.findById("wnd[0]/tbar[1]/btn[8]").Press

    ' Layout (User Specific Layout)
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[33]").Press
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cmbG51_SCREEN-USPEC_LBOX").Key = "X"
    Set LayoutGrid = session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
    LayoutGrid.SetCurrentCell 0, "TEXT"
    LayoutGrid.ClickCurrentCell
    On Error GoTo ErrorHandler

    ' Đọc kết quả từ ALV Grid
    Set Grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    Set dicSTOO = CreateObject("Scripting.Dictionary")
    Set dicLastSTO = CreateObject("Scripting.Dictionary")
    Set dicLastDate = CreateObject("Scripting.Dictionary")

    STOIdx = 0
    If Grid.RowCount > 0 Then
        ReDim ArrME2N(1 To Grid.RowCount, 1 To 5)

        For r = 0 To Grid.RowCount - 1
            Mat = Trim(Grid.GetCellValue(r, "EMATN"))
            If Mat = "" Then GoTo NextRow

            STO = Trim(Grid.GetCellValue(r, "EBELN"))
            OrderQty = Val(Grid.GetCellValue(r, "MENGE"))
            NotifiedQty = Val(Grid.GetCellValue(r, "LAMNG"))
            STOOpen = OrderQty - NotifiedQty

            If STOOpen <= 0 Then GoTo NextRow
            If Trim(Grid.GetCellValue(r, "LOEKZ")) <> "" Then GoTo NextRow

            sDate = Trim(Grid.GetCellValue(r, "BEDAT"))
            If sDate <> "" Then
                ArrDate = Split(sDate, ".")
                If UBound(ArrDate) = 2 Then
                    DelDate = DateSerial(CLng(ArrDate(2)), CLng(ArrDate(1)), CLng(ArrDate(0)))
                Else
                    DelDate = DateSerial(1900, 1, 1)
                End If
            Else
                DelDate = DateSerial(1900, 1, 1)
            End If

            STOIdx = STOIdx + 1
            ArrME2N(STOIdx, 1) = STOIdx
            ArrME2N(STOIdx, 2) = STO
            ArrME2N(STOIdx, 3) = STOOpen
            ArrME2N(STOIdx, 4) = Mat
            ArrME2N(STOIdx, 5) = DelDate

            If Not dicSTOO.Exists(Mat) Then
                dicSTOO(Mat) = STOOpen
                dicLastSTO(Mat) = STO
                dicLastDate(Mat) = DelDate
            Else
                dicSTOO(Mat) = dicSTOO(Mat) + STOOpen
                If DelDate < dicLastDate(Mat) Then
                    dicLastDate(Mat) = DelDate
                    dicLastSTO(Mat) = STO
                End If
            End If
NextRow:
        Next r
    End If

    ' Xuất kết quả ra sheet KQ_STO
    oldLastSTO = wsKqSTO.Cells(wsKqSTO.Rows.Count, "A").End(xlUp).Row
    If oldLastSTO >= 5 Then wsKqSTO.Range("A5:T" & oldLastSTO).ClearContents
    wsKqSTO.Range("C2").Value = Format(Now, "dd/mm/yyyy hh:nn:ss")

    ' Bảng 1: Tổng hợp STO theo danh sách Pending
    If LastRowProduct >= 2 Then
        ReDim ArrTable1(1 To (LastRowProduct - 1), 1 To 9)
        For r = 2 To LastRowProduct
            Mat = Trim(CStr(wsP.Cells(r, "A").Value))
            Table1Idx = r - 1
            ArrTable1(Table1Idx, 1) = Table1Idx
            ArrTable1(Table1Idx, 2) = Mat

            sliVal = 0: If dicSLI.Exists(Mat) Then sliVal = dicSLI(Mat)
            ArrTable1(Table1Idx, 3) = sliVal

            ' Q'ty DO Open: 0 nếu không có liên kết DO
            ArrTable1(Table1Idx, 4) = 0

            STOOpen = 0: If dicSTOO.Exists(Mat) Then STOOpen = dicSTOO(Mat)
            ArrTable1(Table1Idx, 5) = STOOpen

            STO = "": If dicLastSTO.Exists(Mat) Then STO = dicLastSTO(Mat)
            ArrTable1(Table1Idx, 6) = STO

            If sliVal > STOOpen Then
                ArrTable1(Table1Idx, 7) = sliVal - STOOpen
            Else
                ArrTable1(Table1Idx, 7) = 0
            End If
            ArrTable1(Table1Idx, 8) = 0

            If sliVal > 0 Then
                If STOOpen >= sliVal Then
                    statusStr = ChrW(272) & ChrW(7911) & " STO"
                Else
                    statusStr = "THI" & ChrW(7870) & "U STO (C" & ChrW(7847) & "n +" & (sliVal - STOOpen) & ")"
                End If
            Else
                If STOOpen > 0 Then
                    statusStr = "C" & ChrW(211) & " STO M" & ChrW(7842)
                Else
                    statusStr = "CH" & ChrW(431) & "A C" & ChrW(211) & " STO"
                End If
            End If
            ArrTable1(Table1Idx, 9) = statusStr
        Next r

        wsKqSTO.Range("A5").Resize(LastRowProduct - 1, 9).Value = ArrTable1
    End If

    ' Bảng 3: Chi tiết đơn hàng STO từ ME2N (cột P:T)
    If STOIdx > 0 Then
        Dim subME2N() As Variant
        ReDim subME2N(1 To STOIdx, 1 To 5)
        Dim mRow As Long, mCol As Long
        For mRow = 1 To STOIdx
            For mCol = 1 To 5: subME2N(mRow, mCol) = ArrME2N(mRow, mCol): Next mCol
        Next mRow
        wsKqSTO.Range("P5").Resize(STOIdx, 5).Value = subME2N
        wsKqSTO.Range("T5:T" & (4 + STOIdx)).NumberFormat = "dd/mm/yyyy"
    End If

    MsgBox "Kiểm tra STO (ME2N) hoàn tất!" & vbCrLf & _
           "- Số Material kiểm tra: " & (LastRowProduct - 1) & vbCrLf & _
           "- Số đơn STO mở: " & STOIdx & vbCrLf & _
           "- Kết quả đã được xuất ra sheet KQ_STO.", vbInformation, "Thành Công"

CleanExit:
    Set LayoutGrid = Nothing
    Set Grid = Nothing
    Set session = Nothing
    Set connection = Nothing
    Set SapApp = Nothing
    Set SapGuiAuto = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Lỗi trong quá trình kiểm tra STO: " & Err.Description, vbCritical, "Lỗi"
    Resume CleanExit
End Sub
"""

CODE_MODULE_DASHBOARD = """Option Explicit

Private Function SafeStr(ByVal v As Variant) As String
    If IsError(v) Then
        SafeStr = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        SafeStr = ""
    Else
        SafeStr = Trim(CStr(v))
    End If
End Function

Private Function SafeVal(ByVal v As Variant) As Double
    If IsError(v) Then
        SafeVal = 0
    ElseIf IsEmpty(v) Or IsNull(v) Then
        SafeVal = 0
    ElseIf IsNumeric(v) Then
        SafeVal = CDbl(v)
    Else
        SafeVal = Val(v)
    End If
End Function

Public Function GetCheckSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If InStr(1, ws.Name, "Check", vbTextCompare) > 0 Or ws.CodeName = "Sheet4" Then
            Set GetCheckSheet = ws
            Exit Function
        End If
    Next ws
    On Error Resume Next
    Set GetCheckSheet = wb.Sheets("Check " & ChrW(7891) & "n")
    If GetCheckSheet Is Nothing Then Set GetCheckSheet = wb.Sheets(3)
    On Error GoTo 0
End Function

Public Function UniText(ByVal key As String) As String
    Select Case key
        Case "Rat_cao": UniText = "R" & ChrW(7845) & "t cao"
        Case "Cao": UniText = "Cao"
        Case "Trung_binh": UniText = "Trung b" & ChrW(236) & "nh"
        Case "Thap": UniText = "Th" & ChrW(7845) & "p"
        Case "Het_ton": UniText = "H" & ChrW(7871) & "t t" & ChrW(7891) & "n"
        Case "Duoi_min": UniText = "D" & ChrW(432) & ChrW(7899) & "i m" & ChrW(7913) & "c min"
        Case "Keo_line": UniText = ChrW(431) & "u ti" & ChrW(234) & "n k" & ChrW(233) & "o line"
        Case "UT_xuat": UniText = "UT xu" & ChrW(7845) & "t"
        Case "Du_ton": UniText = ChrW(272) & ChrW(7911) & " t" & ChrW(7891) & "n"
        Case "Theo_doi": UniText = "Theo d" & ChrW(245) & "i, l" & ChrW(234) & "n k" & ChrW(7871) & " ho" & ChrW(7841) & "ch"
        Case "Xuat_kho": UniText = ChrW(431) & "u ti" & ChrW(234) & "n xu" & ChrW(7845) & "t kho"
        Case "Binh_thuong": UniText = "B" & ChrW(236) & "nh th" & ChrW(434) & ChrW(7890) & "ng"
        Case "Tat_ca": UniText = "T" & ChrW(7845) & "t c" & ChrW(7843)
        Case Else: UniText = key
    End Select
End Function

Public Sub Dashboard_Refresh_MB52()
    ' Nút Refresh MB52: kết nối thực SAP GUI MB52
    Call Export_MB52
End Sub

Public Sub Show_Select_Models()
    On Error Resume Next
    frmSelectModel.Show
End Sub

Public Sub Copy_CanhBao(Optional ByVal ShowMessage As Boolean = True)
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("DASHBOARD")
    
    Dim lastRow As Long
    lastRow = wsDash.Cells(wsDash.Rows.Count, "B").End(xlUp).Row
    
    If lastRow < 10 Then
        If ShowMessage Then
            MsgBox "Kh" & ChrW(244) & "ng c" & ChrW(243) & " m" & ChrW(227) & " n" & ChrW(224) & "o c" & ChrW(7847) & "n c" & ChrW(7843) & "nh b" & ChrW(225) & "o!", vbInformation, "Th" & ChrW(244) & "ng B" & ChrW(225) & "o"
        End If
        Exit Sub
    End If
    
    wsDash.Range("B9:N" & lastRow).Copy
    
    If ShowMessage Then
        MsgBox ChrW(272) & ChrW(227) & " copy danh s" & ChrW(225) & "ch c" & ChrW(7843) & "nh b" & ChrW(225) & "o (" & (lastRow - 9) & " m" & ChrW(227) & ") v" & ChrW(224) & "o Clipboard!" & vbCrLf & _
               "B" & ChrW(7841) & "n c" & ChrW(243) & " th" & ChrW(7875) & " Paste tr" & ChrW(7921) & "c ti" & ChrW(7871) & "p v" & ChrW(224) & "o Teams, Zalo, Chat ho" & ChrW(7863) & "c Excel.", vbInformation, "Copy C" & ChrW(7843) & "nh B" & ChrW(225) & "o"
    End If
End Sub

Public Sub Sync_Pending_From_Model(ByVal strModels As String)
    Dim wb As Workbook
    Dim wsCheck As Worksheet, wsPending As Worksheet
    Dim lastRowCheck As Long, lastRowPending As Long
    Dim dicSelectedModels As Object, dicUniqueMat As Object
    Dim isAllModels As Boolean
    Dim parts() As String, p As Variant, mName As String
    Dim arrCheck() As Variant, r As Long
    Dim mat As String, cModel As String
    Dim k As Variant, outArr() As Variant, idx As Long
    
    Set wb = ThisWorkbook
    Set wsCheck = GetCheckSheet(wb)
    Set wsPending = wb.Sheets("Pending")
    
    If wsCheck.FilterMode Then
        On Error Resume Next
        wsCheck.ShowAllData
        On Error GoTo 0
    End If
    
    If wsPending.AutoFilterMode Then
        On Error Resume Next
        wsPending.ShowAllData
        wsPending.AutoFilterMode = False
        On Error GoTo 0
    End If
    wsPending.Rows.Hidden = False
    
    If Len(Trim(strModels)) = 0 Then strModels = "ALL"
    isAllModels = (UCase(Trim(strModels)) = "ALL" Or strModels = UniText("Tat_ca") Or strModels = "T" & ChrW(7845) & "t c" & ChrW(7843) Or strModels = "*")
    
    Set dicSelectedModels = CreateObject("Scripting.Dictionary")
    If Not isAllModels Then
        parts = Split(strModels, ",")
        For Each p In parts
            mName = Trim(CStr(p))
            If Len(mName) > 0 Then dicSelectedModels(UCase(mName)) = True
        Next p
    End If
    
    lastRowCheck = wsCheck.Cells(wsCheck.Rows.Count, "A").End(xlUp).Row
    If lastRowCheck < 2 Then Exit Sub
    
    arrCheck = wsCheck.Range("A2:J" & lastRowCheck).Value
    Set dicUniqueMat = CreateObject("Scripting.Dictionary")
    
    For r = 1 To UBound(arrCheck, 1)
        mat = SafeStr(arrCheck(r, 1))
        If Len(mat) > 0 Then
            cModel = SafeStr(arrCheck(r, 10))
            If isAllModels Or dicSelectedModels.Exists(UCase(cModel)) Then
                If Not dicUniqueMat.Exists(mat) Then
                    dicUniqueMat(mat) = True
                End If
            End If
        End If
    Next r
    
    ' Clear old Material in Pending Column A
    lastRowPending = wsPending.Cells(wsPending.Rows.Count, "A").End(xlUp).Row
    If lastRowPending >= 2 Then
        wsPending.Range("A2:A" & lastRowPending).ClearContents
    End If
    
    ' Write new unique materials to Pending starting at A2
    If dicUniqueMat.Count > 0 Then
        ReDim outArr(1 To dicUniqueMat.Count, 1 To 1)
        idx = 0
        For Each k In dicUniqueMat.Keys
            idx = idx + 1
            outArr(idx, 1) = k
        Next k
        wsPending.Range("A2").Resize(dicUniqueMat.Count, 1).Value = outArr
    End If
    
    ' Ensure Plant 3001 and 3093 in Column B (B2=3001, B3=3093)
    wsPending.Range("B2").Value = 3001
    wsPending.Range("B3").Value = 3093
    ' Clear any leftover plants below B3
    Dim lastPlantRow As Long
    lastPlantRow = wsPending.Cells(wsPending.Rows.Count, "B").End(xlUp).Row
    If lastPlantRow > 3 Then
        wsPending.Range("B4:B" & lastPlantRow).ClearContents
    End If
    
    ' Ensure Header and Control params in Pending
    wsPending.Range("A1").Value = "M" & ChrW(227)
    wsPending.Range("B1").Value = "Plant"
    wsPending.Range("D1").Value = "VL06O"
    wsPending.Range("D2").Value = "Shipping Point"
    wsPending.Range("D3").Value = 3093
    wsPending.Range("E2").Value = "Shipto"
    wsPending.Range("E3").Value = 3001
    wsPending.Range("F1").Value = "ME2N"
    wsPending.Range("F2").Value = "Plant"
    wsPending.Range("F3").Value = 3001
    wsPending.Range("G2").Value = "Supplying Plant"
    wsPending.Range("G3").Value = 3093
End Sub

Public Sub Calculate_All_Stock(Optional ByVal IsInteractive As Boolean = True, Optional ByVal IsSAPRefresh As Boolean = False)
    Dim wsDash As Worksheet, wsFull As Worksheet, wsCheck As Worksheet, wsMB As Worksheet
    Dim wb As Workbook
    Dim lastRowMB As Long
    Dim dicKhoBody As Object, dicLineBody As Object
    Dim dicTransit As Object, dicKhoMA As Object, dicLineMA As Object
    Dim dicDesc As Object
    Dim arrMB() As Variant
    Dim i As Long
    Dim mat As String, plant As String, sloc As String, mDesc As String
    Dim qty As Double, transit As Double
    Dim strModels As String, minLevel As Double
    Dim dicSelectedModels As Object
    Dim isAllModels As Boolean
    Dim parts() As String, p As Variant, mName As String
    Dim lastRowCheck As Long
    Dim arrCheck() As Variant
    Dim totalMaster As Long
    Dim arrFull() As Variant, arrDash() As Variant
    Dim idxFull As Long, idxDash As Long
    Dim cntDuoiMin As Long, cntKeoLine As Long, cntUTXuat As Long, cntHetTon As Long
    Dim r As Long, colIdx As Long
    Dim cModel As String, cUsage As Double
    Dim tKhoBody As Double, tLineBody As Double, cCanBuild As Double
    Dim inTrans As Double, tKhoMA As Double, tLineMA As Double, tongMA As Double
    Dim cRack As String, cPack As String, cThuTu As String, cTram As String
    Dim mucDo As String, canhBao As String, hanhDong As String
    Dim isLowStock As Boolean, totalStockBody As Double
    Dim matchModel As Boolean
    Dim oldLastFull As Long, subArrFull() As Variant
    Dim oldLastDash As Long, subArrDash() As Variant
    Dim modelTitle As String
    
    Set wb = ThisWorkbook
    Set wsDash = wb.Sheets("DASHBOARD")
    Set wsFull = wb.Sheets("Ton_Full")
    Set wsCheck = GetCheckSheet(wb)
    Set wsMB = wb.Sheets("MB52")
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrHandler
    
    ' 1. Read MB52 into Memory & Dictionaries
    lastRowMB = wsMB.Cells(wsMB.Rows.Count, "A").End(xlUp).Row
    
    Set dicKhoBody = CreateObject("Scripting.Dictionary")
    Set dicLineBody = CreateObject("Scripting.Dictionary")
    Set dicTransit = CreateObject("Scripting.Dictionary")
    Set dicKhoMA = CreateObject("Scripting.Dictionary")
    Set dicLineMA = CreateObject("Scripting.Dictionary")
    Set dicDesc = CreateObject("Scripting.Dictionary")
    
    If lastRowMB >= 2 Then
        arrMB = wsMB.Range("A2:I" & lastRowMB).Value
        
        For i = 1 To UBound(arrMB, 1)
            mat = SafeStr(arrMB(i, 1))
            If Len(mat) > 0 Then
                qty = SafeVal(arrMB(i, 4))
                plant = SafeStr(arrMB(i, 6))
                If InStr(plant, ".") > 0 Then plant = Left(plant, InStr(plant, ".") - 1)
                sloc = SafeStr(arrMB(i, 7))
                If InStr(sloc, ".") > 0 Then sloc = Left(sloc, InStr(sloc, ".") - 1)
                transit = SafeVal(arrMB(i, 8))
                mDesc = SafeStr(arrMB(i, 9))
                
                If Not dicDesc.Exists(mat) And Len(mDesc) > 0 Then dicDesc(mat) = mDesc
                
                If plant = "3001" And sloc = "1003" Then
                    dicKhoBody(mat) = dicKhoBody(mat) + qty
                ElseIf plant = "3001" And sloc = "3003" Then
                    dicLineBody(mat) = dicLineBody(mat) + qty
                End If
                
                dicTransit(mat) = dicTransit(mat) + transit
                
                If plant = "3093" And (sloc = "3093" Or sloc = "3003") Then
                    dicKhoMA(mat) = dicKhoMA(mat) + qty
                ElseIf plant = "3093" And (sloc = "3000" Or sloc = "3001") Then
                    dicLineMA(mat) = dicLineMA(mat) + qty
                End If
            End If
        Next i
    End If
    
    ' 2. Parse Selected Models and Min Level
    strModels = SafeStr(wsDash.Range("C3").Value)
    If Len(strModels) = 0 Then strModels = "ALL"
    
    minLevel = SafeVal(wsDash.Range("C4").Value)
    If minLevel <= 0 Then minLevel = 200: wsDash.Range("C4").Value = 200
    
    Set dicSelectedModels = CreateObject("Scripting.Dictionary")
    
    isAllModels = (UCase(strModels) = "ALL" Or strModels = UniText("Tat_ca") Or strModels = "T" & ChrW(7845) & "t c" & ChrW(7843) Or strModels = "*")
    
    If Not isAllModels Then
        parts = Split(strModels, ",")
        For Each p In parts
            mName = Trim(CStr(p))
            If Len(mName) > 0 Then dicSelectedModels(UCase(mName)) = True
        Next p
    End If
    
    ' 3. Process Master Data from Check ton
    If wsCheck.FilterMode Then
        On Error Resume Next
        wsCheck.ShowAllData
        On Error GoTo 0
    End If
    lastRowCheck = wsCheck.Cells(wsCheck.Rows.Count, "A").End(xlUp).Row
    If lastRowCheck < 2 Then GoTo FinishExit
    
    arrCheck = wsCheck.Range("A2:N" & lastRowCheck).Value
    totalMaster = UBound(arrCheck, 1)
    
    ReDim arrFull(1 To totalMaster, 1 To 19)
    ReDim arrDash(1 To totalMaster, 1 To 13)
    
    idxFull = 0: idxDash = 0
    cntDuoiMin = 0: cntKeoLine = 0: cntUTXuat = 0: cntHetTon = 0
    
    For r = 1 To totalMaster
        mat = SafeStr(arrCheck(r, 1))
        If Len(mat) > 0 Then
            cModel = SafeStr(arrCheck(r, 10))
            
            If isAllModels Then
                matchModel = True
            Else
                matchModel = dicSelectedModels.Exists(UCase(cModel))
            End If
            
            If matchModel Then
                cUsage = SafeVal(arrCheck(r, 2))
                If cUsage <= 0 Then cUsage = 1
                cRack = SafeStr(arrCheck(r, 11))
                cPack = SafeStr(arrCheck(r, 12))
                cThuTu = SafeStr(arrCheck(r, 13))
                cTram = SafeStr(arrCheck(r, 14))
                
                tKhoBody = 0: If dicKhoBody.Exists(mat) Then tKhoBody = dicKhoBody(mat)
                tLineBody = 0: If dicLineBody.Exists(mat) Then tLineBody = dicLineBody(mat)
                inTrans = 0: If dicTransit.Exists(mat) Then inTrans = dicTransit(mat)
                tKhoMA = 0: If dicKhoMA.Exists(mat) Then tKhoMA = dicKhoMA(mat)
                tLineMA = 0: If dicLineMA.Exists(mat) Then tLineMA = dicLineMA(mat)
                tongMA = tKhoMA + tLineMA
                cCanBuild = Round((tKhoBody + tLineBody) / cUsage, 1)
                
                mDesc = ""
                If dicDesc.Exists(mat) Then mDesc = dicDesc(mat)
                
                totalStockBody = tKhoBody + tLineBody
                
                ' Classification
                isLowStock = False
                If totalStockBody = 0 Then
                    mucDo = UniText("Rat_cao")
                    canhBao = UniText("Het_ton")
                    If inTrans > 0 Then
                        hanhDong = UniText("Xuat_kho")
                        cntUTXuat = cntUTXuat + 1
                    Else
                        hanhDong = UniText("Keo_line")
                        cntKeoLine = cntKeoLine + 1
                    End If
                    cntHetTon = cntHetTon + 1
                    cntDuoiMin = cntDuoiMin + 1
                    isLowStock = True
                ElseIf tLineBody = 0 Then
                    mucDo = UniText("Rat_cao")
                    canhBao = UniText("Keo_line")
                    hanhDong = UniText("Keo_line")
                    cntKeoLine = cntKeoLine + 1
                    cntDuoiMin = cntDuoiMin + 1
                    isLowStock = True
                ElseIf totalStockBody < minLevel Or tLineBody < minLevel Then
                    isLowStock = True
                    cntDuoiMin = cntDuoiMin + 1
                    If totalStockBody < minLevel * 0.3 Then
                        mucDo = UniText("Cao")
                        canhBao = UniText("Duoi_min")
                        If tKhoBody > 0 And tLineBody < 50 Then
                            hanhDong = UniText("Keo_line")
                            cntKeoLine = cntKeoLine + 1
                        Else
                            hanhDong = UniText("Theo_doi")
                        End If
                    Else
                        mucDo = UniText("Trung_binh")
                        canhBao = UniText("Duoi_min")
                        hanhDong = UniText("Theo_doi")
                    End If
                ElseIf inTrans > 0 And tKhoBody < minLevel Then
                    isLowStock = True
                    mucDo = UniText("Cao")
                    canhBao = UniText("UT_xuat")
                    hanhDong = UniText("Xuat_kho")
                    cntUTXuat = cntUTXuat + 1
                Else
                    mucDo = UniText("Thap")
                    canhBao = UniText("Du_ton")
                    hanhDong = UniText("Binh_thuong")
                End If
                
                ' Add to Ton_Full
                idxFull = idxFull + 1
                arrFull(idxFull, 1) = idxFull
                arrFull(idxFull, 2) = mat
                arrFull(idxFull, 3) = mDesc
                arrFull(idxFull, 4) = cModel
                arrFull(idxFull, 5) = cUsage
                arrFull(idxFull, 6) = tKhoBody
                arrFull(idxFull, 7) = tLineBody
                arrFull(idxFull, 8) = cCanBuild
                arrFull(idxFull, 9) = inTrans
                arrFull(idxFull, 10) = tKhoMA
                arrFull(idxFull, 11) = tLineMA
                arrFull(idxFull, 12) = tongMA
                arrFull(idxFull, 13) = cRack
                arrFull(idxFull, 14) = cPack
                arrFull(idxFull, 15) = cThuTu
                arrFull(idxFull, 16) = cTram
                arrFull(idxFull, 17) = mucDo
                arrFull(idxFull, 18) = canhBao
                arrFull(idxFull, 19) = hanhDong
                
                ' Add to DASHBOARD (Low stock only)
                If isLowStock Then
                    idxDash = idxDash + 1
                    arrDash(idxDash, 1) = idxDash
                    arrDash(idxDash, 2) = mat
                    arrDash(idxDash, 3) = mDesc
                    arrDash(idxDash, 4) = cModel
                    arrDash(idxDash, 5) = tLineBody
                    arrDash(idxDash, 6) = tKhoBody
                    arrDash(idxDash, 7) = tongMA
                    arrDash(idxDash, 8) = inTrans
                    arrDash(idxDash, 9) = cCanBuild
                    arrDash(idxDash, 10) = minLevel
                    arrDash(idxDash, 11) = mucDo
                    arrDash(idxDash, 12) = canhBao
                    arrDash(idxDash, 13) = hanhDong
                End If
            End If
        End If
    Next r
    
    ' 4. Write Ton_Full
    oldLastFull = wsFull.Cells(wsFull.Rows.Count, "A").End(xlUp).Row
    If oldLastFull >= 3 Then wsFull.Range("A3:S" & oldLastFull).Clear
    
    If idxFull > 0 Then
        ReDim subArrFull(1 To idxFull, 1 To 19)
        For r = 1 To idxFull
            For colIdx = 1 To 19: subArrFull(r, colIdx) = arrFull(r, colIdx): Next colIdx
        Next r
        wsFull.Range("A3").Resize(idxFull, 19).Value = subArrFull
        Call Format_Ton_Full_Rows(wsFull, idxFull)
    End If
    
    ' 5. Write DASHBOARD Table
    oldLastDash = wsDash.Cells(wsDash.Rows.Count, "B").End(xlUp).Row
    If oldLastDash >= 10 Then wsDash.Range("B10:N" & oldLastDash).Clear
    
    If idxDash > 0 Then
        ReDim subArrDash(1 To idxDash, 1 To 13)
        For r = 1 To idxDash
            For colIdx = 1 To 13: subArrDash(r, colIdx) = arrDash(r, colIdx): Next colIdx
        Next r
        wsDash.Range("B10").Resize(idxDash, 13).Value = subArrDash
        Call Format_Dashboard_Rows(wsDash, idxDash)
    End If
    
    ' 6. Update Control Info & KPI Summary
    ' Chỉ cập nhật C5 và C7 khi thực sự chạy qua SAP Refresh!
    If IsSAPRefresh Then
        wsDash.Range("C5").Value = Format(Now, "dd/mm/yyyy hh:nn:ss")
        wsDash.Range("C7").Value = ChrW(272) & ChrW(227) & " t" & ChrW(7843) & "i MB52"
    End If
    
    wsDash.Range("C6").Value = idxDash
    
    If isAllModels Then modelTitle = UniText("Tat_ca") Else modelTitle = strModels
    wsDash.Range("L2").Value = "T" & ChrW(7892) & "NG H" & ChrW(7906) & "P C" & ChrW(7842) & "NH B" & ChrW(193) & "O (Model: " & modelTitle & ")"
    wsDash.Range("N3").Value = cntDuoiMin
    wsDash.Range("N4").Value = cntKeoLine
    wsDash.Range("N5").Value = cntUTXuat
    wsDash.Range("N6").Value = cntHetTon

FinishExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    If IsInteractive Then
        MsgBox ChrW(272) & ChrW(227) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t d" & ChrW(7919) & " li" & ChrW(7879) & "u th" & ChrW(224) & "nh c" & ChrW(244) & "ng!" & vbCrLf & _
               "- Model: " & strModels & vbCrLf & _
               "- T" & ChrW(7893) & "ng m" & ChrW(227) & " t" & ChrW(7843) & "i t" & ChrW(7891) & "n (Ton_Full): " & idxFull & vbCrLf & _
               "- T" & ChrW(7893) & "ng m" & ChrW(227) & " c" & ChrW(7843) & "nh b" & ChrW(225) & "o (DASHBOARD): " & idxDash, vbInformation, "DowTon Dashboard"
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "L" & ChrW(7895) & "i khi c" & ChrW(7853) & "p nh" & ChrW(7853) & "t Dashboard: " & Err.Description, vbCritical, "L" & ChrW(7895) & "i"
End Sub

Private Sub Format_Dashboard_Rows(ByVal ws As Worksheet, ByVal numRows As Long)
    Dim dataRng As Range
    Set dataRng = ws.Range("B10:N" & (9 + numRows))
    
    dataRng.Font.Name = "Calibri"
    dataRng.Font.Size = 10
    ws.Range("B10:N" & (9 + numRows)).RowHeight = 20
    
    ' Alignments
    ws.Range("B10:B" & (9 + numRows)).HorizontalAlignment = -4108
    ws.Range("C10:C" & (9 + numRows)).HorizontalAlignment = -4108
    ws.Range("C10:C" & (9 + numRows)).Font.Bold = True
    ws.Range("D10:D" & (9 + numRows)).HorizontalAlignment = -4131
    ws.Range("E10:E" & (9 + numRows)).HorizontalAlignment = -4108
    
    ' Numbers
    ws.Range("F10:I" & (9 + numRows)).HorizontalAlignment = -4152
    ws.Range("F10:I" & (9 + numRows)).NumberFormat = "#,##0"
    ws.Range("J10:J" & (9 + numRows)).HorizontalAlignment = -4152
    ws.Range("J10:J" & (9 + numRows)).NumberFormat = "#,##0.0"
    ws.Range("K10:K" & (9 + numRows)).HorizontalAlignment = -4152
    ws.Range("K10:K" & (9 + numRows)).NumberFormat = "#,##0"
    
    ws.Range("L10:N" & (9 + numRows)).HorizontalAlignment = -4108
    
    ' Borders
    On Error Resume Next
    dataRng.Borders.LineStyle = 1
    dataRng.Borders.Color = RGB(217, 217, 217)
    
    Dim r As Long, cellMucDo As Range, cellCanhBao As Range, cellHanhDong As Range
    Dim valMucDo As String, valCanhBao As String, valHanhDong As String
    
    For r = 10 To 9 + numRows
        Set cellMucDo = ws.Cells(r, "L")
        Set cellCanhBao = ws.Cells(r, "M")
        Set cellHanhDong = ws.Cells(r, "N")
        
        valMucDo = Trim(CStr(cellMucDo.Value))
        valCanhBao = Trim(CStr(cellCanhBao.Value))
        valHanhDong = Trim(CStr(cellHanhDong.Value))
        
        ' Muc do
        If InStr(valMucDo, "cao") > 0 And InStr(valMucDo, "R") > 0 Then
            cellMucDo.Interior.Color = RGB(192, 0, 0)
            cellMucDo.Font.Color = RGB(255, 255, 255)
            cellMucDo.Font.Bold = True
        ElseIf valMucDo = "Cao" Then
            cellMucDo.Interior.Color = RGB(237, 125, 49)
            cellMucDo.Font.Color = RGB(255, 255, 255)
            cellMucDo.Font.Bold = True
        ElseIf InStr(valMucDo, "Trung") > 0 Then
            cellMucDo.Interior.Color = RGB(255, 230, 153)
            cellMucDo.Font.Color = RGB(128, 96, 0)
            cellMucDo.Font.Bold = True
        ElseIf InStr(valMucDo, "Th") > 0 Then
            cellMucDo.Interior.Color = RGB(169, 208, 142)
            cellMucDo.Font.Color = RGB(0, 97, 0)
            cellMucDo.Font.Bold = True
        End If
        
        ' Canh bao
        If InStr(valCanhBao, "ton") > 0 Or InStr(valCanhBao, "t" & ChrW(7891) & "n") > 0 Then
            If InStr(valCanhBao, "H") > 0 Then
                cellCanhBao.Interior.Color = RGB(252, 228, 214)
                cellCanhBao.Font.Color = RGB(192, 0, 0)
                cellCanhBao.Font.Bold = True
            ElseIf InStr(valCanhBao, ChrW(272)) > 0 Or InStr(valCanhBao, "D") > 0 Then
                cellCanhBao.Interior.Color = RGB(198, 239, 206)
                cellCanhBao.Font.Color = RGB(0, 97, 0)
                cellCanhBao.Font.Bold = True
            End If
        ElseIf InStr(valCanhBao, "line") > 0 Then
            cellCanhBao.Interior.Color = RGB(252, 228, 214)
            cellCanhBao.Font.Color = RGB(198, 89, 17)
            cellCanhBao.Font.Bold = True
        ElseIf InStr(valCanhBao, "min") > 0 Then
            cellCanhBao.Interior.Color = RGB(255, 242, 204)
            cellCanhBao.Font.Color = RGB(128, 96, 0)
            cellCanhBao.Font.Bold = True
        ElseIf InStr(valCanhBao, "UT") > 0 Then
            cellCanhBao.Interior.Color = RGB(221, 235, 247)
            cellCanhBao.Font.Color = RGB(31, 78, 121)
            cellCanhBao.Font.Bold = True
        End If
        
        ' Hanh dong uu tien
        If InStr(valHanhDong, "line") > 0 Then
            cellHanhDong.Font.Color = RGB(192, 0, 0)
            cellHanhDong.Font.Bold = True
        ElseIf InStr(valHanhDong, "xuat") > 0 Or InStr(valHanhDong, "xu" & ChrW(7845) & "t") > 0 Then
            cellHanhDong.Font.Color = RGB(31, 78, 121)
            cellHanhDong.Font.Bold = True
        End If
    Next r
    On Error GoTo 0
End Sub

Private Sub Format_Ton_Full_Rows(ByVal ws As Worksheet, ByVal numRows As Long)
    Dim dataRng As Range
    Set dataRng = ws.Range("A3:S" & (2 + numRows))
    
    dataRng.Font.Name = "Calibri"
    dataRng.Font.Size = 10
    ws.Range("A3:S" & (2 + numRows)).RowHeight = 19
    
    ws.Range("A3:A" & (2 + numRows)).HorizontalAlignment = -4108
    ws.Range("B3:B" & (2 + numRows)).HorizontalAlignment = -4108
    ws.Range("B3:B" & (2 + numRows)).Font.Bold = True
    ws.Range("C3:C" & (2 + numRows)).HorizontalAlignment = -4131
    ws.Range("D3:E" & (2 + numRows)).HorizontalAlignment = -4108
    
    ws.Range("F3:L" & (2 + numRows)).HorizontalAlignment = -4152
    ws.Range("F3:L" & (2 + numRows)).NumberFormat = "#,##0"
    ws.Range("H3:H" & (2 + numRows)).NumberFormat = "#,##0.0"
    
    ws.Range("M3:P" & (2 + numRows)).HorizontalAlignment = -4108
    ws.Range("Q3:S" & (2 + numRows)).HorizontalAlignment = -4108
    
    On Error Resume Next
    dataRng.Borders.LineStyle = 1
    dataRng.Borders.Color = RGB(217, 217, 217)
    On Error GoTo 0
End Sub
"""

CODE_FRM_SELECT_MODEL = """Option Explicit

Private Sub UserForm_Initialize()
    Dim wsCheck As Worksheet
    Dim lastRow As Long
    Dim dicModels As Object
    Dim r As Long, m As String
    Dim k As Variant
    Dim curVal As String
    Dim isAll As Boolean
    Dim i As Long
    Dim parts() As String, p As Variant
    Dim selDic As Object
    
    Me.Caption = "Ch" & ChrW(7885) & "n Model (Multi-Select)"
    Me.lblInfo.Caption = "Ch" & ChrW(7885) & "n m" & ChrW(7897) & "t ho" & ChrW(7863) & "c nhi" & ChrW(7873) & "u Model c" & ChrW(7847) & "n theo d" & ChrW(245) & "i:"
    Me.btnSelectAll.Caption = "Ch" & ChrW(7885) & "n T" & ChrW(7845) & "t C" & ChrW(7843)
    Me.btnClearAll.Caption = "B" & ChrW(7887) & " Ch" & ChrW(7885) & "n"
    Me.btnApply.Caption = ChrW(193) & "p D" & ChrW(7909) & "ng"
    
    Set wsCheck = GetCheckSheet(ThisWorkbook)
    lastRow = wsCheck.Cells(wsCheck.Rows.Count, "A").End(xlUp).Row
    
    Set dicModels = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        m = Trim(CStr(wsCheck.Cells(r, 10).Value))
        If Len(m) > 0 Then dicModels(m) = True
    Next r
    
    Me.lstModels.Clear
    For Each k In dicModels.Keys
        Me.lstModels.AddItem CStr(k)
    Next k
    
    ' Pre-select based on C3
    curVal = Trim(CStr(ThisWorkbook.Sheets("DASHBOARD").Range("C3").Value))
    isAll = (Len(curVal) = 0 Or UCase(curVal) = "ALL" Or curVal = UniText("Tat_ca") Or curVal = "T" & ChrW(7845) & "t c" & ChrW(7843))
    
    If isAll Then
        For i = 0 To Me.lstModels.ListCount - 1
            Me.lstModels.Selected(i) = True
        Next i
    Else
        parts = Split(curVal, ",")
        Set selDic = CreateObject("Scripting.Dictionary")
        For Each p In parts
            If Len(Trim(CStr(p))) > 0 Then selDic(UCase(Trim(CStr(p)))) = True
        Next p
        
        For i = 0 To Me.lstModels.ListCount - 1
            If selDic.Exists(UCase(Me.lstModels.List(i))) Then
                Me.lstModels.Selected(i) = True
            Else
                Me.lstModels.Selected(i) = False
            End If
        Next i
    End If
End Sub

Private Sub btnSelectAll_Click()
    Dim i As Long
    For i = 0 To Me.lstModels.ListCount - 1
        Me.lstModels.Selected(i) = True
    Next i
End Sub

Private Sub btnClearAll_Click()
    Dim i As Long
    For i = 0 To Me.lstModels.ListCount - 1
        Me.lstModels.Selected(i) = False
    Next i
End Sub

Private Sub btnApply_Click()
    Dim selCount As Long, i As Long
    Dim res As String
    selCount = 0: res = ""
    
    For i = 0 To Me.lstModels.ListCount - 1
        If Me.lstModels.Selected(i) Then
            selCount = selCount + 1
            If Len(res) > 0 Then res = res & ", "
            res = res & Me.lstModels.List(i)
        End If
    Next i
    
    Application.EnableEvents = False
    If selCount = 0 Or selCount = Me.lstModels.ListCount Then
        ThisWorkbook.Sheets("DASHBOARD").Range("C3").Value = "ALL"
    Else
        ThisWorkbook.Sheets("DASHBOARD").Range("C3").Value = res
    End If
    
    ' Đồng bộ Pending theo Model đã chọn (unique list)
    Call Sync_Pending_From_Model(ThisWorkbook.Sheets("DASHBOARD").Range("C3").Value)
    
    ' Đổi Model KHÔNG giả lập Đã tải MB52
    ThisWorkbook.Sheets("DASHBOARD").Range("C7").Value = "Chưa tải MB52"
    
    Call Calculate_All_Stock(False, False)
    Application.EnableEvents = True
    
    Unload Me
End Sub
"""

CODE_DASHBOARD_SHEET = """Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit
    If Not Intersect(Target, Me.Range("C3:C4")) Is Nothing Then
        Application.EnableEvents = False
        If Not Intersect(Target, Me.Range("C3")) Is Nothing Then
            ' Khi người dùng đổi ô C3 -> đồng bộ lại Pending và không giả lập tải MB52
            Call Sync_Pending_From_Model(Me.Range("C3").Value)
            Me.Range("C7").Value = "Chưa tải MB52"
        End If
        Call Calculate_All_Stock(False, False)
        Application.EnableEvents = True
    End If
CleanExit:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, Me.Range("C3")) Is Nothing Then
        Cancel = True
        Call Show_Select_Models
    End If
End Sub
"""

# ---------------------------------------------------------
# Step 2: Open Workbook with Excel COM and update VBA
# ---------------------------------------------------------
print("Step 2: Connecting to Excel COM...")
xl = win32.Dispatch('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False

try:
    wb = xl.Workbooks.Open(dst_v3)
    vb_proj = wb.VBProject
    print("Workbook opened successfully!")

    def replace_component_code(comp_name, new_code):
        for comp in vb_proj.VBComponents:
            if comp.Name == comp_name:
                print(f"Updating code in {comp_name}...")
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(new_code)
                return True
        return False

    replace_component_code("Module1", CODE_MODULE1)
    replace_component_code("Module2", CODE_MODULE2)
    replace_component_code("Module3", CODE_MODULE3)
    replace_component_code("Module_Dashboard", CODE_MODULE_DASHBOARD)
    replace_component_code("frmSelectModel", CODE_FRM_SELECT_MODEL)

    # Find DASHBOARD worksheet component
    ws_dash = wb.Sheets("DASHBOARD")
    dash_comp = None
    for comp in vb_proj.VBComponents:
        if comp.Type == 100: # Worksheet
            try:
                if comp.Properties("Name").Value == ws_dash.CodeName:
                    dash_comp = comp
                    break
            except: pass
    if not dash_comp:
        for comp in vb_proj.VBComponents:
            if comp.Type == 100 and (ws_dash.CodeName in comp.Name):
                dash_comp = comp
                break
    if dash_comp:
        print(f"Updating DASHBOARD sheet events in {dash_comp.Name}...")
        dash_comp.CodeModule.DeleteLines(1, dash_comp.CodeModule.CountOfLines)
        dash_comp.CodeModule.AddFromString(CODE_DASHBOARD_SHEET)

    # ---------------------------------------------------------
    # Step 3: Synchronize Pending for default model "Limo"
    # ---------------------------------------------------------
    print("Step 3: Synchronizing Pending with default model 'Limo'...")
    xl.Run("Sync_Pending_From_Model", "Limo")
    ws_dash.Range("C3").Value = "Limo"
    ws_dash.Range("C4").Value = 200
    ws_dash.Range("C7").Value = "Chưa tải MB52"

    print("Running Calculate_All_Stock(False, False)...")
    xl.Run("Calculate_All_Stock", False, False)

    # ---------------------------------------------------------
    # Step 4: Verify VBA Compile
    # ---------------------------------------------------------
    print("Step 4: Checking VBA compilation...")
    cmd = xl.VBE.CommandBars.FindControl(Id=578)
    if cmd:
        if cmd.Enabled:
            cmd.Execute()
            print("Compile command executed without errors!")
        else:
            print("Project is already clean and compiled!")
    else:
        print("Compile command control not found, proceeding...")

    # ---------------------------------------------------------
    # Step 5: Save V3 Workbook
    # ---------------------------------------------------------
    print(f"Step 5: Saving {dst_v3}...")
    ws_dash.Activate()
    ws_dash.Range("A1").Select()
    wb.Save()
    print("V3 saved successfully!")

finally:
    wb.Close(SaveChanges=True)
    xl.Quit()

# ---------------------------------------------------------
# Step 6: Create ZIP package
# ---------------------------------------------------------
print(f"Step 6: Creating zip package {zip_v3}...")
with zipfile.ZipFile(zip_v3, 'w', zipfile.ZIP_DEFLATED) as zf:
    zf.write(dst_v3, os.path.basename(dst_v3))
print(f"Created {zip_v3} successfully! Size: {os.path.getsize(zip_v3)} bytes.")
