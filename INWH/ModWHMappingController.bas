Attribute VB_Name = "ModWHMappingController"
Option Explicit

' ==============================================================================
' MODULE: ModWHMappingController - SAP IN WH PORTABLE v1.0
' Giao dien Excel dan danh sach WH, gan so don va doi soat tu dong voi PDF SAP
' ==============================================================================

Private Const SHEET_MAIN As String = "IN_PHIEU_WH"

' Khoi tao hoac lam moi bang dieu khien & nut bam tren Sheet
Public Sub KhoiTaoGiaoDien(Optional ByVal isSilent As Boolean = False)
    Dim ws As Worksheet
    Application.ScreenUpdating = False

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SHEET_MAIN
    End If

    ws.Cells.Clear
    ws.Range("B2").Value = "HỆ THỐNG XỬ LÝ & ĐÓNG SỐ PHIẾU KHO SAP THEO BẢNG ĐỐI CHIẾU (PORTABLE v1.0)"
    ws.Range("B2").Font.Size = 13
    ws.Range("B2").Font.Bold = True

    ' Khu vuc cau hinh chung
    ws.Range("B4").Value = "Tiền tố ngày (vd: 06-):"
    ws.Range("C4").NumberFormat = "@"
    ws.Range("C4").Value = Format(Date, "dd") & "-"

    ws.Range("B5").Value = "File PDF gốc từ SAP:"
    ws.Range("C5").Value = ""

    ws.Range("B6").Value = "File PDF kết quả in:"
    ws.Range("C6").Value = ""

    ws.Range("B7").Value = "Trạng thái:"
    ws.Range("C7").Value = "Sẵn sàng"
    ws.Range("C7").Font.Bold = True

    ' Tieu de bang doi chieu (Dong 9)
    Dim headers As Variant
    headers = Array("STT", "Mã Warehouse Order (WH)", "Số đơn gán (Số in)", "Trang gốc", "Trang giữ", "Trang bỏ", "Trạng thái đối soát")
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(9, i + 2).Value = headers(i)
        ws.Cells(9, i + 2).Font.Bold = True
        ws.Cells(9, i + 2).Interior.Color = RGB(220, 230, 241)
    Next i

    ' Do rong cac cot
    ws.Columns("B:B").ColumnWidth = 8   ' STT
    ws.Columns("C:C").ColumnWidth = 24  ' Ma WH (Dan vao day)
    ws.Columns("D:D").ColumnWidth = 20  ' So don gan
    ws.Columns("E:E").ColumnWidth = 12  ' Trang goc
    ws.Columns("F:F").ColumnWidth = 12  ' Trang giu
    ws.Columns("G:G").ColumnWidth = 12  ' Trang bo
    ws.Columns("H:H").ColumnWidth = 45  ' Trang thai

    ' Tao cac nut bam thao tac
    Dim btnPick As Button, btnAutoLabel As Button, btnRun As Button, btnPrint As Button
    ws.Buttons.Delete

    Set btnPick = ws.Buttons.Add(ws.Range("D5").Left, ws.Range("D5").Top, 120, 22)
    btnPick.OnAction = "ChonFilePDF"
    btnPick.Caption = "Chọn File PDF..."

    Set btnAutoLabel = ws.Buttons.Add(ws.Range("D4").Left, ws.Range("D4").Top, 160, 22)
    btnAutoLabel.OnAction = "TuDongDienSoDonNhanh"
    btnAutoLabel.Caption = "Tự điền số đơn nhanh"

    Set btnRun = ws.Buttons.Add(ws.Range("B8").Left, ws.Range("B8").Top, 200, 28)
    btnRun.OnAction = "XuLyVaDongSoTheoMapping"
    btnRun.Caption = "▶ 1. XỬ LÝ & ĐỐI SOÁT PDF"

    Set btnPrint = ws.Buttons.Add(ws.Range("D8").Left, ws.Range("D8").Top, 200, 28)
    btnPrint.OnAction = "MoPDFKetQuaDeIn"
    btnPrint.Caption = "🖨 2. MỞ PDF KẾT QUẢ ĐỂ IN"

    Application.ScreenUpdating = True
    If Not isSilent Then
        MsgBox "Đã khởi tạo giao diện thành công! Hãy copy danh sách WH dán vào Cột C.", vbInformation, "Thông báo"
    End If
End Sub

' Nut [Chon File PDF...]
Public Sub ChonFilePDF()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Chọn file PDF phiếu WH xuất từ SAP"
        .Filters.Clear
        .Filters.Add "PDF Files (*.pdf)", "*.pdf"
        If .Show = -1 Then
            ThisWorkbook.Worksheets(SHEET_MAIN).Range("C5").Value = .SelectedItems(1)
        End If
    End With
End Sub

' Nut [Tu dien so don nhanh]: Lay ma WH cot C va dien cot D dang 06-01, 06-02...
Public Sub TuDongDienSoDonNhanh()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)

    Dim prefix As String
    prefix = Trim(CStr(ws.Range("C4").Value))
    If prefix = "" Then prefix = Format(Date, "dd") & "-"

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastRow < 10 Then
        MsgBox "Chưa có danh sách Mã WH ở cột C. Vui lòng dán danh sách WH trước!", vbExclamation
        Exit Sub
    End If

    Dim r As Long, seq As Long
    seq = 1
    For r = 10 To lastRow
        If Trim(CStr(ws.Cells(r, "C").Value)) <> "" Then
            ws.Cells(r, "B").Value = seq
            ws.Cells(r, "D").NumberFormat = "@"
            ws.Cells(r, "D").Value = prefix & Format(seq, "00")
            seq = seq + 1
        End If
    Next r
End Sub

' Nut [▶ 1. XU LY & DOI SOAT PDF]
Public Sub XuLyVaDongSoTheoMapping()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)

    Dim pdfInput As String
    pdfInput = Trim(CStr(ws.Range("C5").Value))
    If pdfInput = "" Or Dir(pdfInput) = "" Then
        MsgBox "Vui lòng chọn file PDF gốc xuất từ SAP!", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastRow < 10 Then
        MsgBox "Danh sách WH ở Cột C đang trống. Vui lòng dán danh sách WH cần in!", vbExclamation
        Exit Sub
    End If

    ' 1. Xuat mapping sang mapping_wh.csv (UTF-8)
    Dim baseFolder As String, mappingCsv As String, outPdf As String, outCsv As String
    baseFolder = ThisWorkbook.Path & "\"
    mappingCsv = baseFolder & "mapping_wh.csv"
    outPdf = Left(pdfInput, Len(pdfInput) - 4) & "_DA_DONG_SO.pdf"
    outCsv = Left(pdfInput, Len(pdfInput) - 4) & "_DOI_CHIEU.csv"

    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(mappingCsv, True, True)
    ts.WriteLine "WH_NO,LABEL"

    Dim r As Long, whVal As String, labelVal As String
    For r = 10 To lastRow
        whVal = Trim(CStr(ws.Cells(r, "C").Value))
        labelVal = Trim(CStr(ws.Cells(r, "D").Value))
        If whVal <> "" Then
            ts.WriteLine whVal & "," & labelVal
        End If
    Next r
    ts.Close

    ' 2. Xac dinh Engine portable hoac python
    Dim engineExe As String, enginePy As String, cmd As String
    engineExe = baseFolder & "wh_engine.exe"
    enginePy = baseFolder & "wh_engine.py"

    If Dir(engineExe) <> "" Then
        cmd = """" & engineExe & """ --input """ & pdfInput & """ --mapping """ & mappingCsv & """ --out-pdf """ & outPdf & """ --out-csv """ & outCsv & """"
    ElseIf Dir(enginePy) <> "" Then
        cmd = "python """ & enginePy & """ --input """ & pdfInput & """ --mapping """ & mappingCsv & """ --out-pdf """ & outPdf & """ --out-csv """ & outCsv & """"
    Else
        MsgBox "Không tìm thấy wh_engine.exe hoặc wh_engine.py trong thư mục!", vbCritical
        Exit Sub
    End If

    ws.Range("C7").Value = "Đang xử lý phân tích và đóng số PDF..."
    DoEvents

    Dim wsh As Object, exitCode As Long
    Set wsh = CreateObject("WScript.Shell")
    exitCode = wsh.Run(cmd, 0, True)

    If exitCode <> 0 Or Dir(outCsv) = "" Then
        ws.Range("C7").Value = "Lỗi khi xử lý (Mã lỗi: " & exitCode & ")"
        MsgBox "Có lỗi xảy ra khi xử lý file PDF. Vui lòng kiểm tra lại!", vbCritical
        Exit Sub
    End If

    ' 3. Doc file CSV doi chieu va cap nhat vao cac cot
    CapNhatKetQuaVaoSheet outCsv

    ws.Range("C6").Value = outPdf
    ws.Range("C7").Value = "Hoàn tất đối soát & đóng số!"
    MsgBox "Đã xử lý xong! Hãy kiểm tra bảng đối soát ở các cột bên phải.", vbInformation, "Thành công"
End Sub

' Cap nhat trang thai tu CSV vao cot E, F, G, H
Private Sub CapNhatKetQuaVaoSheet(ByVal csvPath As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)

    Dim stream As Object, rawText As String
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile csvPath
    rawText = stream.ReadText(-1)
    stream.Close

    rawText = Replace(rawText, vbCrLf, vbLf)
    Dim lines() As String
    lines = Split(rawText, vbLf)

    Dim maxR As Long
    maxR = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If maxR >= 10 Then
        ws.Range("B10:H" & (maxR + 50)).ClearContents
    End If

    Dim i As Long, curRow As Long
    curRow = 10
    For i = 1 To UBound(lines)
        If Trim(lines(i)) <> "" Then
            Dim cols() As String
            cols = Split(lines(i), ",")
            If UBound(cols) >= 6 Then
                ws.Cells(curRow, "B").Value = cols(0) ' STT
                ws.Cells(curRow, "C").NumberFormat = "@"
                ws.Cells(curRow, "C").Value = cols(2) ' WH
                ws.Cells(curRow, "D").NumberFormat = "@"
                ws.Cells(curRow, "D").Value = cols(1) ' So don
                ws.Cells(curRow, "E").Value = cols(3) ' Trang goc
                ws.Cells(curRow, "F").Value = cols(4) ' Trang giu
                ws.Cells(curRow, "G").Value = cols(5) ' Trang bo
                ws.Cells(curRow, "H").Value = cols(6) ' Trang thai

                If InStr(cols(6), "THIẾU") > 0 Then
                    ws.Range(ws.Cells(curRow, "B"), ws.Cells(curRow, "H")).Interior.Color = RGB(255, 204, 204)
                ElseIf InStr(cols(6), "THỪA") > 0 Then
                    ws.Range(ws.Cells(curRow, "B"), ws.Cells(curRow, "H")).Interior.Color = RGB(255, 255, 204)
                Else
                    ws.Range(ws.Cells(curRow, "B"), ws.Cells(curRow, "H")).Interior.Color = xlNone
                End If

                curRow = curRow + 1
            End If
        End If
    Next i

    ws.Range("B10:H" & (curRow - 1)).Borders.LineStyle = xlContinuous
End Sub

' Nut [🖨 2. MO PDF KET QUA DE IN]
Public Sub MoPDFKetQuaDeIn()
    Dim outPdf As String
    outPdf = Trim(CStr(ThisWorkbook.Worksheets(SHEET_MAIN).Range("C6").Value))

    If outPdf = "" Or Dir(outPdf) = "" Then
        MsgBox "Chưa có file PDF kết quả. Vui lòng bấm [1. XỬ LÝ & ĐỐI SOÁT PDF] trước!", vbExclamation
        Exit Sub
    End If

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run """" & outPdf & """", 1, False
End Sub
