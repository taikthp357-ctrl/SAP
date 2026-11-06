Attribute VB_Name = "ModWHMappingController"
Option Explicit

' ==============================================================================
' MODULE: ModWHMappingController - SAP IN WH PORTABLE v1.1
' Ho tro 2 che do: Tu dong quet PDF 100% HOAC Doi soat theo danh sach WH dan vao
' ==============================================================================

Private Const SHEET_MAIN As String = "IN_PHIEU_WH"
Private Const DATA_START_ROW As Long = 12

' Nut [Chon File PDF...]
Public Sub ChonFilePDF()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Chon file PDF phieu WH xuat tu SAP"
        .Filters.Clear
        .Filters.Add "PDF Files (*.pdf)", "*.pdf"
        .AllowMultiSelect = False
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
    If lastRow < DATA_START_ROW Then
        MsgBox "Cot C dang trong. Ban co the de trong de tool TU DONG QUET truc tiep tu PDF, hoac dan danh sach WH vao day de gan so!", vbInformation, "Thong bao"
        Exit Sub
    End If

    Dim r As Long, seq As Long
    seq = 1
    For r = DATA_START_ROW To lastRow
        Dim whVal As String
        whVal = Trim(Replace(Replace(CStr(ws.Cells(r, "C").Value), Chr(160), " "), vbTab, " "))
        If whVal <> "" And Len(whVal) >= 4 And LCase(whVal) <> "w" Then
            ws.Cells(r, "B").Value = seq
            ws.Cells(r, "D").NumberFormat = "@"
            ws.Cells(r, "D").Value = prefix & Format(seq, "00")
            seq = seq + 1
        End If
    Next r
End Sub

' Nut [1. XU LY & DOI SOAT PDF]
Public Sub XuLyVaDongSoTheoMapping()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)

    Dim pdfInput As String
    pdfInput = Trim(CStr(ws.Range("C5").Value))
    If pdfInput = "" Or Dir(pdfInput) = "" Then
        MsgBox "Vui long chon file PDF goc xuat tu SAP!", vbExclamation, "Thong bao"
        Exit Sub
    End If

    Dim prefix As String
    prefix = Trim(CStr(ws.Range("C4").Value))
    If prefix = "" Then prefix = Format(Date, "dd") & "-"

    ' Kiem tra xem nguoi dung co dan danh sach WH hop le khong
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    Dim baseFolder As String, mappingCsv As String, outPdf As String, outCsv As String
    baseFolder = ThisWorkbook.Path & "\"
    mappingCsv = baseFolder & "mapping_wh.csv"
    outPdf = Left(pdfInput, Len(pdfInput) - 4) & "_DA_DONG_SO.pdf"
    outCsv = Left(pdfInput, Len(pdfInput) - 4) & "_DOI_CHIEU.csv"

    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim validWHCount As Long
    validWHCount = 0

    If lastRow >= DATA_START_ROW Then
        Set ts = fso.CreateTextFile(mappingCsv, True, True)
        ts.WriteLine "WH_NO,LABEL"

        Dim r As Long, whVal As String, labelVal As String
        For r = DATA_START_ROW To lastRow
            whVal = Trim(Replace(Replace(CStr(ws.Cells(r, "C").Value), Chr(160), " "), vbTab, " "))
            labelVal = Trim(CStr(ws.Cells(r, "D").Value))
            ' Chi chap nhan ma WH hop le (do dai >= 4, khong phai chu 'w' don le)
            If whVal <> "" And Len(whVal) >= 4 And LCase(whVal) <> "w" Then
                ts.WriteLine whVal & "," & labelVal
                validWHCount = validWHCount + 1
            End If
        Next r
        ts.Close
    End If

    ' Neu khong co WH hop le nao trong Excel, xoa file mapping de Engine chay che do TU DONG QUET 100%
    If validWHCount = 0 Then
        If fso.FileExists(mappingCsv) Then fso.DeleteFile mappingCsv, True
    End If

    ' Xac dinh engine
    Dim engineExe As String, enginePy As String, cmd As String
    engineExe = baseFolder & "wh_engine.exe"
    enginePy = baseFolder & "wh_engine.py"

    Dim mappingParam As String
    If validWHCount > 0 Then
        mappingParam = " --mapping """ & mappingCsv & """"
    Else
        mappingParam = ""
    End If

    If Dir(engineExe) <> "" Then
        cmd = """" & engineExe & """ --input """ & pdfInput & """" & mappingParam & _
              " --day """ & prefix & """ --start 1 --out-pdf """ & outPdf & """ --out-csv """ & outCsv & """"
    ElseIf Dir(enginePy) <> "" Then
        cmd = "python """ & enginePy & """ --input """ & pdfInput & """" & mappingParam & _
              " --day """ & prefix & """ --start 1 --out-pdf """ & outPdf & """ --out-csv """ & outCsv & """"
    Else
        MsgBox "Khong tim thay wh_engine.exe hoac wh_engine.py trong thu muc!", vbCritical, "Loi"
        Exit Sub
    End If

    ws.Range("C7").Value = "Dang phan tich & dong so PDF..."
    DoEvents

    Dim wsh As Object, exitCode As Long
    Set wsh = CreateObject("WScript.Shell")
    exitCode = wsh.Run(cmd, 0, True)

    If exitCode <> 0 Or Dir(outCsv) = "" Then
        ws.Range("C7").Value = "Loi khi xu ly (Ma loi: " & exitCode & ")"
        MsgBox "Co loi xay ra khi xu ly file PDF. Vui long kiem tra lai file dau vao!", vbCritical, "Loi"
        Exit Sub
    End If

    ' Cap nhat ket qua tu file CSV ra Sheet
    CapNhatKetQuaVaoSheet outCsv

    ws.Range("C6").Value = outPdf
    If validWHCount = 0 Then
        ws.Range("C7").Value = "Hoan tat: Da tu dong quet tu PDF & dong so!"
        MsgBox "Da tu dong quet toan bo ma WH tu PDF, gop nhom va dong so thanh cong!", vbInformation, "Thanh cong"
    Else
        ws.Range("C7").Value = "Hoan tat doi soat va dong so theo danh sach!"
        MsgBox "Da doi soat theo danh sach va dong so hoan tat!", vbInformation, "Thanh cong"
    End If
End Sub

' Cap nhat trang thai tu CSV vao bang
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

    ' Xoa sach bang cu truoc khi dien
    Dim maxR As Long
    maxR = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If maxR < DATA_START_ROW Then maxR = DATA_START_ROW + 50
    ws.Range("B" & DATA_START_ROW & ":H" & (maxR + 50)).ClearContents
    ws.Range("B" & DATA_START_ROW & ":H" & (maxR + 50)).Interior.Color = xlNone

    Dim i As Long, curRow As Long
    curRow = DATA_START_ROW
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

                If InStr(cols(6), "THIẾU") > 0 Or InStr(cols(6), "THIEU") > 0 Then
                    ws.Range(ws.Cells(curRow, "B"), ws.Cells(curRow, "H")).Interior.Color = RGB(255, 204, 204)
                ElseIf InStr(cols(6), "THỪA") > 0 Or InStr(cols(6), "THUA") > 0 Then
                    ws.Range(ws.Cells(curRow, "B"), ws.Cells(curRow, "H")).Interior.Color = RGB(255, 255, 204)
                Else
                    ws.Range(ws.Cells(curRow, "B"), ws.Cells(curRow, "H")).Interior.Color = xlNone
                End If

                curRow = curRow + 1
            End If
        End If
    Next i

    If curRow > DATA_START_ROW Then
        ws.Range("B" & DATA_START_ROW & ":H" & (curRow - 1)).Borders.LineStyle = xlContinuous
    End If
End Sub

' Nut [2. MO PDF KET QUA DE IN]
Public Sub MoPDFKetQuaDeIn()
    Dim outPdf As String
    outPdf = Trim(CStr(ThisWorkbook.Worksheets(SHEET_MAIN).Range("C6").Value))

    If outPdf = "" Or Dir(outPdf) = "" Then
        MsgBox "Chua co file PDF ket qua. Vui long bam [1. XU LY & DOI SOAT PDF] truoc!", vbExclamation, "Thong bao"
        Exit Sub
    End If

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run """" & outPdf & """", 1, False
End Sub
