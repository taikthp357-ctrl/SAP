# -*- coding: utf-8 -*-
"""
Script to create the Excel Tool: 261107_TOOL_ATT_MAIL_MIGO_FAST.xlsm
Features:
- Input: DOC (10 digits), Ticket Number (So phieu)
- Folder selection button + cell
- File matching: matches files in folder containing ticket number or doc number (ignores file type/category/material)
- 1 single Result column: DONE or Lỗi
- Option for 1 Tab or 2 Tabs (cell G3/H3, max 2)
- Function buttons: Điền số phiếu, Att mail & Điền phiếu, Chọn Folder, Xóa kết quả
"""

import sys
import os
import win32com.client

vba_mod_main = '''Attribute VB_Name = "ModMain"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Const SH_NAME As String = "ATT_MAIL"
Public Const FIRST_ROW As Long = 7
Public Const COL_STT As Long = 1
Public Const COL_DOC As Long = 2
Public Const COL_TICKET As Long = 3
Public Const COL_RESULT As Long = 4
Public Const COL_ATTACHED As Long = 5
Public Const COL_TIME As Long = 6

Public Sub ChonFolder()
    Dim fd As Object
    Set fd = Application.FileDialog(4) ' 4 = msoFileDialogFolderPicker
    With fd
        .Title = "Chọn thư mục chứa tệp đính kèm (BBBG / VAT / Hóa đơn)"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Worksheets(SH_NAME).Range("C3").Value
        If .Show = -1 Then
            ThisWorkbook.Worksheets(SH_NAME).Range("C3").Value = .SelectedItems(1)
            ThisWorkbook.Worksheets(SH_NAME).Range("D4").Value = "Đã chọn thư mục: " & .SelectedItems(1)
            ThisWorkbook.Worksheets(SH_NAME).Range("D4").Font.Color = RGB(0, 102, 204)
        End If
    End With
End Sub

Public Sub XoaKetQua()
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Worksheets(SH_NAME)
    lastRow = ws.Cells(ws.Rows.Count, COL_DOC).End(xlUp).Row
    If lastRow < FIRST_ROW Then lastRow = FIRST_ROW
    
    ws.Range(ws.Cells(FIRST_ROW, COL_RESULT), ws.Cells(lastRow + 10, COL_TIME)).ClearContents
    ws.Range(ws.Cells(FIRST_ROW, COL_RESULT), ws.Cells(lastRow + 10, COL_RESULT)).Interior.ColorIndex = xlNone
    ws.Range("D4").Value = "Đã làm sạch kết quả. Sẵn sàng làm việc."
    ws.Range("D4").Font.Color = RGB(80, 80, 80)
End Sub

Public Sub DienSoPhieu()
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim docNo As String, ticketNo As String, yearStr As String
    Dim numTabs As Long, actualTabs As Long, sess1 As Object, sess2 As Object, curSess As Object
    Dim errMsg As String, outMsg As String, tStart As Double
    Dim succCount As Long, failCount As Long, totalCount As Long
    
    Set ws = ThisWorkbook.Worksheets(SH_NAME)
    lastRow = ws.Cells(ws.Rows.Count, COL_DOC).End(xlUp).Row
    If lastRow < FIRST_ROW Then
        MsgBox "Vui lòng nhập ít nhất một số chứng từ DOC vào bảng!", vbExclamation, "Chưa có dữ liệu"
        Exit Sub
    End If
    
    yearStr = Trim(CStr(ws.Range("C4").Value))
    If yearStr = "" Then yearStr = CStr(Year(Date))
    
    numTabs = CLng(ws.Range("H3").Value)
    If numTabs < 1 Then numTabs = 1
    If numTabs > 2 Then numTabs = 2
    
    ws.Range("D4").Value = "Đang kết nối SAP GUI..."
    ws.Range("D4").Font.Color = RGB(255, 128, 0)
    DoEvents
    
    If Not ConnectSAP(numTabs, sess1, sess2, actualTabs, errMsg) Then
        ws.Range("D4").Value = "Lỗi: " & errMsg
        ws.Range("D4").Font.Color = RGB(204, 0, 0)
        MsgBox errMsg, vbCritical, "Lỗi kết nối SAP"
        Exit Sub
    End If
    
    ws.Range("D4").Value = "Đang chạy điền số phiếu (Sử dụng " & actualTabs & " tab)..."
    ws.Range("D4").Font.Color = RGB(0, 102, 204)
    DoEvents
    
    succCount = 0
    failCount = 0
    totalCount = 0
    
    For r = FIRST_ROW To lastRow
        docNo = Trim(CStr(ws.Cells(r, COL_DOC).Value))
        ticketNo = Trim(CStr(ws.Cells(r, COL_TICKET).Value))
        
        If docNo <> "" Then
            totalCount = totalCount + 1
            tStart = Timer
            
            If actualTabs = 2 Then
                If (r Mod 2) = 1 Then
                    Set curSess = sess1
                Else
                    Set curSess = sess2
                End If
            Else
                Set curSess = sess1
            End If
            
            ws.Cells(r, COL_RESULT).Value = "Đang xử lý..."
            ws.Cells(r, COL_RESULT).Interior.Color = RGB(255, 255, 204)
            ws.Cells(r, COL_RESULT).Font.Color = RGB(0, 0, 0)
            DoEvents
            
            ' 1. Nạp MIGO A12
            If OpenMigoA12(curSess, docNo, yearStr, outMsg) Then
                ' 2. Điền số phiếu
                If SetHeadText(curSess, ticketNo, outMsg) Then
                    ws.Cells(r, COL_RESULT).Value = "DONE"
                    ws.Cells(r, COL_RESULT).Interior.Color = RGB(218, 242, 208)
                    ws.Cells(r, COL_RESULT).Font.Color = RGB(0, 128, 0)
                    succCount = succCount + 1
                Else
                    ws.Cells(r, COL_RESULT).Value = "Lỗi: " & outMsg
                    ws.Cells(r, COL_RESULT).Interior.Color = RGB(255, 220, 220)
                    ws.Cells(r, COL_RESULT).Font.Color = RGB(204, 0, 0)
                    failCount = failCount + 1
                End If
            Else
                ws.Cells(r, COL_RESULT).Value = "Lỗi: " & outMsg
                ws.Cells(r, COL_RESULT).Interior.Color = RGB(255, 220, 220)
                ws.Cells(r, COL_RESULT).Font.Color = RGB(204, 0, 0)
                failCount = failCount + 1
            End If
            
            ' Dọn dẹp trạng thái MIGO để sẵn sàng cho dòng tiếp theo
            ResetMigo curSess
            ws.Cells(r, COL_TIME).Value = Format(Timer - tStart, "0.0") & "s"
            DoEvents
        End If
    Next r
    
    ws.Range("D4").Value = "Hoàn tất điền số phiếu! Thành công: " & succCount & "/" & totalCount & " chứng từ."
    ws.Range("D4").Font.Color = RGB(0, 153, 76)
    MsgBox "Đã hoàn tất điền số phiếu!" & vbCrLf & _
           "- Thành công: " & succCount & vbCrLf & _
           "- Lỗi: " & failCount & vbCrLf & _
           "- Tổng cộng: " & totalCount, vbInformation, "Kết quả thực hiện"
End Sub

Public Sub AttMail()
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim docNo As String, ticketNo As String, yearStr As String, folderPath As String
    Dim numTabs As Long, actualTabs As Long, sess1 As Object, sess2 As Object, curSess As Object
    Dim errMsg As String, outMsg As String, tStart As Double
    Dim succCount As Long, failCount As Long, totalCount As Long
    Dim allFolderFiles As Collection, matchedFiles As Collection
    Dim fp As Variant, attOk As Boolean, attSuccCount As Long, attList As String
    
    Set ws = ThisWorkbook.Worksheets(SH_NAME)
    lastRow = ws.Cells(ws.Rows.Count, COL_DOC).End(xlUp).Row
    If lastRow < FIRST_ROW Then
        MsgBox "Vui lòng nhập ít nhất một số chứng từ DOC vào bảng!", vbExclamation, "Chưa có dữ liệu"
        Exit Sub
    End If
    
    folderPath = Trim(CStr(ws.Range("C3").Value))
    If folderPath = "" Then
        MsgBox "Vui lòng chọn Thư mục chứa tệp đính kèm trước khi bấm Att Mail!", vbExclamation, "Thiếu thư mục"
        Exit Sub
    End If
    
    yearStr = Trim(CStr(ws.Range("C4").Value))
    If yearStr = "" Then yearStr = CStr(Year(Date))
    
    numTabs = CLng(ws.Range("H3").Value)
    If numTabs < 1 Then numTabs = 1
    If numTabs > 2 Then numTabs = 2
    
    ws.Range("D4").Value = "Đang quét thư mục tệp..."
    ws.Range("D4").Font.Color = RGB(255, 128, 0)
    DoEvents
    
    Set allFolderFiles = ScanFolderFiles(folderPath)
    If allFolderFiles.Count = 0 Then
        ws.Range("D4").Value = "Thư mục không có tệp nào hoặc không truy cập được!"
        ws.Range("D4").Font.Color = RGB(204, 0, 0)
        MsgBox "Không tìm thấy tệp nào trong thư mục đã chọn, hoặc đường dẫn mạng LAN không truy cập được!" & vbCrLf & folderPath, vbCritical, "Thư mục rỗng / Lỗi mạng"
        Exit Sub
    End If
    
    ws.Range("D4").Value = "Tìm thấy " & allFolderFiles.Count & " tệp trong thư mục. Đang kết nối SAP GUI..."
    DoEvents
    
    If Not ConnectSAP(numTabs, sess1, sess2, actualTabs, errMsg) Then
        ws.Range("D4").Value = "Lỗi: " & errMsg
        ws.Range("D4").Font.Color = RGB(204, 0, 0)
        MsgBox errMsg, vbCritical, "Lỗi kết nối SAP"
        Exit Sub
    End If
    
    ws.Range("D4").Value = "Đang chạy Att Mail & Điền số phiếu (Sử dụng " & actualTabs & " tab)..."
    ws.Range("D4").Font.Color = RGB(0, 102, 204)
    DoEvents
    
    succCount = 0
    failCount = 0
    totalCount = 0
    
    For r = FIRST_ROW To lastRow
        docNo = Trim(CStr(ws.Cells(r, COL_DOC).Value))
        ticketNo = Trim(CStr(ws.Cells(r, COL_TICKET).Value))
        
        If docNo <> "" Then
            totalCount = totalCount + 1
            tStart = Timer
            
            ' Tìm các tệp trong thư mục có tên chứa số giống hệt (không quan tâm loại tệp/mã gì)
            Set matchedFiles = FindMatchingFiles(allFolderFiles, docNo, ticketNo)
            
            If matchedFiles.Count = 0 Then
                ws.Cells(r, COL_RESULT).Value = "Lỗi: Không tìm thấy file"
                ws.Cells(r, COL_RESULT).Interior.Color = RGB(255, 220, 220)
                ws.Cells(r, COL_RESULT).Font.Color = RGB(204, 0, 0)
                ws.Cells(r, COL_ATTACHED).Value = "Không có file khớp số phiếu " & ticketNo
                failCount = failCount + 1
            Else
                If actualTabs = 2 Then
                    If (r Mod 2) = 1 Then
                        Set curSess = sess1
                    Else
                        Set curSess = sess2
                    End If
                Else
                    Set curSess = sess1
                End If
                
                ws.Cells(r, COL_RESULT).Value = "Đang đính kèm " & matchedFiles.Count & " file..."
                ws.Cells(r, COL_RESULT).Interior.Color = RGB(255, 255, 204)
                ws.Cells(r, COL_RESULT).Font.Color = RGB(0, 0, 0)
                DoEvents
                
                ' 1. Mở MIGO A12
                If OpenMigoA12(curSess, docNo, yearStr, outMsg) Then
                    ' 2. Điền số phiếu HeadText nếu có
                    If ticketNo <> "" Then
                        SetHeadText curSess, ticketNo, outMsg
                    End If
                    
                    ' 3. Đính kèm toàn bộ các file khớp vào GOS
                    attSuccCount = 0
                    attList = ""
                    For Each fp In matchedFiles
                        attOk = AttachFileToGOS(curSess, CStr(fp), outMsg)
                        If attOk Then
                            attSuccCount = attSuccCount + 1
                            attList = attList & IIf(attList = "", "", ", ") & Mid(CStr(fp), InStrRev(CStr(fp), "\") + 1)
                        End If
                    Next fp
                    
                    If attSuccCount > 0 Then
                        ws.Cells(r, COL_RESULT).Value = "DONE"
                        ws.Cells(r, COL_RESULT).Interior.Color = RGB(218, 242, 208)
                        ws.Cells(r, COL_RESULT).Font.Color = RGB(0, 128, 0)
                        ws.Cells(r, COL_ATTACHED).Value = attList
                        succCount = succCount + 1
                    Else
                        ws.Cells(r, COL_RESULT).Value = "Lỗi: " & outMsg
                        ws.Cells(r, COL_RESULT).Interior.Color = RGB(255, 220, 220)
                        ws.Cells(r, COL_RESULT).Font.Color = RGB(204, 0, 0)
                        ws.Cells(r, COL_ATTACHED).Value = "Đính kèm thất bại: " & outMsg
                        failCount = failCount + 1
                    End If
                Else
                    ws.Cells(r, COL_RESULT).Value = "Lỗi: " & outMsg
                    ws.Cells(r, COL_RESULT).Interior.Color = RGB(255, 220, 220)
                    ws.Cells(r, COL_RESULT).Font.Color = RGB(204, 0, 0)
                    failCount = failCount + 1
                End If
                
                ' Dọn dẹp trạng thái MIGO để sẵn sàng cho dòng tiếp theo
                ResetMigo curSess
            End If
            
            ws.Cells(r, COL_TIME).Value = Format(Timer - tStart, "0.0") & "s"
            DoEvents
        End If
    Next r
    
    ws.Range("D4").Value = "Hoàn tất Att Mail! Thành công: " & succCount & "/" & totalCount & " chứng từ."
    ws.Range("D4").Font.Color = RGB(0, 153, 76)
    MsgBox "Đã hoàn tất Đính kèm Mail & Điền số phiếu!" & vbCrLf & _
           "- Thành công: " & succCount & vbCrLf & _
           "- Lỗi / Thiếu file: " & failCount & vbCrLf & _
           "- Tổng cộng: " & totalCount, vbInformation, "Kết quả thực hiện"
End Sub
'''

vba_mod_sap = '''Attribute VB_Name = "ModSAP"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Function ConnectSAP(ByVal numTabs As Long, ByRef sess1 As Object, ByRef sess2 As Object, ByRef actualTabs As Long, ByRef errMsg As String) As Boolean
    Dim SapGuiAuto As Object, app As Object, conn As Object
    Dim tStart As Double
    
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If SapGuiAuto Is Nothing Then
        errMsg = "Chưa phát hiện SAP GUI đang mở. Vui lòng đăng nhập SAP trước."
        ConnectSAP = False
        Exit Function
    End If
    
    Set app = SapGuiAuto.GetScriptingEngine
    If app Is Nothing Then
        errMsg = "Không lấy được Scripting Engine. Vui lòng bật SAP GUI Scripting trong Options."
        ConnectSAP = False
        Exit Function
    End If
    
    If app.Children.Count = 0 Then
        errMsg = "Chưa có kết nối Connection nào trong SAP."
        ConnectSAP = False
        Exit Function
    End If
    
    Set conn = app.Children(0)
    If conn.Children.Count = 0 Then
        errMsg = "Không tìm thấy phiên Session nào trong SAP."
        ConnectSAP = False
        Exit Function
    End If
    
    Set sess1 = conn.Children(0)
    actualTabs = 1
    
    If numTabs >= 2 Then
        If conn.Children.Count >= 2 Then
            Set sess2 = conn.Children(1)
            actualTabs = 2
        Else
            ' Tạo thêm Session 2 tự động
            sess1.CreateSession
            tStart = Timer
            Do While Timer - tStart < 3#
                DoEvents
                Sleep 100
                If conn.Children.Count >= 2 Then
                    Set sess2 = conn.Children(1)
                    actualTabs = 2
                    Exit Do
                End If
            Loop
        End If
    End If
    
    ConnectSAP = True
End Function

Public Function OpenMigoA12(session As Object, ByVal docNo As String, ByVal yearStr As String, ByRef outMsg As String) As Boolean
    Dim sbar As Object, sType As String, sMsg As String
    Dim matDocFld As Object, yearFld As Object, titlBar As Object, titlText As String
    
    On Error Resume Next
    
    ' 1. Kiểm tra màn hình MIGO
    If UCase(session.Info.Transaction) <> "MIGO" Then
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nMIGO"
        session.findById("wnd[0]").SendVKey 0
        SAP_Wait session, 1.0
    End If
    
    ' 2. Đặt Action = A12 (Change), RefDoc = R02 (Material Document)
    session.findById("wnd[0]/usr/subSUB_MAIN:SAPLMIGO:0001/cmbGODYNPRO-ACTION").Key = "A12"
    session.findById("wnd[0]/usr/subSUB_MAIN:SAPLMIGO:0001/cmbGODYNPRO-REFDOC").Key = "R02"
    
    ' Thoát chứng từ cũ nếu ô nhập doc đang bị khóa
    Set matDocFld = session.findById("wnd[0]/usr/subSUB_MAIN:SAPLMIGO:0001/ctxtGODYNPRO-MAT_DOC")
    If matDocFld Is Nothing Or Not matDocFld.Changeable Then
        session.findById("wnd[0]").SendVKey 12 ' Esc
        SAP_Wait session, 0.3
        session.findById("wnd[0]/usr/subSUB_MAIN:SAPLMIGO:0001/cmbGODYNPRO-ACTION").Key = "A12"
        session.findById("wnd[0]/usr/subSUB_MAIN:SAPLMIGO:0001/cmbGODYNPRO-REFDOC").Key = "R02"
    End If
    
    Set matDocFld = session.findById("wnd[0]/usr/subSUB_MAIN:SAPLMIGO:0001/ctxtGODYNPRO-MAT_DOC")
    If matDocFld Is Nothing Then
        outMsg = "Không tìm thấy ô nhập số DOC trên MIGO"
        OpenMigoA12 = False
        Exit Function
    End If
    
    matDocFld.Text = Trim(docNo)
    
    Set yearFld = session.findById("wnd[0]/usr/subSUB_MAIN:SAPLMIGO:0001/txtGODYNPRO-DOC_YEAR")
    If Not yearFld Is Nothing Then
        yearFld.Text = Trim(yearStr)
    End If
    
    ' 3. Nhấn Enter nạp chứng từ
    session.findById("wnd[0]").SendVKey 0
    SAP_Wait session, 1.2
    
    ' 4. Kiểm tra thanh trạng thái sbar
    Set sbar = session.findById("wnd[0]/sbar")
    If Not sbar Is Nothing Then
        sType = sbar.MessageType
        sMsg = sbar.Text
        If sType = "E" Or sType = "A" Then
            session.findById("wnd[0]").SendVKey 12 ' Esc
            outMsg = sMsg
            OpenMigoA12 = False
            Exit Function
        End If
    End If
    
    ' 5. Kiểm tra tiêu đề Display vs Change
    Set titlBar = session.findById("wnd[0]/titl")
    If Not titlBar Is Nothing Then
        titlText = UCase(titlBar.Text)
        If InStr(titlText, "DISPLAY") > 0 Then
            session.findById("wnd[0]").SendVKey 12
            outMsg = "MIGO đang ở chế độ Display (không cho sửa)"
            OpenMigoA12 = False
            Exit Function
        End If
    End If
    
    outMsg = "OK"
    OpenMigoA12 = True
End Function

Public Function SetHeadText(session As Object, ByVal ticketNo As String, ByRef outMsg As String) As Boolean
    Dim headFld As Object, sbar As Object, sType As String, sMsg As String
    
    If Trim(ticketNo) = "" Then
        outMsg = "Không có số phiếu"
        SetHeadText = True
        Exit Function
    End If
    
    On Error Resume Next
    session.SendCommand "MIGO_OK_HEADER_OPEN"
    SAP_Wait session, 0.2
    
    Set headFld = session.findById("wnd[0]").findByName("GOHEAD-BKTXT", "GuiTextField")
    If headFld Is Nothing Then
        Set headFld = session.findById("wnd[0]/usr/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_HEADER/tabpOK_HEADER_GENERAL/ssubSUB_HEADER:SAPLMIGO:0102/txtGOHEAD-BKTXT")
    End If
    
    If headFld Is Nothing Then
        outMsg = "Không tìm thấy ô Header Text"
        SetHeadText = False
        Exit Function
    End If
    
    headFld.Text = Trim(ticketNo)
    
    ' Nhấn Lưu (Save / Ctrl+S)
    session.findById("wnd[0]").SendVKey 11
    SAP_Wait session, 0.8
    
    DismissModalPopups session
    
    Set sbar = session.findById("wnd[0]/sbar")
    If Not sbar Is Nothing Then
        sType = sbar.MessageType
        sMsg = sbar.Text
        If sType = "E" Or sType = "A" Then
            outMsg = sMsg
            SetHeadText = False
            Exit Function
        End If
    End If
    
    outMsg = "OK"
    SetHeadText = True
End Function

Public Function AttachFileToGOS(session As Object, ByVal filePath As String, ByRef outMsg As String) As Boolean
    Dim fso As Object, parentFld As String, fName As String
    Dim wnd1 As Object, sbar As Object, sType As String, sMsg As String
    Dim tStart As Double, opened As Boolean
    
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        outMsg = "Tệp không tồn tại"
        AttachFileToGOS = False
        Exit Function
    End If
    
    parentFld = fso.GetParentFolderName(filePath)
    fName = fso.GetFileName(filePath)
    
    opened = False
    ' Cách 1: Nút GOS Toolbox
    session.findById("wnd[0]/titl/shellcont/shell").PressContextButton "%GOS_TOOLBOX"
    session.findById("wnd[0]/titl/shellcont/shell").SelectContextMenuItem "%GOS_PCATTA_CREA"
    If Err.Number = 0 Then opened = True
    Err.Clear
    
    ' Cách 2: Menu System
    If Not opened Then
        session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[0]").Select
        If Err.Number = 0 Then opened = True
        Err.Clear
    End If
    
    ' Cách 3: Command code
    If Not opened Then
        session.findById("wnd[0]/tbar[0]/okcd").Text = "%GOS_PCATTA_CREA"
        session.findById("wnd[0]").SendVKey 0
    End If
    
    ' Chờ cửa sổ wnd[1] hiện ra (Import File dialog)
    tStart = Timer
    Do
        Set wnd1 = session.findById("wnd[1]")
        If Not wnd1 Is Nothing Then Exit Do
        DoEvents
        Sleep 40
    Loop While Timer - tStart < 2.5
    
    If wnd1 Is Nothing Then
        outMsg = "Không mở được hộp thoại GOS"
        AttachFileToGOS = False
        Exit Function
    End If
    
    ' Điền đường dẫn thư mục và tên tệp
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = parentFld
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = fName
    session.findById("wnd[1]/tbar[0]/btn[0]").Press
    
    ' Chờ wnd[1] đóng
    tStart = Timer
    Do
        Set wnd1 = Nothing
        Set wnd1 = session.findById("wnd[1]")
        If wnd1 Is Nothing Then Exit Do
        DoEvents
        Sleep 40
    Loop While Timer - tStart < 2.5
    
    Set sbar = session.findById("wnd[0]/sbar")
    If Not sbar Is Nothing Then
        sType = sbar.MessageType
        sMsg = sbar.Text
        If sType = "E" Or sType = "A" Then
            outMsg = sMsg
            AttachFileToGOS = False
            Exit Function
        End If
    End If
    
    outMsg = "OK"
    AttachFileToGOS = True
End Function

Public Sub ResetMigo(session As Object)
    On Error Resume Next
    Dim w0 As Object
    Set w0 = session.findById("wnd[0]")
    If Not w0 Is Nothing Then
        w0.SendVKey 12 ' Esc thoát chứng từ
        SAP_Wait session, 0.2
    End If
    DismissModalPopups session
End Sub

Public Sub DismissModalPopups(session As Object)
    On Error Resume Next
    Dim wnd As Object, idx As Long
    For idx = 3 To 1 Step -1
        Set wnd = session.findById("wnd[" & idx & "]")
        If Not wnd Is Nothing Then
            wnd.Close
            If Err.Number <> 0 Then
                Err.Clear
                wnd.SendVKey 12
            End If
        End If
    Next idx
End Sub

Public Sub SAP_Wait(session As Object, ByVal maxSeconds As Double)
    Dim tStart As Double
    tStart = Timer
    Do While Timer - tStart < maxSeconds
        On Error Resume Next
        If Not session.Busy Then Exit Do
        DoEvents
        Sleep 50
    Loop
End Sub
'''

vba_mod_file = '''Attribute VB_Name = "ModFileMatch"
Option Explicit

Public Function ScanFolderFiles(ByVal folderPath As String) As Collection
    Dim fso As Object, fld As Object, f As Object
    Dim col As New Collection
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        Set ScanFolderFiles = col
        Exit Function
    End If
    Set fld = fso.GetFolder(folderPath)
    For Each f In fld.Files
        col.Add f.Path
    Next f
    Set ScanFolderFiles = col
End Function

Public Function FindMatchingFiles(ByVal allFiles As Collection, ByVal docNo As String, ByVal ticketNo As String) As Collection
    Dim res As New Collection
    Dim fPath As Variant, fName As String, fNameClean As String
    Dim tok As Variant, tokens As Collection
    Dim isMatch As Boolean
    
    Set tokens = ExtractSearchTokens(ticketNo, docNo)
    If tokens.Count = 0 Then
        Set FindMatchingFiles = res
        Exit Function
    End If
    
    For Each fPath In allFiles
        fName = Mid(CStr(fPath), InStrRev(CStr(fPath), "\") + 1)
        fNameClean = CleanString(fName)
        isMatch = False
        
        For Each tok In tokens
            If InStr(1, fName, CStr(tok), vbTextCompare) > 0 Or _
               InStr(1, fNameClean, CleanString(CStr(tok)), vbTextCompare) > 0 Then
                isMatch = True
                Exit For
            End If
        Next tok
        
        If isMatch Then
            res.Add CStr(fPath)
        End If
    Next fPath
    
    Set FindMatchingFiles = res
End Function

Public Function ExtractSearchTokens(ByVal ticketNo As String, ByVal docNo As String) As Collection
    Dim col As New Collection
    Dim raw As String, i As Long, ch As String, curNum As String
    
    raw = Trim(ticketNo)
    If raw = "" Then raw = Trim(docNo)
    
    curNum = ""
    For i = 1 To Len(raw)
        ch = Mid(raw, i, 1)
        If ch >= "0" And ch <= "9" Then
            curNum = curNum & ch
        Else
            If Len(curNum) >= 3 Then
                AddUnique col, curNum
            End If
            curNum = ""
        End If
    Next i
    If Len(curNum) >= 3 Then
        AddUnique col, curNum
    End If
    
    ' Nếu số phiếu có ký tự chữ và số ngắn
    Dim cleanTick As String
    cleanTick = CleanString(ticketNo)
    If Len(cleanTick) >= 4 Then
        AddUnique col, cleanTick
    End If
    
    Set ExtractSearchTokens = col
End Function

Private Sub AddUnique(ByRef col As Collection, ByVal valStr As String)
    Dim v As Variant
    For Each v In col
        If CStr(v) = valStr Then Exit Sub
    Next v
    col.Add valStr
End Sub

Public Function CleanString(ByVal s As String) As String
    Dim res As String, i As Long, ch As String
    res = ""
    For i = 1 To Len(s)
        ch = UCase(Mid(s, i, 1))
        If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Then
            res = res & ch
        End If
    Next i
    CleanString = res
End Function
'''

def build_excel_tool():
    out_dir = r"G:\SAP"
    out_file = os.path.join(out_dir, "261107_TOOL_ATT_MAIL_MIGO_FAST.xlsm")
    
    print(f"Khoi tao Excel de tao file: {out_file}...")
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "ATT_MAIL"
        
        # 1. Page & Grid setup
        ws.Cells.Font.Name = "Calibri"
        ws.Cells.Font.Size = 11
        
        # 2. Header Banner
        ws.Range("A1:F1").Merge()
        ws.Range("A1").Value = "CÔNG CỤ ĐÍNH KÈM FILE GOS & ĐIỀN SỐ PHIẾU MIGO SAP"
        ws.Range("A1").Font.Size = 16
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Color = 0xFFFFFF
        ws.Range("A1").Interior.Color = 0x794E1F  # Dark Blue #1F4E79
        ws.Range("A1").HorizontalAlignment = -4108  # Center
        ws.Range("A1").VerticalAlignment = -4108
        ws.Rows(1).RowHeight = 35
        
        # Subtitle
        ws.Range("A2:F2").Merge()
        ws.Range("A2").Value = "Tự động so khớp số phiếu/DOC trong thư mục • Hỗ trợ Turbo 2 Tab song song • Nhẹ & Tối ưu tốc độ"
        ws.Range("A2").Font.Size = 10
        ws.Range("A2").Font.Italic = True
        ws.Range("A2").Font.Color = 0x505050
        ws.Range("A2").HorizontalAlignment = -4108
        ws.Rows(2).RowHeight = 18
        
        # 3. Config Area (Row 3 & 4)
        ws.Range("B3").Value = "Thư mục tệp:"
        ws.Range("B3").Font.Bold = True
        ws.Range("B3").HorizontalAlignment = -4152  # Right
        
        default_folder = r"\\10.128.10.72\fvn$\03.Khối sản xuất\01.Nhà máy VinFast Hải Phòng\04.Xưởng lắp ráp kim loại - cơ khí NMHP\01. Lắp ráp kim loại\5.LOGISTIC\III. Internal\III. Kho P part\3. Theo doi hang nhap P\240101_FT-P\260802_ File ATT BBGH"
        ws.Range("C3:E3").Merge()
        ws.Range("C3").Value = default_folder
        ws.Range("C3").Interior.Color = 0xF5F5F5
        ws.Range("C3").Borders.LineStyle = 1
        ws.Range("C3").Borders.Color = 0xD0D0D0
        ws.Range("C3").Font.Size = 10
        
        ws.Range("G3").Value = "Số Tab SAP:"
        ws.Range("G3").Font.Bold = True
        ws.Range("G3").HorizontalAlignment = -4152
        
        ws.Range("H3").Value = 2
        ws.Range("H3").Font.Bold = True
        ws.Range("H3").HorizontalAlignment = -4108
        ws.Range("H3").Interior.Color = 0xE6F2FF
        ws.Range("H3").Borders.LineStyle = 1
        ws.Range("H3").Borders.Color = 0x3399FF
        # Data validation for H3 (1, 2)
        h3_val = ws.Range("H3").Validation
        h3_val.Delete()
        h3_val.Add(Type=3, AlertStyle=1, Operator=1, Formula1="1, 2")  # 3 = xlValidateList
        h3_val.IgnoreBlank = True
        h3_val.InCellDropdown = True
        
        ws.Range("B4").Value = "Năm chứng từ:"
        ws.Range("B4").Font.Bold = True
        ws.Range("B4").HorizontalAlignment = -4152
        
        ws.Range("C4").Value = "2026"
        ws.Range("C4").HorizontalAlignment = -4108
        ws.Range("C4").Interior.Color = 0xF5F5F5
        ws.Range("C4").Borders.LineStyle = 1
        ws.Range("C4").Borders.Color = 0xD0D0D0
        
        ws.Range("D4:F4").Merge()
        ws.Range("D4").Value = "Sẵn sàng làm việc."
        ws.Range("D4").Font.Italic = True
        ws.Range("D4").Font.Color = 0x505050
        ws.Range("D4").HorizontalAlignment = -4131  # Left
        
        ws.Rows(3).RowHeight = 24
        ws.Rows(4).RowHeight = 22
        ws.Rows(5).RowHeight = 32  # Button space
        
        # 4. Table Headers (Row 6)
        headers = [
            ("A6", "STT", 6),
            ("B6", "Số DOC (10 số)", 18),
            ("C6", "Số phiếu", 22),
            ("D6", "Kết quả", 25),
            ("E6", "Tệp đã đính kèm", 40),
            ("F6", "Thời gian", 12)
        ]
        
        for cell, title, width in headers:
            rng = ws.Range(cell)
            rng.Value = title
            rng.Font.Bold = True
            rng.Font.Color = 0xFFFFFF
            rng.Interior.Color = 0x4A2500  # Dark navy/blue
            rng.HorizontalAlignment = -4108
            rng.VerticalAlignment = -4108
            rng.Borders.LineStyle = 1
            rng.Borders.Color = 0x202020
            col_letter = cell[0]
            ws.Columns(col_letter).ColumnWidth = width
            
        ws.Columns("G").ColumnWidth = 14
        ws.Columns("H").ColumnWidth = 8
        ws.Rows(6).RowHeight = 26
        
        # 5. Sample Rows (Row 7 to 11)
        sample_data = [
            (1, "5007186634", "6691 + 6714"),
            (2, "5007186547", "6691 + 6714"),
            (3, "5007194960", "VAT_6736"),
            (4, "5007195402", "VAT_6738"),
            (5, "5007151342", "6538")
        ]
        
        FIRST_ROW = 7
        for idx, (stt, doc, ticket) in enumerate(sample_data, start=FIRST_ROW):
            ws.Cells(idx, 1).Value = stt
            ws.Cells(idx, 1).HorizontalAlignment = -4108
            ws.Cells(idx, 2).Value = f"'{doc}"
            ws.Cells(idx, 2).HorizontalAlignment = -4108
            ws.Cells(idx, 3).Value = ticket
            ws.Cells(idx, 3).HorizontalAlignment = -4108
            
            # Format row
            for col in range(1, 7):
                c = ws.Cells(idx, col)
                c.Borders.LineStyle = 1
                c.Borders.Color = 0xD9D9D9
                if (idx % 2) == 0:
                    c.Interior.Color = 0xF9F9F9
            ws.Rows(idx).RowHeight = 22
            
        # Empty rows up to row 30 with borders
        for idx in range(12, 35):
            ws.Cells(idx, 1).Value = idx - 6
            ws.Cells(idx, 1).HorizontalAlignment = -4108
            for col in range(1, 7):
                c = ws.Cells(idx, col)
                c.Borders.LineStyle = 1
                c.Borders.Color = 0xE8E8E8
                if (idx % 2) == 0:
                    c.Interior.Color = 0xFAFAFA
            ws.Rows(idx).RowHeight = 20

        # 6. Add Buttons using Excel Shapes
        # Shape 1: Nút "Chọn Folder" (Row 3, Column F)
        btn_folder = ws.Shapes.AddShape(1, 480, 48, 110, 26)  # 1 = msoShapeRectangle
        btn_folder.TextFrame.Characters().Text = "📁 Chọn Folder"
        btn_folder.Fill.ForeColor.RGB = 0x0066CC  # Blue
        btn_folder.Line.Visible = False
        btn_folder.TextFrame.Characters().Font.Bold = True
        btn_folder.TextFrame.Characters().Font.Color = 0xFFFFFF
        btn_folder.TextFrame.Characters().Font.Size = 10
        btn_folder.OnAction = "ChonFolder"

        # Shape 2: Nút "Att Mail & Điền phiếu" (Green Action Button)
        btn_att = ws.Shapes.AddShape(5, 75, 96, 210, 32)  # 5 = msoShapeRoundedRectangle
        btn_att.TextFrame.Characters().Text = "🚀 ATT MAIL & ĐIỀN PHIẾU"
        btn_att.Fill.ForeColor.RGB = 0x228B22  # Forest Green
        btn_att.Line.Visible = False
        btn_att.TextFrame.Characters().Font.Bold = True
        btn_att.TextFrame.Characters().Font.Color = 0xFFFFFF
        btn_att.TextFrame.Characters().Font.Size = 11
        btn_att.OnAction = "AttMail"

        # Shape 3: Nút "Điền số phiếu" (Blue Action Button)
        btn_head = ws.Shapes.AddShape(5, 305, 96, 180, 32)
        btn_head.TextFrame.Characters().Text = "✍️ ĐIỀN SỐ PHIẾU (HEADTEXT)"
        btn_head.Fill.ForeColor.RGB = 0xCC7A00  # Dark Cyan/Blue
        btn_head.Line.Visible = False
        btn_head.TextFrame.Characters().Font.Bold = True
        btn_head.TextFrame.Characters().Font.Color = 0xFFFFFF
        btn_head.TextFrame.Characters().Font.Size = 10.5
        btn_head.OnAction = "DienSoPhieu"

        # Shape 4: Nút "Làm sạch kết quả" (Gray Button)
        btn_clear = ws.Shapes.AddShape(5, 505, 96, 140, 32)
        btn_clear.TextFrame.Characters().Text = "🧹 Xóa kết quả"
        btn_clear.Fill.ForeColor.RGB = 0x708090  # Slate Gray
        btn_clear.Line.Visible = False
        btn_clear.TextFrame.Characters().Font.Bold = True
        btn_clear.TextFrame.Characters().Font.Color = 0xFFFFFF
        btn_clear.TextFrame.Characters().Font.Size = 10
        btn_clear.OnAction = "XoaKetQua"

        # 7. Add VBA Modules
        vba_proj = wb.VBProject
        
        # Module ModMain
        mod1 = vba_proj.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        mod1.Name = "ModMain"
        mod1.CodeModule.AddFromString(vba_mod_main)
        
        # Module ModSAP
        mod2 = vba_proj.VBComponents.Add(1)
        mod2.Name = "ModSAP"
        mod2.CodeModule.AddFromString(vba_mod_sap)
        
        # Module ModFileMatch
        mod3 = vba_proj.VBComponents.Add(1)
        mod3.Name = "ModFileMatch"
        mod3.CodeModule.AddFromString(vba_mod_file)
        
        # 8. Save as XLSM
        if os.path.exists(out_file):
            os.remove(out_file)
            
        wb.SaveAs(out_file, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
        print(f"Da tao thanh cong: {out_file}")
        
        # Copy to Downloads
        downloads_dir = r"C:\Users\taikt\Downloads"
        if os.path.exists(downloads_dir):
            import shutil
            dest_copy = os.path.join(downloads_dir, "261107_TOOL_ATT_MAIL_MIGO_FAST.xlsm")
            shutil.copy2(out_file, dest_copy)
            print(f"Da dong bo vao Downloads: {dest_copy}")
            
    finally:
        wb.Close(False)
        xl.Quit()

if __name__ == "__main__":
    build_excel_tool()
