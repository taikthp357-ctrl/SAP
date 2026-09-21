Attribute VB_Name = "Module3"
Option Explicit

Sub Run_ME2N()
    Dim wsI As Worksheet, wsP As Worksheet, wsKqSTO As Worksheet
    Dim LastRowProduct As Long, LastRowInput As Long, r As Long
    Dim SapGuiAuto As Object, SapApp As Object, connection As Object, session As Object
    Dim Grid As Object, LayoutGrid As Object
    Dim dicSLI As Object, dicSTOO As Object, dicLastSTO As Object, dicLastDate As Object
    Dim mat As String, STO As String, OrderQty As Double, NotifiedQty As Double, STOOpen As Double
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
        MsgBox "Không có Material trong sheet Pending. Vui lòng ch?n Model tru?c!", vbExclamation, "Thông báo"
        Exit Sub
    End If

    ' Plant & Supplying Plant from Pending
    plantVal = Trim(CStr(wsP.Range("F3").Value))
    If Len(plantVal) = 0 Then plantVal = "3001"
    reswkVal = Trim(CStr(wsP.Range("G3").Value))
    If Len(reswkVal) = 0 Then reswkVal = "3093"

    ' L?y SLI t? sheet INPUT (n?u có)
    Set dicSLI = CreateObject("Scripting.Dictionary")
    If Not wsI Is Nothing Then
        LastRowInput = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row
        For r = 3 To LastRowInput
            mat = Trim(CStr(wsI.Cells(r, "A").Value))
            If Len(mat) > 0 Then
                dicSLI(mat) = dicSLI(mat) + Val(wsI.Cells(r, "D").Value)
            End If
        Next r
    End If

    ' K?t n?i SAP GUI
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Or SapGuiAuto Is Nothing Then
        MsgBox "Không th? k?t n?i SAP GUI!" & vbCrLf & _
               "Vui lòng m? SAP Logon và dang nh?p vào h? th?ng tru?c khi ki?m tra STO.", vbCritical, "L?i k?t n?i SAP GUI"
        Err.Clear
        Exit Sub
    End If

    Set SapApp = SapGuiAuto.GetScriptingEngine
    If SapApp Is Nothing Or SapApp.Children.Count = 0 Then
        MsgBox "SAP GUI dang m? nhung chua có phiên làm vi?c nào!", vbCritical, "L?i k?t n?i SAP GUI"
        Exit Sub
    End If

    Set connection = SapApp.Children(0)
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

    ' Vào giao d?ch ME2N
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
    application.CutCopyMode = False

    ' Supplying Plant (G3 trong Pending)
    session.findById("wnd[0]/usr/btn%_S_RESWK_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    wsP.Range("G3").Copy
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    application.CutCopyMode = False

    ' Material t? Pending (A2:A...)
    session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").Press
    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    wsP.Range("A2:A" & LastRowProduct).Copy
    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    application.CutCopyMode = False

    ' Execute
    session.findById("wnd[0]/tbar[1]/btn[8]").Press

    ' Layout (User Specific Layout)
    On Error Resume Next
    session.findById("wnd[0]/tbar[1]/btn[33]").Press
    session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cmbG51_SCREEN-USPEC_LBOX").key = "X"
    Set LayoutGrid = session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
    LayoutGrid.SetCurrentCell 0, "TEXT"
    LayoutGrid.ClickCurrentCell
    On Error GoTo ErrorHandler

    ' Ð?c k?t qu? t? ALV Grid
    Set Grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    Set dicSTOO = CreateObject("Scripting.Dictionary")
    Set dicLastSTO = CreateObject("Scripting.Dictionary")
    Set dicLastDate = CreateObject("Scripting.Dictionary")

    STOIdx = 0
    If Grid.RowCount > 0 Then
        ReDim ArrME2N(1 To Grid.RowCount, 1 To 5)

        For r = 0 To Grid.RowCount - 1
            mat = Trim(Grid.GetCellValue(r, "EMATN"))
            If mat = "" Then GoTo NextRow

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
            ArrME2N(STOIdx, 4) = mat
            ArrME2N(STOIdx, 5) = DelDate

            If Not dicSTOO.Exists(mat) Then
                dicSTOO(mat) = STOOpen
                dicLastSTO(mat) = STO
                dicLastDate(mat) = DelDate
            Else
                dicSTOO(mat) = dicSTOO(mat) + STOOpen
                If DelDate < dicLastDate(mat) Then
                    dicLastDate(mat) = DelDate
                    dicLastSTO(mat) = STO
                End If
            End If
NextRow:
        Next r
    End If

    ' Xu?t k?t qu? ra sheet KQ_STO
    oldLastSTO = wsKqSTO.Cells(wsKqSTO.Rows.Count, "A").End(xlUp).Row
    If oldLastSTO >= 5 Then wsKqSTO.Range("A5:T" & oldLastSTO).ClearContents
    wsKqSTO.Range("C2").Value = Format(Now, "dd/mm/yyyy hh:nn:ss")

    ' B?ng 1: T?ng h?p STO theo danh sách Pending
    If LastRowProduct >= 2 Then
        ReDim ArrTable1(1 To (LastRowProduct - 1), 1 To 9)
        For r = 2 To LastRowProduct
            mat = Trim(CStr(wsP.Cells(r, "A").Value))
            Table1Idx = r - 1
            ArrTable1(Table1Idx, 1) = Table1Idx
            ArrTable1(Table1Idx, 2) = mat

            sliVal = 0: If dicSLI.Exists(mat) Then sliVal = dicSLI(mat)
            ArrTable1(Table1Idx, 3) = sliVal

            ' Q'ty DO Open: 0 n?u không có liên k?t DO
            ArrTable1(Table1Idx, 4) = 0

            STOOpen = 0: If dicSTOO.Exists(mat) Then STOOpen = dicSTOO(mat)
            ArrTable1(Table1Idx, 5) = STOOpen

            STO = "": If dicLastSTO.Exists(mat) Then STO = dicLastSTO(mat)
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

    ' B?ng 3: Chi ti?t don hàng STO t? ME2N (c?t P:T)
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

    MsgBox "Ki?m tra STO (ME2N) hoàn t?t!" & vbCrLf & _
           "- S? Material ki?m tra: " & (LastRowProduct - 1) & vbCrLf & _
           "- S? don STO m?: " & STOIdx & vbCrLf & _
           "- K?t qu? dã du?c xu?t ra sheet KQ_STO.", vbInformation, "Thành Công"

CleanExit:
    Set LayoutGrid = Nothing
    Set Grid = Nothing
    Set session = Nothing
    Set connection = Nothing
    Set SapApp = Nothing
    Set SapGuiAuto = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "L?i trong quá trình ki?m tra STO: " & Err.Description, vbCritical, "L?i"
    Resume CleanExit
End Sub


