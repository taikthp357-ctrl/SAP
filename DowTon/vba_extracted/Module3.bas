Option Explicit

Sub Run_ME2N()
On Error GoTo ExitHandler


    Dim wsI As Worksheet
    Dim wsP As Worksheet

    Dim SapGuiAuto As Object
    Dim SapApp As Object
    Dim connection As Object
    Dim session As Object

    Dim Grid As Object
    Dim LayoutGrid As Object

    Dim dicSTOO As Object
    Dim dicLastSTO As Object
    Dim dicLastDate As Object

    Dim LastME2NRow As Long
    Dim LastReportRow As Long

    Dim r As Long

    Dim Mat As String
    Dim STO As String

    Dim OrderQty As Double
    Dim NotifiedQty As Double
    Dim STOOpen As Double

    Dim SLI As Double
    Dim DOO As Double

    Dim DelDate As Date
Dim colSTO As Collection

Dim DetailRow As Long

Dim NeedDO As Double
Dim PushQty As Double

Dim Item As Variant

Dim STOInfo As String
Dim Tmp() As String
Dim STOQty As Double

Dim STODataRow As Long
Dim i As Long
Dim j As Long

Dim ArrSTO() As String
Dim TmpSTO As String
Dim A1() As String
Dim A2() As String
Dim sDate As String
Dim ArrDate() As String
Dim dicDOBySTO As Object

Dim LastVLRow As Long

Dim DOExist As Double
Dim STOPushAble As Double
Dim dicSTOByMat As Object
Dim XLApp As application
Dim ArrPush() As Variant
Dim PushIdx As Long
Dim ArrME2N() As Variant
Dim STOIdx As Long
Dim wsT As Worksheet
Dim dicRemain As Object
Dim dicStatus As Object

Dim LastInputRow As Long
Dim BatchQty As Double
Dim StatusTxt As String

Set wsT = Sheets("DATA_TEMP")


Set XLApp = Excel.application
XLApp.ScreenUpdating = False
XLApp.EnableEvents = False
XLApp.Calculation = xlCalculationManual

    Set wsI = Sheets("INPUT")
    Set wsP = Sheets("Pending")

    Set dicSTOO = CreateObject("Scripting.Dictionary")
    Set dicLastSTO = CreateObject("Scripting.Dictionary")
    Set dicLastDate = CreateObject("Scripting.Dictionary")
Set colSTO = New Collection

Set dicDOBySTO = CreateObject("Scripting.Dictionary")
Set dicSTOByMat = CreateObject("Scripting.Dictionary")
Set dicRemain = CreateObject("Scripting.Dictionary")
Set dicStatus = CreateObject("Scripting.Dictionary")
    '=========================
    ' CHECK MATERIAL
    '=========================

    LastME2NRow = wsT.Cells(wsT.Rows.Count, "G").End(xlUp).Row

    If LastME2NRow < 3 Then

        MsgBox "No Material Need STO Check", vbInformation
        Exit Sub

    End If
    LastVLRow = wsT.Cells(wsT.Rows.Count, "A").End(xlUp).Row

For r = 3 To LastVLRow

    STO = Trim(wsT.Cells(r, "A").Value)

    If STO <> "" Then

        If Not dicDOBySTO.Exists(STO) Then

            dicDOBySTO.Add STO, Val(wsT.Cells(r, "B").Value)

        Else

            dicDOBySTO(STO) = dicDOBySTO(STO) + _
                              Val(wsT.Cells(r, "B").Value)

        End If

    End If

Next r


    '=========================
    ' SAP CONNECT
    '=========================

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapApp = SapGuiAuto.GetScriptingEngine
Set connection = SapApp.Children(0)
    Set session = connection.Children(0)

    session.findById("wnd[0]").maximize

    '=========================
    ' OPEN ME2N
    '=========================

    session.findById("wnd[0]/tbar[0]/btn[12]").Press

    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme2n"
    session.findById("wnd[0]").sendVKey 0

    '=========================
    ' CLEAR STO NUMBER
    '=========================

    session.findById("wnd[0]/usr/btn%_EN_EBELN_%_APP_%-VALU_PUSH").Press

    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    '=========================
    ' ALV
    '=========================

    session.findById("wnd[0]/usr/ctxtLISTU").Text = "ALV"

    '=========================
    ' PLANT
    '=========================

    session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").Press

    session.findById("wnd[1]/tbar[0]/btn[16]").Press

    wsP.Range("F3").Copy

    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    '=========================
    ' SUPPLYING PLANT
    '=========================

    session.findById("wnd[0]/usr/btn%_S_RESWK_%_APP_%-VALU_PUSH").Press

    session.findById("wnd[1]/tbar[0]/btn[16]").Press

    wsP.Range("G3").Copy

    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    '=========================
    ' MATERIAL
    '=========================

    session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").Press

    session.findById("wnd[1]/tbar[0]/btn[16]").Press

    wsT.Range("G3:G" & LastME2NRow).Copy

    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    '=========================
    ' EXECUTE
    '=========================

    session.findById("wnd[0]/tbar[1]/btn[8]").Press

    '=========================
    ' USER SPECIFIC LAYOUT
    '=========================

    session.findById("wnd[0]/tbar[1]/btn[33]").Press

    session.findById( _
    "wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/" & _
    "cmbG51_SCREEN-USPEC_LBOX").Key = "X"

    Set LayoutGrid = session.findById( _
    "wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/" & _
    "cntlD500_CONTAINER/shellcont/shell")

    LayoutGrid.SetCurrentCell 0, "TEXT"
    LayoutGrid.ClickCurrentCell

    '=========================
    ' RESULT GRID
    '=========================

    Set Grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
ReDim ArrME2N(1 To Grid.RowCount, 1 To 4)

STOIdx = 1


wsI.Range("N3:T" & wsI.Rows.Count).ClearContents
wsI.Range("AD3:AG" & wsI.Rows.Count).ClearContents

wsI.Range("N2").Value = "Q'ty STO Open"
wsI.Range("O2").Value = "STO Number"
wsI.Range("P2").Value = "Qty Need Increase/Create"
wsI.Range("Q2").Value = "Qty Need DO Create"

wsI.Range("R2").Value = "Material"
wsI.Range("S2").Value = "STO Number"
wsI.Range("T2").Value = "Qty Need Push DO"

wsI.Range("AD2").Value = "STO Number"
wsI.Range("AE2").Value = "STO Open Qty"
wsI.Range("AF2").Value = "Material"
wsI.Range("AG2").Value = "BEDAT"

DetailRow = 3
STODataRow = 3

    '=========================
    ' READ STO DATA
    '=========================

    For r = 0 To Grid.RowCount - 1

        Mat = Trim(Grid.GetCellValue(r, "EMATN"))

        If Mat = "" Then GoTo NextRow

        STO = Trim(Grid.GetCellValue(r, "EBELN"))

        OrderQty = Val(Grid.GetCellValue(r, "MENGE"))

        NotifiedQty = Val(Grid.GetCellValue(r, "LAMNG"))

        STOOpen = OrderQty - NotifiedQty



       If STOOpen <= 0 Then GoTo NextRow

If Trim(Grid.GetCellValue(r, "LOEKZ")) <> "" Then GoTo NextRow

ArrME2N(STOIdx, 1) = STO
ArrME2N(STOIdx, 2) = STOOpen
ArrME2N(STOIdx, 3) = Mat
sDate = Trim(Grid.GetCellValue(r, "BEDAT"))

If sDate <> "" Then

    ArrDate = Split(sDate, ".")

    If UBound(ArrDate) = 2 Then

        DelDate = DateSerial( _
                    CLng(ArrDate(2)), _
                    CLng(ArrDate(1)), _
                    CLng(ArrDate(0)))

    Else

        DelDate = DateSerial(1900, 1, 1)

    End If

Else

    DelDate = DateSerial(1900, 1, 1)

End If

       
ArrME2N(STOIdx, 4) = DelDate

STOIdx = STOIdx + 1

STOInfo = Mat & "|" & _
STO & "|" & _
Format(DelDate, "yyyymmdd") & "|" & _
STOOpen

colSTO.Add STOInfo
If Not dicSTOByMat.Exists(Mat) Then

    Set dicSTOByMat(Mat) = New Collection

End If

dicSTOByMat(Mat).Add STOInfo

        If Not dicSTOO.Exists(Mat) Then

            dicSTOO.Add Mat, STOOpen
            dicLastSTO.Add Mat, STO
            dicLastDate.Add Mat, DelDate

        Else

            dicSTOO(Mat) = dicSTOO(Mat) + STOOpen

            If DelDate < dicLastDate(Mat) Then

                dicLastDate(Mat) = DelDate
                dicLastSTO(Mat) = STO

            End If

        End If

NextRow:

    Next r
    If STOIdx > 1 Then

    wsI.Range("AD3").Resize(STOIdx - 1, 4).Value = ArrME2N

    wsI.Range("AG3:AG" & STOIdx + 1).NumberFormat = "dd/mm/yyyy"

End If
    '=========================
' SORT STO BY MATERIAL + BEDATE
'=========================

If colSTO.Count > 1 Then

    ReDim ArrSTO(1 To colSTO.Count)

    For i = 1 To colSTO.Count
        ArrSTO(i) = colSTO(i)
    Next i

    For i = 1 To UBound(ArrSTO) - 1

        For j = i + 1 To UBound(ArrSTO)


            A1 = Split(ArrSTO(i), "|")
            A2 = Split(ArrSTO(j), "|")

            If A1(0) > A2(0) Then

                TmpSTO = ArrSTO(i)
                ArrSTO(i) = ArrSTO(j)
                ArrSTO(j) = TmpSTO

            ElseIf A1(0) = A2(0) Then

                If CLng(A1(2)) > CLng(A2(2)) Then

                    TmpSTO = ArrSTO(i)
                    ArrSTO(i) = ArrSTO(j)
                    ArrSTO(j) = TmpSTO

                End If

            End If

        Next j

    Next i

    Set colSTO = New Collection

    For i = 1 To UBound(ArrSTO)
        colSTO.Add ArrSTO(i)
    Next i

End If

    '=========================
    ' CLEAR OLD RESULT
    '=========================

    '=========================
    ' UPDATE REPORT
    '=========================

    LastReportRow = wsI.Cells(wsI.Rows.Count, "H").End(xlUp).Row
ReDim ArrPush(1 To LastReportRow * 10, 1 To 3)

PushIdx = 1

    For r = 3 To LastReportRow

        If wsI.Cells(r, "M").Value <> "NEED STO CHECK" Then GoTo NextReport

        Mat = Trim(wsI.Cells(r, "H").Value)

        SLI = Val(wsI.Cells(r, "I").Value)

        DOO = Val(wsI.Cells(r, "J").Value)
If Not dicSTOO.Exists(Mat) Then

    STOOpen = 0

    wsI.Cells(r, "N").Value = 0
    wsI.Cells(r, "O").Value = ""

    wsI.Cells(r, "P").Value = SLI - STOOpen
    wsI.Cells(r, "Q").Value = ""

    wsI.Cells(r, "M").Value = "CREATE STO"

    GoTo NextReport

End If

STOOpen = dicSTOO(Mat)

wsI.Cells(r, "N").Value = STOOpen

If STOOpen = DOO And STOOpen > 0 Then

    wsI.Cells(r, "O").Value = dicLastSTO(Mat)

    wsI.Cells(r, "P").Value = SLI - STOOpen
    wsI.Cells(r, "Q").Value = ""

    wsI.Cells(r, "M").Value = "INCREASE STO"

ElseIf STOOpen > DOO And STOOpen < SLI Then

    wsI.Cells(r, "O").Value = dicLastSTO(Mat)

    wsI.Cells(r, "P").Value = SLI - STOOpen
    wsI.Cells(r, "Q").Value = STOOpen - DOO

    wsI.Cells(r, "M").Value = "PUSH DO + INCREASE STO"

ElseIf STOOpen > DOO And STOOpen >= SLI Then

    wsI.Cells(r, "O").Value = ""

    wsI.Cells(r, "P").Value = ""
    wsI.Cells(r, "Q").Value = SLI - DOO

    wsI.Cells(r, "M").Value = "PUSH STO TO DO"

Else

    wsI.Cells(r, "O").Value = ""
    wsI.Cells(r, "P").Value = ""
    wsI.Cells(r, "Q").Value = ""

End If
       
NextReport:

    Next r


For r = 3 To LastReportRow

    NeedDO = Val(wsI.Cells(r, "Q").Value)
If wsI.Cells(r, "M").Value <> "PUSH DO + INCREASE STO" _
And wsI.Cells(r, "M").Value <> "PUSH STO TO DO" Then
GoTo NextMaterial
End If

    If NeedDO <= 0 Then GoTo NextMaterial

    Mat = Trim(wsI.Cells(r, "H").Value)

   For Each Item In colSTO

        Tmp = Split(Item, "|")

        If Tmp(0) <> Mat Then GoTo NextSTO

        If NeedDO <= 0 Then Exit For

       STOQty = CDbl(Tmp(3))

If dicDOBySTO.Exists(Tmp(1)) Then

    DOExist = dicDOBySTO(Tmp(1))

Else

    DOExist = 0

End If

STOPushAble = STOQty - DOExist

If STOPushAble <= 0 Then GoTo NextSTO

PushQty = STOPushAble


ArrPush(PushIdx, 1) = Mat
ArrPush(PushIdx, 2) = Tmp(1)
ArrPush(PushIdx, 3) = PushQty

PushIdx = PushIdx + 1

NeedDO = NeedDO - STOPushAble

If NeedDO <= 0 Then Exit For

NextSTO:

    Next Item

NextMaterial:

Next r
If PushIdx > 1 Then

    wsI.Range("R3").Resize(PushIdx - 1, 3).Value = ArrPush

End If
'=========================
' SHIP STATUS BY BATCH
'=========================

For r = 3 To LastReportRow

    Mat = Trim(wsI.Cells(r, "H").Value)

    If Mat <> "" Then

        dicStatus(Mat) = Trim(wsI.Cells(r, "M").Value)

        If IsNumeric(wsI.Cells(r, "N").Value) Then
            dicRemain(Mat) = CDbl(wsI.Cells(r, "N").Value)
        Else
            dicRemain(Mat) = 0
        End If

    End If

Next r

LastInputRow = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row

wsI.Range("F3:F" & LastInputRow).ClearContents

For r = 3 To LastInputRow

    Mat = Trim(wsI.Cells(r, "A").Value)

    BatchQty = Val(wsI.Cells(r, "D").Value)

    If dicStatus.Exists(Mat) Then

        StatusTxt = dicStatus(Mat)

        If StatusTxt = "OK TO SHIP" _
        Or StatusTxt = "PUSH STO TO DO" Then

            wsI.Cells(r, "F").Value = "OK TO SHIP"

        Else

            If dicRemain(Mat) >= BatchQty Then

                wsI.Cells(r, "F").Value = "OK TO SHIP"

                dicRemain(Mat) = dicRemain(Mat) - BatchQty

            Else

                wsI.Cells(r, "F").Value = "NOT OK TO SHIP"

            End If

        End If

    Else

        wsI.Cells(r, "F").Value = "NOT OK TO SHIP"

    End If

Next r

XLApp.ScreenUpdating = True
XLApp.EnableEvents = True
XLApp.Calculation = xlCalculationAutomatic

MsgBox "ME2N Completed", vbInformation

Exit Sub

ExitHandler:

XLApp.ScreenUpdating = True
XLApp.EnableEvents = True
XLApp.Calculation = xlCalculationAutomatic

MsgBox Err.Description, vbExclamation

End Sub
