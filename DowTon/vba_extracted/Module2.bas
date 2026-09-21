Option Explicit

Sub Run_VL06O()

    Dim wsI As Worksheet
    Dim wsP As Worksheet

    Dim LastRow As Long

    Dim r As Long

    Dim SapGuiAuto As Object
    Dim application As Object
    Dim connection As Object
    Dim session As Object
    Dim Grid As Object

    Dim dicSLI As Object
    Dim dicDOO As Object
    Dim dicSlog As Object
    Dim dicBatchFilled As Object
Dim dicDOBySTO As Object

Dim STO As String


    Dim Mat As String
    Dim Qty As Double
    Dim DOO As Double
    Dim SlogQty As Double
    Dim BatchFilledQty As Double

    Dim Batch As String
    Dim Bin As String

    Dim Key As Variant
    
    Dim dicShip As Object
Dim ArrSTO() As Variant
Dim ArrIndex As Long
Dim ArrReport() As Variant
Dim ArrNeed() As Variant

Dim ReportIdx As Long
Dim NeedIdx As Long

Dim wsT As Worksheet


Set wsT = Sheets("DATA_TEMP")
    Set wsI = Sheets("INPUT")
    Set wsP = Sheets("Pending")

    Set dicSLI = CreateObject("Scripting.Dictionary")
    Set dicDOO = CreateObject("Scripting.Dictionary")
    Set dicSlog = CreateObject("Scripting.Dictionary")
    Set dicBatchFilled = CreateObject("Scripting.Dictionary")
Set dicDOBySTO = CreateObject("Scripting.Dictionary")
Set dicShip = CreateObject("Scripting.Dictionary")

    '=========================
    ' BUILD SLI
    '=========================

    LastRow = wsI.Cells(wsI.Rows.Count, "A").End(xlUp).Row

    For r = 3 To LastRow

        If Trim(wsI.Cells(r, "A").Value) <> "" Then

            Mat = Trim(wsI.Cells(r, "A").Value)
            Qty = CDbl(wsI.Cells(r, "D").Value)

            If Not dicSLI.Exists(Mat) Then
                dicSLI.Add Mat, Qty
            Else
                dicSLI(Mat) = dicSLI(Mat) + Qty
            End If

        End If

    Next r

    '=========================
    ' SAP CONNECT
    '=========================

    Set SapGuiAuto = GetObject("SAPGUI")
    Set application = SapGuiAuto.GetScriptingEngine
    Set connection = application.Children(0)
    Set session = connection.Children(0)

    session.findById("wnd[0]").maximize

    '=========================
    ' OPEN VL06O
    '=========================

    session.findById("wnd[0]/tbar[0]/btn[12]").Press

    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nvl06o"
    session.findById("wnd[0]").sendVKey 0

    session.findById("wnd[0]/usr/btnBUTTON6").Press

    '=========================
    ' SHIPPING POINT
    '=========================

    session.findById("wnd[0]/usr/btn%_IF_VSTEL_%_APP_%-VALU_PUSH").Press

    session.findById("wnd[1]/tbar[0]/btn[16]").Press

    wsP.Range("D3").Copy

    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    '=========================
    ' CLEAR DATE
    '=========================

    session.findById("wnd[0]/usr/btn%_IT_WADAT_%_APP_%-VALU_PUSH").Press

    session.findById("wnd[1]/tbar[0]/btn[16]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    '=========================
    ' MATERIAL
    '=========================

    session.findById("wnd[0]/usr/btn%_IT_MATNR_%_APP_%-VALU_PUSH").Press

    session.findById("wnd[1]/tbar[0]/btn[16]").Press

    wsI.Range("A3:A" & LastRow).Copy

    session.findById("wnd[1]/tbar[0]/btn[24]").Press
    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    '=========================
    ' WBSTK = A
    '=========================

    session.findById("wnd[0]/usr/btn%_IT_WBSTK_%_APP_%-VALU_PUSH").Press

    session.findById("wnd[1]/tbar[0]/btn[16]").Press

    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "A"

    session.findById("wnd[1]/tbar[0]/btn[8]").Press

    '=========================
    ' EXECUTE
    '=========================

    session.findById("wnd[0]/tbar[1]/btn[8]").Press

    '=========================
    ' ITEM VIEW
    '=========================

    session.findById("wnd[0]/tbar[1]/btn[18]").Press

    '=========================
    ' FILTER SHIP TO
    '=========================

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

    '=========================
    ' READ GRID
    '=========================

    Set Grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
ReDim ArrSTO(1 To Grid.RowCount, 1 To 3)

ArrIndex = 1

   For r = 0 To Grid.RowCount - 1

   If r Mod 20 = 0 Then
    Grid.FirstVisibleRow = r
    DoEvents
End If

    Mat = Trim(Grid.GetCellValue(r, "MATNR"))

        Qty = Val(Grid.GetCellValue(r, "S_LGMNG"))

STO = Trim(Grid.GetCellValue(r, "EBELN"))

        Batch = ""
        Bin = ""

        On Error Resume Next
        Batch = Trim(Grid.GetCellValue(r, "S_CHARG"))
        Bin = Trim(Grid.GetCellValue(r, "LGORT"))
        On Error GoTo 0


        'DO OPEN
        If Batch = "" And Bin = "" Then

    If Not dicDOO.Exists(Mat) Then
        dicDOO.Add Mat, Qty
    Else
        dicDOO(Mat) = dicDOO(Mat) + Qty
    End If

    If STO <> "" Then

        If Not dicDOBySTO.Exists(STO) Then
            dicDOBySTO.Add STO, Qty
        Else
            dicDOBySTO(STO) = dicDOBySTO(STO) + Qty
        End If

    End If

    If STO <> "" And Mat <> "" Then

        ArrSTO(ArrIndex, 1) = STO
        ArrSTO(ArrIndex, 2) = Qty
        ArrSTO(ArrIndex, 3) = Mat

        ArrIndex = ArrIndex + 1

    End If

End If

        'SLOG FILLED
        If Batch = "" And Bin <> "" Then

            If Not dicSlog.Exists(Mat) Then
                dicSlog.Add Mat, Qty
            Else
                dicSlog(Mat) = dicSlog(Mat) + Qty
            End If

        End If

        'BATCH + SLOG FILLED
        If Batch <> "" And Bin <> "" Then

            If Not dicBatchFilled.Exists(Mat) Then
                dicBatchFilled.Add Mat, Qty
            Else
                dicBatchFilled(Mat) = dicBatchFilled(Mat) + Qty
            End If

        End If

    Next r




    '=========================
    ' REPORT
    '=========================

    
wsI.Range("H3:M" & wsI.Rows.Count).ClearContents
wsI.Range("AG3:AG" & wsI.Rows.Count).ClearContents
wsT.Range("A:C").ClearContents

    wsI.Range("H2") = "Product"
wsI.Range("I2") = "SLI"
wsI.Range("J2") = "Q'ty DO Open"
wsI.Range("K2") = "Q'ty Slog Filled"
wsI.Range("L2") = "Q'ty Batch+Slog Filled"
wsI.Range("M2") = "Status"

    wsI.Range("AG2") = "Material_ME2N"
wsT.Range("A2") = "STO Number"
wsT.Range("B2") = "DO Qty Open"
wsT.Range("C2") = "Material"

 
ReDim ArrReport(1 To dicSLI.Count, 1 To 6)
ReDim ArrNeed(1 To dicSLI.Count, 1 To 1)

ReportIdx = 1
NeedIdx = 1
If ArrIndex > 1 Then

    wsT.Range("A3").Resize(ArrIndex - 1, 3).Value = ArrSTO

End If

    For Each Key In dicSLI.Keys



        If dicDOO.Exists(Key) Then
            DOO = dicDOO(Key)
        Else
            DOO = 0
        End If

        If dicSlog.Exists(Key) Then
            SlogQty = dicSlog(Key)
        Else
            SlogQty = 0
        End If

        If dicBatchFilled.Exists(Key) Then
            BatchFilledQty = dicBatchFilled(Key)
        Else
            BatchFilledQty = 0
        End If

        ArrReport(ReportIdx, 1) = Key
ArrReport(ReportIdx, 2) = dicSLI(Key)
ArrReport(ReportIdx, 3) = DOO
ArrReport(ReportIdx, 4) = SlogQty
ArrReport(ReportIdx, 5) = BatchFilledQty

If dicSLI(Key) <= DOO Then

    ArrReport(ReportIdx, 6) = "OK TO SHIP"

Else

    ArrReport(ReportIdx, 6) = "NEED STO CHECK"

    ArrNeed(NeedIdx, 1) = Key

    NeedIdx = NeedIdx + 1

End If

ReportIdx = ReportIdx + 1

    Next Key
    If ReportIdx > 1 Then

    wsI.Range("H3").Resize(ReportIdx - 1, 6).Value = ArrReport

End If
'=========================
' SHIP STATUS FROM VL06O
'=========================

For r = 3 To wsI.Cells(wsI.Rows.Count, "H").End(xlUp).Row

    If Trim(wsI.Cells(r, "M").Value) = "OK TO SHIP" Then

        dicShip(Trim(wsI.Cells(r, "H").Value)) = True

    End If

Next r

wsI.Range("F3:F" & LastRow).ClearContents

For r = 3 To LastRow

    Mat = Trim(wsI.Cells(r, "A").Value)

    If dicShip.Exists(Mat) Then

        wsI.Cells(r, "F").Value = "OK TO SHIP"

    End If

Next r
If NeedIdx > 1 Then
    wsT.Range("G3").Resize(NeedIdx - 1, 1).Value = ArrNeed
End If
Sheets("DATA_TEMP").Visible = xlSheetVeryHidden
MsgBox "Completed"
End Sub