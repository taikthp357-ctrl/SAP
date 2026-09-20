Attribute VB_Name = "ModuleSAP"
Option Explicit
Private mSAPXFilterError As String
#If VBA7 Then
Private Declare PtrSafe Sub SAPX_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub SAPX_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If
Public Function NormalizeHeader(ByVal s As String) As String
    Dim r As String, i As Long, ch As String, mapped As String
    Dim code As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code >= 768 And code <= 879 Then
        ElseIf ch = " " Or ch = "." Or ch = "_" Or ch = "-" Then
        Else
            mapped = StripVNChar(AscW(LCase$(ch)))
            If mapped <> "" Then
                r = r & mapped
            Else
                r = r & ch
            End If
        End If
    Next i
    NormalizeHeader = UCase$(r)
End Function
Private Function StripVNChar(ByVal code As Long) As String
    Select Case code
        Case &HE0, &HE1, &H1EA3, &HE3, &H1EA1, &H103, &H1EB1, &H1EAF, &H1EB3, &H1EB5, &H1EB7, _
             &HE2, &H1EA7, &H1EA5, &H1EA9, &H1EAB, &H1EAD
            StripVNChar = "a"
        Case &HE8, &HE9, &H1EBB, &H1EBD, &H1EB9, &HEA, &H1EC1, &H1EBF, &H1EC3, &H1EC5, &H1EC7
            StripVNChar = "e"
        Case &HEC, &HED, &H1EC9, &H129, &H1ECB
            StripVNChar = "i"
        Case &HF2, &HF3, &H1ECF, &HF5, &H1ECD, &HF4, &H1ED3, &H1ED1, &H1ED5, &H1ED7, &H1ED9, _
             &H1A1, &H1EDD, &H1EDB, &H1EDF, &H1EE1, &H1EE3
            StripVNChar = "o"
        Case &HF9, &HFA, &H1EE7, &H169, &H1EE5, &H1B0, &H1EEB, &H1EE9, &H1EED, &H1EEF, &H1EF1
            StripVNChar = "u"
        Case &H1EF3, &HFD, &H1EF7, &H1EF9, &H1EF5
            StripVNChar = "y"
        Case &H111
            StripVNChar = "d"
        Case Else
            StripVNChar = ""
    End Select
End Function
Public Function FindHeaderCol(ws As Worksheet, ByVal aliases As Variant) As Long
    Dim lastCol As Long, c As Long, h As String, a As Variant
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = NormalizeHeader(ModuleHelper.CleanStr(ws.Cells(1, c).value))
        If h <> "" Then
            For Each a In aliases
                If h = NormalizeHeader(CStr(a)) Then
                    FindHeaderCol = c
                    Exit Function
                End If
            Next a
        End If
    Next c
    FindHeaderCol = 0
End Function
Private Sub MapCol(ws As Worksheet, colMap As Object, ByVal key As String, ByVal aliases As Variant, _
                    ByVal required As Boolean, missingRequired As Collection)
    Dim col As Long
    col = FindHeaderCol(ws, aliases)
    colMap(key) = col
    If required Then
        If col = 0 Then
            If missingRequired Is Nothing Then
                Err.Raise vbObjectError + 2410, "ModuleSAP.MapCol", "missingRequired collection chua duoc khoi tao."
            End If
            missingRequired.Add key
        End If
    End If
End Sub
Public Sub MapPostingColumns(ws As Worksheet, colMap As Object, missingRequired As Collection)
    MapCol ws, colMap, "PRODUCT", Array("Product", "Material"), True, missingRequired
    MapCol ws, colMap, "BATCH", Array("Batch"), True, missingRequired
    MapCol ws, colMap, "TOTALSTK", Array("Total Stk", "Total Stock"), False, missingRequired
    MapCol ws, colMap, "AQTY", Array("A.Qty AUoM", "AQty AUoM", "A Qty AUoM", "A.Qty"), True, missingRequired
    MapCol ws, colMap, "GRDATE", Array("GR Date"), True, missingRequired
    MapCol ws, colMap, "GRTIME", Array("GR Time"), True, missingRequired
    MapCol ws, colMap, "SOURCEBIN", Array("Source Storage Bin", "Source Bin"), False, missingRequired
    MapCol ws, colMap, "DESTBIN", Array("Dest. Bin", "Dest Bin"), False, missingRequired
    MapCol ws, colMap, "TST", Array("TST"), False, missingRequired
    MapCol ws, colMap, "QTYPOSTCHG", Array("QtyPostChg"), False, missingRequired
    MapCol ws, colMap, "DESTBATCH", Array("Dest.Batch", "Dest Batch"), False, missingRequired
    MapCol ws, colMap, "ST", Array("ST"), False, missingRequired
    MapCol ws, colMap, "PROCTYPE", Array("Proc.Type", "Proc Type"), False, missingRequired
End Sub
Public Sub MapOrderColumns(ws As Worksheet, colMap As Object, missingRequired As Collection)
    MapCol ws, colMap, "PRODUCT", Array("Product", "Ma NVL", "Material"), True, missingRequired
    MapCol ws, colMap, "QTY", Array("Qty", "SL oder", "SL Order", "SL dat"), True, missingRequired
    MapCol ws, colMap, "TRAM", Array("Tram", "Text"), True, missingRequired
End Sub
Public Sub ReadPosting(ByRef dataArr As Variant, ByRef colMap As Object, ByRef missingRequired As Collection)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("POSTING")
    Set colMap = CreateObject("Scripting.Dictionary")
    Set missingRequired = New Collection
    MapPostingColumns ws, colMap, missingRequired
    dataArr = Empty
    If missingRequired.Count > 0 Then Exit Sub
    Dim lastRow As Long, lastCol As Long
    lastRow = ModuleHelper.GetLastRow(ws, colMap("PRODUCT"))
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then Exit Sub
    dataArr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).value
End Sub
Public Sub ReadOrder(ByRef dataArr As Variant, ByRef colMap As Object, ByRef missingRequired As Collection)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ORDER")
    Set colMap = CreateObject("Scripting.Dictionary")
    Set missingRequired = New Collection
    MapOrderColumns ws, colMap, missingRequired
    dataArr = Empty
    If missingRequired.Count > 0 Then Exit Sub
    Dim lastRow As Long, lastCol As Long
    lastRow = ModuleHelper.GetLastRow(ws, colMap("PRODUCT"))
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then Exit Sub
    dataArr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).value
End Sub
Public Sub ReadSpecialProducts(ByRef specialDict As Object)
    Set specialDict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("DM_CHU_Y")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim colMap As Object, missing As Collection
    Set colMap = CreateObject("Scripting.Dictionary")
    Set missing = New Collection
    MapCol ws, colMap, "PRODUCT", Array("Product", "Ma NVL", "Material"), True, missing
    MapCol ws, colMap, "LOAI", Array("Loai", "Type"), True, missing
    If missing.Count > 0 Then
        Dim missTxt As String, itm As Variant
        For Each itm In missing
            If missTxt <> "" Then missTxt = missTxt & ", "
            missTxt = missTxt & CStr(itm)
        Next itm
        ModuleLog.WriteLog "CANH BAO", "Sheet DM_CHU_Y thieu header bat buoc, bo qua danh sach dac biet: " & missTxt
        Exit Sub
    End If
    Dim lastRow As Long
    lastRow = ModuleHelper.GetLastRow(ws, colMap("PRODUCT"))
    If lastRow < 2 Then Exit Sub
    Dim i As Long, p As String, l As String
    For i = 2 To lastRow
        p = ModuleHelper.CleanStr(ws.Cells(i, colMap("PRODUCT")).value)
        l = UCase$(ModuleHelper.CleanStr(ws.Cells(i, colMap("LOAI")).value))
        If p <> "" And l <> "" Then
            specialDict(UCase$(p)) = l
        End If
    Next i
End Sub
Public Sub DownloadPOSTOnly()
Dim oldSU As Boolean, oldEv As Boolean, oldCalc As XlCalculation, ok As Boolean
If Not MIGO311_RunPost_V24_ErrorGuard.ToolTryBegin("TAI POSTING") Then Exit Sub
oldSU = Application.ScreenUpdating: oldEv = Application.EnableEvents: oldCalc = Application.Calculation
On Error GoTo EH
Application.ScreenUpdating = False: Application.EnableEvents = False: Application.Calculation = xlCalculationManual
ok = DownloadPostingForOrder(True)
Application.ScreenUpdating = oldSU: Application.EnableEvents = oldEv: Application.Calculation = oldCalc
MIGO311_RunPost_V24_ErrorGuard.ToolEnd
Exit Sub
EH:
Application.ScreenUpdating = oldSU: Application.EnableEvents = oldEv: Application.Calculation = oldCalc
MIGO311_RunPost_V24_ErrorGuard.ToolEnd
MsgBox "TAI POSTING loi " & CStr(Err.Number) & ": " & Err.Description, vbCritical, "TAI POSTING"
End Sub
Public Function DownloadPostingForOrder(Optional ByVal showMessage As Boolean = True) As Boolean
Dim ses As Object, products As Object, rows As Collection, p As Variant, grid As Object
Dim cProd As String, cQty As String, cBatch As String, cBin As String, cTst As String, colDate As String, cTime As String, cTotal As String
Dim cDestBin As String, cQtyPost As String, cDestBatch As String, cST As String, cProcType As String, cAvlBU As String
Dim r As Long, n As Long, rowCount As Long, lastP As String, pv As String, qv As String, bv As String, tv As String, oneRow As Variant
Dim ws As Worksheet, oldLast As Long, outArr() As Variant, i As Long, msg As String, stage As String
On Error GoTo Fail
stage = "READ_ORDER"
Set products = POST_OrderProducts()
If products.Count = 0 Then Err.Raise vbObjectError + 2900, , "ORDER khong co Product hop le."
Set rows = New Collection
stage = "OPEN_POST"
Set ses = SAPX_GetSessionForTcode(ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST"))
If ses Is Nothing Then Err.Raise vbObjectError + 2901, , "Khong ket noi duoc SAP session."
stage = "OPEN_ADVANCED_SEARCH"
Application.StatusBar = "POSTING: dang gui " & products.Count & " Product vao Advanced Search..."
If Not POST_OpenAdvancedProductResult(ses, products, grid, msg) Then Err.Raise vbObjectError + 2902, , msg
If grid Is Nothing Then Err.Raise vbObjectError + 2903, , "Advanced Search khong tra result grid."
stage = "MAP_COLUMNS"
cProd = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("POST_COL_PRODUCT", ""), "MATNR,PRODUCT,MATERIAL", True)
cQty = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("POST_COL_AQTY", ""), "AQUAN,AQTY,A.QTY AUOM,AVAILABLE QTY,AVAILABLE QUANTITY,AVLQTY", True)
cBatch = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("POST_COL_BATCH", ""), "CHARG,BATCH", True)
cBin = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("POST_COL_SOURCE_BIN", ""), "LGPLA,SOURCE STORAGE BIN,SOURCE BIN,STORAGE BIN", True)
cTst = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("POST_COL_TST", ""), "CAT,TST,STOCK TYPE,ST", True)
colDate = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("POST_COL_GR_DATE", ""), "WDATU,GR DATE,GOODS RECEIPT DATE", True)
cTime = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("POST_COL_GR_TIME", ""), "WTIME,GR TIME,GOODS RECEIPT TIME", True)
cTotal = SAPX_ResolveColumn(grid, "", "QUAN,TOTAL STK,TOTAL STOCK", False)
cDestBin = SAPX_ResolveColumn(grid, "", "NLPLA,DEST. BIN,DEST BIN,DESTINATION STORAGE BIN", False)
cQtyPost = SAPX_ResolveColumn(grid, "", "POST_QUAN,QTYPOSTCHG,QTY POST CHG", False)
cDestBatch = SAPX_ResolveColumn(grid, "", "TO_CHARG,DEST.BATCH,DEST BATCH", False)
cST = SAPX_ResolveColumn(grid, "", "CAT,ST", True)
cProcType = SAPX_ResolveColumn(grid, "", "PROCTY,PROC. TYPE,PROCESS TYPE", False)
cAvlBU = SAPX_ResolveColumn(grid, "", "AVAIL_QUAN,AVLQTYBUOM,AVAILABLE QTY BUOM", False)
stage = "READ_RESULT_GRID"
rowCount = CLng(grid.rowCount): lastP = ""
Application.StatusBar = "POSTING: dang doc " & rowCount & " dong ton..."
For r = 0 To rowCount - 1
    pv = SAPX_GridText(grid, r, cProd)
    If pv <> "" Then lastP = pv Else pv = lastP
    If products.Exists(Trim$(pv)) Then
        qv = SAPX_GridText(grid, r, cQty): bv = SAPX_GridText(grid, r, cBatch): tv = SAPX_GridText(grid, r, cTst)
        If POST_ParseIntegerQty(qv) > 0 And Trim$(bv) <> "" Then
            oneRow = Array(Trim$(pv), POST_ParseIntegerQty(qv), bv, SAPX_GridText(grid, r, cBin), tv, _
                IIf(cTotal <> "", POST_ParseIntegerQty(SAPX_GridText(grid, r, cTotal)), POST_ParseIntegerQty(qv)), _
                SAPX_GridText(grid, r, cDestBin), POST_ParseIntegerQty(SAPX_GridText(grid, r, cQtyPost)), _
                SAPX_GridText(grid, r, cDestBatch), SAPX_GridText(grid, r, cST), _
                SAPX_GridText(grid, r, colDate), SAPX_GridText(grid, r, cTime), _
                SAPX_GridText(grid, r, cProcType), _
                IIf(cAvlBU <> "", POST_ParseIntegerQty(SAPX_GridText(grid, r, cAvlBU)), POST_ParseIntegerQty(qv)), "", "", 0)
            rows.Add oneRow
        End If
    End If
Next r
stage = "COMMIT_POSTING"
Set ws = ThisWorkbook.Worksheets("POSTING")
oldLast = POST_LastDataRow(ws, 17)
If oldLast < 2 Then oldLast = 2
If rows.Count = 0 Then
    ws.Range("A2:Q" & CStr(oldLast)).ClearContents
    DownloadPostingForOrder = True
    Application.StatusBar = "POSTING: SAP tra 0 dong ton cho ORDER."
    If showMessage Then MsgBox "SAP da tim xong nhung khong co dong ton hop le cho ORDER. POSTING cu da duoc xoa de tranh dung ton cu.", vbInformation, "TAI POSTING"
    Exit Function
End If
ReDim outArr(1 To rows.Count, 1 To 17)
For i = 1 To rows.Count
    oneRow = rows(i)
    For n = 1 To 17: outArr(i, n) = oneRow(n - 1): Next n
Next i
ws.Range("A2:Q" & CStr(oldLast)).ClearContents
ws.Range("C2:C" & CStr(rows.Count + 1)).NumberFormat = "@"
ws.Range("I2:I" & CStr(rows.Count + 1)).NumberFormat = "@"
ws.Range("A2").Resize(rows.Count, 17).Value2 = outArr
ws.Range("B2:B" & CStr(rows.Count + 1)).NumberFormat = "#,##0"
ws.Range("F2:F" & CStr(rows.Count + 1)).NumberFormat = "#,##0"
ws.Range("H2:H" & CStr(rows.Count + 1)).NumberFormat = "#,##0"
ws.Range("N2:N" & CStr(rows.Count + 1)).NumberFormat = "#,##0"
ws.Range("Q2:Q" & CStr(rows.Count + 1)).NumberFormat = "#,##0"
DownloadPostingForOrder = True
Application.StatusBar = "POSTING: da tai " & rows.Count & " dong."
If showMessage Then MsgBox "Da tai " & rows.Count & " dong POSTING.", vbInformation, "TAI POSTING"
Exit Function
Fail:
Application.StatusBar = False
ModuleLog.WriteLog "LOI", "TAI POSTING [" & stage & "]: " & Err.Description
If showMessage Then MsgBox "TAI POSTING loi tai " & stage & ":" & vbCrLf & Err.Description & vbCrLf & vbCrLf & "POSTING cu duoc giu nguyen.", vbCritical, "TAI POSTING"
End Function
Private Function POST_LastDataRow(ByVal ws As Worksheet, Optional ByVal maxCol As Long = 17) As Long
Dim c As Long, candidate As Long, lastRow As Long
lastRow = 1
For c = 1 To maxCol
    candidate = ws.Cells(ws.rows.Count, c).End(xlUp).Row
    If candidate > lastRow Then lastRow = candidate
Next c
POST_LastDataRow = lastRow
End Function
Private Function POST_ParseIntegerQty(ByVal rawValue As Variant) As Double
Dim s As String, negativeQty As Boolean
If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
If IsNumeric(rawValue) And VarType(rawValue) <> vbString Then
    POST_ParseIntegerQty = Fix(CDbl(rawValue))
    Exit Function
End If
s = Trim$(CStr(rawValue & ""))
If s = "" Then Exit Function
negativeQty = (Left$(s, 1) = "-" Or (Left$(s, 1) = "(" And Right$(s, 1) = ")"))
s = Replace(s, "-", "")
s = Replace(s, "(", "")
s = Replace(s, ")", "")
s = Replace(s, " ", "")
s = Replace(s, ChrW(160), "")
s = Replace(s, ".", "")
s = Replace(s, ",", "")
If s = "" Or Not IsNumeric(s) Then Exit Function
POST_ParseIntegerQty = CDbl(s)
If negativeQty Then POST_ParseIntegerQty = -POST_ParseIntegerQty
End Function
Private Function POST_OrderProducts() As Object
Dim d As Object, ws As Worksheet, c As Long, lr As Long, a As Variant, i As Long, s As String
Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
Set ws = ThisWorkbook.Worksheets("ORDER")
c = FindHeaderCol(ws, Array("Product", "Material", "Ma NVL"))
If c = 0 Then Set POST_OrderProducts = d: Exit Function
lr = ws.Cells(ws.rows.Count, c).End(xlUp).Row
If lr < 2 Then Set POST_OrderProducts = d: Exit Function
a = ws.Range(ws.Cells(2, c), ws.Cells(lr, c)).Value2
For i = 1 To UBound(a, 1): s = Trim$(CStr(a(i, 1) & "")): If s <> "" Then d(s) = True
Next i
Set POST_OrderProducts = d
End Function
Private Function POST_OpenAdvancedProductResult(ByVal ses As Object, ByVal products As Object, ByRef grid As Object, ByRef msg As String) As Boolean
Const ID_ADV_TOGGLE As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/btnGV_BUTTON_TEXT"
Const ID_ADV_RESET As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_AS_BUTTONS:/SCMB/SAPLSERVICES:1000/btnCMD_CLEAR_ADVANCED"
Const ID_ADV_START As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_AS_BUTTONS:/SCMB/SAPLSERVICES:1000/btnCMD_START_ADVANCED"
Const ID_PRODUCT_MULTI As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_SELECTION_SCREEN:/SCWM/SAPLUI_POST:2500/btn%_P_MATNR_%_APP_%-VALU_PUSH"
Const ID_OWNER_LOW As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_SELECTION_SCREEN:/SCWM/SAPLUI_POST:2500/ctxtP_OWNER-LOW"
Const ID_STOCK_GRID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/subSUB_OIP_1_DATA:/SCWM/SAPLUI_POST:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell"
Dim ctl As Object, pop As Object, ws As Worksheet, c As Long, lr As Long
Dim ownerValue As String, statusText As String, statusType As String
On Error GoTo Fail
msg = "": Set grid = Nothing
On Error Resume Next
Set pop = ses.findById("wnd[1]", False)
If Not pop Is Nothing Then
    Set ctl = ses.findById("wnd[1]/tbar[0]/btn[12]", False)
    If Not ctl Is Nothing Then ctl.press Else pop.Close
End If
On Error GoTo Fail
SAPX_Navigate ses, ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST")
Set ctl = SAPX_FindById(ses, ID_OWNER_LOW)
If ctl Is Nothing Then
    Set ctl = SAPX_FindById(ses, ID_ADV_TOGGLE)
    If ctl Is Nothing Then Err.Raise vbObjectError + 2950, , "Khong tim thay nut Open Advanced Search."
    ctl.press
    If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 2951, , "SAP Busy khi mo Advanced Search."
End If
Set ctl = SAPX_FindById(ses, ID_ADV_RESET)
If Not ctl Is Nothing Then ctl.press: SAPX_WaitIdle ses, 10
ownerValue = ModuleHelper.ToolConfig("POST_OWNER", "3093")
Set ctl = SAPX_FindById(ses, ID_OWNER_LOW)
If ctl Is Nothing Then Err.Raise vbObjectError + 2952, , "Khong tim thay Owner trong Advanced Search sau khi mo lai."
SAPX_SetValue ctl, ownerValue
If StrComp(SAPX_ControlValue(ctl), ownerValue, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 2953, , "Owner readback khong khop " & ownerValue
Set ctl = SAPX_FindById(ses, ID_PRODUCT_MULTI)
If ctl Is Nothing Then Err.Raise vbObjectError + 2954, , "Khong tim thay Multiple Selection cua Product."
ctl.press
If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 2955, , "SAP Busy khi mo Product list."
Set pop = SAPX_FindById(ses, "wnd[1]")
If pop Is Nothing Then Err.Raise vbObjectError + 2956, , "Khong mo duoc popup Product list."
On Error Resume Next
Set ctl = ses.findById("wnd[1]/tbar[0]/btn[16]", False)
If Not ctl Is Nothing Then ctl.press
On Error GoTo Fail
Set ws = ThisWorkbook.Worksheets("ORDER")
c = FindHeaderCol(ws, Array("Product", "Material", "Ma NVL"))
If c = 0 Then Err.Raise vbObjectError + 2957, , "ORDER khong co cot Product."
lr = ws.Cells(ws.rows.Count, c).End(xlUp).Row
If lr < 2 Then Err.Raise vbObjectError + 2957, , "ORDER khong co list Product de copy."
ws.Range(ws.Cells(2, c), ws.Cells(lr, c)).Copy
On Error Resume Next
Set ctl = ses.findById("wnd[1]/tbar[0]/btn[24]", False)
On Error GoTo Fail
If ctl Is Nothing Then Err.Raise vbObjectError + 2958, , "Popup Product khong co nut Paste from Clipboard."
ctl.press
SAPX_WaitIdle ses, 10
Set ctl = SAPX_FindById(ses, "wnd[1]/tbar[0]/btn[8]")
If ctl Is Nothing Then Err.Raise vbObjectError + 2959, , "Popup Product khong co nut OK."
ctl.press
If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 2960, , "SAP Busy khi nhan list Product."
Set ctl = SAPX_FindById(ses, ID_ADV_START)
If ctl Is Nothing Then Err.Raise vbObjectError + 2961, , "Khong tim thay nut Advanced Search."
ctl.press
If Not SAPX_WaitIdle(ses, 60) Then Err.Raise vbObjectError + 2962, , "SAP Busy qua lau khi tai " & products.Count & " Product."
statusText = SAPX_StatusText(ses): statusType = SAPX_StatusType(ses)
If statusType = "E" Or statusType = "A" Then Err.Raise vbObjectError + 2963, , "SAP: " & statusText
Set grid = SAPX_FindById(ses, ID_STOCK_GRID)
If grid Is Nothing Then Set grid = SAPX_WaitGrid(ses, 25)
If grid Is Nothing Then Err.Raise vbObjectError + 2964, , "Khong tim thay stock grid sau Advanced Search."
POST_OpenAdvancedProductResult = True
Application.CutCopyMode = False
Exit Function
Fail:
msg = "Advanced Search: " & Err.Description
Application.CutCopyMode = False
End Function
Private Function POST_OpenProductResult(ByVal ses As Object, ByVal product As String, ByRef grid As Object, ByRef msg As String) As Boolean
Dim ctl As Object, crit As Object, statusText As String, statusType As String
On Error GoTo Fail
If Not SAPX_TcodeIs(ses, ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST")) Then SAPX_Navigate ses, ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST")
POST_FindProductControls ses, crit, ctl
If ctl Is Nothing Then
    SAPX_Navigate ses, ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST")
    POST_FindProductControls ses, crit, ctl
End If
If ctl Is Nothing Then Err.Raise vbObjectError + 2910, , "Khong tim thay o Product tren /SCWM/POST theo layout hien tai."
If Not POST_EnsureProductCriterion(crit) Then Err.Raise vbObjectError + 2914, , "Criterion /SCWM/POST khong o che do Product."
POST_SetOptionalOwner ses
SAPX_SetValue ctl, product
If StrComp(Trim$(SAPX_ControlValue(ctl)), Trim$(product), vbTextCompare) <> 0 Then Err.Raise vbObjectError + 2911, , "Product readback khong khop: " & product
On Error Resume Next
ses.findById(ModuleHelper.ToolConfig("POST_EXECUTE_ID", "wnd[0]/tbar[1]/btn[8]"), False).press
If Err.Number <> 0 Then Err.Clear: ses.findById("wnd[0]").sendVKey 8
On Error GoTo Fail
If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 2912, , "SAP Busy qua lau khi tai " & product
statusText = SAPX_StatusText(ses): statusType = SAPX_StatusType(ses)
If statusType = "E" Or statusType = "A" Then Err.Raise vbObjectError + 2915, , "SAP: " & statusText
Set grid = SAPX_WaitGrid(ses, 6)
If grid Is Nothing Then
    If statusText <> "" Then
        POST_OpenProductResult = True
        msg = "Product " & product & ": " & statusText
        Exit Function
    End If
    Set grid = SAPX_WaitGrid(ses, 14)
End If
If grid Is Nothing Then Err.Raise vbObjectError + 2913, , "Khong tim thay grid ket qua /SCWM/POST cho " & product & "."
POST_OpenProductResult = True
Exit Function
Fail:
msg = "Product " & product & ": " & Err.Description
End Function
Private Function POST_ObjectText(ByVal obj As Object) As String
On Error Resume Next
POST_ObjectText = Trim$(CStr(obj.text))
If POST_ObjectText = "" Then POST_ObjectText = Trim$(CStr(obj.value))
If POST_ObjectText = "" Then POST_ObjectText = Trim$(CStr(obj.key))
On Error GoTo 0
End Function
Private Function POST_FindProductComboRec(ByVal obj As Object, ByVal depth As Long) As Object
Dim typ As String, txt As String, i As Long, ch As Object, f As Object
If obj Is Nothing Or depth > 16 Then Exit Function
On Error Resume Next: typ = UCase$(CStr(obj.Type)): txt = UCase$(POST_ObjectText(obj)): On Error GoTo 0
If InStr(typ, "COMBOBOX") > 0 And InStr(txt, "PRODUCT") > 0 Then Set POST_FindProductComboRec = obj: Exit Function
On Error Resume Next
For i = 0 To obj.Children.Count - 1
    Set ch = SAPX_CollectionObjectAt(obj.Children, i): Set f = POST_FindProductComboRec(ch, depth + 1)
    If Not f Is Nothing Then Set POST_FindProductComboRec = f: Exit Function
Next i
On Error GoTo 0
End Function
Private Sub POST_ScoreProductInputRec(ByVal obj As Object, ByVal combo As Object, ByRef best As Object, ByRef bestScore As Double, ByVal depth As Long)
Dim typ As String, idt As String, l As Double, t As Double, cr As Double, ct As Double, dx As Double, dy As Double, score As Double, i As Long, ch As Object
If obj Is Nothing Or combo Is Nothing Or depth > 16 Then Exit Sub
On Error Resume Next: typ = UCase$(CStr(obj.Type)): idt = UCase$(CStr(obj.id)): On Error GoTo 0
If InStr(typ, "TEXTFIELD") > 0 And InStr(idt, "OKCD") = 0 Then
    On Error Resume Next: l = CDbl(obj.ScreenLeft): t = CDbl(obj.ScreenTop): cr = CDbl(combo.ScreenLeft) + CDbl(combo.Width): ct = CDbl(combo.ScreenTop): On Error GoTo 0
    dx = l - cr: dy = Abs(t - ct)
    If dx >= -5 And dx <= 600 And dy <= 25 Then
        score = 100000# - dx - dy * 20#
        If InStr(idt, "PRODUCT") > 0 Or InStr(idt, "MATNR") > 0 Or InStr(idt, "MATID") > 0 Then score = score + 10000#
        If score > bestScore Then bestScore = score: Set best = obj
    End If
End If
On Error Resume Next
For i = 0 To obj.Children.Count - 1: Set ch = SAPX_CollectionObjectAt(obj.Children, i): POST_ScoreProductInputRec ch, combo, best, bestScore, depth + 1: Next i
On Error GoTo 0
End Sub
Private Sub POST_FindProductControls(ByVal ses As Object, ByRef crit As Object, ByRef productCtl As Object)
Dim root As Object, score As Double
Set crit = SAPX_FindById(ses, ModuleHelper.ToolConfig("POST_CRITERION_ID", ""))
Set productCtl = SAPX_FindById(ses, ModuleHelper.ToolConfig("POST_PRODUCT_ID", ""))
On Error Resume Next: Set root = ses.findById("wnd[0]", False): On Error GoTo 0
If crit Is Nothing Then Set crit = POST_FindProductComboRec(root, 0)
If productCtl Is Nothing And Not crit Is Nothing Then POST_ScoreProductInputRec root, crit, productCtl, score, 0
If productCtl Is Nothing Then Set productCtl = SAPX_FindByTokens(ses, "ASPQ_STOCK-MATNR|MATNR|MATID|PRODUCT", True)
End Sub
Private Function POST_EnsureProductCriterion(ByVal ctl As Object) As Boolean
Dim txt As String
If ctl Is Nothing Then POST_EnsureProductCriterion = True: Exit Function
txt = UCase$(POST_ObjectText(ctl))
If InStr(txt, "PRODUCT") > 0 Then POST_EnsureProductCriterion = True: Exit Function
POST_SetCriterion ctl
txt = UCase$(POST_ObjectText(ctl))
POST_EnsureProductCriterion = (InStr(txt, "PRODUCT") > 0 Or UCase$(Trim$(SAPX_ControlValue(ctl))) = "PROD")
End Function
Private Sub POST_SetOptionalOwner(ByVal ses As Object)
Dim ctl As Object, v As String
v = ModuleHelper.ToolConfig("POST_OWNER", "3093")
If Trim$(v) = "" Then Exit Sub
Set ctl = SAPX_FindByTokens(ses, "OWNER|ENTITLED", True)
If Not ctl Is Nothing Then SAPX_SetValue ctl, v
End Sub
Private Function SAPX_StatusText(ByVal ses As Object) As String
On Error Resume Next: SAPX_StatusText = Trim$(CStr(ses.findById("wnd[0]/sbar", False).text)): On Error GoTo 0
End Function
Private Function SAPX_StatusType(ByVal ses As Object) As String
On Error Resume Next: SAPX_StatusType = UCase$(Trim$(CStr(ses.findById("wnd[0]/sbar", False).messageType))): On Error GoTo 0
End Function
Private Sub POST_SetCriterion(ByVal ctl As Object)
On Error Resume Next
ctl.key = "PROD"
If Err.Number <> 0 Then Err.Clear: ctl.text = "Product"
If Err.Number <> 0 Then Err.Clear: ctl.key = "1"
On Error GoTo 0
End Sub
Public Sub CheckOBTreoVL06O()
Dim oldSU As Boolean, oldEv As Boolean, oldCalc As XlCalculation
If Not MIGO311_RunPost_V24_ErrorGuard.ToolTryBegin("CHECK OB TREO") Then Exit Sub
oldSU = Application.ScreenUpdating: oldEv = Application.EnableEvents: oldCalc = Application.Calculation
On Error GoTo EH
Application.ScreenUpdating = False: Application.EnableEvents = False: Application.Calculation = xlCalculationManual
    VL6_RunPendingReport True
Application.ScreenUpdating = oldSU: Application.EnableEvents = oldEv: Application.Calculation = oldCalc
MIGO311_RunPost_V24_ErrorGuard.ToolEnd
Exit Sub
EH:
Application.ScreenUpdating = oldSU: Application.EnableEvents = oldEv: Application.Calculation = oldCalc: Application.StatusBar = False
MIGO311_RunPost_V24_ErrorGuard.ToolEnd
MsgBox "CHECK OB TREO loi " & CStr(Err.Number) & ": " & Err.Description, vbCritical, "VL06O"
End Sub
Public Function TestVL06O_NoMessage() As Boolean
On Error GoTo Failed
VL6_RunPendingReport False
TestVL06O_NoMessage = True
Exit Function
Failed:
ModuleLog.WriteLog "LOI", "VL06O TEST: " & Err.Description
End Function
Public Function TestVL06O_ControlDump() As String
Dim ses As Object, root As Object, dumpText As String
On Error GoTo Failed
Set ses = SAPX_GetSessionForTcode("VL06O")
Set root = ses.findById("wnd[0]", False)
SAPX_DumpControlIds root, dumpText, 0
TestVL06O_ControlDump = dumpText
Exit Function
Failed:
TestVL06O_ControlDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function
Public Function TestPOST_ControlDump() As String
Dim ses As Object, root As Object, dumpText As String
On Error GoTo Failed
Set ses = SAPX_GetSessionForTcode(ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST"))
Set root = ses.findById("wnd[0]", False)
SAPX_DumpControlIds root, dumpText, 0
TestPOST_ControlDump = dumpText
Exit Function
Failed:
TestPOST_ControlDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function
Public Function TestPOST_ColumnDump() As String
Const ID_STOCK_GRID As String = "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/subSUB_OIP_1_DATA:/SCWM/SAPLUI_POST:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell"
Dim ses As Object, grid As Object, ids As Collection, id As Variant, dumpText As String
On Error GoTo Failed
Set ses = SAPX_GetSessionForTcode(ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST"))
Set grid = SAPX_FindById(ses, ID_STOCK_GRID)
If grid Is Nothing Then Err.Raise vbObjectError + 2990, , "Khong tim thay POST stock grid."
Set ids = SAPX_ColumnIds(grid)
For Each id In ids
    dumpText = dumpText & CStr(id) & " | " & SAPX_ColumnCaption(grid, CStr(id)) & vbLf
Next id
TestPOST_ColumnDump = dumpText
Exit Function
Failed:
TestPOST_ColumnDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function
Public Function TestPOST_AdvancedDump() As String
Dim ses As Object, root As Object, btn As Object, dumpText As String
On Error GoTo Failed
Set ses = SAPX_GetSessionForTcode(ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST"))
Set btn = SAPX_FindByTokens(ses, "GV_BUTTON_TEXT|ADVANCED SEARCH", False)
If btn Is Nothing Then Err.Raise vbObjectError + 2991, , "Khong tim thay Open Advanced Search."
btn.press
SAPX_WaitIdle ses, 10
On Error Resume Next
Set root = ses.findById("wnd[1]", False)
If root Is Nothing Then Set root = ses.findById("wnd[0]", False)
On Error GoTo Failed
SAPX_DumpControlIds root, dumpText, 0
TestPOST_AdvancedDump = dumpText
Exit Function
Failed:
TestPOST_AdvancedDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function
Public Function TestPOST_ProductMultiDump() As String
Dim ses As Object, root As Object, btn As Object, dumpText As String
On Error GoTo Failed
Set ses = SAPX_GetSessionForTcode(ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST"))
Set btn = SAPX_FindByTokens(ses, "%_P_MATNR_%_APP_%-VALU_PUSH", False)
If btn Is Nothing Then Err.Raise vbObjectError + 2992, , "Khong tim thay Multiple Selection Product."
btn.press
SAPX_WaitIdle ses, 10
Set root = ses.findById("wnd[1]", False)
If root Is Nothing Then Err.Raise vbObjectError + 2993, , "Khong mo duoc popup Product list."
SAPX_DumpControlIds root, dumpText, 0
TestPOST_ProductMultiDump = dumpText
Exit Function
Failed:
TestPOST_ProductMultiDump = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function
Private Sub SAPX_DumpControlIds(ByVal obj As Object, ByRef dumpText As String, ByVal depth As Long)
Dim i As Long, ch As Object, idText As String, typeText As String, labelText As String
If obj Is Nothing Or depth > 14 Or Len(dumpText) > 30000 Then Exit Sub
On Error Resume Next
idText = CStr(obj.id): typeText = CStr(obj.Type): labelText = SAPX_ControlText(obj)
On Error GoTo 0
If InStr(1, UCase$(idText & " " & labelText), "DATE", vbTextCompare) > 0 Or _
   InStr(1, UCase$(idText & " " & labelText), "WADAT", vbTextCompare) > 0 Or _
   InStr(1, UCase$(idText & " " & labelText), "GI", vbTextCompare) > 0 Or _
   InStr(1, UCase$(typeText), "TEXTFIELD", vbTextCompare) > 0 Or _
   InStr(1, UCase$(typeText), "BUTTON", vbTextCompare) > 0 Or _
   InStr(1, UCase$(typeText), "RADIO", vbTextCompare) > 0 Or _
   InStr(1, UCase$(typeText), "LABEL", vbTextCompare) > 0 Or _
   InStr(1, UCase$(typeText), "SHELL", vbTextCompare) > 0 Or _
   InStr(1, UCase$(typeText), "GRID", vbTextCompare) > 0 Then
    dumpText = dumpText & idText & " | " & typeText & " | " & labelText & vbLf
End If
On Error Resume Next
For i = 0 To obj.Children.Count - 1
    Set ch = SAPX_CollectionObjectAt(obj.Children, i)
    SAPX_DumpControlIds ch, dumpText, depth + 1
Next i
Err.Clear
On Error GoTo 0
End Sub
Private Sub VL6_RunPendingReport(Optional ByVal showMessage As Boolean = True)
Dim ses As Object, grid As Object, ctl As Object, fromD As Date, toD As Date
Dim cGM As String, cDel As String, cMat As String, cQty As String, cBatch As String, cSloc As String, cShip As String, colDate As String, cTime As String, cBy As String, cShpt As String
Dim shipValue As String, slocValue As String, r As Long, rowCount As Long, outCount As Long, arr() As Variant, shipPrefilled As Boolean
Dim lastDel As String, lastShip As String, lastSloc As String, lastCreated As String, vDel As String, vShip As String, vSloc As String, vMat As String, vQty As String, vCreated As String, d As Date
Dim ws As Worksheet, oldLast As Long, clearLast As Long, oldCount As Long, unique As Object, i As Long, finalArr() As Variant, sortedData As Variant, j As Long
Dim failNo As Long, failText As String, shipFilterOK As Boolean, slocFilterOK As Boolean
On Error GoTo Fail
shipValue = ModuleHelper.ToolConfig("VL06O_SHIP_TO", "3093"): slocValue = ModuleHelper.ToolConfig("VL06O_SLOC", "1005")
fromD = DateAdd("m", -ModuleHelper.ToolConfigLong("VL06O_LOOKBACK_MONTHS", 1, 1, 12), Date): toD = Date
Set ses = SAPX_GetSessionForTcode("VL06O")
If ses Is Nothing Then Err.Raise vbObjectError + 2920, , "Khong ket noi duoc SAP session."
Set grid = SAPX_FindLargestGrid(ses)
If Not grid Is Nothing Then
    Application.StatusBar = "VL06O: dang Shift+F8 lam moi tab ket qua hien tai..."
    If Not SAPX_RefreshVL06OAndCatchReady(ses, grid, 30) Then Err.Raise vbObjectError + 2924, , "Khong bat duoc ket qua moi sau Shift+F8 VL06O."
    GoTo ResultReady
End If
Application.StatusBar = "VL06O: dang tai su dung man hinh chon hien tai..."
VL6_OpenSelection ses
shipPrefilled = VL6_SetSelectionShipTo(ses, shipValue)
If Not shipPrefilled Then Err.Raise vbObjectError + 2942, , "Khong tim thay/khong dien duoc Ship-to=" & shipValue & " tai man hinh chon VL06O. Luong moi khong loc Ship-to lan hai tai Item View."
Set ctl = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_DATE_FROM_ID", "wnd[0]/usr/ctxtIT_WADAT-LOW"))
If ctl Is Nothing Then Set ctl = SAPX_FindByTokens(ses, "WADAT-LOW|WADAT_LOW|WADATLOW", True)
If ctl Is Nothing Then Err.Raise vbObjectError + 2921, , "Thieu Planned GI Date From."
SAPX_SetValue ctl, Format$(fromD, "dd.mm.yyyy")
Set ctl = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_DATE_TO_ID", "wnd[0]/usr/ctxtIT_WADAT-HIGH"))
If ctl Is Nothing Then Set ctl = SAPX_FindByTokens(ses, "WADAT-HIGH|WADAT_HIGH|WADATHIGH", True)
If ctl Is Nothing Then Err.Raise vbObjectError + 2922, , "Thieu Planned GI Date To."
SAPX_SetValue ctl, Format$(toD, "dd.mm.yyyy")
Set ctl = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_GM_STATUS_ID", "wnd[0]/usr/ctxtIT_WBSTK-LOW"))
If ctl Is Nothing Then Set ctl = SAPX_FindByTokens(ses, "WBSTK-LOW|WBSTK_LOW|WBSTKLOW|WBSTA-LOW", True)
If ctl Is Nothing Then Err.Raise vbObjectError + 2923, , "Thieu Goods Movement Status."
SAPX_SetValue ctl, "A"
On Error Resume Next: ses.findById(ModuleHelper.ToolConfig("VL06O_EXECUTE_ID", "wnd[0]/tbar[1]/btn[8]"), False).press: If Err.Number <> 0 Then Err.Clear: ses.findById("wnd[0]").sendVKey 8
On Error GoTo Fail
If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 2924, , "SAP Busy qua lau sau Execute VL06O."
Set grid = SAPX_WaitGrid(ses, 20): If grid Is Nothing Then Err.Raise vbObjectError + 2925, , "Khong tim thay VL06O result grid."
ResultReady:
If SAPX_ResolveColumn(grid, "", "MATNR,MATERIAL", False) = "" Then
    VL6_PressItemView ses
    Set grid = SAPX_WaitGrid(ses, 20): If grid Is Nothing Then Err.Raise vbObjectError + 2926, , "Khong mo duoc Item View."
End If
cGM = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_GM", "WBSTA"), "WBSTA,WBSTK,GM,GOODS MOVEMENT STATUS", False)
cDel = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_DELIVERY", "VBELN"), "VBELN,DELIVERY,OUTBOUND DELIVERY", True)
cMat = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_MATERIAL", "MATNR"), "MATNR,MATERIAL", True)
cQty = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_QTY", "LFIMG"), "LFIMG,DLV.QTY,DELIVERY QTY", True)
cBatch = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_BATCH", "S_CHARG"), "S_CHARG,CHARG,BATCH", False)
cSloc = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_SLOC", "LGORT"), "LGORT,SLOC,STORAGE LOCATION", True)
cShip = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_SHIPTO", "KUNNR"), "KUNNR,KUNWE,SHIP-TO PARTY,SHIP TO", True)
colDate = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_CREATED_ON", "ERDAT"), "ERDAT,CREATED ON,CREATION DATE", True)
cTime = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_CREATED_TIME", "ERZET"), "ERZET,TIME,CREATED TIME", False)
cBy = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_CREATED_BY", "ERNAM"), "ERNAM,CREATED BY", False)
cShpt = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_SHPT", "VSTEL"), "VSTEL,SHIPPING POINT", False)
Application.StatusBar = "VL06O: dang kiem tra filter SLoc hien co..."
slocFilterOK = SAPX_VerifyFilter(grid, cSloc, slocValue)
If slocFilterOK Then
    Application.StatusBar = "VL06O: filter SLoc hien co da dung; bo qua loc lai."
Else
    SAPX_ClearGridFiltersFast ses, grid
    Set grid = SAPX_FindLargestGrid(ses)
    If grid Is Nothing Then Err.Raise vbObjectError + 2926, , "VL06O grid mat sau khi xoa filter cu."
    If Not SAPX_ApplyGridFilter(ses, grid, cSloc, slocValue, "SLOC,STORAGE LOCATION,LGORT") Then Err.Raise vbObjectError + 2929, , "Khong dien duoc filter SLoc=" & slocValue & " tai Item View. " & mSAPXFilterError
    Set grid = SAPX_FindLargestGrid(ses)
    If grid Is Nothing Or Not SAPX_VerifyFilter(grid, cSloc, slocValue) Then Err.Raise vbObjectError + 2930, , "Filter SLoc=" & slocValue & " khong co hieu luc tai Item View."
End If
rowCount = CLng(grid.rowCount): If rowCount > 20000 Then Err.Raise vbObjectError + 2931, , "VL06O co qua 20000 dong."
If rowCount > 0 Then ReDim arr(1 To rowCount, 1 To 11)
lastDel = "": lastShip = "": lastSloc = "": lastCreated = ""
For r = 0 To rowCount - 1
    SAPX_EnsureGridRowVisible grid, r
    vDel = SAPX_GridText(grid, r, cDel): If vDel <> "" Then lastDel = vDel Else vDel = lastDel
    vShip = SAPX_GridText(grid, r, cShip): If vShip <> "" Then lastShip = vShip Else vShip = lastShip
    vSloc = SAPX_GridText(grid, r, cSloc): If vSloc <> "" Then lastSloc = vSloc Else vSloc = lastSloc
    vCreated = SAPX_GridText(grid, r, colDate): If vCreated <> "" Then lastCreated = vCreated Else vCreated = lastCreated
    vMat = SAPX_GridText(grid, r, cMat)
    vQty = SAPX_GridText(grid, r, cQty)
    If Trim$(vDel) <> "" And Trim$(vMat) <> "" And ModuleHelper.CleanDouble(vQty) > 0 And _
       StrComp(Trim$(vShip), shipValue, vbTextCompare) = 0 And _
       StrComp(Trim$(vSloc), slocValue, vbTextCompare) = 0 Then
        outCount = outCount + 1
        arr(outCount, 1) = IIf(cGM = "", "A", SAPX_GridText(grid, r, cGM))
        arr(outCount, 2) = vDel: arr(outCount, 3) = vMat: arr(outCount, 4) = vQty
        arr(outCount, 5) = SAPX_GridText(grid, r, cBatch): arr(outCount, 6) = vSloc: arr(outCount, 7) = vShip
        arr(outCount, 8) = vCreated: arr(outCount, 9) = SAPX_GridText(grid, r, cTime)
        arr(outCount, 10) = SAPX_GridText(grid, r, cBy): arr(outCount, 11) = SAPX_GridText(grid, r, cShpt)
    End If
Next r
Set ws = ThisWorkbook.Worksheets("OB_TREO_VL06O")
clearLast = 6
For j = 1 To 11
    oldLast = ws.Cells(ws.rows.Count, j).End(xlUp).Row
    If oldLast > clearLast Then clearLast = oldLast
Next j
ws.Range("A6:K" & CStr(clearLast)).ClearContents
ws.Range("A6:K" & CStr(clearLast)).Interior.pattern = xlNone
If outCount > 0 Then
    ReDim finalArr(1 To outCount, 1 To 11)
    For i = 1 To outCount
        For j = 1 To 11: finalArr(i, j) = arr(i, j): Next j
        d = ModuleHelper.ParseSapDate(arr(i, 8))
        If d > 0 Then finalArr(i, 8) = d
    Next i
    ws.Range("E6:E" & CStr(outCount + 5)).NumberFormat = "@"
    ws.Range("H6:H" & CStr(outCount + 5)).NumberFormat = "dd.mm.yyyy"
    ws.Range("A6").Resize(outCount, 11).Value2 = finalArr
    ws.Range("A5:K" & CStr(outCount + 5)).Sort Key1:=ws.Range("H6"), Order1:=xlAscending, header:=xlYes
End If
Set unique = CreateObject("Scripting.Dictionary")
If outCount > 0 Then sortedData = ws.Range("A6:K" & CStr(outCount + 5)).Value2
For i = 1 To outCount
    vDel = Trim$(CStr(sortedData(i, 2) & "")): If vDel <> "" Then unique(vDel) = True
    d = ModuleHelper.ParseSapDate(sortedData(i, 8))
    If d > 0 Then
        If Date - d >= 2 Then
            ws.Range(ws.Cells(i + 5, 1), ws.Cells(i + 5, 11)).Interior.Color = RGB(255, 199, 206)
            oldCount = oldCount + 1
        ElseIf Date - d = 1 Then
            ws.Range(ws.Cells(i + 5, 1), ws.Cells(i + 5, 11)).Interior.Color = RGB(255, 235, 156)
        End If
    End If
Next i
ws.Range("A2").value = "Lan kiem tra": ws.Range("B2").value = Now: ws.Range("B2").NumberFormat = "dd.mm.yyyy hh:mm:ss"
ws.Range("D2").value = "So item": ws.Range("E2").value = outCount: ws.Range("G2").value = "So OB": ws.Range("H2").value = unique.Count: ws.Range("J2").value = "OB >=2 ngay": ws.Range("K2").value = oldCount
ws.Range("A3").value = "Date: " & Format$(fromD, "dd.mm.yyyy") & " - " & Format$(toD, "dd.mm.yyyy") & " | GM=A | Ship-to=" & shipValue & " | SLoc=" & slocValue & " | Khong dien Shipping Point/Receiving Pt."
Application.StatusBar = False
If showMessage Then MsgBox "Da check OB treo: " & unique.Count & " OB / " & outCount & " item.", vbInformation, "VL06O"
Exit Sub
Fail:
failNo = Err.Number: failText = Err.Description
Application.StatusBar = False
ModuleLog.WriteLog "LOI", "VL06O: " & failText
Err.Raise failNo, , failText
End Sub
Private Sub VL6_OpenSelection(ByVal ses As Object)
Dim ctl As Object, btn As Object
If Not SAPX_TcodeIs(ses, "VL06O") Then SAPX_Navigate ses, "VL06O"
Set ctl = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_DATE_FROM_ID", "wnd[0]/usr/ctxtIT_WADAT-LOW"))
If Not ctl Is Nothing Then Exit Sub
Set btn = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_LIST_OUTBOUND_BUTTON_ID", "wnd[0]/usr/btnBUTTON6"))
If Not btn Is Nothing Then btn.press: SAPX_WaitIdle ses, 15
Set ctl = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_DATE_FROM_ID", "wnd[0]/usr/ctxtIT_WADAT-LOW"))
If ctl Is Nothing Then Set ctl = SAPX_FindByTokens(ses, "WADAT-LOW|WADAT_LOW|WADATLOW", True)
If Not ctl Is Nothing Then Exit Sub
Set btn = SAPX_FindButtonText(ses.findById("wnd[0]/usr", False), "LIST OUTBOUND DELIVERIES", 0)
If Not btn Is Nothing Then btn.press: SAPX_WaitIdle ses, 15
Set ctl = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_DATE_FROM_ID", "wnd[0]/usr/ctxtIT_WADAT-LOW"))
If ctl Is Nothing Then Set ctl = SAPX_FindByTokens(ses, "WADAT-LOW|WADAT_LOW|WADATLOW", True)
End Sub
Private Function VL6_SetSelectionShipTo(ByVal ses As Object, ByVal shipValue As String) As Boolean
Dim ctl As Object, configuredId As String
configuredId = ModuleHelper.ToolConfig("VL06O_SHIP_TO_ID", "wnd[0]/usr/ctxtIT_KUNWE-LOW")
Set ctl = SAPX_FindById(ses, configuredId)
If ctl Is Nothing Then Set ctl = SAPX_FindByTokens(ses, "IT_KUNWE-LOW|IT_KUNNR-LOW|S_KUNWE-LOW|S_KUNNR-LOW|KUNWE-LOW|KUNNR-LOW|SHIP-TO PARTY|SHIP TO PARTY", True)
If ctl Is Nothing Then Exit Function
SAPX_SetValue ctl, shipValue
On Error Resume Next
VL6_SetSelectionShipTo = (StrComp(Trim$(CStr(ctl.text)), Trim$(shipValue), vbTextCompare) = 0)
On Error GoTo 0
If Not VL6_SetSelectionShipTo Then Err.Raise vbObjectError + 2942, , "Khong dien duoc Ship-to=" & shipValue & " tai man hinh chon VL06O."
End Function
Private Sub VL6_PressItemView(ByVal ses As Object)
Dim btn As Object
On Error Resume Next
Set btn = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_ITEM_VIEW_BUTTON_ID", "wnd[0]/tbar[1]/btn[18]"))
If Not btn Is Nothing Then btn.press Else Set btn = SAPX_FindButtonText(ses.findById("wnd[0]", False), "ITEM VIEW", 0): If Not btn Is Nothing Then btn.press
On Error GoTo 0
SAPX_WaitIdle ses, 15
End Sub
Private Function SAPX_GetSessionForTcode(ByVal tcode As String) As Object
Dim seed As Object, app As Object, con As Object, ses As Object, cand As Object, ids As Object
Dim i As Long, j As Long, n As Long, sid As String, t As Double
Set seed = modSupports.Get_SAPSession()
If seed Is Nothing Then Err.Raise vbObjectError + 2938, , "Khong co SAP session."

' User profiles do not keep the same SAP connection/session index.
' Search all connections before creating a new tab.
On Error Resume Next
Set app = seed.parent.parent
On Error GoTo 0
If Not app Is Nothing Then
    For i = 0 To app.Children.Count - 1
        Set con = SAPX_CollectionObjectAt(app.Children, i)
        If Not con Is Nothing Then
            For j = 0 To con.Children.Count - 1
                Set ses = SAPX_CollectionObjectAt(con.Children, j)
                If Not ses Is Nothing Then
                    If SAPX_TcodeIs(ses, tcode) Then
                        If SAPX_WaitIdle(ses, 20) Then
                            Set SAPX_GetSessionForTcode = ses
                            Exit Function
                        End If
                    End If
                End If
            Next j
        End If
    Next i
End If

' No existing tcode: create a new session in the active connection.
Set con = seed.parent
Set ids = CreateObject("Scripting.Dictionary"): n = con.Children.Count
For i = 0 To n - 1
    Set ses = SAPX_CollectionObjectAt(con.Children, i)
    If Not ses Is Nothing Then sid = CStr(ses.id): If sid <> "" Then ids(sid) = True
Next i
seed.CreateSession: t = Timer
Do
    DoEvents: SAPX_Sleep 25
    If con.Children.Count > n Then Exit Do
    If SAPX_Elapsed(t) > 8 Then Err.Raise vbObjectError + 2941, , "Tao SAP session bi timeout."
Loop
For i = 0 To con.Children.Count - 1
    Set ses = SAPX_CollectionObjectAt(con.Children, i)
    If Not ses Is Nothing Then
        sid = CStr(ses.id)
        If Not ids.Exists(sid) Then Set cand = ses: Exit For
    End If
Next i
If cand Is Nothing Then Err.Raise vbObjectError + 2942, , "Khong nhan duoc SAP session moi."
If Not SAPX_WaitIdle(cand, 15) Then Err.Raise vbObjectError + 2943, , "SAP session moi dang Busy."
SAPX_Navigate cand, tcode
If Not SAPX_WaitIdle(cand, 20) Then Err.Raise vbObjectError + 2943, , "SAP chua san sang sau khi dieu huong."
If Not SAPX_TcodeIs(cand, tcode) Then Err.Raise vbObjectError + 2944, , "SAP session moi sai TCode."
Set SAPX_GetSessionForTcode = cand
End Function
Private Function SAPX_TcodeIs(ByVal ses As Object, ByVal tcode As String) As Boolean
Dim a As String, b As String
On Error Resume Next: a = UCase$(Trim$(CStr(ses.info.Transaction))): On Error GoTo 0
b = UCase$(Trim$(tcode))
If Left$(a, 1) = "/" Then a = Mid$(a, 2)
If Left$(b, 1) = "/" Then b = Mid$(b, 2)
SAPX_TcodeIs = (a = b)
End Function
Private Sub SAPX_Navigate(ByVal ses As Object, ByVal tcode As String)
Dim cmd As String
cmd = tcode: If Left$(cmd, 1) = "/" Then cmd = "/n" & cmd Else cmd = "/n" & cmd
ses.findById("wnd[0]/tbar[0]/okcd").text = cmd: ses.findById("wnd[0]").sendVKey 0
If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 2940, , "SAP Busy khi mo " & tcode
End Sub
Private Function SAPX_WaitIdle(ByVal ses As Object, ByVal timeoutSec As Double) As Boolean
Dim t As Double: t = Timer
Do
    On Error Resume Next
    If Not ses.Busy Then SAPX_WaitIdle = True: Exit Function
    On Error GoTo 0
    DoEvents: SAPX_Sleep 25
Loop While SAPX_Elapsed(t) < timeoutSec
End Function
Private Function SAPX_Elapsed(ByVal t As Double) As Double
If Timer >= t Then SAPX_Elapsed = Timer - t Else SAPX_Elapsed = 86400# - t + Timer
End Function
Private Function SAPX_FindById(ByVal ses As Object, ByVal idText As String) As Object
If Trim$(idText) = "" Then Exit Function
On Error Resume Next: Set SAPX_FindById = ses.findById(idText, False): On Error GoTo 0
End Function
Private Function SAPX_CollectionObjectAt(ByVal items As Object, ByVal indexNo As Long) As Object
Dim idx As Variant
If items Is Nothing Then Exit Function
idx = CLng(indexNo)
On Error Resume Next
Err.Clear
Set SAPX_CollectionObjectAt = items.ElementAt(idx)
If SAPX_CollectionObjectAt Is Nothing Or Err.Number <> 0 Then
    Err.Clear
    Set SAPX_CollectionObjectAt = items.Item(idx)
End If
If SAPX_CollectionObjectAt Is Nothing Or Err.Number <> 0 Then
    Err.Clear
    Set SAPX_CollectionObjectAt = items(CLng(indexNo))
End If
If SAPX_CollectionObjectAt Is Nothing Or Err.Number <> 0 Then
    Err.Clear
    Set SAPX_CollectionObjectAt = items(CInt(indexNo))
End If
Err.Clear
On Error GoTo 0
End Function
Private Function SAPX_CollectionValueAt(ByVal items As Object, ByVal indexNo As Long) As Variant
Dim idx As Variant
If items Is Nothing Then Exit Function
idx = CLng(indexNo)
On Error Resume Next
Err.Clear
SAPX_CollectionValueAt = items.ElementAt(idx)
If Err.Number <> 0 Then Err.Clear: SAPX_CollectionValueAt = items.Item(idx)
If Err.Number <> 0 Then Err.Clear: SAPX_CollectionValueAt = items(CLng(indexNo))
If Err.Number <> 0 Then Err.Clear: SAPX_CollectionValueAt = items(CInt(indexNo))
Err.Clear
On Error GoTo 0
End Function
Private Function SAPX_FindByTokens(ByVal ses As Object, ByVal tokens As String, ByVal textOnly As Boolean) As Object
Dim root As Object
On Error Resume Next: Set root = ses.findById("wnd[0]/usr", False): On Error GoTo 0
Set SAPX_FindByTokens = SAPX_FindTokensRec(root, UCase$(tokens), textOnly, 0)
End Function
Private Function SAPX_FindTokensRec(ByVal obj As Object, ByVal tokens As String, ByVal textOnly As Boolean, ByVal depth As Long) As Object
Dim idt As String, typ As String, a As Variant, ch As Object, i As Long
If obj Is Nothing Or depth > 12 Then Exit Function
On Error Resume Next: idt = UCase$(CStr(obj.id)): typ = UCase$(CStr(obj.Type)): On Error GoTo 0
If (Not textOnly Or InStr(typ, "TEXTFIELD") > 0 Or InStr(typ, "COMBOBOX") > 0) Then
    For Each a In Split(tokens, "|"): If Trim$(CStr(a)) <> "" And InStr(idt, UCase$(Trim$(CStr(a)))) > 0 Then Set SAPX_FindTokensRec = obj: Exit Function
    Next a
End If
On Error Resume Next
For i = 0 To obj.Children.Count - 1
    Set ch = SAPX_CollectionObjectAt(obj.Children, i)
    Set SAPX_FindTokensRec = SAPX_FindTokensRec(ch, tokens, textOnly, depth + 1)
    If Not SAPX_FindTokensRec Is Nothing Then Exit Function
Next i
On Error GoTo 0
End Function
Private Sub SAPX_SetValue(ByVal ctl As Object, ByVal valueText As String)
Dim typ As String
On Error Resume Next: typ = UCase$(CStr(ctl.Type)): On Error GoTo 0
On Error Resume Next
If InStr(typ, "COMBOBOX") > 0 Then ctl.key = valueText
If Err.Number <> 0 Or InStr(typ, "COMBOBOX") = 0 Then Err.Clear: ctl.text = valueText
If Err.Number <> 0 Then Err.Clear: ctl.key = valueText
On Error GoTo 0
End Sub
Private Function SAPX_ControlValue(ByVal ctl As Object) As String
On Error Resume Next: SAPX_ControlValue = Trim$(CStr(ctl.text)): If SAPX_ControlValue = "" Then SAPX_ControlValue = Trim$(CStr(ctl.key)): On Error GoTo 0
End Function
Private Function SAPX_RefreshVL06OAndCatchReady(ByVal ses As Object, ByRef grid As Object, ByVal timeoutSec As Double) As Boolean
Dim started As Double, busyNow As Boolean, eventSeen As Boolean
Dim oldStatus As String, newStatus As String, candidate As Object
On Error GoTo Failed
oldStatus = SAPX_StatusText(ses)
ses.findById("wnd[0]").sendVKey ModuleHelper.ToolConfigLong("VL06O_REFRESH_VKEY", 20, 0, 99)
eventSeen = True
started = Timer
Do
    DoEvents
    busyNow = CBool(ses.Busy)
    newStatus = SAPX_StatusText(ses)
    If busyNow Then eventSeen = True
    If StrComp(newStatus, oldStatus, vbTextCompare) <> 0 Then eventSeen = True
    If InStr(1, newStatus, "refreshed", vbTextCompare) > 0 Then eventSeen = True
    If eventSeen And Not busyNow Then
        Set candidate = SAPX_FindLargestGrid(ses)
        If Not candidate Is Nothing Then
            Set grid = candidate
            SAPX_RefreshVL06OAndCatchReady = True
            Exit Function
        End If
    End If
    SAPX_Sleep 25
Loop While SAPX_Elapsed(started) < timeoutSec
Exit Function
Failed:
SAPX_RefreshVL06OAndCatchReady = False
End Function
Private Function SAPX_WaitGrid(ByVal ses As Object, ByVal timeoutSec As Double) As Object
Dim started As Double, grid As Object, rowCount As Long, lastRowCount As Long
Dim stableHits As Long, busyNow As Boolean
started = Timer: lastRowCount = -1
Do
    Set grid = SAPX_FindLargestGrid(ses)
    busyNow = True
    On Error Resume Next
    busyNow = CBool(ses.Busy)
    If Not grid Is Nothing Then rowCount = CLng(grid.rowCount) Else rowCount = -1
    If Err.Number <> 0 Then Err.Clear: rowCount = -1
    On Error GoTo 0
    If Not grid Is Nothing And Not busyNow And rowCount >= 0 Then
        If rowCount = lastRowCount Then
            stableHits = stableHits + 1
        Else
            stableHits = 0
            lastRowCount = rowCount
        End If
        If stableHits >= 1 Then
            Set SAPX_WaitGrid = grid
            Exit Function
        End If
    Else
        stableHits = 0
    End If
    SAPX_Sleep 25
    DoEvents
Loop While SAPX_Elapsed(started) < timeoutSec
If Not grid Is Nothing And Not busyNow And rowCount >= 0 Then Set SAPX_WaitGrid = grid
End Function
Private Function SAPX_FindLargestGrid(ByVal ses As Object) As Object
Dim root As Object, best As Object, bestRows As Long
On Error Resume Next: Set root = ses.findById("wnd[0]", False): On Error GoTo 0
SAPX_FindGridRec root, best, bestRows, 0
Set SAPX_FindLargestGrid = best
End Function
Private Sub SAPX_FindGridRec(ByVal obj As Object, ByRef best As Object, ByRef bestRows As Long, ByVal depth As Long)
Dim typ As String, subTyp As String, n As Long, i As Long, ch As Object, hasRowCount As Boolean
If obj Is Nothing Or depth > 14 Then Exit Sub
On Error Resume Next
typ = UCase$(CStr(obj.Type))
subTyp = UCase$(CStr(obj.SubType))
Err.Clear
n = CLng(obj.rowCount)
hasRowCount = (Err.Number = 0)
Err.Clear
If InStr(typ, "GRIDVIEW") > 0 Or InStr(subTyp, "GRIDVIEW") > 0 Or _
   (InStr(typ, "GUISHELL") > 0 And hasRowCount) Then
    If best Is Nothing Or n > bestRows Then Set best = obj: bestRows = n
End If
For i = 0 To obj.Children.Count - 1: Set ch = SAPX_CollectionObjectAt(obj.Children, i): SAPX_FindGridRec ch, best, bestRows, depth + 1: Next i
On Error GoTo 0
End Sub
Private Function SAPX_ResolveColumn(ByVal grid As Object, ByVal configured As String, ByVal aliasesCsv As String, ByVal required As Boolean) As String
Dim ids As Collection, id As Variant, cap As String, a As Variant, nc As String, na As String, ni As String
If Trim$(configured) <> "" Then If SAPX_ColumnExists(grid, configured) Then SAPX_ResolveColumn = configured: Exit Function
Set ids = SAPX_ColumnIds(grid)
' Pass 1: exact normalized ID or caption. This prevents short aliases such
' as ST from matching SOURCE STORAGE BIN.
For Each id In ids
    ni = NormalizeHeader(CStr(id))
    nc = NormalizeHeader(SAPX_ColumnCaption(grid, CStr(id)))
    For Each a In Split(aliasesCsv, ",")
        na = NormalizeHeader(Trim$(CStr(a)))
        If na <> "" And (ni = na Or nc = na) Then SAPX_ResolveColumn = CStr(id): Exit Function
    Next a
Next id
' Pass 2: fuzzy match only for meaningful aliases (4+ normalized chars).
For Each id In ids
    cap = UCase$(CStr(id) & " " & SAPX_ColumnCaption(grid, CStr(id)))
    nc = NormalizeHeader(cap)
    For Each a In Split(aliasesCsv, ",")
        na = NormalizeHeader(Trim$(CStr(a)))
        If Len(na) >= 4 And InStr(1, nc, na, vbTextCompare) > 0 Then SAPX_ResolveColumn = CStr(id): Exit Function
    Next a
Next id
If required Then Err.Raise vbObjectError + 2941, , "Thieu cot SAP: " & aliasesCsv
End Function
Private Function SAPX_ColumnIds(ByVal grid As Object) As Collection
Dim c As New Collection, o As Object, i As Long, v As Variant
On Error Resume Next: Set o = grid.ColumnOrder
If Not o Is Nothing Then
    For i = 0 To o.Count - 1
        v = SAPX_CollectionValueAt(o, i)
        If Trim$(CStr(v & "")) <> "" Then c.Add CStr(v)
    Next i
End If
On Error GoTo 0
Set SAPX_ColumnIds = c
End Function
Private Function SAPX_ColumnCaption(ByVal grid As Object, ByVal colId As String) As String
Dim v As Variant, o As Object, i As Long, s As String
On Error Resume Next
v = grid.GetDisplayedColumnTitle(colId): s = CStr(v & "")
If s = "" Then
    v = grid.GetColumnTitles(colId)
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            If Trim$(CStr(v(i) & "")) <> "" Then s = s & " " & CStr(v(i))
        Next i
    End If
End If
If s = "" And IsObject(v) Then
    Set o = v
    For i = 0 To o.Count - 1
        s = s & " " & CStr(o(i))
    Next i
End If
On Error GoTo 0
SAPX_ColumnCaption = Trim$(s)
End Function
Private Function SAPX_ColumnExists(ByVal grid As Object, ByVal colId As String) As Boolean
Dim ids As Collection, v As Variant
Set ids = SAPX_ColumnIds(grid)
For Each v In ids: If StrComp(CStr(v), colId, vbTextCompare) = 0 Then SAPX_ColumnExists = True: Exit Function
Next v
End Function
Private Function SAPX_GridText(ByVal grid As Object, ByVal rowNo As Long, ByVal colId As String) As String
If colId = "" Then Exit Function
On Error Resume Next: SAPX_GridText = Trim$(CStr(grid.GetCellValue(rowNo, colId) & "")): On Error GoTo 0
End Function
Private Sub SAPX_EnsureGridRowVisible(ByVal grid As Object, ByVal rowNo As Long)
Dim firstRow As Long, visibleRows As Long
' GetCellValue can read rows outside the viewport. Only touch the scrollbar
' periodically so large VL06O lists do not make two COM calls for every row.
If rowNo > 0 And rowNo Mod 20 <> 0 Then Exit Sub
On Error Resume Next
firstRow = CLng(grid.FirstVisibleRow)
visibleRows = CLng(grid.VisibleRowCount)
If visibleRows < 1 Then visibleRows = 1
If rowNo < firstRow Or rowNo >= firstRow + visibleRows Then
    grid.FirstVisibleRow = rowNo
    DoEvents
End If
On Error GoTo 0
End Sub
Private Function SAPX_SortGridAscending(ByVal ses As Object, ByVal grid As Object, ByVal colId As String) As Boolean
If grid Is Nothing Or Trim$(colId) = "" Then Exit Function
On Error Resume Next
Err.Clear
grid.SetCurrentCell 0, colId
grid.SelectColumn colId
grid.PressToolbarButton "&SORT_ASC"
If Err.Number <> 0 Then
    Err.Clear
    grid.ContextMenu
    grid.selectContextMenuItem "&SORT_ASC"
End If
If Err.Number = 0 Then SAPX_SortGridAscending = SAPX_WaitIdle(ses, 30)
Err.Clear
On Error GoTo 0
End Function
Private Sub SAPX_ClearGridFiltersFast(ByVal ses As Object, ByVal grid As Object)
On Error Resume Next
Err.Clear
grid.PressToolbarContextButton "&MB_FILTER"
grid.SelectToolbarMenuItem "&DELETE_FILTER"
If Err.Number <> 0 Then
    Err.Clear
    grid.PressToolbarButton "&DELETE_FILTER"
End If
If Err.Number <> 0 Then
    Err.Clear
    grid.PressToolbarButton "&DEL_FILTER"
End If
Err.Clear
On Error GoTo 0
SAPX_WaitIdle ses, 10
End Sub
Private Function SAPX_FindButtonText(ByVal obj As Object, ByVal wanted As String, ByVal depth As Long) As Object
Dim typ As String, txt As String, i As Long, ch As Object
If obj Is Nothing Or depth > 12 Then Exit Function
On Error Resume Next: typ = UCase$(CStr(obj.Type)): txt = UCase$(CStr(obj.text)): On Error GoTo 0
If InStr(typ, "BUTTON") > 0 And InStr(txt, UCase$(wanted)) > 0 Then Set SAPX_FindButtonText = obj: Exit Function
On Error Resume Next
For i = 0 To obj.Children.Count - 1: Set ch = SAPX_CollectionObjectAt(obj.Children, i): Set SAPX_FindButtonText = SAPX_FindButtonText(ch, wanted, depth + 1): If Not SAPX_FindButtonText Is Nothing Then Exit Function
Next i
On Error GoTo 0
End Function
Private Function SAPX_ApplyGridFilter(ByVal ses As Object, ByVal grid As Object, ByVal colId As String, ByVal valueText As String, ByVal labelAliases As String) As Boolean
Dim t As Double, low As Object, ok As Boolean, pop As Object, popupRoot As Object
Dim filterStage As String, failNo As Long, failText As String, controlDump As String
On Error GoTo Failed
filterStage = "SELECT_COLUMN": mSAPXFilterError = ""
If CLng(grid.rowCount) > 0 Then grid.SetCurrentCell 0, colId
grid.SelectColumn colId
filterStage = "OPEN_FILTER"
On Error Resume Next
Err.Clear
grid.ContextMenu
grid.selectContextMenuItem "&FILTER"
SAPX_Sleep 25: DoEvents
ok = SAPX_FilterPopupOpen(ses)
Err.Clear
If Not ok Then
    grid.PressToolbarContextButton "&MB_FILTER"
    grid.SelectToolbarMenuItem "&FILTER"
    SAPX_Sleep 25: DoEvents
    ok = SAPX_FilterPopupOpen(ses)
End If
Err.Clear
If Not ok Then
    grid.PressToolbarButton "&FILTER"
    SAPX_Sleep 25: DoEvents
    ok = SAPX_FilterPopupOpen(ses)
End If
Err.Clear
If Not ok Then
    grid.PressToolbarButton "&MB_FILTER"
    SAPX_Sleep 25: DoEvents
    ok = SAPX_FilterPopupOpen(ses)
End If
Err.Clear
If Not ok Then
    grid.PressToolbarContextButton "&MB_FILTER"
    grid.SelectToolbarMenuItem "&MB_FILTER"
    SAPX_Sleep 25: DoEvents
    ok = SAPX_FilterPopupOpen(ses)
End If
Err.Clear
On Error GoTo Failed
If Not ok Then GoTo Failed
filterStage = "FIND_INPUT"
t = Timer
Do
    On Error Resume Next
    Set low = ses.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxtRSCSEL_255-SLOW_I[1,0]", False)
    Set popupRoot = ses.findById("wnd[1]/usr", False)
    On Error GoTo Failed
    If low Is Nothing And Not popupRoot Is Nothing Then Set low = SAPX_FindTokensRec(popupRoot, "SLOW_I[1,0]|SLOW_I|LOW_I", True, 0)
    If low Is Nothing Then Set low = SAPX_FindFilterInput(ses, labelAliases)
    If Not low Is Nothing Then Exit Do
    SAPX_Sleep 25: DoEvents
Loop While SAPX_Elapsed(t) < 5
If low Is Nothing Then GoTo Failed
filterStage = "SET_VALUE"
SAPX_SetValue low, valueText
If StrComp(SAPX_ControlValue(low), Trim$(valueText), vbTextCompare) <> 0 Then GoTo Failed
filterStage = "CONFIRM"
On Error Resume Next: ses.findById("wnd[1]/tbar[0]/btn[0]", False).press: If Err.Number <> 0 Then Err.Clear: ses.findById("wnd[1]").sendVKey 0
On Error GoTo Failed
If Not SAPX_WaitIdle(ses, 20) Then GoTo Failed
SAPX_ApplyGridFilter = True: Exit Function
Failed:
failNo = Err.Number: failText = Err.Description
mSAPXFilterError = filterStage & " [" & CStr(failNo) & "] " & failText
If filterStage = "FIND_INPUT" And Not popupRoot Is Nothing Then
    SAPX_DumpControlIds popupRoot, controlDump, 0
    mSAPXFilterError = mSAPXFilterError & " Controls=" & Replace(controlDump, vbLf, "; ")
End If
On Error Resume Next
Set pop = ses.findById("wnd[1]", False)
If Not pop Is Nothing Then pop.Close
On Error GoTo 0
End Function
Private Function SAPX_FilterPopupOpen(ByVal ses As Object) As Boolean
Dim pop As Object
On Error Resume Next
Set pop = ses.findById("wnd[1]", False)
SAPX_FilterPopupOpen = Not pop Is Nothing
On Error GoTo 0
End Function
Private Function SAPX_FindFilterInput(ByVal ses As Object, ByVal aliasesCsv As String) As Object
Dim root As Object, labels As New Collection, inputs As New Collection, o As Object, l As Object, inp As Object
Dim lr As Long, ir As Long, bestD As Long, d As Long, a As Variant, txt As String, dynToken As String
On Error Resume Next: Set root = ses.findById("wnd[1]/usr", False): On Error GoTo 0
If root Is Nothing Then Exit Function
SAPX_CollectFilterControls root, labels, inputs, 0
For Each l In labels
    txt = NormalizeHeader(SAPX_ControlText(l))
    For Each a In Split(aliasesCsv, ",")
        If InStr(txt, NormalizeHeader(Trim$(CStr(a)))) > 0 Then
            dynToken = SAPX_DynToken(CStr(l.id))
            If dynToken <> "" Then
                For Each o In inputs
                    If InStr(1, CStr(o.id), dynToken, vbTextCompare) > 0 Then Set SAPX_FindFilterInput = o: Exit Function
                Next o
            End If
            lr = SAPX_IdRow(CStr(l.id)): bestD = 9999: Set inp = Nothing
            For Each o In inputs: ir = SAPX_IdRow(CStr(o.id)): d = Abs(ir - lr): If d < bestD Then bestD = d: Set inp = o
            Next o
            If Not inp Is Nothing And bestD <= 1 Then Set SAPX_FindFilterInput = inp: Exit Function
        End If
    Next a
Next l
If inputs.Count = 1 Then Set SAPX_FindFilterInput = inputs(1)
End Function
Private Function SAPX_DynToken(ByVal idText As String) As String
Dim p As Long, i As Long, ch As String, token As String
p = InStr(1, idText, "%%DYN", vbTextCompare)
If p = 0 Then Exit Function
For i = p To Len(idText)
    ch = Mid$(idText, i, 1)
    If (ch >= "0" And ch <= "9") Or ch = "%" Or UCase$(ch) = "D" Or UCase$(ch) = "Y" Or UCase$(ch) = "N" Then
        token = token & ch
    Else
        Exit For
    End If
Next i
SAPX_DynToken = token
End Function
Private Sub SAPX_CollectFilterControls(ByVal obj As Object, ByRef labels As Collection, ByRef inputs As Collection, ByVal depth As Long)
Dim typ As String, idt As String, i As Long, ch As Object
If obj Is Nothing Or depth > 12 Then Exit Sub
On Error Resume Next: typ = UCase$(CStr(obj.Type)): idt = UCase$(CStr(obj.id)): On Error GoTo 0
If InStr(typ, "LABEL") > 0 Or ((InStr(typ, "TEXTFIELD") > 0) And InStr(idt, "-TEXT") > 0) Then labels.Add obj
If (InStr(typ, "TEXTFIELD") > 0 Or InStr(typ, "COMBOBOX") > 0) And InStr(idt, "-LOW") > 0 Then inputs.Add obj
On Error Resume Next
For i = 0 To obj.Children.Count - 1: Set ch = SAPX_CollectionObjectAt(obj.Children, i): SAPX_CollectFilterControls ch, labels, inputs, depth + 1: Next i
On Error GoTo 0
End Sub
Private Function SAPX_ControlText(ByVal obj As Object) As String
On Error Resume Next: SAPX_ControlText = CStr(obj.text & ""): On Error GoTo 0
End Function
Private Function SAPX_IdRow(ByVal idText As String) As Long
Dim p As Long, q As Long, z As Long, s As String
p = InStrRev(idText, "["): q = InStrRev(idText, "]"): If p = 0 Or q <= p Then SAPX_IdRow = -999: Exit Function
s = Mid$(idText, p + 1, q - p - 1): z = InStr(s, ","): If z > 0 Then SAPX_IdRow = CLng(Val(Mid$(s, z + 1))) Else SAPX_IdRow = CLng(Val(s))
End Function
Private Function SAPX_VerifyFilter(ByVal grid As Object, ByVal colId As String, ByVal expected As String) As Boolean
Dim r As Long, v As String, lastV As String, n As Long, seen As Boolean
On Error GoTo Nope
n = CLng(grid.rowCount): If n = 0 Then SAPX_VerifyFilter = True: Exit Function
For r = 0 To n - 1
    v = SAPX_GridText(grid, r, colId)
    If v <> "" Then lastV = v: seen = True Else v = lastV
    If v <> "" And StrComp(Trim$(v), Trim$(expected), vbTextCompare) <> 0 Then Exit Function
Next r
SAPX_VerifyFilter = seen
Nope:
End Function
Private Sub SAPX_VerifyTwoFilters(ByVal grid As Object, _
ByVal colId1 As String, ByVal expected1 As String, ByRef ok1 As Boolean, _
ByVal colId2 As String, ByVal expected2 As String, ByRef ok2 As Boolean)
Dim r As Long, n As Long, v1 As String, v2 As String
Dim last1 As String, last2 As String, seen1 As Boolean, seen2 As Boolean
ok1 = True: ok2 = True
On Error GoTo Failed
n = CLng(grid.rowCount)
If n = 0 Then Exit Sub
For r = 0 To n - 1
    If Not ok1 And Not ok2 Then Exit For
    SAPX_EnsureGridRowVisible grid, r
    If ok1 Then
        v1 = SAPX_GridText(grid, r, colId1)
        If v1 <> "" Then last1 = v1: seen1 = True Else v1 = last1
        If v1 <> "" And StrComp(Trim$(v1), Trim$(expected1), vbTextCompare) <> 0 Then ok1 = False
    End If
    If ok2 Then
        v2 = SAPX_GridText(grid, r, colId2)
        If v2 <> "" Then last2 = v2: seen2 = True Else v2 = last2
        If v2 <> "" And StrComp(Trim$(v2), Trim$(expected2), vbTextCompare) <> 0 Then ok2 = False
    End If
Next r
If Not seen1 Then ok1 = False
If Not seen2 Then ok2 = False
Exit Sub
Failed:
ok1 = False: ok2 = False
End Sub

Public Function ATT_GetMigoSession() As Object
    On Error GoTo EH
    Set ATT_GetMigoSession = SAPX_GetSessionForTcode("MIGO")
    Exit Function
EH:
    Set ATT_GetMigoSession = Nothing
End Function
