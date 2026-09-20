Option Explicit
Public LastPostingError As String
Public LastPostingRowsRead As Long
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
    lastCol = ws.Cells(1, ws.columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = NormalizeHeader(ModuleHelper.CleanStr(ws.Cells(1, c).Value))
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
    If missingRequired.count > 0 Then Exit Sub
    Dim lastRow As Long, lastCol As Long
    lastRow = ModuleHelper.GetLastRow(ws, colMap("PRODUCT"))
    lastCol = ws.Cells(1, ws.columns.count).End(xlToLeft).Column
    If lastRow < 2 Then Exit Sub
    dataArr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value
End Sub
Public Sub ReadOrder(ByRef dataArr As Variant, ByRef colMap As Object, ByRef missingRequired As Collection)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ORDER")
    Set colMap = CreateObject("Scripting.Dictionary")
    Set missingRequired = New Collection
    MapOrderColumns ws, colMap, missingRequired
    dataArr = Empty
    If missingRequired.count > 0 Then Exit Sub
    Dim lastRow As Long, lastCol As Long
    lastRow = ModuleHelper.GetLastRow(ws, colMap("PRODUCT"))
    lastCol = ws.Cells(1, ws.columns.count).End(xlToLeft).Column
    If lastRow < 2 Then Exit Sub
    dataArr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value
End Sub
Public Sub ReadSpecialProducts(ByRef specialDict As Object)
    Set specialDict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("DANH_MUC")
    On Error GoTo 0
    If ws Is Nothing Then
        ModuleLog.WriteLog "CANH BAO", "Khong tim thay sheet DANH_MUC; bo qua danh sach ma chu y/FEFO"
        Exit Sub
    End If

    ' DANH_MUC: A=Product, B=Loai; header row 4, data row 5+
    Dim lastRow As Long
    lastRow = ModuleHelper.GetLastRow(ws, 1)
    If lastRow < 5 Then Exit Sub

    Dim i As Long, p As String, l As String
    For i = 5 To lastRow
        p = ModuleHelper.CleanStr(ws.Cells(i, 1).Value)
        l = UCase$(ModuleHelper.CleanStr(ws.Cells(i, 2).Value))
        If p <> "" And l <> "" Then specialDict(UCase$(p)) = l
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
 Dim r As Long, n As Long, rowCount As Long, expectedRows As Long, capturedRows As Long, validRows As Long
 Dim lastP As String, pv As String, qv As String, bv As String, tv As String, oneRow As Variant
 Dim captureClass As String, captureReason As String
Dim ws As Worksheet, oldLast As Long, outArr() As Variant, i As Long, msg As String, stage As String, guardMessage As String
Dim fastColIds As Collection, fastNormIds As Object, fastNormCaps As Object
On Error GoTo Fail
LastPostingError = ""
LastPostingRowsRead = 0
If Not SAPRun_RequireRuntime("POSTING_LOAD", _
    "POST_GRID_ID,POST_PRODUCT_COL,POST_AVAILABLE_QTY_COL,POST_BATCH_COL,POST_SOURCE_BIN_COL,POST_TST_COL,POST_TOTAL_STOCK_COL,POST_DEST_BIN_COL,POST_QTY_CHG_COL,POST_DEST_BATCH_COL,POST_RESULT_COUNT_ID,POST_RESULT_COUNT_PATTERN,POST_ADVANCED_TOGGLE_ID,POST_ADVANCED_RESET_ID,POST_ADVANCED_START_ID,POST_PRODUCT_MULTI_ID,POST_OWNER_ID,POST_PRODUCT_POPUP_CLOSE_ID,POST_PRODUCT_POPUP_CLEAR_ID,POST_PRODUCT_PASTE_ID,POST_PRODUCT_POPUP_OK_ID", guardMessage) Then
    Err.Raise vbObjectError + 2899, , guardMessage
End If
 gMPOST_ExpectedRows = -1
 gMPOST_SAPExpectedRows = -1
 gMPOST_CapturedRawRows = -1
 gMPOST_ValidStockRows = -1
gMPOST_CaptureComplete = False
stage = "READ_ORDER"
Set products = POST_OrderProducts()
If products.count = 0 Then Err.Raise vbObjectError + 2900, , "ORDER khong co Product hop le."
Set rows = New Collection
stage = "OPEN_POST"
Set ses = SAPX_GetSessionForTcode(ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST"), False)
If ses Is Nothing Then Err.Raise vbObjectError + 2901, , "Khong ket noi duoc SAP session."
stage = "OPEN_ADVANCED_SEARCH"
Application.StatusBar = "POSTING: dang gui " & products.count & " Product vao Advanced Search..."
If Not POST_OpenAdvancedProductResult(ses, products, grid, msg) Then Err.Raise vbObjectError + 2902, , msg
If grid Is Nothing Then Err.Raise vbObjectError + 2903, , "Advanced Search khong tra result grid."
stage = "MAP_COLUMNS"
' Production contract is locked in CONFIG. Do not fall back to captions/aliases.
SAPX_BuildColumnCatalogFast grid, fastColIds, fastNormIds, fastNormCaps
cProd = SAPX_RequireExactColumn(grid, ModuleHelper.ToolConfig("POST_PRODUCT_COL", ""), "POST_PRODUCT_COL")
cQty = SAPX_RequireExactColumn(grid, ModuleHelper.ToolConfig("POST_AVAILABLE_QTY_COL", ""), "POST_AVAILABLE_QTY_COL")
cBatch = SAPX_RequireExactColumn(grid, ModuleHelper.ToolConfig("POST_BATCH_COL", ""), "POST_BATCH_COL")
cBin = SAPX_RequireExactColumn(grid, ModuleHelper.ToolConfig("POST_SOURCE_BIN_COL", ""), "POST_SOURCE_BIN_COL")
cTst = SAPX_RequireExactColumn(grid, ModuleHelper.ToolConfig("POST_TST_COL", ""), "POST_TST_COL")
cTotal = SAPX_RequireExactColumn(grid, ModuleHelper.ToolConfig("POST_TOTAL_STOCK_COL", ""), "POST_TOTAL_STOCK_COL")
cDestBin = SAPX_RequireExactColumn(grid, ModuleHelper.ToolConfig("POST_DEST_BIN_COL", ""), "POST_DEST_BIN_COL")
cQtyPost = SAPX_RequireExactColumn(grid, ModuleHelper.ToolConfig("POST_QTY_CHG_COL", ""), "POST_QTY_CHG_COL")
cDestBatch = SAPX_RequireExactColumn(grid, ModuleHelper.ToolConfig("POST_DEST_BATCH_COL", ""), "POST_DEST_BATCH_COL")
cST = cTst
colDate = SAPX_OptionalExactColumn(grid, ModuleHelper.ToolConfig("POST_GR_DATE_COL", ""))
cTime = SAPX_OptionalExactColumn(grid, ModuleHelper.ToolConfig("POST_GR_TIME_COL", ""))
cProcType = SAPX_OptionalExactColumn(grid, ModuleHelper.ToolConfig("POST_PROC_TYPE_COL", ""))
cAvlBU = SAPX_OptionalExactColumn(grid, ModuleHelper.ToolConfig("POST_AVL_BU_QTY_COL", ""))
 stage = "READ_RESULT_COUNT"
 If Not SAPX_ReadIndependentResultCount(ses, expectedRows, msg) Then _
     Err.Raise vbObjectError + 2942, , msg
 If expectedRows < 0 Then Err.Raise vbObjectError + 2942, , _
     "STOP_POST_EXPECTED_COUNT_UNPROVEN: SAP result count is negative."
 gMPOST_SAPExpectedRows = expectedRows
 gMPOST_ExpectedRows = expectedRows
 stage = "EXPOSE_RESULT_GRID"
 If Not SAPX_ExposeExpectedRows(grid, expectedRows, cProd, cQty, cBatch, cBin, cTst, cTotal, _
                                ModuleHelper.ToolConfigDouble("POST_CAPTURE_TIMEOUT_SEC", 30#, 5#, 120#), msg) Then _
     Err.Raise vbObjectError + 2943, , msg
 rowCount = CLng(grid.rowCount)
 If Not SAPX_ValidateExpectedDataRows(grid, cProd, cQty, cBatch, cBin, cTst, cTotal, _
                                      expectedRows, rowCount, msg) Then _
     Err.Raise vbObjectError + 2944, , msg
 stage = "READ_RESULT_GRID"
 Application.StatusBar = "POSTING: dang doc " & expectedRows & " dong ton..."
 lastP = ""
 For r = 0 To expectedRows - 1
     SAPX_EnsureGridRowVisible grid, r
     pv = SAPX_GridText(grid, r, cProd)
     If pv <> "" Then lastP = pv Else pv = lastP
     qv = SAPX_GridText(grid, r, cQty): bv = SAPX_GridText(grid, r, cBatch): tv = SAPX_GridText(grid, r, cTst)
     captureClass = "NON_STOCK": captureReason = ""
     If Trim$(pv) = "" Then
         captureReason = "MATERIAL_EMPTY"
     ElseIf Not products.Exists(Trim$(pv)) Then
         captureReason = "MATERIAL_OUTSIDE_REQUEST"
    ElseIf Trim$(bv) = "" Then
        captureReason = "BATCH_EMPTY"
    ElseIf UCase$(Trim$(tv)) <> "FT" Then
        captureReason = "STOCK_TYPE_NOT_USABLE"
     Else
         captureClass = "VALID_STOCK"
         validRows = validRows + 1
     End If
     oneRow = Array(Trim$(pv), POST_ParseIntegerQty(qv), bv, SAPX_GridText(grid, r, cBin), tv, _
         IIf(cTotal <> "", POST_ParseIntegerQty(SAPX_GridText(grid, r, cTotal)), POST_ParseIntegerQty(qv)), _
         SAPX_GridText(grid, r, cDestBin), POST_ParseIntegerQty(SAPX_GridText(grid, r, cQtyPost)), _
         SAPX_GridText(grid, r, cDestBatch), SAPX_GridText(grid, r, cST), _
         SAPX_GridText(grid, r, colDate), SAPX_GridText(grid, r, cTime), _
         SAPX_GridText(grid, r, cProcType), _
         IIf(cAvlBU <> "", POST_ParseIntegerQty(SAPX_GridText(grid, r, cAvlBU)), POST_ParseIntegerQty(qv)), _
         captureClass, captureReason, 0)
     rows.Add oneRow
     capturedRows = capturedRows + 1
 Next r
 If Not SAPX_ValidateExpectedDataRows(grid, cProd, cQty, cBatch, cBin, cTst, cTotal, _
                                      expectedRows, rowCount, msg) Then _
     Err.Raise vbObjectError + 2945, , msg
 gMPOST_CapturedRawRows = capturedRows
 gMPOST_ValidStockRows = validRows
 gMPOST_CaptureComplete = (gMPOST_CapturedRawRows = gMPOST_SAPExpectedRows)
 If Not gMPOST_CaptureComplete Then Err.Raise vbObjectError + 2946, , "STOP_POST_CAPTURE_INCOMPLETE: captured rows do not equal SAP expected rows."
 stage = "COMMIT_POSTING"
 LastPostingRowsRead = gMPOST_ValidStockRows
Set ws = ThisWorkbook.Worksheets("POSTING")
oldLast = POST_LastDataRow(ws, 17)
If oldLast < 2 Then oldLast = 2
If rows.count = 0 Then
    ws.Range("A2:Q" & CStr(oldLast)).ClearContents
    DownloadPostingForOrder = True
    Application.StatusBar = "POSTING: SAP tra 0 dong ton cho ORDER."
    If showMessage Then MsgBox "SAP da tim xong nhung khong co dong ton hop le cho ORDER. POSTING cu da duoc xoa de tranh dung ton cu.", vbInformation, "TAI POSTING"
    Exit Function
End If
ReDim outArr(1 To rows.count, 1 To 17)
For i = 1 To rows.count
    oneRow = rows(i)
    For n = 1 To 17: outArr(i, n) = oneRow(n - 1): Next n
Next i
ws.Range("A2:Q" & CStr(oldLast)).ClearContents
ws.Range("C2:C" & CStr(rows.count + 1)).NumberFormat = "@"
ws.Range("I2:I" & CStr(rows.count + 1)).NumberFormat = "@"
ws.Range("A2").Resize(rows.count, 17).Value2 = outArr
ws.Range("B2:B" & CStr(rows.count + 1)).NumberFormat = "#,##0"
ws.Range("F2:F" & CStr(rows.count + 1)).NumberFormat = "#,##0"
ws.Range("H2:H" & CStr(rows.count + 1)).NumberFormat = "#,##0"
ws.Range("N2:N" & CStr(rows.count + 1)).NumberFormat = "#,##0"
ws.Range("Q2:Q" & CStr(rows.count + 1)).NumberFormat = "#,##0"
DownloadPostingForOrder = True
Application.StatusBar = "POSTING: da tai " & rows.count & " dong."
If showMessage Then MsgBox "Da tai " & rows.count & " dong POSTING.", vbInformation, "TAI POSTING"
Exit Function
Fail:
Application.StatusBar = False
LastPostingError = "Stage=" & stage & "; Err " & CStr(Err.Number) & ": " & Err.Description
ModuleLog.WriteLog "LOI", "TAI POSTING [" & stage & "]: " & LastPostingError
If showMessage Then MsgBox "TAI POSTING loi tai " & stage & ":" & vbCrLf & Err.Description & vbCrLf & vbCrLf & "POSTING cu duoc giu nguyen.", vbCritical, "TAI POSTING"
End Function
Private Function POST_LastDataRow(ByVal ws As Worksheet, Optional ByVal maxCol As Long = 17) As Long
Dim c As Long, candidate As Long, lastRow As Long
lastRow = 1
For c = 1 To maxCol
    candidate = ws.Cells(ws.rows.count, c).End(xlUp).row
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
On Error Resume Next
Set ws = ThisWorkbook.Worksheets("ORDER_QUERY")
If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("ORDER")
On Error GoTo 0
If ws Is Nothing Then Set POST_OrderProducts = d: Exit Function
c = FindHeaderCol(ws, Array("Product", "Material", "Ma NVL"))
If c = 0 Then Set POST_OrderProducts = d: Exit Function
lr = ws.Cells(ws.rows.count, c).End(xlUp).row
If lr < 2 Then Set POST_OrderProducts = d: Exit Function
a = ws.Range(ws.Cells(2, c), ws.Cells(lr, c)).Value2
If lr = 2 Then
    s = Trim$(CStr(a & ""))
    If s <> "" Then d(s) = True
Else
    For i = 1 To UBound(a, 1): s = Trim$(CStr(a(i, 1) & "")): If s <> "" Then d(s) = True
    Next i
End If
Set POST_OrderProducts = d
End Function
Private Function POST_OpenAdvancedProductResult(ByVal ses As Object, ByVal products As Object, ByRef grid As Object, ByRef msg As String) As Boolean
Dim idAdvToggle As String, idAdvReset As String, idAdvStart As String, idProductMulti As String, idOwnerLow As String, idStockGrid As String
Dim idPopupClose As String, idPopupClear As String, idPopupPaste As String, idPopupOK As String
Dim ctl As Object, pop As Object, ws As Worksheet, c As Long, lr As Long
Dim ownerValue As String, statusText As String, statusType As String
On Error GoTo Fail
msg = "": Set grid = Nothing
idAdvToggle = POST_ExactID("POST_ADVANCED_TOGGLE_ID", "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/btnGV_BUTTON_TEXT")
idAdvReset = POST_ExactID("POST_ADVANCED_RESET_ID", "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_AS_BUTTONS:/SCMB/SAPLSERVICES:1000/btnCMD_CLEAR_ADVANCED")
idAdvStart = POST_ExactID("POST_ADVANCED_START_ID", "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_AS_BUTTONS:/SCMB/SAPLSERVICES:1000/btnCMD_START_ADVANCED")
idProductMulti = POST_ExactID("POST_PRODUCT_MULTI_ID", "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_SELECTION_SCREEN:/SCWM/SAPLUI_POST:2500/btn%_P_MATNR_%_APP_%-VALU_PUSH")
idOwnerLow = POST_ExactID("POST_OWNER_ID", "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_ADVANCED_SEARCH:/SCWM/SAPLUI_POST:2300/subSUB_SELECTION_SCREEN:/SCWM/SAPLUI_POST:2500/ctxtP_OWNER-LOW")
idStockGrid = POST_ExactID("POST_GRID_ID", "wnd[0]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_POST:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_POST:2210/subSUB_OIP_1_DATA:/SCWM/SAPLUI_POST:2211/cntlCONTAINER_ALV_OIP_1/shellcont/shell")
idPopupClose = POST_ExactID("POST_PRODUCT_POPUP_CLOSE_ID", "wnd[1]/tbar[0]/btn[12]")
idPopupClear = POST_ExactID("POST_PRODUCT_POPUP_CLEAR_ID", "wnd[1]/tbar[0]/btn[16]")
idPopupPaste = POST_ExactID("POST_PRODUCT_PASTE_ID", "wnd[1]/tbar[0]/btn[24]")
idPopupOK = POST_ExactID("POST_PRODUCT_POPUP_OK_ID", "wnd[1]/tbar[0]/btn[8]")
If idAdvToggle = "" Or idAdvReset = "" Or idAdvStart = "" Or idProductMulti = "" Or idOwnerLow = "" Or idStockGrid = "" Or idPopupClose = "" Or idPopupClear = "" Or idPopupPaste = "" Or idPopupOK = "" Then Err.Raise vbObjectError + 2947, , "STOP_CONTROL_NOT_FOUND: POST exact CONFIG IDs are incomplete."
On Error Resume Next
Set pop = ses.FindById("wnd[1]", False)
On Error GoTo Fail
If Not pop Is Nothing Then
    Set ctl = ses.FindById(idPopupClose, False)
    If ctl Is Nothing Then Err.Raise vbObjectError + 2948, , "STOP_CONTROL_NOT_FOUND: POST_PRODUCT_POPUP_CLOSE_ID"
    ctl.Press
End If
If SAPRun_IsRunMode() Then
    If Not SAPX_TcodeIs(ses, ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST")) Then _
        Err.Raise vbObjectError + 2952, , "STOP_SESSION_CONTEXT_MISMATCH: exact /SCWM/POST session is not active."
Else
    SAPX_Navigate ses, ModuleHelper.ToolConfig("POST_TCODE", "/SCWM/POST")
End If
Set ctl = SAPX_FindById(ses, idOwnerLow)
If ctl Is Nothing Then
    Set ctl = SAPX_FindById(ses, idAdvToggle)
    If ctl Is Nothing Then Err.Raise vbObjectError + 2950, , "Khong tim thay nut Open Advanced Search."
    ctl.Press
    If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 2951, , "SAP Busy khi mo Advanced Search."
End If
Set ctl = SAPX_FindById(ses, idAdvReset)
If ctl Is Nothing Then Err.Raise vbObjectError + 2949, , "STOP_CONTROL_NOT_FOUND: POST_ADVANCED_RESET_ID"
ctl.Press: SAPX_WaitIdle ses, 10
ownerValue = ModuleHelper.ToolConfig("POST_OWNER", "3093")
Set ctl = SAPX_FindById(ses, idOwnerLow)
If ctl Is Nothing Then Err.Raise vbObjectError + 2952, , "Khong tim thay Owner trong Advanced Search sau khi mo lai."
SAPX_SetValue ctl, ownerValue
If StrComp(SAPX_ControlValue(ctl), ownerValue, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 2953, , "Owner readback khong khop " & ownerValue
Set ctl = SAPX_FindById(ses, idProductMulti)
If ctl Is Nothing Then Err.Raise vbObjectError + 2954, , "Khong tim thay Multiple Selection cua Product."
ctl.Press
If Not SAPX_WaitIdle(ses, 10) Then Err.Raise vbObjectError + 2955, , "SAP Busy khi mo Product list."
Set pop = SAPX_FindById(ses, "wnd[1]")
If pop Is Nothing Then Err.Raise vbObjectError + 2956, , "Khong mo duoc popup Product list."
On Error Resume Next
Set ctl = ses.FindById(idPopupClear, False)
On Error GoTo Fail
If ctl Is Nothing Then Err.Raise vbObjectError + 2950, , "STOP_CONTROL_NOT_FOUND: POST_PRODUCT_POPUP_CLEAR_ID"
ctl.Press
On Error Resume Next
Set ws = ThisWorkbook.Worksheets("ORDER_QUERY")
If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("ORDER")
On Error GoTo Fail
If ws Is Nothing Then Err.Raise vbObjectError + 2956, , "Khong tim thay ORDER_QUERY/ORDER de nap Product."
c = FindHeaderCol(ws, Array("Product", "Material", "Ma NVL"))
If c = 0 Then Err.Raise vbObjectError + 2957, , "ORDER khong co cot Product."
lr = ws.Cells(ws.rows.count, c).End(xlUp).row
If lr < 2 Then Err.Raise vbObjectError + 2957, , "ORDER khong co list Product de copy."
ws.Range(ws.Cells(2, c), ws.Cells(lr, c)).Copy
On Error Resume Next
Set ctl = ses.FindById(idPopupPaste, False)
On Error GoTo Fail
If ctl Is Nothing Then Err.Raise vbObjectError + 2958, , "Popup Product khong co nut Paste from Clipboard."
ctl.Press
SAPX_WaitIdle ses, 10
Set ctl = SAPX_FindById(ses, idPopupOK)
If ctl Is Nothing Then Err.Raise vbObjectError + 2959, , "Popup Product khong co nut OK."
ctl.Press
If Not SAPX_WaitIdle(ses, 15) Then Err.Raise vbObjectError + 2960, , "SAP Busy khi nhan list Product."
Set ctl = SAPX_FindById(ses, idAdvStart)
If ctl Is Nothing Then Err.Raise vbObjectError + 2961, , "Khong tim thay nut Advanced Search."
ctl.Press
If Not SAPX_WaitIdle(ses, 60) Then Err.Raise vbObjectError + 2962, , "SAP Busy qua lau khi tai " & products.count & " Product."
statusText = SAPX_StatusText(ses): statusType = SAPX_StatusType(ses)
If statusType = "E" Or statusType = "A" Then Err.Raise vbObjectError + 2963, , "SAP: " & statusText
Set grid = SAPX_FindById(ses, idStockGrid)
If grid Is Nothing And SAPRun_IsMapMode() Then Set grid = SAPX_WaitGrid(ses, 25)
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
ses.FindById(ModuleHelper.ToolConfig("POST_EXECUTE_ID", "wnd[0]/tbar[1]/btn[8]"), False).Press
If Err.Number <> 0 Then Err.Clear: ses.FindById("wnd[0]").sendVKey 8
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
If POST_ObjectText = "" Then POST_ObjectText = Trim$(CStr(obj.Value))
If POST_ObjectText = "" Then POST_ObjectText = Trim$(CStr(obj.key))
On Error GoTo 0
End Function
Private Function POST_FindProductComboRec(ByVal obj As Object, ByVal depth As Long) As Object
Dim typ As String, txt As String, i As Long, ch As Object, f As Object
If obj Is Nothing Or depth > 16 Then Exit Function
On Error Resume Next: typ = UCase$(CStr(obj.Type)): txt = UCase$(POST_ObjectText(obj)): On Error GoTo 0
If InStr(typ, "COMBOBOX") > 0 And InStr(txt, "PRODUCT") > 0 Then Set POST_FindProductComboRec = obj: Exit Function
On Error Resume Next
For i = 0 To obj.children.count - 1
    Set ch = SAPX_CollectionObjectAt(obj.children, i): Set f = POST_FindProductComboRec(ch, depth + 1)
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
For i = 0 To obj.children.count - 1: Set ch = SAPX_CollectionObjectAt(obj.children, i): POST_ScoreProductInputRec ch, combo, best, bestScore, depth + 1: Next i
On Error GoTo 0
End Sub
Private Sub POST_FindProductControls(ByVal ses As Object, ByRef crit As Object, ByRef productCtl As Object)
Dim root As Object, score As Double
Set crit = SAPX_FindById(ses, ModuleHelper.ToolConfig("POST_CRITERION_ID", ""))
Set productCtl = SAPX_FindById(ses, ModuleHelper.ToolConfig("POST_PRODUCT_ID", ""))
On Error Resume Next: Set root = ses.FindById("wnd[0]", False): On Error GoTo 0
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
On Error Resume Next: SAPX_StatusText = Trim$(CStr(ses.FindById("wnd[0]/sbar", False).text)): On Error GoTo 0
End Function
Private Function SAPX_StatusType(ByVal ses As Object) As String
On Error Resume Next: SAPX_StatusType = UCase$(Trim$(CStr(ses.FindById("wnd[0]/sbar", False).messageType))): On Error GoTo 0
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
Set root = ses.FindById("wnd[0]", False)
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
Set root = ses.FindById("wnd[0]", False)
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
btn.Press
SAPX_WaitIdle ses, 10
On Error Resume Next
Set root = ses.FindById("wnd[1]", False)
If root Is Nothing Then Set root = ses.FindById("wnd[0]", False)
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
btn.Press
SAPX_WaitIdle ses, 10
Set root = ses.FindById("wnd[1]", False)
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
For i = 0 To obj.children.count - 1
    Set ch = SAPX_CollectionObjectAt(obj.children, i)
    SAPX_DumpControlIds ch, dumpText, depth + 1
Next i
Err.Clear
On Error GoTo 0
End Sub
Private Sub VL6_RunPendingReport(Optional ByVal showMessage As Boolean = True)
Dim ses As Object, attempt As Long, failNo As Long, failText As String
attempt = 1
RetryVL06O:
On Error GoTo AttemptFail
If attempt = 1 Then
    Set ses = SAPX_GetSessionForTcode("VL06O")
Else
    Set ses = SAPX_ReopenFreshSessionForTcode(ses, "VL06O")
End If
If ses Is Nothing Then Err.Raise vbObjectError + 2920, , "Khong ket noi duoc SAP session VL06O."
Application.StatusBar = "VL06O: lan thu " & CStr(attempt) & "/3..."
VL6_RunPendingReportCore ses, showMessage
Exit Sub
AttemptFail:
failNo = Err.Number: failText = Err.Description
Application.StatusBar = False
If attempt < 3 Then
    ModuleLog.WriteLog "CANH BAO", "VL06O lan " & CStr(attempt) & " loi: " & failText & ". Dong tab cu, mo tab moi va thu lai."
    attempt = attempt + 1
    Err.Clear
    On Error GoTo 0
    GoTo RetryVL06O
End If
Err.Raise failNo, , "VL06O that bai sau 1 lan dau va 2 lan mo tab moi. Lan cuoi: " & failText
End Sub
Private Sub VL6_RunPendingReportCore(ByVal ses As Object, Optional ByVal showMessage As Boolean = True)
Dim grid As Object, ctl As Object, fromD As Date, toD As Date
Dim cGM As String, cDel As String, cMat As String, cQty As String, cBatch As String, cSloc As String, cShip As String, colDate As String, cTime As String, cBy As String, cShpt As String
Dim shipValue As String, slocValue As String, r As Long, rowCount As Long, outCount As Long, arr() As Variant, shipPrefilled As Boolean
Dim lastDel As String, lastShip As String, lastSloc As String, lastCreated As String, vDel As String, vShip As String, vSloc As String, vMat As String, vQty As String, vCreated As String, d As Date
Dim ws As Worksheet, oldLast As Long, clearLast As Long, oldCount As Long, unique As Object, i As Long, finalArr() As Variant, sortedData As Variant, j As Long
Dim failNo As Long, failText As String, shipFilterOK As Boolean, slocFilterOK As Boolean, slocVerifyDetail As String
Dim vlColIds As Collection, vlNormIds As Object, vlNormCaps As Object
On Error GoTo Fail
shipValue = ModuleHelper.ToolConfig("VL06O_SHIP_TO", "3093"): slocValue = ModuleHelper.ToolConfig("VL06O_SLOC", "1005")
fromD = DateAdd("m", -ModuleHelper.ToolConfigLong("VL06O_LOOKBACK_MONTHS", 1, 1, 12), Date): toD = Date
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
On Error Resume Next: ses.FindById(ModuleHelper.ToolConfig("VL06O_EXECUTE_ID", "wnd[0]/tbar[1]/btn[8]"), False).Press: If Err.Number <> 0 Then Err.Clear: ses.FindById("wnd[0]").sendVKey 8
On Error GoTo Fail
If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 2924, , "SAP Busy qua lau sau Execute VL06O."
Set grid = SAPX_WaitGrid(ses, 20): If grid Is Nothing Then Err.Raise vbObjectError + 2925, , "Khong tim thay VL06O result grid."
ResultReady:
SAPX_BuildColumnCatalogFast grid, vlColIds, vlNormIds, vlNormCaps
If SAPX_ResolveColumnFast("", "MATNR,MATERIAL", False, vlColIds, vlNormIds, vlNormCaps) = "" Then
    VL6_PressItemView ses
    Set grid = SAPX_WaitGrid(ses, 20): If grid Is Nothing Then Err.Raise vbObjectError + 2926, , "Khong mo duoc Item View."
    SAPX_BuildColumnCatalogFast grid, vlColIds, vlNormIds, vlNormCaps
End If
cGM = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_GM", "WBSTA"), "WBSTA,WBSTK,GM,GOODS MOVEMENT STATUS", False, vlColIds, vlNormIds, vlNormCaps)
cDel = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_DELIVERY", "VBELN"), "VBELN,DELIVERY,OUTBOUND DELIVERY", True, vlColIds, vlNormIds, vlNormCaps)
cMat = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_MATERIAL", "MATNR"), "MATNR,MATERIAL", True, vlColIds, vlNormIds, vlNormCaps)
cQty = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_QTY", "LFIMG"), "LFIMG,DLV.QTY,DELIVERY QTY", True, vlColIds, vlNormIds, vlNormCaps)
cBatch = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_BATCH", "S_CHARG"), "S_CHARG,CHARG,BATCH", False, vlColIds, vlNormIds, vlNormCaps)
cSloc = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_SLOC", "LGORT"), "LGORT,SLOC,STORAGE LOCATION", True, vlColIds, vlNormIds, vlNormCaps)
cShip = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_SHIPTO", "KUNNR"), "KUNNR,KUNWE,SHIP-TO PARTY,SHIP TO", True, vlColIds, vlNormIds, vlNormCaps)
colDate = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_CREATED_ON", "ERDAT"), "ERDAT,CREATED ON,CREATION DATE", True, vlColIds, vlNormIds, vlNormCaps)
cTime = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_CREATED_TIME", "ERZET"), "ERZET,TIME,CREATED TIME", False, vlColIds, vlNormIds, vlNormCaps)
cBy = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_CREATED_BY", "ERNAM"), "ERNAM,CREATED BY", False, vlColIds, vlNormIds, vlNormCaps)
cShpt = SAPX_ResolveColumnFast(ModuleHelper.ToolConfig("VL06O_COL_SHPT", "VSTEL"), "VSTEL,SHIPPING POINT", False, vlColIds, vlNormIds, vlNormCaps)
Application.StatusBar = "VL06O: bat buoc loc lai SLoc=" & slocValue & "..."
SAPX_ClearGridFiltersFast ses, grid
Set grid = SAPX_WaitGrid(ses, 20)
If grid Is Nothing Then Err.Raise vbObjectError + 2926, , "VL06O grid mat sau khi xoa filter cu."
cSloc = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_SLOC", "LGORT"), "LGORT,SLOC,STORAGE LOCATION", True)
Application.StatusBar = "VL06O: dang mo Filter va dien SLoc=" & slocValue & "..."
If Not SAPX_ApplyGridFilter(ses, grid, cSloc, slocValue, "SLOC,STORAGE LOCATION,LGORT") Then Err.Raise vbObjectError + 2929, , "Khong dien duoc filter SLoc=" & slocValue & " tai Item View. " & mSAPXFilterError
Set grid = SAPX_WaitGrid(ses, 20)
If grid Is Nothing Then Err.Raise vbObjectError + 2930, , "Filter SLoc=" & slocValue & " da Apply nhung khong bat lai duoc VL06O grid."
cSloc = SAPX_ResolveColumn(grid, ModuleHelper.ToolConfig("VL06O_COL_SLOC", "LGORT"), "LGORT,SLOC,STORAGE LOCATION", True)
slocVerifyDetail = ""
If Not SAPX_VerifyFilterDetailed(grid, cSloc, slocValue, slocVerifyDetail) Then Err.Raise vbObjectError + 2930, , "Filter SLoc=" & slocValue & " khong co hieu luc tai Item View. " & slocVerifyDetail
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
    oldLast = ws.Cells(ws.rows.count, j).End(xlUp).row
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
    ws.Range("A5:K" & CStr(outCount + 5)).Sort Key1:=ws.Range("H6"), Order1:=xlAscending, Header:=xlYes
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
ws.Range("A2").Value = "Lan kiem tra": ws.Range("B2").Value = Now: ws.Range("B2").NumberFormat = "dd.mm.yyyy hh:mm:ss"
ws.Range("D2").Value = "So item": ws.Range("E2").Value = outCount: ws.Range("G2").Value = "So OB": ws.Range("H2").Value = unique.count: ws.Range("J2").Value = "OB >=2 ngay": ws.Range("K2").Value = oldCount
ws.Range("A3").Value = "Date: " & Format$(fromD, "dd.mm.yyyy") & " - " & Format$(toD, "dd.mm.yyyy") & " | GM=A | Ship-to=" & shipValue & " | SLoc=" & slocValue & " | Khong dien Shipping Point/Receiving Pt."
Application.StatusBar = False
If showMessage Then MsgBox "Da check OB treo: " & unique.count & " OB / " & outCount & " item.", vbInformation, "VL06O"
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
If Not btn Is Nothing Then btn.Press: SAPX_WaitIdle ses, 15
Set ctl = SAPX_FindById(ses, ModuleHelper.ToolConfig("VL06O_DATE_FROM_ID", "wnd[0]/usr/ctxtIT_WADAT-LOW"))
If ctl Is Nothing Then Set ctl = SAPX_FindByTokens(ses, "WADAT-LOW|WADAT_LOW|WADATLOW", True)
If Not ctl Is Nothing Then Exit Sub
Set btn = SAPX_FindButtonText(ses.FindById("wnd[0]/usr", False), "LIST OUTBOUND DELIVERIES", 0)
If Not btn Is Nothing Then btn.Press: SAPX_WaitIdle ses, 15
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
If Not btn Is Nothing Then btn.Press Else Set btn = SAPX_FindButtonText(ses.FindById("wnd[0]", False), "ITEM VIEW", 0): If Not btn Is Nothing Then btn.Press
On Error GoTo 0
SAPX_WaitIdle ses, 15
End Sub
Public Function CF_OpenTCodeSession(ByVal tcode As String, ByRef errorText As String) As Object
On Error GoTo Failed
errorText = ""
Set CF_OpenTCodeSession = SAPX_GetSessionForTcode(tcode)
If CF_OpenTCodeSession Is Nothing Then errorText = "Khong nhan duoc SAP session cho TCode=" & tcode
Exit Function
Failed:
errorText = "Err " & CStr(Err.Number) & ": " & Err.Description
Set CF_OpenTCodeSession = Nothing
End Function

Public Function CF_DumpTCodeControls(ByVal tcode As String) As String
Dim ses As Object, root As Object, dumpText As String, openMsg As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession(tcode, openMsg)
If ses Is Nothing Then
    CF_DumpTCodeControls = "OPEN_FAILED " & openMsg
    Exit Function
End If
Set root = SAPX_FindById(ses, "wnd[0]")
If root Is Nothing Then
    CF_DumpTCodeControls = "ROOT_NOT_FOUND"
    Exit Function
End If
SAPX_DumpControlIds root, dumpText, 0
CF_DumpTCodeControls = dumpText
Exit Function
Failed:
CF_DumpTCodeControls = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpAllSAPWindows(ByVal tcode As String) As String
Dim ses As Object, root As Object, dumpText As String, openMsg As String
Dim i As Long
On Error GoTo Failed
Set ses = CF_OpenTCodeSession(tcode, openMsg)
If ses Is Nothing Then
    CF_DumpAllSAPWindows = "OPEN_FAILED " & openMsg
    Exit Function
End If
For i = 0 To 5
    Set root = SAPX_FindById(ses, "wnd[" & CStr(i) & "]")
    If Not root Is Nothing Then
        dumpText = dumpText & "=== WINDOW wnd[" & CStr(i) & "] ===" & vbLf
        SAPX_DumpControlIds root, dumpText, 0
    End If
    Set root = Nothing
Next i
CF_DumpAllSAPWindows = dumpText
Exit Function
Failed:
CF_DumpAllSAPWindows = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpMONTree() As String
Dim ses As Object, tree As Object, keysObj As Object, key As Variant
Dim openMsg As String, txt As String, n As Long, i As Long
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpMONTree = "OPEN_FAILED " & openMsg: Exit Function
Set tree = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[0]/shell")
If tree Is Nothing Then CF_DumpMONTree = "TREE_NOT_FOUND": Exit Function
On Error Resume Next
Err.Clear
Set keysObj = tree.GetAllNodeKeys
If Err.Number <> 0 Or keysObj Is Nothing Then
    CF_DumpMONTree = "GET_ALL_NODE_KEYS_ERR " & CStr(Err.Number) & ": " & Err.Description
    Err.Clear
    On Error GoTo 0
    Exit Function
End If
On Error GoTo Failed
n = SAPX_CollectionCount(keysObj)
For i = 0 To n - 1
    key = SAPX_CollectionValueAt(keysObj, i)
    txt = "KEY=" & CStr(key)
    On Error Resume Next
    Err.Clear
    txt = txt & " | TEXT=" & CStr(tree.GetNodeTextByKey(CStr(key)))
    If Err.Number <> 0 Then txt = txt & " | TEXT_ERR=" & CStr(Err.Number) & ": " & Err.Description: Err.Clear
    On Error GoTo 0
    CF_DumpMONTree = CF_DumpMONTree & txt & vbLf
    If i >= 299 Then Exit For
Next i
If n = 0 Then CF_DumpMONTree = "NO_TREE_NODES"
Exit Function
Failed:
CF_DumpMONTree = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpMONState() As String
Dim ses As Object, wnd As Object, st As Object, openMsg As String, dumpText As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpMONState = "OPEN_FAILED " & openMsg: Exit Function
Set wnd = SAPX_FindById(ses, "wnd[0]")
If wnd Is Nothing Then CF_DumpMONState = "WINDOW_NOT_FOUND": Exit Function
Set st = wnd.DumpState("")
If st Is Nothing Then CF_DumpMONState = "DUMPSTATE_EMPTY": Exit Function
SAPX_DumpAnyCollection st, dumpText, 0
CF_DumpMONState = dumpText
Exit Function
Failed:
CF_DumpMONState = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub SAPX_DumpAnyCollection(ByVal obj As Object, ByRef dumpText As String, ByVal depth As Long)
Dim i As Long, n As Long, item As Variant, child As Object, lineText As String
If obj Is Nothing Or depth > 5 Or Len(dumpText) > 25000 Then Exit Sub
On Error Resume Next
n = CLng(obj.count)
If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Sub
On Error GoTo 0
For i = 0 To n - 1
    If Len(dumpText) > 25000 Then Exit For
    On Error Resume Next
    Err.Clear
    Set child = obj.ElementAt(i)
    If Err.Number <> 0 Then Err.Clear: Set child = obj.item(i)
    If Err.Number = 0 And Not child Is Nothing Then
        dumpText = dumpText & String$(depth * 2, " ") & "[" & CStr(i) & "] Object" & vbLf
        SAPX_DumpAnyCollection child, dumpText, depth + 1
    Else
        Err.Clear
        item = obj.ElementAt(i)
        If Err.Number <> 0 Then Err.Clear: item = obj.item(i)
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit For
        lineText = CStr(item & "")
        dumpText = dumpText & String$(depth * 2, " ") & "[" & CStr(i) & "] " & lineText & vbLf
    End If
    Set child = Nothing
    On Error GoTo 0
Next i
End Sub

Public Function CF_DumpMONUIA() As String
Const UIA_TreeScope_Descendants As Long = 4
Const UIA_NamePropertyId As Long = 30005
Dim uia As Object, root As Object, cond As Object, win As Object, all As Object, el As Object
Dim i As Long, n As Long, lineText As String, nameText As String, aidText As String, classText As String, typeNo As Long
On Error GoTo Failed
On Error Resume Next
Set uia = CreateObject("CUIAutomation8")
If uia Is Nothing Then Set uia = CreateObject("CUIAutomation")
On Error GoTo Failed
If uia Is Nothing Then CF_DumpMONUIA = "UIA_NOT_AVAILABLE": Exit Function
Set root = uia.GetRootElement
Set cond = uia.CreatePropertyCondition(UIA_NamePropertyId, "Warehouse Management Monitor SAP - Whse. VF02 (Time Zone UTC+7)")
Set win = root.FindFirst(UIA_TreeScope_Descendants, cond)
If win Is Nothing Then CF_DumpMONUIA = "SAP_WINDOW_NOT_FOUND": Exit Function
Set all = win.FindAll(UIA_TreeScope_Descendants, uia.CreateTrueCondition)
n = all.Length
For i = 0 To n - 1
    If i > 1000 Then Exit For
    Set el = all.GetElement(i)
    On Error Resume Next
    nameText = CStr(el.CurrentName & "")
    aidText = CStr(el.CurrentAutomationId & "")
    classText = CStr(el.CurrentClassName & "")
    typeNo = CLng(el.CurrentControlType)
    On Error GoTo Failed
    If nameText <> "" Or aidText <> "" Or typeNo = 50004 Then
        lineText = "N=" & nameText & " | A=" & aidText & " | C=" & classText & " | T=" & CStr(typeNo)
        CF_DumpMONUIA = CF_DumpMONUIA & lineText & vbLf
    End If
Next i
If CF_DumpMONUIA = "" Then CF_DumpMONUIA = "SAP_WINDOW_FOUND_BUT_NO_UIA_CHILDREN"
Exit Function
Failed:
CF_DumpMONUIA = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpMONShellState() As String
Dim ses As Object, ctl As Object, st As Object, openMsg As String, dumpText As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpMONShellState = "OPEN_FAILED " & openMsg: Exit Function
Set ctl = SAPX_FindById(ses, "wnd[0]/usr/shell")
If ctl Is Nothing Then CF_DumpMONShellState = "SHELL_NOT_FOUND": Exit Function
Set st = ctl.DumpState("")
If st Is Nothing Then CF_DumpMONShellState = "DUMPSTATE_EMPTY": Exit Function
SAPX_DumpAnyCollection st, dumpText, 0
CF_DumpMONShellState = dumpText
Exit Function
Failed:
CF_DumpMONShellState = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpMONObjectTree() As String
Dim ses As Object, openMsg As String, props As Variant
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpMONObjectTree = "OPEN_FAILED " & openMsg: Exit Function
props = Array("Id", "Type", "Name", "Text", "Tooltip", "AccText", "ScreenLeft", "ScreenTop", "Width", "Height")
CF_DumpMONObjectTree = CStr(ses.GetObjectTree("wnd[0]", props))
Exit Function
Failed:
CF_DumpMONObjectTree = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpMONWindowTree(ByVal windowId As String) As String
Dim ses As Object, openMsg As String, props As Variant
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpMONWindowTree = "OPEN_FAILED " & openMsg: Exit Function
props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "ScreenLeft", "ScreenTop", "Width", "Height")
CF_DumpMONWindowTree = CStr(ses.GetObjectTree(windowId, props))
Exit Function
Failed:
CF_DumpMONWindowTree = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpMONTreeColumns() As String
CF_DumpMONTreeColumns = "SKIPPED: SAP.TableTreeControl.1 does not expose a stable column collection in this MON layout. Use CF_MapMONSAP to map controls by ID/Type/Text." End Function
End Function

Public Function CF_MON_OpenSelectionMap() As String
Dim ses As Object, menuItem As Object, openMsg As String, treeText As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_MON_OpenSelectionMap = "OPEN_FAILED " & openMsg: Exit Function
Set menuItem = SAPX_FindById(ses, "wnd[0]/mbar/menu[1]/menu[5]")
If menuItem Is Nothing Then CF_MON_OpenSelectionMap = "SET_SELECTION_CRITERIA_MENU_NOT_FOUND": Exit Function
menuItem.Select
SAPX_WaitIdle ses, 10
On Error Resume Next
treeText = CStr(ses.GetObjectTree("wnd[0]", Array("Id", "Type", "Name", "Text", "Tooltip", "ScreenLeft", "ScreenTop", "Width", "Height")))
If Err.Number <> 0 Then Err.Clear
On Error GoTo Failed
CF_MON_OpenSelectionMap = treeText
Exit Function
Failed:
CF_MON_OpenSelectionMap = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Sub CF_MapMONSAP()
Dim ses As Object, root As Object, ws As Worksheet, NextRow As Long, openMsg As String
Dim i As Long, headers As Variant, h As Long
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then Err.Raise vbObjectError + 2981, , "OPEN_MON_FAILED: " & openMsg
On Error Resume Next
Set ws = ThisWorkbook.Worksheets("SAP_MAP")
On Error GoTo Failed
If ws Is Nothing Then
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.Name = "SAP_MAP"
End If
headers = Array("Time", "Window", "ID", "Type", "Name", "Text", "Tooltip", "Changeable", "Depth")
If Trim$(CStr(ws.Cells(1, 1).Value2)) = "" Then
    For h = LBound(headers) To UBound(headers): ws.Cells(1, h + 1).Value = headers(h): Next h
    ws.rows(1).Font.Bold = True
End If
NextRow = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
For i = 0 To 3
    Set root = SAPX_FindById(ses, "wnd[" & CStr(i) & "]")
    If Not root Is Nothing Then CF_MON_DumpControlTree root, ws, NextRow, 0, "wnd[" & CStr(i) & "]"
Next i
ws.columns("A:I").AutoFit
MsgBox "Da map SAP MON vao sheet SAP_MAP. Rows=" & CStr(NextRow - 2), vbInformation, "SAP MAP MON"
Exit Sub
Failed:
MsgBox "CF_MapMONSAP failed: Err " & CStr(Err.Number) & ": " & Err.Description, vbExclamation, "SAP MAP MON"
End Sub

Private Sub CF_MON_DumpControlTree(ByVal obj As Object, ByVal ws As Worksheet, ByRef NextRow As Long, ByVal depth As Long, ByVal windowName As String)
Dim children As Object, child As Object, i As Long, n As Long
If obj Is Nothing Or depth > 18 Then Exit Sub
On Error Resume Next
ws.Cells(NextRow, 1).Value = Now
ws.Cells(NextRow, 2).Value = windowName
ws.Cells(NextRow, 3).Value = CStr(obj.id & "")
ws.Cells(NextRow, 4).Value = CStr(obj.Type & "")
ws.Cells(NextRow, 5).Value = CStr(obj.Name & "")
ws.Cells(NextRow, 6).Value = CStr(obj.text & "")
ws.Cells(NextRow, 7).Value = CStr(obj.Tooltip & "")
ws.Cells(NextRow, 8).Value = CStr(obj.Changeable & "")
ws.Cells(NextRow, 9).Value = depth
NextRow = NextRow + 1
Set children = obj.children
If Not children Is Nothing Then n = CLng(children.count)
On Error GoTo 0
For i = 0 To n - 1
    Set child = SAPX_CollectionObjectAt(children, i)
    If Not child Is Nothing Then CF_MON_DumpControlTree child, ws, NextRow, depth + 1, windowName
    Set child = Nothing
Next i
End Sub

Public Function CF_MON_SetLEDelivery(ByVal ses As Object, ByVal ob As String, ByRef detailText As String) As Boolean
Dim menuItem As Object, root As Object, ctl As Object, btn As Object
Dim actualText As String, windowName As String, mapDump As String, openDetail As String
Dim stepText As String, tree As Object, nodeKey As String, runMessage As String
On Error GoTo Failed
detailText = ""
If ses Is Nothing Then detailText = "SAP_SESSION_NOTHING": Exit Function

If SAPRun_IsRunMode() Then
    stepText = "OPEN_MAPPED_OUTBOUND_NODE"
    Set tree = SAPRun_FindExact(ses, "MON_TREE_ID", runMessage)
    If tree Is Nothing Then detailText = runMessage: Exit Function
    nodeKey = Trim$(ModuleHelper.ToolConfig("MON_OUTBOUND_NODE_KEY", ""))
    If nodeKey = "" Then detailText = "STOP_CONTROL_NOT_FOUND: CONFIG MON_OUTBOUND_NODE_KEY is blank.": Exit Function
    On Error Resume Next
    Err.Clear
    CallByName tree, "DoubleClickNode", VbMethod, nodeKey
    If Err.Number <> 0 Then
        detailText = "STOP_CONTROL_NOT_FOUND: MON_OUTBOUND_NODE_KEY=" & nodeKey & "; Err " & CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        On Error GoTo Failed
        Exit Function
    End If
    On Error GoTo Failed
    SAPX_WaitIdle ses, 10

    stepText = "FIND_MAPPED_LE_DELIVERY_FIELD"
    Set ctl = SAPRun_FindExact(ses, "MON_OB_INPUT_ID", runMessage)
    If ctl Is Nothing Then detailText = runMessage: Exit Function
    SAPX_SetValue ctl, Trim$(ob)
    actualText = SAPX_ControlValue(ctl)
    If StrComp(Trim$(actualText), Trim$(ob), vbTextCompare) <> 0 Then
        detailText = "STOP_FILTER_NOT_PROVEN: MON_OB_INPUT_ID read-back mismatch; Actual=" & actualText & "; Expected=" & ob
        Exit Function
    End If

    stepText = "PRESS_MAPPED_EXECUTE"
    Set btn = SAPRun_FindExact(ses, "MON_EXECUTE_ID", runMessage)
    If btn Is Nothing Then detailText = runMessage: Exit Function
    btn.Press
    SAPX_WaitIdle ses, 30
    detailText = "MAP_OK; ID=" & CStr(ctl.id) & "; Type=" & CStr(ctl.Type) & "; Value=" & actualText & "; RUN_EXACT=TRUE"
    CF_MON_SetLEDelivery = True
    Exit Function
End If

' First open the exact MON node by SAP scripting. The menu fallback is kept
' for layouts where the left tree is collapsed or the node is hidden.
stepText = "OPEN_OUTBOUND_DELIVERY_ORDER"
If Not CF_MON_OpenOutboundDeliveryOrder(ses, openDetail) Then
    On Error Resume Next
    Set menuItem = ses.FindById("wnd[0]/mbar/menu[1]/menu[5]", False)
    If Not menuItem Is Nothing Then menuItem.Select
    Err.Clear
    On Error GoTo Failed
End If
SAPX_WaitIdle ses, 10

stepText = "FIND_SELECTION_WINDOW"
Set root = SAPX_FindById(ses, "wnd[1]")
If Not root Is Nothing Then windowName = "wnd[1]"
If root Is Nothing Then Set root = SAPX_FindById(ses, "wnd[0]"): windowName = "wnd[0]"
If root Is Nothing Then detailText = "MON_ACTIVE_WINDOW_NOT_FOUND": Exit Function

' The selection screen has a text field named S_ERP_H-LOW next to the
' visible label LE Delivery.  S_DOCNO-LOW belongs to the separate Outb.
' Delivery Order criterion and must never receive the OB here.
stepText = "FIND_LE_DELIVERY_FIELD"
Set ctl = SAPX_FindById(ses, "wnd[1]/usr/txtS_ERP_H-LOW")
If ctl Is Nothing Then Set ctl = SAPX_FindById(ses, "wnd[0]/usr/txtS_ERP_H-LOW")
If ctl Is Nothing Then Set ctl = CF_MON_FindBestInput(root, "LE DELIVERY")
If ctl Is Nothing Then
    Set ctl = CF_MON_FindBestInput(root, "OUTBOUND DELIVERY")
End If
If ctl Is Nothing Then
    detailText = "MAP_LE_DELIVERY_NOT_FOUND; Window=" & windowName & "; RootType=" & CStr(root.Type)
    Exit Function
End If

stepText = "SET_LE_DELIVERY_VALUE"
SAPX_SetValue ctl, Trim$(ob)
stepText = "READ_LE_DELIVERY_VALUE"
actualText = SAPX_ControlValue(ctl)
If StrComp(Trim$(actualText), Trim$(ob), vbTextCompare) <> 0 Then
    detailText = "MAP_LE_DELIVERY_READBACK_MISMATCH; ID=" & CStr(ctl.id) & "; Type=" & CStr(ctl.Type) & "; Actual=" & actualText & "; Expected=" & ob
    Exit Function
End If

stepText = "FIND_EXECUTE_BUTTON"
Set btn = SAPX_FindById(ses, "wnd[1]/tbar[0]/btn[8]")
If btn Is Nothing Then Set btn = CF_MON_FindBestButton(root, "EXECUTE")
If btn Is Nothing Then Set btn = CF_MON_FindBestButton(root, "CHECK")
stepText = "PRESS_EXECUTE_BUTTON"
If Not btn Is Nothing Then
    btn.Press
Else
    ' The selection screen is a modal wnd[1]. Sending F8 to wnd[0] leaves
    ' the criteria dialog untouched and reloads the old MON view instead.
    If windowName = "wnd[1]" Then
        ses.FindById("wnd[1]", False).sendVKey 8
    Else
        ses.FindById("wnd[0]", False).sendVKey 8
    End If
End If
stepText = "WAIT_AFTER_EXECUTE"
SAPX_WaitIdle ses, 30
detailText = "MAP_OK; ID=" & CStr(ctl.id) & "; Type=" & CStr(ctl.Type) & "; Value=" & actualText
If openDetail <> "" Then detailText = detailText & "; Open=" & openDetail
CF_MON_SetLEDelivery = True
Exit Function
Failed:
detailText = "MON_MAP_ERROR; Step=" & stepText & "; Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function CF_MON_OpenOutboundDeliveryOrder(ByVal ses As Object, ByRef detailText As String) As Boolean
Dim tree As Object, keysObj As Object, keyText As String, nodeText As String, pathText As Variant
Dim paths As Variant
Dim i As Long, keyCount As Long
On Error GoTo Failed
detailText = ""
Set tree = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[0]/shell")
If tree Is Nothing Then detailText = "MON_TREE_NOT_FOUND": Exit Function

' Prefer the actual node key returned by the SAP tree instead of guessing a
' numeric path.  MON contains repeated captions; GetAllNodeKeys preserves
' the tree order, so the first exact Outbound Delivery Order is the node
' under Outbound > Documents that the user selects manually.
On Error Resume Next
Err.Clear
Set keysObj = CallByName(tree, "GetAllNodeKeys", VbMethod)
If Err.Number = 0 Then keyCount = SAPX_CollectionCount(keysObj)
Err.Clear
On Error GoTo Failed
For i = 0 To keyCount - 1
    keyText = CStr(SAPX_CollectionValueAt(keysObj, i) & "")
    nodeText = ""
    On Error Resume Next
    Err.Clear
    nodeText = CStr(CallByName(tree, "GetNodeTextByKey", VbMethod, keyText))
    If Err.Number <> 0 Then Err.Clear: nodeText = ""
    On Error GoTo Failed
    If StrComp(Trim$(nodeText), "OUTBOUND DELIVERY ORDER", vbTextCompare) = 0 Then
        On Error Resume Next
        Err.Clear
        CallByName tree, "DoubleClickNode", VbMethod, keyText
        If Err.Number = 0 Then
            CF_MON_OpenOutboundDeliveryOrder = True
            detailText = "NodeKey=" & keyText & "; NodeText=" & nodeText & "; Source=GetAllNodeKeys"
        Else
            detailText = "TREE_DOUBLECLICK_ERR " & CStr(Err.Number) & ": " & Err.Description & "; NodeKey=" & keyText
        End If
        Err.Clear
        On Error GoTo Failed
        Exit Function
    End If
Next i

' Fallback for older SAP GUI builds where GetAllNodeKeys is not available.
paths = Array("1/1/1", "1/1/2", "1/2/1", "1/2/2", "2/1/1", "2/1/2", "/1/1/1", "/1/1/2", "/1/2/1", "/1/2/2")
For Each pathText In paths
    keyText = ""
    On Error Resume Next
    Err.Clear
    keyText = CStr(CallByName(tree, "FindNodeKeyByPath", VbMethod, CStr(pathText)))
    If Err.Number <> 0 Then Err.Clear: keyText = ""
    If keyText <> "" Then nodeText = CStr(CallByName(tree, "GetNodeTextByKey", VbMethod, keyText)) Else nodeText = ""
    If Err.Number <> 0 Then Err.Clear: nodeText = ""
    On Error GoTo Failed
    If InStr(1, UCase$(nodeText), "OUTBOUND", vbTextCompare) > 0 And _
       InStr(1, UCase$(nodeText), "DELIVERY", vbTextCompare) > 0 Then
        On Error Resume Next
        Err.Clear
        CallByName tree, "DoubleClickNode", VbMethod, keyText
        If Err.Number = 0 Then
            CF_MON_OpenOutboundDeliveryOrder = True
            detailText = "TreePath=" & CStr(pathText) & "; NodeKey=" & keyText & "; NodeText=" & nodeText
        Else
            detailText = "TREE_DOUBLECLICK_ERR " & CStr(Err.Number) & ": " & Err.Description
        End If
        Err.Clear
        On Error GoTo Failed
        Exit Function
    End If
Next pathText
detailText = "MON_OUTBOUND_DELIVERY_ORDER_NODE_NOT_FOUND"
Exit Function
Failed:
detailText = "MON_TREE_MAP_ERROR Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpMONTreeNodes() As String
Dim ses As Object, tree As Object, keysObj As Object, keyValue As Variant, keyText As String
Dim openMsg As String, nodeText As String, lineText As String, n As Long, i As Long, keyCount As Long
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpMONTreeNodes = "OPEN_FAILED " & openMsg: Exit Function
Set tree = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[0]/shell")
If tree Is Nothing Then CF_DumpMONTreeNodes = "TREE_NOT_FOUND": Exit Function
On Error Resume Next
Err.Clear
Set keysObj = CallByName(tree, "GetAllNodeKeys", VbMethod)
If Err.Number <> 0 Then
    CF_DumpMONTreeNodes = "GET_NODE_KEYS_FAILED Err " & CStr(Err.Number) & ": " & Err.Description
    Err.Clear
    On Error GoTo Failed
    Exit Function
End If
On Error GoTo Failed
keyCount = SAPX_CollectionCount(keysObj)
For i = 0 To keyCount - 1
    keyValue = SAPX_CollectionValueAt(keysObj, i)
    keyText = CStr(keyValue & "")
    nodeText = ""
    On Error Resume Next
    Err.Clear
    nodeText = CStr(CallByName(tree, "GetNodeTextByKey", VbMethod, CStr(keyText)))
    Err.Clear
    On Error GoTo Failed
    n = n + 1
    lineText = lineText & "KEY=" & CStr(keyText) & "; TEXT=" & nodeText & vbLf
    If n >= 200 Then Exit For
Next i
CF_DumpMONTreeNodes = "NODE_COUNT=" & CStr(n) & vbLf & lineText
Exit Function
Failed:
CF_DumpMONTreeNodes = "TREE_NODE_DUMP_ERROR Err " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_TestMONMapOB(ByVal ob As String) As String
Dim ses As Object, openMsg As String, detailText As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_TestMONMapOB = "OPEN_FAILED " & openMsg: Exit Function
If CF_MON_SetLEDelivery(ses, ob, detailText) Then
    CF_TestMONMapOB = "PASS " & detailText
Else
    CF_TestMONMapOB = "FAIL " & detailText
End If
Exit Function
Failed:
CF_TestMONMapOB = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpMONSelectionFields() As String
Dim ses As Object, root As Object, openMsg As String, openDetail As String
Dim windowName As String, dumpText As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpMONSelectionFields = "OPEN_FAILED " & openMsg: Exit Function
If Not CF_MON_OpenOutboundDeliveryOrder(ses, openDetail) Then
    CF_DumpMONSelectionFields = "OPEN_NODE_FAILED " & openDetail
    Exit Function
End If
SAPX_WaitIdle ses, 10
Set root = SAPX_FindById(ses, "wnd[1]")
If Not root Is Nothing Then windowName = "wnd[1]"
If root Is Nothing Then Set root = SAPX_FindById(ses, "wnd[0]"): windowName = "wnd[0]"
If root Is Nothing Then CF_DumpMONSelectionFields = "SELECTION_WINDOW_NOT_FOUND": Exit Function
CF_MON_DumpSelectionFieldsTree root, dumpText, 0, windowName
CF_DumpMONSelectionFields = "OPEN=" & openDetail & vbLf & dumpText
Exit Function
Failed:
CF_DumpMONSelectionFields = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub CF_MON_DumpSelectionFieldsTree(ByVal obj As Object, ByRef dumpText As String, ByVal depth As Long, ByVal windowName As String)
Dim typ As String, idText As String, nameText As String, labelText As String, tipText As String, valueText As String
Dim children As Object, child As Object, i As Long, n As Long
If obj Is Nothing Or depth > 18 Then Exit Sub
On Error Resume Next
typ = CStr(obj.Type & "")
idText = CStr(obj.id & "")
nameText = CStr(obj.Name & "")
labelText = CStr(obj.text & "")
tipText = CStr(obj.Tooltip & "")
If InStr(UCase$(typ), "TEXTFIELD") > 0 Or InStr(UCase$(typ), "COMBOBOX") > 0 Or InStr(UCase$(typ), "CTEXTFIELD") > 0 Then
    valueText = SAPX_ControlValue(obj)
    dumpText = dumpText & "ID=" & idText & "; Type=" & typ & "; Name=" & nameText & "; Text=" & labelText & "; Tooltip=" & tipText & "; Value=" & valueText & vbLf
ElseIf InStr(UCase$(typ), "BUTTON") > 0 Then
    dumpText = dumpText & "ID=" & idText & "; Type=" & typ & "; Name=" & nameText & "; Text=" & labelText & "; Tooltip=" & tipText & vbLf
End If
Set children = obj.children
If Not children Is Nothing Then n = CLng(children.count)
Err.Clear
On Error GoTo 0
For i = 0 To n - 1
    Set child = SAPX_CollectionObjectAt(children, i)
    If Not child Is Nothing Then CF_MON_DumpSelectionFieldsTree child, dumpText, depth + 1, windowName
    Set child = Nothing
Next i
End Sub

Public Function CF_DumpWarehouseTaskControls() As String
Dim ses As Object, root As Object, openMsg As String, dumpText As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpWarehouseTaskControls = "OPEN_FAILED " & openMsg: Exit Function
Set root = SAPX_FindById(ses, "wnd[0]")
If root Is Nothing Then CF_DumpWarehouseTaskControls = "MAIN_WINDOW_NOT_FOUND": Exit Function
CF_MON_FindWarehouseTaskControls root, dumpText, 0
If dumpText = "" Then dumpText = "WAREHOUSE_TASK_CONTROL_NOT_FOUND"
CF_DumpWarehouseTaskControls = dumpText
Exit Function
Failed:
CF_DumpWarehouseTaskControls = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpWarehouseTaskControlsAfterFilter(ByVal ob As String) As String
Dim ses As Object, grid As Object, root As Object, openMsg As String, mapMsg As String, dumpText As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpWarehouseTaskControlsAfterFilter = "OPEN_FAILED " & openMsg: Exit Function
If Not CF_MON_SetLEDelivery(ses, ob, mapMsg) Then
    CF_DumpWarehouseTaskControlsAfterFilter = "FILTER_FAILED " & mapMsg
    Exit Function
End If
Set grid = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
If grid Is Nothing Then CF_DumpWarehouseTaskControlsAfterFilter = "UPPER_GRID_NOT_FOUND; " & mapMsg: Exit Function
On Error Resume Next
grid.selectedRows = "0"
grid.SetCurrentCell 0, "DOCNO_H"
grid.ClickCurrentCell
Err.Clear
On Error GoTo Failed
SAPX_WaitIdle ses, 10
Set root = SAPX_FindById(ses, "wnd[0]")
If root Is Nothing Then CF_DumpWarehouseTaskControlsAfterFilter = "MAIN_WINDOW_NOT_FOUND_AFTER_SELECT": Exit Function
CF_MON_FindWarehouseTaskControls root, dumpText, 0
If dumpText = "" Then dumpText = "WAREHOUSE_TASK_CONTROL_NOT_FOUND_AFTER_SELECT"
CF_DumpWarehouseTaskControlsAfterFilter = "FILTER=" & mapMsg & vbLf & dumpText
Exit Function
Failed:
CF_DumpWarehouseTaskControlsAfterFilter = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpMONObjectTreeAfterFilter(ByVal ob As String) As String
Dim ses As Object, grid As Object, openMsg As String, mapMsg As String, props As Variant
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpMONObjectTreeAfterFilter = "OPEN_FAILED " & openMsg: Exit Function
If Not CF_MON_SetLEDelivery(ses, ob, mapMsg) Then
    CF_DumpMONObjectTreeAfterFilter = "FILTER_FAILED " & mapMsg
    Exit Function
End If
Set grid = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
If grid Is Nothing Then CF_DumpMONObjectTreeAfterFilter = "UPPER_GRID_NOT_FOUND; " & mapMsg: Exit Function
On Error Resume Next
grid.selectedRows = "0"
grid.SetCurrentCell 0, "DOCNO_H"
grid.ClickCurrentCell
Err.Clear
On Error GoTo Failed
SAPX_WaitIdle ses, 10
props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable", "ScreenLeft", "ScreenTop", "Width", "Height")
CF_DumpMONObjectTreeAfterFilter = "FILTER=" & mapMsg & vbLf & CStr(ses.GetObjectTree("wnd[0]", props))
Exit Function
Failed:
CF_DumpMONObjectTreeAfterFilter = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpMONClickableControlsAfterFilter(ByVal ob As String) As String
Dim ses As Object, grid As Object, root As Object, openMsg As String, mapMsg As String, dumpText As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpMONClickableControlsAfterFilter = "OPEN_FAILED " & openMsg: Exit Function
If Not CF_MON_SetLEDelivery(ses, ob, mapMsg) Then
    CF_DumpMONClickableControlsAfterFilter = "FILTER_FAILED " & mapMsg
    Exit Function
End If
Set grid = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
If grid Is Nothing Then CF_DumpMONClickableControlsAfterFilter = "UPPER_GRID_NOT_FOUND; " & mapMsg: Exit Function
On Error Resume Next
grid.selectedRows = "0"
grid.SetCurrentCell 0, "DOCNO_H"
grid.ClickCurrentCell
Err.Clear
On Error GoTo Failed
SAPX_WaitIdle ses, 10
Set root = SAPX_FindById(ses, "wnd[0]")
If root Is Nothing Then CF_DumpMONClickableControlsAfterFilter = "MAIN_WINDOW_NOT_FOUND_AFTER_SELECT": Exit Function
CF_MON_DumpClickableControls root, dumpText, 0
CF_DumpMONClickableControlsAfterFilter = "FILTER=" & mapMsg & vbLf & dumpText
Exit Function
Failed:
CF_DumpMONClickableControlsAfterFilter = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpFollowOnFunctionsAfterFilter(ByVal ob As String) As String
Dim ses As Object, grid As Object, openMsg As String, mapMsg As String, treeText As String
Dim props As Variant, keys As Variant, k As Variant, p As Long, resultText As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpFollowOnFunctionsAfterFilter = "OPEN_FAILED " & openMsg: Exit Function
If Not CF_MON_SetLEDelivery(ses, ob, mapMsg) Then
    CF_DumpFollowOnFunctionsAfterFilter = "FILTER_FAILED " & mapMsg
    Exit Function
End If
Set grid = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
If grid Is Nothing Then CF_DumpFollowOnFunctionsAfterFilter = "UPPER_GRID_NOT_FOUND; " & mapMsg: Exit Function
On Error Resume Next
grid.selectedRows = "0"
grid.SetCurrentCell 0, "ERP_NUMBER_H"
grid.ClickCurrentCell
Err.Clear
On Error GoTo Failed
SAPX_WaitIdle ses, 10
props = Array("Id", "Type", "Name", "Text", "Tooltip")
treeText = CStr(ses.GetObjectTree("wnd[0]", props))
keys = Array("FOLLOW", "WAREHOUSE", "TASK", "FUNCTION")
For Each k In keys
    p = InStr(1, UCase$(treeText), CStr(k), vbTextCompare)
    If p > 0 Then resultText = resultText & CStr(k) & ": " & Mid$(treeText, MaxLong(1, p - 350), 900) & vbLf
Next k
If resultText = "" Then resultText = "NO_FOLLOW_ON_KEYWORDS; TREE_LENGTH=" & CStr(Len(treeText)) & "; TREE_HEAD=" & Left$(treeText, 1200)
CF_DumpFollowOnFunctionsAfterFilter = "FILTER=" & mapMsg & vbLf & resultText
Exit Function
Failed:
CF_DumpFollowOnFunctionsAfterFilter = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpExactWarehouseTaskAfterFilter(ByVal ob As String) As String
Dim ses As Object, grid As Object, openMsg As String, mapMsg As String, treeText As String
Dim props As Variant, p As Long, startAt As Long, hitCount As Long, piece As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpExactWarehouseTaskAfterFilter = "OPEN_FAILED " & openMsg: Exit Function
If Not CF_MON_SetLEDelivery(ses, ob, mapMsg) Then
    CF_DumpExactWarehouseTaskAfterFilter = "FILTER_FAILED " & mapMsg
    Exit Function
End If
Set grid = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
If grid Is Nothing Then CF_DumpExactWarehouseTaskAfterFilter = "UPPER_GRID_NOT_FOUND; " & mapMsg: Exit Function
On Error Resume Next
grid.selectedRows = "0"
grid.SetCurrentCell 0, "ERP_NUMBER_H"
grid.ClickCurrentCell
Err.Clear
On Error GoTo Failed
SAPX_WaitIdle ses, 10
props = Array("Id", "Type", "Name", "Text", "Tooltip", "Changeable")
treeText = CStr(ses.GetObjectTree("wnd[0]", props))
startAt = 1
Do
    p = InStr(startAt, UCase$(treeText), "WAREHOUSE TASK", vbTextCompare)
    If p = 0 Then Exit Do
    hitCount = hitCount + 1
    piece = Mid$(treeText, MaxLong(1, p - 500), 1200)
    CF_DumpExactWarehouseTaskAfterFilter = CF_DumpExactWarehouseTaskAfterFilter & "HIT=" & CStr(hitCount) & "; " & piece & vbLf
    startAt = p + 15
    If hitCount >= 10 Then Exit Do
Loop
If hitCount = 0 Then CF_DumpExactWarehouseTaskAfterFilter = "NO_EXACT_WAREHOUSE_TASK; TREE_LENGTH=" & CStr(Len(treeText))
CF_DumpExactWarehouseTaskAfterFilter = "FILTER=" & mapMsg & vbLf & CF_DumpExactWarehouseTaskAfterFilter
Exit Function
Failed:
CF_DumpExactWarehouseTaskAfterFilter = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpUpperGridStateAfterFilter(ByVal ob As String) As String
Dim ses As Object, grid As Object, stateObj As Object, openMsg As String, mapMsg As String, dumpText As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpUpperGridStateAfterFilter = "OPEN_FAILED " & openMsg: Exit Function
If Not CF_MON_SetLEDelivery(ses, ob, mapMsg) Then
    CF_DumpUpperGridStateAfterFilter = "FILTER_FAILED " & mapMsg
    Exit Function
End If
Set grid = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
If grid Is Nothing Then CF_DumpUpperGridStateAfterFilter = "UPPER_GRID_NOT_FOUND; " & mapMsg: Exit Function
On Error Resume Next
grid.selectedRows = "0"
grid.SetCurrentCell 0, "ERP_NUMBER_H"
grid.ClickCurrentCell
Err.Clear
Set stateObj = grid.DumpState("")
If Err.Number <> 0 Then
    Err.Clear
    Set stateObj = grid.DumpState
End If
On Error GoTo Failed
If stateObj Is Nothing Then CF_DumpUpperGridStateAfterFilter = "DUMPSTATE_EMPTY; " & mapMsg: Exit Function
SAPX_DumpAnyCollection stateObj, dumpText, 0
CF_DumpUpperGridStateAfterFilter = "FILTER=" & mapMsg & vbLf & dumpText
Exit Function
Failed:
CF_DumpUpperGridStateAfterFilter = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpMONGridsAfterWarehouseTask(ByVal ob As String) As String
Dim ses As Object, upper As Object, root As Object, openMsg As String, mapMsg As String, diag As String, dumpText As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpMONGridsAfterWarehouseTask = "OPEN_FAILED " & openMsg: Exit Function
If Not CF_MON_SetLEDelivery(ses, ob, mapMsg) Then
    CF_DumpMONGridsAfterWarehouseTask = "FILTER_FAILED " & mapMsg
    Exit Function
End If
Set upper = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
If upper Is Nothing Then CF_DumpMONGridsAfterWarehouseTask = "UPPER_GRID_NOT_FOUND": Exit Function
On Error Resume Next
upper.selectedRows = "0"
upper.SetCurrentCell 0, "DOCNO_H"
Err.Clear
On Error GoTo Failed
If Not CFWT_SelectWarehouseTask(ses, 15#) Then diag = CFWT_LastDiagnostic Else diag = "WAREHOUSE_TASK_ACTIVATED"
Set root = SAPX_FindById(ses, "wnd[0]")
CF_MON_DumpGridLikeControls root, dumpText, 0
CF_DumpMONGridsAfterWarehouseTask = "FILTER=" & mapMsg & "; ACTIVATE=" & diag & vbLf & dumpText
Exit Function
Failed:
CF_DumpMONGridsAfterWarehouseTask = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function CF_DumpWarehouseTaskColumns(ByVal ob As String) As String
Dim ses As Object, upper As Object, wt As Object, order As Object
Dim openMsg As String, mapMsg As String, diag As String, resultText As String
Dim i As Long, n As Long, colKey As String, r As Long
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpWarehouseTaskColumns = "OPEN_FAILED " & openMsg: Exit Function
If Not CF_MON_SetLEDelivery(ses, ob, mapMsg) Then CF_DumpWarehouseTaskColumns = "FILTER_FAILED " & mapMsg: Exit Function
Set upper = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
If upper Is Nothing Then CF_DumpWarehouseTaskColumns = "UPPER_GRID_NOT_FOUND": Exit Function
On Error Resume Next
upper.selectedRows = "0"
upper.SetCurrentCell 0, "DOCNO_H"
Err.Clear
On Error GoTo Failed
If Not CFWT_SelectWarehouseTask(ses, 15#) Then diag = CFWT_LastDiagnostic Else diag = "WAREHOUSE_TASK_ACTIVATED"
Set wt = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell")
If wt Is Nothing Then CF_DumpWarehouseTaskColumns = "WT_GRID_NOT_FOUND; ACTIVATE=" & diag: Exit Function
Set order = wt.columnOrder
n = SAPX_CollectionCount(order)
resultText = "ACTIVATE=" & diag & "; ROWS=" & CStr(wt.rowCount) & "; COLS=" & CStr(n) & vbLf
For i = 0 To n - 1
    colKey = CStr(SAPX_CollectionValueAt(order, i) & "")
    resultText = resultText & CStr(i) & "=" & colKey
    If i < 25 Then
        For r = 0 To CLng(wt.rowCount) - 1
            resultText = resultText & "; r" & CStr(r) & "=" & SAPX_GridText(wt, r, colKey)
        Next r
    End If
    resultText = resultText & vbLf
Next i
CF_DumpWarehouseTaskColumns = resultText
Exit Function
Failed:
CF_DumpWarehouseTaskColumns = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub CF_MON_DumpGridLikeControls(ByVal obj As Object, ByRef dumpText As String, ByVal depth As Long)
Dim typ As String, idText As String, nameText As String, textValue As String, rows As String, cols As String
Dim order As Object, children As Object, child As Object, i As Long, n As Long
If obj Is Nothing Or depth > 18 Then Exit Sub
On Error Resume Next
typ = CStr(obj.Type & "")
idText = CStr(obj.id & "")
nameText = CStr(obj.Name & "")
textValue = CStr(obj.text & "")
rows = CStr(obj.rowCount)
Set order = obj.columnOrder
If Not order Is Nothing Then cols = CStr(SAPX_CollectionCount(order))
If InStr(UCase$(typ), "SHELL") > 0 Or InStr(UCase$(typ), "GRID") > 0 Or _
   InStr(UCase$(typ), "TABLE") > 0 Then
    dumpText = dumpText & "ID=" & idText & "; Type=" & typ & "; Name=" & nameText & "; Text=" & textValue & "; Rows=" & rows & "; Cols=" & cols & vbLf
End If
Set children = obj.children
If Not children Is Nothing Then n = CLng(children.count)
Err.Clear
On Error GoTo 0
For i = 0 To n - 1
    Set child = SAPX_CollectionObjectAt(children, i)
    If Not child Is Nothing Then CF_MON_DumpGridLikeControls child, dumpText, depth + 1
    Set child = Nothing
Next i
End Sub

Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
If a >= b Then MaxLong = a Else MaxLong = b
End Function

Private Sub CF_MON_DumpClickableControls(ByVal obj As Object, ByRef dumpText As String, ByVal depth As Long)
Dim typ As String, idText As String, nameText As String, textValue As String, tipText As String, upperType As String
Dim children As Object, child As Object, i As Long, n As Long
If obj Is Nothing Or depth > 18 Then Exit Sub
On Error Resume Next
typ = CStr(obj.Type & "")
upperType = UCase$(typ)
idText = CStr(obj.id & "")
nameText = CStr(obj.Name & "")
textValue = CStr(obj.text & "")
tipText = CStr(obj.Tooltip & "")
If InStr(upperType, "BUTTON") > 0 Or InStr(upperType, "TAB") > 0 Or _
   InStr(upperType, "TOOLBAR") > 0 Or InStr(upperType, "SHELL") > 0 Or _
   textValue <> "" Or tipText <> "" Then
    dumpText = dumpText & "ID=" & idText & "; Type=" & typ & "; Name=" & nameText & "; Text=" & textValue & "; Tooltip=" & tipText & vbLf
End If
Set children = obj.children
If Not children Is Nothing Then n = CLng(children.count)
Err.Clear
On Error GoTo 0
For i = 0 To n - 1
    Set child = SAPX_CollectionObjectAt(children, i)
    If Not child Is Nothing Then CF_MON_DumpClickableControls child, dumpText, depth + 1
    Set child = Nothing
Next i
End Sub

Private Sub CF_MON_FindWarehouseTaskControls(ByVal obj As Object, ByRef dumpText As String, ByVal depth As Long)
Dim typ As String, idText As String, nameText As String, textValue As String, tipText As String
Dim children As Object, child As Object, i As Long, n As Long, allText As String
If obj Is Nothing Or depth > 18 Then Exit Sub
On Error Resume Next
typ = CStr(obj.Type & "")
idText = CStr(obj.id & "")
nameText = CStr(obj.Name & "")
textValue = CStr(obj.text & "")
tipText = CStr(obj.Tooltip & "")
allText = UCase$(idText & "|" & nameText & "|" & textValue & "|" & tipText)
If InStr(allText, "WAREHOUSE TASK") > 0 Then
    dumpText = dumpText & "ID=" & idText & "; Type=" & typ & "; Name=" & nameText & "; Text=" & textValue & "; Tooltip=" & tipText & vbLf
End If
Set children = obj.children
If Not children Is Nothing Then n = CLng(children.count)
Err.Clear
On Error GoTo 0
For i = 0 To n - 1
    Set child = SAPX_CollectionObjectAt(children, i)
    If Not child Is Nothing Then CF_MON_FindWarehouseTaskControls child, dumpText, depth + 1
    Set child = Nothing
Next i
End Sub

Private Function CF_MON_FindBestInput(ByVal root As Object, ByVal aliasText As String) As Object
Dim best As Object, bestScore As Long
CF_MON_ScoreInput root, aliasText, best, bestScore, 0
Set CF_MON_FindBestInput = best
End Function

Private Sub CF_MON_ScoreInput(ByVal obj As Object, ByVal aliasText As String, ByRef best As Object, ByRef bestScore As Long, ByVal depth As Long)
Dim typ As String, allText As String, score As Long, children As Object, child As Object
Dim i As Long, n As Long
If obj Is Nothing Or depth > 18 Then Exit Sub
On Error Resume Next
typ = UCase$(CStr(obj.Type & ""))
allText = UCase$(CStr(obj.id & "") & "|" & CStr(obj.Name & "") & "|" & CStr(obj.text & "") & "|" & CStr(obj.Tooltip & ""))
If (InStr(typ, "TEXTFIELD") > 0 Or InStr(typ, "COMBOBOX") > 0 Or InStr(typ, "CTEXTFIELD") > 0) And _
   InStr(allText, "-TEXT") = 0 And InStr(allText, "%-TEXT") = 0 Then
    ' MON selection controls are generated as text fields. Labels are also
    ' exposed as GuiTextField. For LE Delivery, require the exact ERP_H
    ' technical field so the adjacent Outb. Delivery Order field is ignored.
    If InStr(UCase$(aliasText), "LE DELIVERY") > 0 Then
        If InStr(allText, "S_ERP_H") = 0 And InStr(allText, "ERP_H") = 0 Then GoTo SkipInput
    ElseIf InStr(allText, "S_DOCNO") = 0 And InStr(allText, "DOCNO") = 0 And _
           InStr(allText, "-LOW") = 0 And InStr(allText, "%-LOW") = 0 Then GoTo SkipInput
    If InStr(allText, UCase$(aliasText)) > 0 Then score = 100
    If InStr(allText, "S_DOCNO") > 0 Or InStr(allText, "DOCNO") > 0 Then score = score + 200
    If InStr(allText, "S_ERP_H") > 0 Or InStr(allText, "ERP_H") > 0 Then score = score + 400
    If InStr(allText, "-LOW") > 0 Then score = score + 150
    If InStr(allText, "LE_DELIVERY") > 0 Then score = score + 40
    If InStr(allText, "VBELN") > 0 Then score = score + 25
    If InStr(allText, "DELIVERY") > 0 Then score = score + 15
    If score > bestScore Then Set best = obj: bestScore = score
End If
SkipInput:
Set children = obj.children
If Not children Is Nothing Then n = CLng(children.count)
On Error GoTo 0
For i = 0 To n - 1
    Set child = SAPX_CollectionObjectAt(children, i)
    If Not child Is Nothing Then CF_MON_ScoreInput child, aliasText, best, bestScore, depth + 1
    Set child = Nothing
Next i
End Sub

Private Function CF_MON_FindBestButton(ByVal root As Object, ByVal aliasText As String) As Object
Dim best As Object, children As Object, child As Object, i As Long, n As Long, typ As String, txt As String
If root Is Nothing Then Exit Function
On Error Resume Next
typ = UCase$(CStr(root.Type & "")): txt = UCase$(CStr(root.text & "") & "|" & CStr(root.Tooltip & ""))
If InStr(typ, "BUTTON") > 0 And InStr(txt, UCase$(aliasText)) > 0 Then Set best = root
Set children = root.children
If Not children Is Nothing Then n = CLng(children.count)
On Error GoTo 0
If Not best Is Nothing Then Set CF_MON_FindBestButton = best: Exit Function
For i = 0 To n - 1
    Set child = SAPX_CollectionObjectAt(children, i)
    If Not child Is Nothing Then Set best = CF_MON_FindBestButton(child, aliasText): If Not best Is Nothing Then Set CF_MON_FindBestButton = best: Exit Function
Next i
End Function

Public Function CF_DumpMONData() As String
Dim ses As Object, grid As Object, order As Object, openMsg As String
Dim i As Long, r As Long, n As Long, rowCount As Long, colId As String, lineText As String
On Error GoTo Failed
Set ses = CF_OpenTCodeSession("/SCWM/MON", openMsg)
If ses Is Nothing Then CF_DumpMONData = "OPEN_FAILED " & openMsg: Exit Function
Set grid = SAPX_FindById(ses, "wnd[0]/usr/shell/shellcont[1]/shell/shellcont[0]/shell")
If grid Is Nothing Then CF_DumpMONData = "UPPER_GRID_NOT_FOUND": Exit Function
Set order = grid.columnOrder
n = SAPX_CollectionCount(order)
CF_DumpMONData = "UPPER_ROWS=" & CStr(CLng(grid.rowCount)) & " COLS=" & CStr(n) & vbLf
For i = 0 To n - 1
    colId = CStr(SAPX_CollectionValueAt(order, i) & "")
    CF_DumpMONData = CF_DumpMONData & "COL[" & CStr(i) & "]=" & colId & " | " & SAPX_ColumnCaption(grid, colId) & vbLf
Next i
rowCount = CLng(grid.rowCount)
If rowCount > 10 Then rowCount = 10
For r = 0 To rowCount - 1
    lineText = "ROW[" & CStr(r) & "]"
    For i = 0 To n - 1
        colId = CStr(SAPX_CollectionValueAt(order, i) & "")
        lineText = lineText & " | " & colId & "=" & SAPX_GridText(grid, r, colId)
    Next i
    CF_DumpMONData = CF_DumpMONData & lineText & vbLf
Next r
Exit Function
Failed:
CF_DumpMONData = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Private Function SAPX_GetSessionForTcode(ByVal tcode As String, Optional ByVal allowCreate As Boolean = True) As Object
Dim seed As Object, con As Object, ses As Object, cand As Object, ids As Object
Dim i As Long, n As Long, sid As String, t As Double
If SAPRun_IsRunMode() Then allowCreate = False
Set seed = modSupports.Get_SAPSession()
If seed Is Nothing Then Err.Raise vbObjectError + 2938, , "Khong co SAP session."
Set con = seed.parent
For i = 0 To con.children.count - 1
    Set ses = SAPX_CollectionObjectAt(con.children, i)
    If Not ses Is Nothing Then If SAPX_TcodeIs(ses, tcode) Then Set SAPX_GetSessionForTcode = ses: Exit Function
Next i
If Not allowCreate Then Err.Raise vbObjectError + 2940, , "SAP_EXISTING_SESSION_NOT_FOUND: " & tcode
Set ids = CreateObject("Scripting.Dictionary"): n = con.children.count
For i = 0 To n - 1
    Set ses = SAPX_CollectionObjectAt(con.children, i)
    If Not ses Is Nothing Then sid = CStr(ses.id): If sid <> "" Then ids(sid) = True
Next i
seed.CreateSession: t = Timer
Do
    DoEvents: SAPX_Sleep 25
    If con.children.count > n Then Exit Do
    If SAPX_Elapsed(t) > 8 Then Err.Raise vbObjectError + 2941, , "Tao SAP session bi timeout."
Loop
For i = 0 To con.children.count - 1
    Set ses = SAPX_CollectionObjectAt(con.children, i)
    If Not ses Is Nothing Then
        sid = CStr(ses.id)
        If Not ids.Exists(sid) Then Set cand = ses: Exit For
    End If
Next i
If cand Is Nothing Then Err.Raise vbObjectError + 2942, , "Khong nhan duoc SAP session moi."
If Not SAPX_WaitIdle(cand, 15) Then Err.Raise vbObjectError + 2943, , "SAP session moi dang Busy."
SAPX_Navigate cand, tcode
If Not SAPX_TcodeIs(cand, tcode) Then Err.Raise vbObjectError + 2944, , "SAP session moi sai TCode."
Set SAPX_GetSessionForTcode = cand
End Function
Private Function SAPX_ReopenFreshSessionForTcode(ByVal oldSes As Object, ByVal tcode As String) As Object
Dim con As Object, ses As Object, cand As Object, ids As Object, pop As Object
Dim i As Long, n As Long, sid As String, oldId As String, t As Double, failNo As Long, failText As String
On Error GoTo Failed
If oldSes Is Nothing Then Set SAPX_ReopenFreshSessionForTcode = SAPX_GetSessionForTcode(tcode): Exit Function
Set con = oldSes.parent
If con Is Nothing Then Err.Raise vbObjectError + 2945, , "Khong lay duoc SAP connection de mo lai tab."
Set ids = CreateObject("Scripting.Dictionary"): n = con.children.count
For i = 0 To n - 1
    Set ses = SAPX_CollectionObjectAt(con.children, i)
    If Not ses Is Nothing Then sid = CStr(ses.id): If sid <> "" Then ids(sid) = True
Next i
oldId = CStr(oldSes.id)
oldSes.CreateSession: t = Timer
Do
    DoEvents: SAPX_Sleep 25
    If con.children.count > n Then Exit Do
    If SAPX_Elapsed(t) > 8 Then Err.Raise vbObjectError + 2946, , "Mo tab SAP moi bi timeout."
Loop
For i = 0 To con.children.count - 1
    Set ses = SAPX_CollectionObjectAt(con.children, i)
    If Not ses Is Nothing Then
        sid = CStr(ses.id)
        If Not ids.Exists(sid) Then Set cand = ses: Exit For
    End If
Next i
If cand Is Nothing Then Err.Raise vbObjectError + 2947, , "Khong nhan duoc tab SAP moi."
If Not SAPX_WaitIdle(cand, 15) Then Err.Raise vbObjectError + 2948, , "Tab SAP moi dang Busy."
SAPX_Navigate cand, tcode
If Not SAPX_TcodeIs(cand, tcode) Then Err.Raise vbObjectError + 2949, , "Tab SAP moi khong vao dung " & tcode & "."
On Error Resume Next
Err.Clear
con.CloseSession oldId
If Err.Number <> 0 Then
    Err.Clear
    oldSes.FindById("wnd[0]", False).Close
    Set pop = oldSes.FindById("wnd[1]", False)
    If Not pop Is Nothing Then oldSes.FindById("wnd[1]/usr/btnSPOP-OPTION1", False).Press
End If
Err.Clear
On Error GoTo Failed
Set SAPX_ReopenFreshSessionForTcode = cand
Exit Function
Failed:
failNo = Err.Number: failText = Err.Description
Err.Raise failNo, , "Khong the dong tab VL06O cu/mo tab moi: " & failText
End Function
Private Function SAPX_TcodeIs(ByVal ses As Object, ByVal tcode As String) As Boolean
Dim a As String, b As String
On Error Resume Next: a = UCase$(Trim$(CStr(ses.Info.Transaction))): On Error GoTo 0
b = UCase$(Trim$(tcode))
If Left$(a, 1) = "/" Then a = Mid$(a, 2)
If Left$(b, 1) = "/" Then b = Mid$(b, 2)
SAPX_TcodeIs = (a = b)
End Function
Private Sub SAPX_Navigate(ByVal ses As Object, ByVal tcode As String)
Dim cmd As String
If SAPRun_IsRunMode() Then
    Err.Raise vbObjectError + 2951, , "STOP_SCREEN_CONTROL_FORBIDDEN: navigation by OK-code is not allowed in RUN."
End If
cmd = tcode: If Left$(cmd, 1) = "/" Then cmd = "/n" & cmd Else cmd = "/n" & cmd
ses.FindById("wnd[0]/tbar[0]/okcd").text = cmd: ses.FindById("wnd[0]").sendVKey 0
If Not SAPX_WaitIdle(ses, 30) Then Err.Raise vbObjectError + 2940, , "SAP Busy khi mo " & tcode
End Sub
Private Function SAPX_WaitIdle(ByVal ses As Object, ByVal timeoutSec As Double) As Boolean
Dim t As Double: t = Timer
Do
    On Error Resume Next
    If Not ses.busy Then SAPX_WaitIdle = True: Exit Function
    On Error GoTo 0
    DoEvents: SAPX_Sleep 25
Loop While SAPX_Elapsed(t) < timeoutSec
End Function
Private Function SAPX_Elapsed(ByVal t As Double) As Double
If Timer >= t Then SAPX_Elapsed = Timer - t Else SAPX_Elapsed = 86400# - t + Timer
End Function
Private Function SAPX_FindById(ByVal ses As Object, ByVal idText As String) As Object
If Trim$(idText) = "" Then Exit Function
On Error Resume Next: Set SAPX_FindById = ses.FindById(idText, False): On Error GoTo 0
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
    Set SAPX_CollectionObjectAt = items.item(idx)
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
Private Function SAPX_CollectionCount(ByVal items As Object) As Long
If items Is Nothing Then Exit Function
On Error Resume Next
Err.Clear
SAPX_CollectionCount = CLng(CallByName(items, "Count", VbGet))
If Err.Number <> 0 Then
    Err.Clear
    SAPX_CollectionCount = CLng(CallByName(items, "Length", VbGet))
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
If Err.Number <> 0 Then Err.Clear: SAPX_CollectionValueAt = items.item(idx)
If Err.Number <> 0 Then Err.Clear: SAPX_CollectionValueAt = items(CLng(indexNo))
If Err.Number <> 0 Then Err.Clear: SAPX_CollectionValueAt = items(CInt(indexNo))
Err.Clear
On Error GoTo 0
End Function
Private Function SAPX_FindByTokens(ByVal ses As Object, ByVal tokens As String, ByVal textOnly As Boolean) As Object
Dim root As Object
On Error Resume Next: Set root = ses.FindById("wnd[0]/usr", False): On Error GoTo 0
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
For i = 0 To obj.children.count - 1
    Set ch = SAPX_CollectionObjectAt(obj.children, i)
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
ses.FindById("wnd[0]").sendVKey ModuleHelper.ToolConfigLong("VL06O_REFRESH_VKEY", 20, 0, 99)
eventSeen = True
started = Timer
Do
    DoEvents
    busyNow = CBool(ses.busy)
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
    busyNow = CBool(ses.busy)
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
On Error Resume Next: Set root = ses.FindById("wnd[0]", False): On Error GoTo 0
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
For i = 0 To obj.children.count - 1: Set ch = SAPX_CollectionObjectAt(obj.children, i): SAPX_FindGridRec ch, best, bestRows, depth + 1: Next i
On Error GoTo 0
End Sub
Private Sub SAPX_BuildColumnCatalogFast(ByVal grid As Object, ByRef ids As Collection, ByRef normIds As Object, ByRef normCaps As Object)
Dim id As Variant, key As String
Set ids = SAPX_ColumnIds(grid)
Set normIds = CreateObject("Scripting.Dictionary")
Set normCaps = CreateObject("Scripting.Dictionary")
normIds.CompareMode = vbTextCompare
normCaps.CompareMode = vbTextCompare
For Each id In ids
    key = CStr(id)
    normIds(key) = NormalizeHeader(key)
    normCaps(key) = NormalizeHeader(SAPX_ColumnCaption(grid, key))
Next id
End Sub
Private Function SAPX_ResolveColumnFast(ByVal configured As String, ByVal aliasesCsv As String, ByVal required As Boolean, _
ByVal ids As Collection, ByVal normIds As Object, ByVal normCaps As Object) As String
Dim id As Variant, a As Variant, key As String, ni As String, nc As String, na As String
If Not ids Is Nothing Then
    If Trim$(configured) <> "" Then
        For Each id In ids
            If StrComp(CStr(id), configured, vbTextCompare) = 0 Then SAPX_ResolveColumnFast = CStr(id): Exit Function
        Next id
    End If
    ' Pass 1: exact normalized ID/caption.
    For Each id In ids
        key = CStr(id)
        ni = CStr(normIds(key))
        nc = CStr(normCaps(key))
        For Each a In Split(aliasesCsv, ",")
            na = NormalizeHeader(Trim$(CStr(a)))
            If na <> "" And (ni = na Or nc = na) Then SAPX_ResolveColumnFast = key: Exit Function
        Next a
    Next id
    ' Pass 2: fuzzy only for meaningful aliases.
    For Each id In ids
        key = CStr(id)
        ni = CStr(normIds(key))
        nc = CStr(normCaps(key))
        For Each a In Split(aliasesCsv, ",")
            na = NormalizeHeader(Trim$(CStr(a)))
            If Len(na) >= 4 And (InStr(1, ni, na, vbTextCompare) > 0 Or InStr(1, nc, na, vbTextCompare) > 0) Then
                SAPX_ResolveColumnFast = key
                Exit Function
            End If
        Next a
    Next id
End If
If required Then Err.Raise vbObjectError + 2941, , "Thieu cot SAP: " & aliasesCsv
End Function
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
On Error Resume Next: Set o = grid.columnOrder
If Not o Is Nothing Then
    For i = 0 To o.count - 1
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
    For i = 0 To o.count - 1
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

Private Function POST_ExactID(ByVal configKey As String, ByVal mapDefault As String) As String
    If SAPRun_IsRunMode() Then
        POST_ExactID = Trim$(ModuleHelper.ToolConfig(configKey, ""))
    Else
        POST_ExactID = Trim$(ModuleHelper.ToolConfig(configKey, mapDefault))
    End If
End Function

Private Function SAPX_RequireExactColumn(ByVal grid As Object, ByVal configuredId As String, _
                                         ByVal configKey As String) As String
    configuredId = Trim$(configuredId)
    If configuredId = "" Then Err.Raise vbObjectError + 2940, , _
        "STOP_POST_LAYOUT_MISMATCH: missing CONFIG " & configKey
    If Not SAPX_ColumnExists(grid, configuredId) Then Err.Raise vbObjectError + 2941, , _
        "STOP_POST_LAYOUT_MISMATCH: required technical column " & configuredId & _
        " (" & configKey & ") is not present."
    SAPX_RequireExactColumn = configuredId
End Function

Private Function SAPX_OptionalExactColumn(ByVal grid As Object, ByVal configuredId As String) As String
    configuredId = Trim$(configuredId)
    If configuredId = "" Then Exit Function
    If SAPX_ColumnExists(grid, configuredId) Then SAPX_OptionalExactColumn = configuredId
End Function

Private Function SAPX_ReadIndependentResultCount(ByVal ses As Object, ByRef expectedRows As Long, _
                                                  ByRef messageText As String) As Boolean
    Dim controlID As String, patternText As String, ctl As Object, rawText As String
    controlID = Trim$(ModuleHelper.ToolConfig("POST_RESULT_COUNT_ID", ""))
    patternText = Trim$(ModuleHelper.ToolConfig("POST_RESULT_COUNT_PATTERN", ""))
    If controlID = "" Or patternText = "" Then
        messageText = "STOP_POST_EXPECTED_COUNT_UNPROVEN: map POST_RESULT_COUNT_ID and " & _
                      "POST_RESULT_COUNT_PATTERN before TAI POSTING. Grid.RowCount alone is not accepted."
        Exit Function
    End If
    Set ctl = SAPX_FindById(ses, controlID)
    If ctl Is Nothing Then
        messageText = "STOP_POST_EXPECTED_COUNT_UNPROVEN: mapped result-count control was not found: " & controlID
        Exit Function
    End If
    rawText = SAPX_ControlText(ctl)
    If Not SAPX_ParseResultCount(rawText, patternText, expectedRows) Then
        messageText = "STOP_POST_EXPECTED_COUNT_UNPROVEN: cannot parse result count from mapped control; " & _
                      "Text='" & rawText & "'; Pattern='" & patternText & "'."
        Exit Function
    End If
    RT_Log "POST_RESULT_COUNT", "", "", "INFO", _
           "Control=" & controlID & "|Pattern=" & SAPX_PostTraceText(patternText) & _
           "|Text=" & SAPX_PostTraceText(rawText) & "|Parsed=" & CStr(expectedRows)
    SAPX_ReadIndependentResultCount = True
End Function

Private Function SAPX_ParseResultCount(ByVal rawText As String, ByVal patternText As String, _
                                       ByRef resultCount As Long) As Boolean
    Dim p As Long, tailText As String, digits As String, i As Long, ch As String
    rawText = Trim$(rawText)
    patternText = Trim$(patternText)
    If rawText = "" Or patternText = "" Then Exit Function
    p = InStr(1, rawText, patternText, vbTextCompare)
    If p = 0 Then Exit Function
    tailText = Mid$(rawText, p + Len(patternText))
    For i = 1 To Len(tailText)
        ch = Mid$(tailText, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf digits <> "" Then
            Exit For
        End If
    Next i
    If digits = "" Then Exit Function
    On Error GoTo Failed
    resultCount = CLng(digits)
    SAPX_ParseResultCount = (resultCount >= 0)
    Exit Function
Failed:
    resultCount = -1
End Function

Private Function SAPX_ExposeExpectedRows(ByVal grid As Object, ByVal expectedRows As Long, _
                                         ByVal materialColumn As String, ByVal availableQtyColumn As String, _
                                         ByVal batchColumn As String, ByVal sourceBinColumn As String, _
                                         ByVal stockTypeColumn As String, ByVal totalStockColumn As String, _
                                         ByVal timeoutSeconds As Double, ByRef messageText As String) As Boolean
    Dim started As Double, rowCount As Long, visibleRows As Long, firstRow As Long
    Dim maxExpectedRows As Long
    maxExpectedRows = ModuleHelper.ToolConfigLong("POST_MAX_EXPECTED_ROWS", 5000, 1, 50000)
    If expectedRows > maxExpectedRows Then
        messageText = "STOP_POST_CAPTURE_INCOMPLETE: SAP result count=" & CStr(expectedRows) & _
                      " exceeds POST_MAX_EXPECTED_ROWS=" & CStr(maxExpectedRows)
        Exit Function
    End If
    If expectedRows = 0 Then
        If CLng(grid.rowCount) = 0 Then SAPX_ExposeExpectedRows = True Else _
            messageText = "STOP_POST_CAPTURE_INCOMPLETE: SAP result count is zero but grid is not zero."
        Exit Function
    End If
    started = Timer
    Do
        On Error GoTo Failed
        rowCount = CLng(grid.rowCount)
        If rowCount = expectedRows Then
            SAPX_ExposeExpectedRows = True
            Exit Function
        End If
        If rowCount > expectedRows Then
            If SAPX_ValidateExpectedDataRows(grid, materialColumn, availableQtyColumn, batchColumn, _
                                              sourceBinColumn, stockTypeColumn, totalStockColumn, _
                                              expectedRows, rowCount, messageText) Then
                SAPX_ExposeExpectedRows = True
            End If
            Exit Function
        End If
        visibleRows = 1
        On Error Resume Next
        visibleRows = CLng(grid.VisibleRowCount)
        If visibleRows < 1 Then visibleRows = 1
        firstRow = rowCount - visibleRows
        If firstRow < 0 Then firstRow = 0
        grid.FirstVisibleRow = firstRow
        On Error GoTo Failed
        DoEvents
        SAPX_Sleep 150
    Loop While SAPX_Elapsed(started) < timeoutSeconds
    messageText = "STOP_POST_CAPTURE_INCOMPLETE: grid did not expose all rows; GridRowCount=" & _
                  CStr(CLng(grid.rowCount)) & "; SAPExpectedRows=" & CStr(expectedRows)
    Exit Function
Failed:
    messageText = "STOP_POST_CAPTURE_INCOMPLETE: error while exposing SAP grid rows. Err " & _
                  CStr(Err.Number) & ": " & Err.Description
End Function

' /SCWM/POST can append totals, blank suffixes, or grouped rows.  Before a
' Posting cache may be published, classify the entire bounded result grid and
' prove that concrete stock rows exactly match SAP's independent result count.
Private Function SAPX_ValidateExpectedDataRows(ByVal grid As Object, ByVal materialColumn As String, _
                                               ByVal availableQtyColumn As String, ByVal batchColumn As String, _
                                               ByVal sourceBinColumn As String, ByVal stockTypeColumn As String, _
                                               ByVal totalStockColumn As String, ByVal expectedRows As Long, _
                                               ByRef actualRows As Long, ByRef messageText As String) As Boolean
    Dim r As Long, dataRows As Long, auxRows As Long, unknownRows As Long
    Dim sawAux As Boolean, orderBroken As Boolean, rowClass() As String
    Dim materialText As String, availableQtyText As String, batchText As String
    Dim sourceBinText As String, stockTypeText As String, totalStockText As String
    On Error GoTo Failed
    messageText = ""
    actualRows = CLng(grid.rowCount)
    gMPOST_GridRowCount = actualRows
    gMPOST_ClassifiedDataRows = 0
    gMPOST_AuxRows = 0
    gMPOST_UnknownRows = 0
    gMPOST_CaptureProof = "FAIL|POST_GRID_NOT_YET_CLASSIFIED"
    If actualRows < expectedRows Then
        gMPOST_CaptureProof = "FAIL|GRID_SHORT|Expected=" & CStr(expectedRows) & "|Grid=" & CStr(actualRows)
        messageText = "STOP_POST_CAPTURE_INCOMPLETE: grid RowCount=" & CStr(actualRows) & _
                      "; independent SAP result count=" & CStr(expectedRows) & "."
        Exit Function
    End If
    If actualRows = 0 Then
        gMPOST_CaptureProof = "PASS|Expected=0|Grid=0|Data=0|Aux=0|Unknown=0"
        SAPX_ValidateExpectedDataRows = True
        Exit Function
    End If

    ReDim rowClass(0 To actualRows - 1)
    For r = 0 To actualRows - 1
        SAPX_EnsureGridRowVisible grid, r
        materialText = SAPX_GridText(grid, r, materialColumn)
        availableQtyText = SAPX_GridText(grid, r, availableQtyColumn)
        batchText = SAPX_GridText(grid, r, batchColumn)
        sourceBinText = SAPX_GridText(grid, r, sourceBinColumn)
        stockTypeText = SAPX_GridText(grid, r, stockTypeColumn)
        totalStockText = SAPX_GridText(grid, r, totalStockColumn)
        rowClass(r) = SAPX_ClassifyPostRow(materialText, availableQtyText, batchText, _
                                            sourceBinText, stockTypeText, totalStockText, dataRows > 0)
        Select Case rowClass(r)
            Case "DATA_STOCK"
                If sawAux Then orderBroken = True
                dataRows = dataRows + 1
            Case "UNKNOWN"
                unknownRows = unknownRows + 1
            Case Else
                auxRows = auxRows + 1
                sawAux = True
        End Select
    Next r

    gMPOST_ClassifiedDataRows = dataRows
    gMPOST_AuxRows = auxRows
    gMPOST_UnknownRows = unknownRows
    SAPX_LogPostRowDiagnostics grid, rowClass, materialColumn, availableQtyColumn, batchColumn, _
                               sourceBinColumn, stockTypeColumn, totalStockColumn
    gMPOST_CaptureProof = "P0|Expected=" & CStr(expectedRows) & "|Grid=" & CStr(actualRows) & _
                          "|Data=" & CStr(dataRows) & "|Aux=" & CStr(auxRows) & _
                          "|Unknown=" & CStr(unknownRows)
    If unknownRows > 0 Then
        gMPOST_CaptureProof = "FAIL|UNKNOWN_ROWS|" & Mid$(gMPOST_CaptureProof, 4)
        messageText = "STOP_POST_CAPTURE_INCOMPLETE: unknown data-like grid rows=" & CStr(unknownRows) & _
                      "; Expected=" & CStr(expectedRows) & "; Data=" & CStr(dataRows) & _
                      "; Aux=" & CStr(auxRows) & "; Grid=" & CStr(actualRows) & "."
        Exit Function
    End If
    If orderBroken Then
        gMPOST_CaptureProof = "FAIL|DATA_AFTER_AUX|" & Mid$(gMPOST_CaptureProof, 4)
        messageText = "STOP_POST_CAPTURE_INCOMPLETE: data row occurs after an auxiliary row; " & _
                      "Expected=" & CStr(expectedRows) & "; Data=" & CStr(dataRows) & _
                      "; Aux=" & CStr(auxRows) & "; Grid=" & CStr(actualRows) & "."
        Exit Function
    End If
    If dataRows <> expectedRows Then
        gMPOST_CaptureProof = "FAIL|DATA_COUNT_MISMATCH|" & Mid$(gMPOST_CaptureProof, 4)
        messageText = "STOP_POST_CAPTURE_INCOMPLETE: classified data rows=" & CStr(dataRows) & _
                      "; independent SAP result count=" & CStr(expectedRows) & _
                      "; auxiliary rows=" & CStr(auxRows) & "; grid RowCount=" & CStr(actualRows) & "."
        Exit Function
    End If
    gMPOST_CaptureProof = "PASS|" & Mid$(gMPOST_CaptureProof, 4)
    SAPX_ValidateExpectedDataRows = True
    Exit Function
Failed:
    gMPOST_CaptureProof = "FAIL|CLASSIFIER_ERROR|Err=" & CStr(Err.Number)
    messageText = "STOP_POST_CAPTURE_INCOMPLETE: cannot classify POST grid rows. Err " & _
                  CStr(Err.Number) & ": " & Err.Description
End Function

Private Function SAPX_ClassifyPostRow(ByVal materialText As String, ByVal availableQtyText As String, _
                                      ByVal batchText As String, ByVal sourceBinText As String, _
                                      ByVal stockTypeText As String, ByVal totalStockText As String, _
                                      ByVal sawData As Boolean) As String
    Dim hasStockEvidence As Boolean
    materialText = Trim$(materialText)
    availableQtyText = Trim$(availableQtyText)
    batchText = Trim$(batchText)
    sourceBinText = Trim$(sourceBinText)
    stockTypeText = Trim$(stockTypeText)
    totalStockText = Trim$(totalStockText)
    If materialText <> "" Then
        SAPX_ClassifyPostRow = "DATA_STOCK"
        Exit Function
    End If
    hasStockEvidence = (availableQtyText <> "" Or batchText <> "" Or sourceBinText <> "" Or stockTypeText <> "")
    ' A footer total in /SCWM/POST has aggregate quantities but no stock
    ' identity. It is auxiliary only when Batch, Source Bin, and TST are all
    ' blank; any missing-Material row carrying an identity remains UNKNOWN.
    If batchText = "" And sourceBinText = "" And stockTypeText = "" And totalStockText <> "" Then
        SAPX_ClassifyPostRow = "TOTAL_SUBTOTAL"
    ElseIf hasStockEvidence Then
        SAPX_ClassifyPostRow = "UNKNOWN"
    ElseIf totalStockText <> "" Then
        SAPX_ClassifyPostRow = "TOTAL_SUBTOTAL"
    ElseIf sawData Then
        SAPX_ClassifyPostRow = "BLANK_SUFFIX"
    Else
        SAPX_ClassifyPostRow = "GROUP_HEADER"
    End If
End Function

Public Sub SAPX_POST_CompileProbe()
    If SAPX_ClassifyPostRow("MAT-1", "1", "B1", "BIN-1", "FT", "1", False) <> "DATA_STOCK" Then _
        Err.Raise vbObjectError + 2947, , "POST classifier failed DATA_STOCK probe."
    If SAPX_ClassifyPostRow("", "12.389", "", "", "", "13.649", True) <> "TOTAL_SUBTOTAL" Then _
        Err.Raise vbObjectError + 2947, , "POST classifier failed total-row probe."
    If SAPX_ClassifyPostRow("", "1", "B1", "", "", "", True) <> "UNKNOWN" Then _
        Err.Raise vbObjectError + 2947, , "POST classifier failed unknown-row probe."
End Sub

Private Sub SAPX_LogPostRowDiagnostics(ByVal grid As Object, ByRef rowClass() As String, _
                                       ByVal materialColumn As String, ByVal availableQtyColumn As String, _
                                       ByVal batchColumn As String, ByVal sourceBinColumn As String, _
                                       ByVal stockTypeColumn As String, ByVal totalStockColumn As String)
    Dim rowsToLog As Object, r As Long, n As Long, maxRow As Long
    Dim materialText As String, availableQtyText As String, batchText As String
    Dim sourceBinText As String, stockTypeText As String, totalStockText As String
    On Error GoTo SafeExit
    maxRow = UBound(rowClass)
    Set rowsToLog = CreateObject("Scripting.Dictionary")
    For r = 0 To maxRow
        If rowClass(r) <> "DATA_STOCK" Then
            For n = r - 2 To r + 2
                If n >= 0 And n <= maxRow Then rowsToLog(CStr(n)) = True
            Next n
        End If
    Next r
    For r = 0 To maxRow
        If rowsToLog.Exists(CStr(r)) Then
            SAPX_EnsureGridRowVisible grid, r
            materialText = SAPX_GridText(grid, r, materialColumn)
            availableQtyText = SAPX_GridText(grid, r, availableQtyColumn)
            batchText = SAPX_GridText(grid, r, batchColumn)
            sourceBinText = SAPX_GridText(grid, r, sourceBinColumn)
            stockTypeText = SAPX_GridText(grid, r, stockTypeColumn)
            totalStockText = SAPX_GridText(grid, r, totalStockColumn)
            RT_Log "POST_ROW_CLASS", "", "", "INFO", _
                   "Row=" & CStr(r + 1) & "|Class=" & rowClass(r) & _
                   "|Material=" & SAPX_PostTraceText(materialText) & _
                   "|AvlQty=" & SAPX_PostTraceText(availableQtyText) & _
                   "|Batch=" & SAPX_PostTraceText(batchText) & _
                   "|Bin=" & SAPX_PostTraceText(sourceBinText) & _
                   "|TST=" & SAPX_PostTraceText(stockTypeText) & _
                   "|TotalStock=" & SAPX_PostTraceText(totalStockText)
        End If
    Next r
SafeExit:
End Sub

Private Function SAPX_PostTraceText(ByVal valueText As String) As String
    valueText = Replace(valueText, vbCr, " ")
    valueText = Replace(valueText, vbLf, " ")
    valueText = Replace(valueText, "|", "/")
    SAPX_PostTraceText = Trim$(valueText)
End Function

Private Function SAPX_GetStableRowCount(ByVal grid As Object, ByVal timeoutSeconds As Double, _
                                        ByRef finalRows As Long) As Long
    Dim started As Double, previousRows As Long, currentRows As Long, stableHits As Long
    On Error GoTo Failed
    started = Timer
    previousRows = -1
    Do
        currentRows = CLng(grid.rowCount)
        If currentRows < 0 Then Exit Function
        If currentRows = previousRows Then
            stableHits = stableHits + 1
        Else
            stableHits = 0
            previousRows = currentRows
        End If
        If stableHits >= 3 Then
            finalRows = currentRows
            SAPX_GetStableRowCount = currentRows
            Exit Function
        End If
        DoEvents
        SAPX_Sleep 250
    Loop While SAPX_Elapsed(started) < timeoutSeconds
Failed:
    SAPX_GetStableRowCount = -1
End Function

Private Sub SAPX_EnsureGridRowVisible(ByVal grid As Object, ByVal rowNo As Long)
Dim firstRow As Long, visibleRows As Long, targetFirstRow As Long
' /SCWM/POST may expose RowCount before cells outside the viewport are
' materialized. Move by viewport pages so a blank cell is never mistaken for
' a total or suffix merely because the ALV row has not been rendered yet.
On Error Resume Next
firstRow = CLng(grid.FirstVisibleRow)
visibleRows = CLng(grid.VisibleRowCount)
If visibleRows < 1 Then visibleRows = 1
If rowNo < firstRow Or rowNo >= firstRow + visibleRows Then
    targetFirstRow = (rowNo \ visibleRows) * visibleRows
    If targetFirstRow < 0 Then targetFirstRow = 0
    grid.FirstVisibleRow = targetFirstRow
    DoEvents
    SAPX_Sleep 50
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
    grid.SelectContextMenuItem "&SORT_ASC"
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
For i = 0 To obj.children.count - 1: Set ch = SAPX_CollectionObjectAt(obj.children, i): Set SAPX_FindButtonText = SAPX_FindButtonText(ch, wanted, depth + 1): If Not SAPX_FindButtonText Is Nothing Then Exit Function
Next i
On Error GoTo 0
End Function
Private Sub SAPX_ClearGridSelection(ByVal grid As Object)
' A saved ALV variant can leave more than one column selected. Clear the
' selection before SelectColumn so the filter popup belongs only to colId.
If grid Is Nothing Then Exit Sub
On Error Resume Next
Err.Clear
grid.ClearSelection
Err.Clear
grid.selectedRows = ""
Err.Clear
grid.SelectedColumns = ""
Err.Clear
grid.SelectedCells = ""
Err.Clear
On Error GoTo 0
End Sub
Private Function SAPX_ApplyGridFilter(ByVal ses As Object, ByVal grid As Object, ByVal colId As String, ByVal valueText As String, ByVal labelAliases As String) As Boolean
Dim t As Double, low As Object, ok As Boolean, pop As Object, popupRoot As Object
Dim filterStage As String, failNo As Long, failText As String, controlDump As String, lowCount As Long
On Error GoTo Failed
filterStage = "SELECT_COLUMN": mSAPXFilterError = ""
SAPX_ClearGridSelection grid
If CLng(grid.rowCount) > 0 Then grid.SetCurrentCell 0, colId
grid.SelectColumn colId
filterStage = "OPEN_FILTER"
On Error Resume Next
Err.Clear
grid.ContextMenu
grid.SelectContextMenuItem "&FILTER"
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
    Set low = Nothing
    Set popupRoot = Nothing
    On Error Resume Next
    Set popupRoot = ses.FindById("wnd[1]/usr", False)
    On Error GoTo Failed
    ' Bind to SLoc/Storage Location/LGORT first. Do not pick the first LOW
    ' field before the target label because saved variants can expose more
    ' than one filter condition.
    If Not popupRoot Is Nothing Then Set low = SAPX_FindFilterInput(ses, labelAliases)
    ' Fixed-ID fallback is safe only when the popup exposes zero/one known LOW
    ' input. With multiple LOW inputs, fail instead of writing 1005 to the
    ' wrong condition.
    If low Is Nothing And Not popupRoot Is Nothing Then
        lowCount = SAPX_FilterLowInputCount(popupRoot)
        If lowCount <= 1 Then
            On Error Resume Next
            Set low = ses.FindById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxtRSCSEL_255-SLOW_I[1,0]", False)
            On Error GoTo Failed
        End If
    End If
    If Not low Is Nothing Then Exit Do
    SAPX_Sleep 25: DoEvents
Loop While SAPX_Elapsed(t) < 5
If low Is Nothing Then GoTo Failed
filterStage = "SET_VALUE"
SAPX_SetValue low, valueText
If StrComp(SAPX_ControlValue(low), Trim$(valueText), vbTextCompare) <> 0 Then GoTo Failed
filterStage = "CONFIRM"
On Error Resume Next
Err.Clear
ses.FindById("wnd[1]/tbar[0]/btn[0]", False).Press
If Err.Number <> 0 Then Err.Clear: ses.FindById("wnd[1]").sendVKey 0
On Error GoTo Failed
filterStage = "WAIT_POPUP_CLOSE"
If Not SAPX_WaitFilterPopupClosed(ses, 20) Then GoTo Failed
filterStage = "WAIT_IDLE"
If Not SAPX_WaitIdle(ses, 20) Then GoTo Failed
SAPX_ApplyGridFilter = True: Exit Function
Failed:
failNo = Err.Number: failText = Err.Description
If failNo = 0 And Trim$(failText) = "" Then failText = "Khong hoan tat buoc " & filterStage
mSAPXFilterError = filterStage & " [" & CStr(failNo) & "] " & failText
If filterStage = "FIND_INPUT" And Not popupRoot Is Nothing Then
    SAPX_DumpControlIds popupRoot, controlDump, 0
    mSAPXFilterError = mSAPXFilterError & " Controls=" & Replace(controlDump, vbLf, "; ")
End If
On Error Resume Next
Set pop = ses.FindById("wnd[1]", False)
If Not pop Is Nothing Then pop.Close
On Error GoTo 0
End Function
Private Function SAPX_FilterPopupOpen(ByVal ses As Object) As Boolean
Dim pop As Object
On Error Resume Next
Set pop = ses.FindById("wnd[1]", False)
SAPX_FilterPopupOpen = Not pop Is Nothing
On Error GoTo 0
End Function
Private Function SAPX_WaitFilterPopupClosed(ByVal ses As Object, ByVal timeoutSec As Double) As Boolean
Dim started As Double
started = Timer
Do
    If Not SAPX_FilterPopupOpen(ses) Then
        SAPX_WaitFilterPopupClosed = True
        Exit Function
    End If
    SAPX_Sleep 25
    DoEvents
Loop While SAPX_Elapsed(started) < timeoutSec
End Function
Private Function SAPX_FindFilterInput(ByVal ses As Object, ByVal aliasesCsv As String) As Object
Dim root As Object, labels As New Collection, inputs As New Collection, o As Object, l As Object, inp As Object
Dim lr As Long, ir As Long, bestD As Long, d As Long, a As Variant, txt As String, dynToken As String
On Error Resume Next: Set root = ses.FindById("wnd[1]/usr", False): On Error GoTo 0
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
If inputs.count = 1 Then Set SAPX_FindFilterInput = inputs(1)
End Function
Private Function SAPX_FilterLowInputCount(ByVal root As Object) As Long
Dim labels As New Collection, inputs As New Collection
If root Is Nothing Then Exit Function
SAPX_CollectFilterControls root, labels, inputs, 0
SAPX_FilterLowInputCount = inputs.count
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
For i = 0 To obj.children.count - 1: Set ch = SAPX_CollectionObjectAt(obj.children, i): SAPX_CollectFilterControls ch, labels, inputs, depth + 1: Next i
On Error GoTo 0
End Sub
Private Function SAPX_ControlText(ByVal obj As Object) As String
On Error Resume Next: SAPX_ControlText = CStr(obj.text & ""): On Error GoTo 0
End Function
Private Function SAPX_IdRow(ByVal idText As String) As Long
Dim p As Long, q As Long, z As Long, s As String
p = InStrRev(idText, "["): q = InStrRev(idText, "]"): If p = 0 Or q <= p Then SAPX_IdRow = -999: Exit Function
s = Mid$(idText, p + 1, q - p - 1): z = InStr(s, ","): If z > 0 Then SAPX_IdRow = CLng(val(Mid$(s, z + 1))) Else SAPX_IdRow = CLng(val(s))
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
Private Function SAPX_VerifyFilterDetailed(ByVal grid As Object, ByVal colId As String, ByVal expected As String, ByRef detail As String) As Boolean
Dim r As Long, v As String, lastV As String, n As Long, seen As Boolean
detail = ""
On Error GoTo Failed
n = CLng(grid.rowCount)
If n = 0 Then
    detail = "Column=" & colId & "; RowCount=0"
    SAPX_VerifyFilterDetailed = True
    Exit Function
End If
For r = 0 To n - 1
    v = SAPX_GridText(grid, r, colId)
    If v <> "" Then
        lastV = v
        seen = True
    Else
        v = lastV
    End If
    If v <> "" And StrComp(Trim$(v), Trim$(expected), vbTextCompare) <> 0 Then
        detail = "Column=" & colId & "; Expected=" & expected & "; Row=" & CStr(r + 1) & "; Actual=" & v & "; RowCount=" & CStr(n)
        Exit Function
    End If
Next r
If Not seen Then
    detail = "Column=" & colId & "; Expected=" & expected & "; khong doc duoc gia tri SLoc; RowCount=" & CStr(n)
    Exit Function
End If
detail = "Column=" & colId & "; Expected=" & expected & "; RowCount=" & CStr(n) & "; AllRowsMatched"
SAPX_VerifyFilterDetailed = True
Exit Function
Failed:
detail = "Column=" & colId & "; VerifyError=" & CStr(Err.Number) & " " & Err.Description
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





