Option Explicit

Public Const RESULT_SHEET_NAME As String = "RESULT"
Public Const OB_SHEET_NAME As String = "OB_GROUPS"
Private Const DEFAULT_MAX_PRODUCTS_PER_OB As Long = 4
Private Const DEFAULT_MIN_TARGET_PRODUCTS_PER_OB As Long = 3
Public LastOBGroupSucceeded As Boolean
Public LastOBGroupCount As Long
Public LastOBGroupMessage As String

Public Sub GroupResultIntoOB_ProductBased()
If ModuleMain.gOneClickMode Then
GroupResultIntoOB_ProductBased_Core
Exit Sub
End If
If Not MIGO311_RunPost_V24_ErrorGuard.ToolTryBegin("GOM OB THU CONG") Then Exit Sub
On Error GoTo WrapperFail
GroupResultIntoOB_ProductBased_Core
MIGO311_RunPost_V24_ErrorGuard.ToolEnd
Exit Sub
WrapperFail:
MIGO311_RunPost_V24_ErrorGuard.ToolEnd
MsgBox "Loi khi gom OB thu cong: " & Err.description, vbCritical, "Gom OB"
End Sub

Private Sub GroupResultIntoOB_ProductBased_Core()
Dim wsResult As Worksheet, wsOB As Worksheet
Dim maxProductsPerOB As Long, minTargetProductsPerOB As Long, maxRowsPerOB As Long
Dim splitByArea As Boolean, splitByZone As Boolean, groupByTL As Boolean, allowRebuildHistory As Boolean

LastOBGroupSucceeded = False
LastOBGroupCount = 0
LastOBGroupMessage = "Khong the tao OB_GROUPS. Kiem tra RESULT va LOG."
ModuleHelper.ResetToolConfigCache
maxProductsPerOB = ModuleHelper.ToolConfigLong("MAX_PRODUCTS_PER_OB", DEFAULT_MAX_PRODUCTS_PER_OB, 1, 20)
minTargetProductsPerOB = ModuleHelper.ToolConfigLong("MIN_TARGET_PRODUCTS_PER_OB", DEFAULT_MIN_TARGET_PRODUCTS_PER_OB, 1, maxProductsPerOB)
maxRowsPerOB = ModuleHelper.ToolConfigLong("MAX_ROWS_PER_OB", 20, 1, 50)
splitByArea = ModuleHelper.ToolConfigBool("GROUP_SPLIT_BY_AREA", False)
splitByZone = ModuleHelper.ToolConfigBool("GROUP_SPLIT_BY_ZONE", False)
groupByTL = ModuleHelper.ToolConfigBool("GROUP_BY_TL", True)
allowRebuildHistory = ModuleHelper.ToolConfigBool("ALLOW_REBUILD_POSTED_HISTORY", False)
On Error Resume Next
Set wsResult = ThisWorkbook.Worksheets(RESULT_SHEET_NAME)
On Error GoTo 0

If wsResult Is Nothing Then
LastOBGroupMessage = "Khong tim thay sheet RESULT."
If Not ModuleMain.gOneClickMode Then MsgBox LastOBGroupMessage, vbCritical, "Thieu sheet"
Exit Sub
End If

If Not CheckResultTemplate(wsResult) Then Exit Sub

Dim lastRow As Long
lastRow = GetLastRowLocal(wsResult, 1)
If lastRow < 2 Then
LastOBGroupMessage = "Sheet RESULT chua co du lieu."
If Not ModuleMain.gOneClickMode Then MsgBox LastOBGroupMessage, vbExclamation, "Chua co du lieu"
Exit Sub
End If

If Not ValidateResultRows(wsResult, lastRow) Then
LastOBGroupMessage = "RESULT dang co loi. Kiem tra cac dong bi to mau do / cot LOG truoc khi gom OB."
If Not ModuleMain.gOneClickMode Then MsgBox LastOBGroupMessage, vbExclamation, "Du lieu chua PASS"
Exit Sub
End If

If HasBlockingRunLog() Then
LastOBGroupMessage = "LOG dang con loi du lieu nghiem trong (LOI/ERROR). Hay sua loi truoc khi Gom OB."
If Not ModuleMain.gOneClickMode Then
MsgBox LastOBGroupMessage & vbCrLf & _
"THIEU TON chi la canh bao va khong chan cac dong RESULT da cap hop le.", _
vbExclamation, "LOG dang chan Gom OB"
End If
Exit Sub
End If

Dim oldScreenUpdating As Boolean
oldScreenUpdating = Application.ScreenUpdating
Application.ScreenUpdating = False
On Error GoTo CleanFail

Dim src As Variant
src = wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(lastRow, 10)).value

Dim n As Long
n = UBound(src, 1)

Dim dictGroupRows As Object, dictGroupOrder As Object
Set dictGroupRows = CreateObject("Scripting.Dictionary")
Set dictGroupOrder = CreateObject("Scripting.Dictionary")

Dim i As Long, groupKey As String
For i = 1 To n
groupKey = BuildGroupKey(CleanStrLocal(src(i, 3)), CleanStrLocal(src(i, 4)), _
CleanStrLocal(src(i, 5)), CleanStrLocal(src(i, 6)), _
CleanStrLocal(src(i, 8)), CleanStrLocal(src(i, 9)), _
splitByArea, splitByZone, groupByTL)
If Not dictGroupRows.Exists(groupKey) Then
Dim colRows As Collection
Set colRows = New Collection
dictGroupRows.Add groupKey, colRows
dictGroupOrder.Add groupKey, dictGroupOrder.count
End If
dictGroupRows(groupKey).Add i
Next i

Dim groupKeys As Variant
groupKeys = SortKeysByFirstAppearance(dictGroupRows.keys, dictGroupOrder)

On Error Resume Next
Set wsOB = ThisWorkbook.Worksheets(OB_SHEET_NAME)
On Error GoTo CleanFail
If wsOB Is Nothing Then
Application.ScreenUpdating = oldScreenUpdating
LastOBGroupMessage = "Khong tim thay sheet OB_GROUPS trong template. Tool khong tu tao sheet moi."
If Not ModuleMain.gOneClickMode Then MsgBox LastOBGroupMessage, vbCritical, "Thieu OB_GROUPS"
Exit Sub
End If

If OBHasPostedHistory(wsOB) Then
Application.ScreenUpdating = oldScreenUpdating
If ModuleMain.gOneClickMode Then
If Not allowRebuildHistory Then
LastOBGroupMessage = "OB_GROUPS dang co Delivery/Status da Post. Dat ALLOW_REBUILD_POSTED_HISTORY=TRUE neu chac chan muon tao lai."
Exit Sub
End If
Else
If MsgBox("OB_GROUPS dang co Delivery/Status da Post." & vbCrLf & _
"Gom lai se XOA toan bo lich su nay va co nguy co post trung." & vbCrLf & _
"Chi bam Yes khi chac chan muon tao lai OB_GROUPS.", _
vbExclamation + vbYesNo + vbDefaultButton2, "Canh bao xoa lich su OB") <> vbYes Then Exit Sub
End If
Application.ScreenUpdating = False
End If
PrepareOBSheet wsOB

Dim outBuf() As Variant
ReDim outBuf(1 To n, 1 To 17)
Dim outRow As Long, obNo As Long
outRow = 0
obNo = 0

Dim obProductCount As Object, obRowCount As Object, obFirstBufRow As Object, obLastBufRow As Object
Set obProductCount = CreateObject("Scripting.Dictionary")
Set obRowCount = CreateObject("Scripting.Dictionary")
Set obFirstBufRow = CreateObject("Scripting.Dictionary")
Set obLastBufRow = CreateObject("Scripting.Dictionary")

Dim runStamp As String
runStamp = Format$(Now, "yyyymmdd_hhnnss")

Dim t As Long
For t = LBound(groupKeys) To UBound(groupKeys)
Dim rowsCol As Collection
Set rowsCol = dictGroupRows(CStr(groupKeys(t)))

Dim dictProductRows As Object, productOrder As Collection
Set dictProductRows = CreateObject("Scripting.Dictionary")
Set productOrder = New Collection

Dim r As Variant, srcRow As Long, productKey As String
For Each r In rowsCol
srcRow = CLng(r)
productKey = UCase$(CleanStrLocal(src(srcRow, 1)))
If Not dictProductRows.Exists(productKey) Then
Dim pRows As Collection
Set pRows = New Collection
dictProductRows.Add productKey, pRows
productOrder.Add productKey
End If
dictProductRows(productKey).Add srcRow
Next r

Dim currentOB As String, currentProductCount As Long, currentRowCount As Long
currentOB = ""
currentProductCount = 0
currentRowCount = 0

Dim sortedProducts As Variant
sortedProducts = SortProductsByAreaZone(productOrder, dictProductRows, src)

Dim p As Variant
For Each p In sortedProducts
Dim productRows As Collection, productRowCount As Long
Set productRows = dictProductRows(CStr(p))
productRowCount = productRows.count
If productRowCount > maxRowsPerOB Then _
Err.Raise vbObjectError + 2657, , "Product " & CStr(p) & " co " & productRowCount & _
" dong Batch, vuot MAX_ROWS_PER_OB=" & maxRowsPerOB & "."

If currentProductCount = 0 Or currentProductCount >= maxProductsPerOB Or _
currentRowCount + productRowCount > maxRowsPerOB Then
obNo = obNo + 1
currentOB = "OB" & Format$(obNo, "0000")
currentProductCount = 0
currentRowCount = 0
obProductCount.Add currentOB, 0
obRowCount.Add currentOB, 0
End If

currentProductCount = currentProductCount + 1
currentRowCount = currentRowCount + productRowCount
obProductCount(currentOB) = CLng(obProductCount(currentOB)) + 1

For Each r In productRows
srcRow = CLng(r)
outRow = outRow + 1

If Not obFirstBufRow.Exists(currentOB) Then obFirstBufRow.Add currentOB, outRow
obLastBufRow(currentOB) = outRow

outBuf(outRow, 1) = "'" & CStr(currentOB)               ' OB - ma neo nhom
outBuf(outRow, 2) = "'" & CleanStrLocal(src(srcRow, 3)) ' Tram/Text
outBuf(outRow, 3) = "'" & CleanStrLocal(src(srcRow, 1)) ' Product
outBuf(outRow, 4) = src(srcRow, 2)                      ' Qty
outBuf(outRow, 5) = "'" & CleanStrLocal(src(srcRow, 4)) ' Plant
outBuf(outRow, 6) = "'" & CleanStrLocal(src(srcRow, 5)) ' SLoc
outBuf(outRow, 7) = "'" & CleanStrLocal(src(srcRow, 6)) ' SLoc Transfer
If CleanStrLocal(src(srcRow, 7)) = "" Then
outBuf(outRow, 8) = Empty                            ' NO_BATCH: giu o that su trong
Else
outBuf(outRow, 8) = "'" & CleanStrLocal(src(srcRow, 7))
End If
outBuf(outRow, 9) = "Ready"
outBuf(outRow, 13) = runStamp
outBuf(outRow, 16) = CStr(groupKeys(t))

obRowCount(currentOB) = CLng(obRowCount(currentOB)) + 1
Next r
Next p
Next t

Dim obKey As Variant, firstBufRow As Long, lastBufRow As Long, prodCnt As Long, warnText As String, bRow As Long
For Each obKey In obProductCount.keys
firstBufRow = CLng(obFirstBufRow(CStr(obKey)))
lastBufRow = CLng(obLastBufRow(CStr(obKey)))
prodCnt = CLng(obProductCount(CStr(obKey)))
If prodCnt < minTargetProductsPerOB Then
warnText = "WARNING: Don chi co " & prodCnt & " Product khac nhau do cuoi Tram khong con du ma. Van hop le, KHONG gom Tram khac de bu."
Else
warnText = ""
End If
For bRow = firstBufRow To lastBufRow
outBuf(bRow, 14) = prodCnt
outBuf(bRow, 15) = CLng(obRowCount(CStr(obKey)))
outBuf(bRow, 17) = warnText
Next bRow
Next obKey

If outRow > 0 Then
wsOB.Range(wsOB.Cells(2, 1), wsOB.Cells(1 + outRow, 17)).value = ResizeBuffer(outBuf, outRow)
End If

FormatOBSheet wsOB
ProtectOBSheet wsOB

Application.ScreenUpdating = oldScreenUpdating

LastOBGroupSucceeded = True
LastOBGroupCount = obNo
LastOBGroupMessage = "Da tao " & obNo & " OB."
If Not ModuleMain.gOneClickMode Then
MsgBox "Da tao " & obNo & " don trong sheet OB_GROUPS." & vbCrLf & _
"Rule: toi da " & maxProductsPerOB & " Product va " & maxRowsPerOB & _
" dong/OB; tat ca Batch cua cung Product luon o chung 1 OB.", _
vbInformation, "Hoan thanh"
End If
Exit Sub

CleanFail:
Application.ScreenUpdating = oldScreenUpdating
LastOBGroupMessage = "Co loi khi tao OB_GROUPS: " & Err.description
If Not ModuleMain.gOneClickMode Then MsgBox LastOBGroupMessage, vbCritical, "Loi"
End Sub

Private Function ResizeBuffer(ByVal buf As Variant, ByVal usedRows As Long) As Variant
If usedRows = UBound(buf, 1) Then
ResizeBuffer = buf
Exit Function
End If
Dim outp() As Variant, i As Long, c As Long
ReDim outp(1 To usedRows, 1 To 17)
For i = 1 To usedRows
For c = 1 To 17
outp(i, c) = buf(i, c)
Next c
Next i
ResizeBuffer = outp
End Function

Private Function CheckResultTemplate(ByVal ws As Worksheet) As Boolean
Dim headers As Variant
headers = Array("Product", "Qty", "Text", "Plant", "SLoc", "SLoc Transfer", "Batch", "Area", "Zone", "LOG")

Dim i As Long, actualValue As String
For i = LBound(headers) To UBound(headers)
actualValue = CleanStrLocal(ws.Cells(1, i + 1).value)
If UCase$(actualValue) <> UCase$(CStr(headers(i))) Then
LastOBGroupMessage = "Sai layout RESULT. Cot " & ColumnLetter(i + 1) & _
" phai la: " & headers(i) & "; hien tai: " & actualValue
If Not ModuleMain.gOneClickMode Then MsgBox LastOBGroupMessage, vbCritical, "Sai template"
CheckResultTemplate = False
Exit Function
End If
Next i

CheckResultTemplate = True
End Function

Private Function ValidateResultRows(ByVal ws As Worksheet, ByVal lastRow As Long) As Boolean
Dim hasError As Boolean
hasError = False

Dim dictBatchState As Object
Set dictBatchState = CreateObject("Scripting.Dictionary")

Dim r As Long, errMsg As String
For r = 2 To lastRow
errMsg = ""

If CleanStrLocal(ws.Cells(r, 1).value) = "" Then errMsg = errMsg & "Thieu Product; "

If CleanStrLocal(ws.Cells(r, 2).value) = "" Then
errMsg = errMsg & "Thieu Qty; "
ElseIf Not IsNumeric(ws.Cells(r, 2).value) Then
errMsg = errMsg & "Qty khong phai so; "
ElseIf CDbl(ws.Cells(r, 2).value) <= 0 Then
errMsg = errMsg & "Qty phai > 0; "
End If

If CleanStrLocal(ws.Cells(r, 3).value) = "" Then errMsg = errMsg & "Thieu Text/Tram; "
If CleanStrLocal(ws.Cells(r, 4).value) = "" Then errMsg = errMsg & "Thieu Plant; "
If CleanStrLocal(ws.Cells(r, 5).value) = "" Then errMsg = errMsg & "Thieu SLoc; "
If CleanStrLocal(ws.Cells(r, 6).value) = "" Then errMsg = errMsg & "Thieu SLoc Transfer; "
If CleanStrLocal(ws.Cells(r, 8).value) = "" Then errMsg = errMsg & "Thieu Area; "
If CleanStrLocal(ws.Cells(r, 9).value) = "" Then errMsg = errMsg & "Thieu Zone; "

If CleanStrLocal(ws.Cells(r, 5).value) <> "" And CleanStrLocal(ws.Cells(r, 6).value) <> "" Then
If CleanStrLocal(ws.Cells(r, 5).value) = CleanStrLocal(ws.Cells(r, 6).value) Then
errMsg = errMsg & "SLoc va SLoc Transfer khong duoc giong nhau; "
End If
End If

If IsBlockingResultLog(CleanStrLocal(ws.Cells(r, 10).value)) Then
errMsg = errMsg & "LOG dang bao loi; "
End If

Dim batchStateKey As String, rowBatchState As Long, oldBatchState As Long
batchStateKey = UCase$(CleanStrLocal(ws.Cells(r, 1).value)) & "|~|" & _
UCase$(CleanStrLocal(ws.Cells(r, 3).value)) & "|~|" & _
UCase$(CleanStrLocal(ws.Cells(r, 4).value)) & "|~|" & _
UCase$(CleanStrLocal(ws.Cells(r, 5).value)) & "|~|" & _
UCase$(CleanStrLocal(ws.Cells(r, 6).value))
If CleanStrLocal(ws.Cells(r, 7).value) = "" Then rowBatchState = 1 Else rowBatchState = 2
If Not dictBatchState.Exists(batchStateKey) Then
dictBatchState.Add batchStateKey, rowBatchState
Else
oldBatchState = CLng(dictBatchState(batchStateKey))
If oldBatchState <> rowBatchState Then dictBatchState(batchStateKey) = 3
End If

If errMsg <> "" Then
ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)).Interior.Color = RGB(255, 199, 206)
ws.Cells(r, 10).value = Trim$(CStr(ws.Cells(r, 10).value) & " " & errMsg)
hasError = True
Else
ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)).Interior.ColorIndex = xlNone
End If
Next r

Dim stateKey As Variant
For Each stateKey In dictBatchState.keys
If CLng(dictBatchState(CStr(stateKey))) = 3 Then
For r = 2 To lastRow
batchStateKey = UCase$(CleanStrLocal(ws.Cells(r, 1).value)) & "|~|" & _
UCase$(CleanStrLocal(ws.Cells(r, 3).value)) & "|~|" & _
UCase$(CleanStrLocal(ws.Cells(r, 4).value)) & "|~|" & _
UCase$(CleanStrLocal(ws.Cells(r, 5).value)) & "|~|" & _
UCase$(CleanStrLocal(ws.Cells(r, 6).value))
If batchStateKey = CStr(stateKey) Then
ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)).Interior.Color = RGB(255, 199, 206)
ws.Cells(r, 10).value = Trim$(CStr(ws.Cells(r, 10).value) & " LOI: Product vua co Batch vua NO_BATCH")
End If
Next r
hasError = True
End If
Next stateKey

ValidateResultRows = Not hasError
End Function

Private Function IsBlockingResultLog(ByVal logText As String) As Boolean
Dim t As String
t = UCase$(Trim$(logText))
If t = "" Or t = "NO_BATCH_OK" Then Exit Function

If Left$(t, 4) = "LOI:" Or Left$(t, 6) = "ERROR:" Or Left$(t, 5) = "FAIL:" Or _
Left$(t, 5) = "HOLD:" Then
IsBlockingResultLog = True
End If
End Function

Private Function HasBlockingRunLog() As Boolean
Dim ws As Worksheet
On Error Resume Next
Set ws = ThisWorkbook.Worksheets("LOG")
On Error GoTo 0
If ws Is Nothing Then Exit Function

Dim lastRow As Long, r As Long, logType As String
lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).Row
For r = 2 To lastRow
logType = UCase$(CleanStrLocal(ws.Cells(r, 2).value))
Select Case logType
Case "LOI", "ERROR", "SYSTEM ERROR", "CONFIG ERROR"
HasBlockingRunLog = True
Exit Function
Case "THIEU HANG", "THIEU TON", "CANH BAO", "WARNING"
End Select
Next r
End Function

Private Function OBHasPostedHistory(ByVal ws As Worksheet) As Boolean
Dim lastRow As Long, r As Long, statusText As String
lastRow = GetLastRowLocal(ws, 1)
If lastRow < 2 Then Exit Function

For r = 2 To lastRow
statusText = UCase$(CleanStrLocal(ws.Cells(r, 9).value))
If CleanStrLocal(ws.Cells(r, 11).value) <> "" Or InStr(1, statusText, "POSTED", vbTextCompare) > 0 Then
OBHasPostedHistory = True
Exit Function
End If
Next r
End Function

Private Sub PrepareOBSheet(ByVal ws As Worksheet)
On Error Resume Next
ws.Unprotect
On Error GoTo 0

Dim lastRow As Long
lastRow = GetLastRowLocal(ws, 1)
If lastRow < 2 Then lastRow = ws.Cells(ws.rows.count, 17).End(xlUp).Row
If lastRow >= 2 Then ws.Range("A2:Q" & CStr(lastRow)).ClearContents

Dim headers As Variant
headers = Array("OB", "Tram", "Product", "Qty", "Plant", "SLoc", "SLoc Transfer", "Batch", _
"Status", "SAP Message", "Delivery", "Post Time", "Run ID", "Product Count", _
"Row Count", "Group Key", "Warning")

ws.Range(ws.Cells(1, 1), ws.Cells(1, 17)).value = headers

ws.Columns("A:A").NumberFormat = "@"
ws.Columns("B:C").NumberFormat = "@"
ws.Columns("E:H").NumberFormat = "@"
ws.Columns("K:K").NumberFormat = "@"
End Sub
Private Sub FormatOBSheet(ByVal ws As Worksheet)
Dim lastRow As Long, lastCol As Long
lastRow = GetLastRowLocal(ws, 1)
lastCol = 17

With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
.Font.Bold = True
.Font.Color = RGB(255, 255, 255)
.Interior.Color = RGB(31, 78, 120)
.HorizontalAlignment = xlCenter
.VerticalAlignment = xlCenter
End With

On Error Resume Next
If ws.AutoFilterMode Then ws.AutoFilterMode = False
On Error GoTo 0

If lastRow >= 2 Then
With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Borders
.LineStyle = xlContinuous
.Weight = xlThin
.Color = RGB(210, 220, 230)
End With
End If
ws.Range(ws.Cells(1, 1), ws.Cells(IIf(lastRow < 2, 2, lastRow), lastCol)).AutoFilter

ws.Columns("A:A").NumberFormat = "@"
ws.Columns("B:C").NumberFormat = "@"
ws.Columns("E:H").NumberFormat = "@"
ws.Columns("K:K").NumberFormat = "@"
ws.Columns("L:L").NumberFormat = "dd/mm/yyyy hh:mm:ss"

ws.Columns("A:A").ColumnWidth = 14
ws.Columns("B:B").ColumnWidth = 28
ws.Columns("C:C").ColumnWidth = 18
ws.Columns("D:D").ColumnWidth = 12
ws.Columns("E:G").ColumnWidth = 12
ws.Columns("H:H").ColumnWidth = 16
ws.Columns("I:I").ColumnWidth = 18
ws.Columns("J:J").ColumnWidth = 42
ws.Columns("K:K").ColumnWidth = 16
ws.Columns("L:L").ColumnWidth = 20
ws.Columns("Q:Q").ColumnWidth = 45

ws.Columns("M:P").Hidden = True
ws.Columns("Q:Q").Hidden = False
End Sub
Private Sub ProtectOBSheet(ByVal ws As Worksheet)
On Error Resume Next
ws.Cells.Locked = False
ws.Protect DrawingObjects:=True, Contents:=False, Scenarios:=True, _
UserInterfaceOnly:=True, AllowFiltering:=True
On Error GoTo 0
End Sub

Private Function BuildGroupKey(ByVal tram As String, ByVal plant As String, _
ByVal sloc As String, ByVal slocTransfer As String, _
ByVal area As String, ByVal zone As String, _
ByVal splitByArea As Boolean, ByVal splitByZone As Boolean, _
ByVal groupByTL As Boolean) As String
Dim keyText As String, tramKey As String
If groupByTL Then tramKey = ModuleHelper.TramTLKey(tram) Else tramKey = UCase$(Trim$(tram))
keyText = tramKey & "|" & UCase$(Trim$(plant)) & "|" & _
UCase$(Trim$(sloc)) & "|" & UCase$(Trim$(slocTransfer))
If splitByArea Then keyText = keyText & "|AREA=" & UCase$(Trim$(area))
If splitByZone Then keyText = keyText & "|ZONE=" & UCase$(Trim$(zone))
BuildGroupKey = keyText
End Function

Private Function SortProductsByAreaZone(ByVal productOrder As Collection, ByVal dictProductRows As Object, ByVal src As Variant) As Variant
Dim n As Long, i As Long, j As Long
n = productOrder.count

Dim arr() As Variant, areaRank() As Long, zoneKey() As String, firstRow() As Long
ReDim arr(1 To n)
ReDim areaRank(1 To n)
ReDim zoneKey(1 To n)
ReDim firstRow(1 To n)

Dim dominantClass As Variant, currentRows As Collection
For i = 1 To n
arr(i) = CStr(productOrder(i))
dominantClass = ProductDominantAreaZone(CStr(arr(i)), dictProductRows, src)
areaRank(i) = AreaSortRank(CStr(dominantClass(0)))
zoneKey(i) = UCase$(CStr(dominantClass(1)))
Set currentRows = dictProductRows(CStr(arr(i)))
firstRow(i) = CLng(currentRows(1))
Next i

Dim tmp As Variant, tmpRank As Long, tmpZone As String, tmpRow As Long
For i = 1 To n - 1
For j = i + 1 To n
If areaRank(j) < areaRank(i) Or _
(areaRank(j) = areaRank(i) And zoneKey(j) < zoneKey(i)) Or _
(areaRank(j) = areaRank(i) And zoneKey(j) = zoneKey(i) And firstRow(j) < firstRow(i)) Then
tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
tmpRank = areaRank(i): areaRank(i) = areaRank(j): areaRank(j) = tmpRank
tmpZone = zoneKey(i): zoneKey(i) = zoneKey(j): zoneKey(j) = tmpZone
tmpRow = firstRow(i): firstRow(i) = firstRow(j): firstRow(j) = tmpRow
End If
Next j
Next i

SortProductsByAreaZone = arr
End Function

Private Function ProductDominantAreaZone(ByVal productKey As String, ByVal dictProductRows As Object, ByVal src As Variant) As Variant
Const SEP As String = "|~|"

Dim qtyByClass As Object, firstOrder As Object, displayClass As Object
Set qtyByClass = CreateObject("Scripting.Dictionary")
Set firstOrder = CreateObject("Scripting.Dictionary")
Set displayClass = CreateObject("Scripting.Dictionary")

Dim rowsCol As Collection, r As Variant, srcRow As Long
Dim areaText As String, zoneText As String, classKey As String, qty As Double
Set rowsCol = dictProductRows(productKey)

For Each r In rowsCol
srcRow = CLng(r)
areaText = CleanStrLocal(src(srcRow, 8))
zoneText = CleanStrLocal(src(srcRow, 9))
classKey = UCase$(areaText) & SEP & UCase$(zoneText)

If Not qtyByClass.Exists(classKey) Then
qtyByClass.Add classKey, 0#
firstOrder.Add classKey, firstOrder.count
displayClass.Add classKey, Array(areaText, zoneText)
End If

If IsNumeric(src(srcRow, 2)) Then qty = CDbl(src(srcRow, 2)) Else qty = 0#
qtyByClass(classKey) = CDbl(qtyByClass(classKey)) + qty
Next r

Dim k As Variant, bestKey As String, bestQty As Double, isFirst As Boolean
isFirst = True
For Each k In qtyByClass.keys
If isFirst Then
bestKey = CStr(k)
bestQty = CDbl(qtyByClass(CStr(k)))
isFirst = False
ElseIf CDbl(qtyByClass(CStr(k))) > bestQty Then
bestKey = CStr(k)
bestQty = CDbl(qtyByClass(CStr(k)))
ElseIf CDbl(qtyByClass(CStr(k))) = bestQty Then
If CLng(firstOrder(CStr(k))) < CLng(firstOrder(bestKey)) Then bestKey = CStr(k)
End If
Next k

ProductDominantAreaZone = displayClass(bestKey)
End Function

Private Function AreaSortRank(ByVal areaText As String) As Long
Select Case UCase$(Trim$(areaText))
Case "P": AreaSortRank = 1
Case "ZF": AreaSortRank = 2
Case "SPECIAL": AreaSortRank = 3
Case Else: AreaSortRank = 9
End Select
End Function

Private Function SortKeysByFirstAppearance(ByVal keys As Variant, ByVal dictOrder As Object) As Variant
Dim arr As Variant
arr = keys

Dim i As Long, j As Long, tmp As Variant
If UBound(arr) > LBound(arr) Then
For i = LBound(arr) To UBound(arr) - 1
For j = LBound(arr) To UBound(arr) - 1 - (i - LBound(arr))
If CLng(dictOrder(CStr(arr(j)))) > CLng(dictOrder(CStr(arr(j + 1)))) Then
tmp = arr(j)
arr(j) = arr(j + 1)
arr(j + 1) = tmp
End If
Next j
Next i
End If

SortKeysByFirstAppearance = arr
End Function

Private Function GetLastRowLocal(ByVal ws As Worksheet, ByVal colNo As Long) As Long
GetLastRowLocal = ws.Cells(ws.rows.count, colNo).End(xlUp).Row
End Function

Private Function CleanStrLocal(ByVal v As Variant) As String
If IsError(v) Or IsNull(v) Then
CleanStrLocal = ""
Else
CleanStrLocal = Trim$(CStr(v))
End If
End Function

Private Function ColumnLetter(ByVal colNo As Long) As String
ColumnLetter = Split(Cells(1, colNo).Address(False, False), "1")(0)
End Function
