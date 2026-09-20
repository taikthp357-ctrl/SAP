Attribute VB_Name = "ModuleFIFO"
Option Explicit

Public LastAllocationSucceeded As Boolean
Public LastAllocationHasBlockingIssues As Boolean
Public LastShortageCount As Long
Public LastShortageSummary As String

Public Sub ProcessAllocation()
LastAllocationSucceeded = False
LastAllocationHasBlockingIssues = False
ModuleHelper.ResetBinConfigCache
LastShortageCount = 0
LastShortageSummary = ""

Dim wsResult As Worksheet, wsPosting As Worksheet
Dim defaultPlant As String, defaultSLoc As String, defaultSLocTransfer As String

Set wsResult = ThisWorkbook.Sheets("RESULT")
Set wsPosting = ThisWorkbook.Sheets("POSTING")

ModuleHelper.ResetToolConfigCache
defaultPlant = Trim$(ModuleHelper.ToolConfig("DEFAULT_PLANT", "3093"))
defaultSLoc = Trim$(ModuleHelper.ToolConfig("DEFAULT_SLOC", "1005"))
defaultSLocTransfer = Trim$(ModuleHelper.ToolConfig("DEFAULT_SLOC_TRANSFER", "3001"))

PrepareResultSheet wsResult
MarkExistingOBGroupsStale

Dim postingData As Variant, orderData As Variant
Dim postingCol As Object, orderCol As Object
Dim postingMissing As Collection, orderMissing As Collection

ModuleSAP.ReadPosting postingData, postingCol, postingMissing
If postingMissing.count > 0 Then
ModuleLog.WriteLog "LOI", "Thieu header bat buoc tren sheet POSTING: " & JoinCollection(postingMissing)
Exit Sub
End If
If IsEmpty(postingData) Then
ModuleLog.WriteLog "LOI", "Sheet POSTING khong co du lieu"
Exit Sub
End If

ModuleSAP.ReadOrder orderData, orderCol, orderMissing
If orderMissing.count > 0 Then
ModuleLog.WriteLog "LOI", "Thieu header bat buoc tren sheet ORDER: " & JoinCollection(orderMissing)
Exit Sub
End If
If IsEmpty(orderData) Then
ModuleLog.WriteLog "LOI", "Sheet ORDER khong co du lieu"
Exit Sub
End If

Dim specialDict As Object
ModuleSAP.ReadSpecialProducts specialDict

Dim cProduct As Long, cBatch As Long, cAQty As Long, cGRDate As Long, cGRTime As Long
Dim cSourceBin As Long, cTst As Long

cProduct = postingCol("PRODUCT")
cBatch = postingCol("BATCH")
cAQty = postingCol("AQTY")
cGRDate = postingCol("GRDATE")
cGRTime = postingCol("GRTIME")
cSourceBin = postingCol("SOURCEBIN")

cTst = postingCol("TST")
If cTst = 0 Then cTst = postingCol("ST")

If cSourceBin = 0 Then
ModuleLog.WriteLog "LOI", "Thieu header Source Storage Bin/Source Bin tren sheet POSTING - khong the chia theo BIN_CONFIG"
Exit Sub
End If

If cTst = 0 Then
ModuleLog.WriteLog "LOI", _
"Thieu header TST/ST tren sheet POSTING - khong the xac dinh ton FT duoc phep xuat"
Exit Sub
End If

If Not ModuleHelper.ValidateBinConfig() Then
ModuleLog.WriteLog "LOI", "BIN_CONFIG co loi; dung cap hang de tranh gom sai Zone"
Exit Sub
End If

Dim nPosting As Long
nPosting = UBound(postingData, 1)

Dim dictProductBatchState As Object, dictMixedBatchLogged As Object
Set dictProductBatchState = BuildPostingBatchState(postingData, cProduct, cBatch, cAQty, cTst, cSourceBin)
Set dictMixedBatchLogged = CreateObject("Scripting.Dictionary")

Dim rowAvail() As Double, rowAlloc() As Double
Dim rowGRDateNum() As Double, rowGRTimeNum() As Double   ' cache GR Date/Time da parse (dong hop le)
ReDim rowAvail(1 To nPosting)
ReDim rowAlloc(1 To nPosting)
ReDim rowGRDateNum(1 To nPosting)
ReDim rowGRTimeNum(1 To nPosting)

Dim dictGroups As Object      ' key Product|Batch|Area|Zone -> Collection(row index)
Dim dictUnknownBins As Object
Dim dictQualityBlocked As Object
Dim dictSystemBlockedBins As Object

Set dictGroups = CreateObject("Scripting.Dictionary")
Set dictUnknownBins = CreateObject("Scripting.Dictionary")
Set dictQualityBlocked = CreateObject("Scripting.Dictionary")
Set dictSystemBlockedBins = CreateObject("Scripting.Dictionary")

Dim i As Long
For i = 1 To nPosting
rowAlloc(i) = 0
rowAvail(i) = 0

Dim prod As String, batch As String, sourceBin As String, stockStatus As String
prod = ModuleHelper.CleanStr(postingData(i, cProduct))
batch = ModuleHelper.CleanStr(postingData(i, cBatch))
sourceBin = ModuleHelper.CleanStr(postingData(i, cSourceBin))
stockStatus = UCase$(ModuleHelper.CleanStr(postingData(i, cTst)))

If prod = "" And batch = "" Then
GoTo ContinueRow ' dong tong cong cua SAP ALV (Grand Total) - bo qua am tham, khong phai loi
End If

If prod = "" Then
ModuleLog.WriteLog "LOI", "Khong tim thay ma Product", "POSTING dong " & (i + 1)
GoTo ContinueRow
End If

Dim aQty As Double, avail As Double
aQty = ModuleHelper.CleanDouble(postingData(i, cAQty))
avail = aQty   ' Available = A.Qty AUoM (SAP da tinh san so luong kha dung, khong tru Total Stk nua)

If avail <= 0 Then GoTo ContinueRow ' bo dong, KHONG log (tranh LOG rac)

If IsBlockedSourceBin(sourceBin) Then
Dim blockedBinKey As String
blockedBinKey = UCase$(prod & "|" & batch & "|" & sourceBin)

If Not dictSystemBlockedBins.Exists(blockedBinKey) Then
dictSystemBlockedBins.Add blockedBinKey, True
ModuleLog.WriteLog "CANH BAO", _
"BIN BLOCK HE THONG - BO QUA FIFO/OB | Source Bin " & sourceBin & _
" | Qty " & CStr(avail), _
"POSTING dong " & (i + 1) & " - Batch " & batch, prod
End If
GoTo ContinueRow
End If

If Not IsWarehouseAvailableStatus(stockStatus) Then
If IsQualityBlockedStatus(stockStatus) Then
Dim qualityKey As String
qualityKey = UCase$(prod & "|" & batch & "|" & stockStatus & "|" & sourceBin)

If Not dictQualityBlocked.Exists(qualityKey) Then
dictQualityBlocked.Add qualityKey, True
ModuleLog.WriteLog "CANH BAO", _
"HANG CHAT LUONG BI BLOCK | TST " & stockStatus & _
" | Khong dua vao FIFO/OB | Qty " & CStr(avail), _
"POSTING dong " & (i + 1) & " - Batch " & batch & _
" - Bin " & sourceBin, prod
End If
End If

GoTo ContinueRow
End If

Dim productBatchState As Long, productStateKey As String
productBatchState = 0
productStateKey = UCase$(prod)
If dictProductBatchState.Exists(productStateKey) Then productBatchState = CLng(dictProductBatchState(productStateKey))
If productBatchState = 3 Then
If Not dictMixedBatchLogged.Exists(productStateKey) Then
dictMixedBatchLogged.Add productStateKey, True
ModuleLog.WriteLog "LOI", "Product vua co dong Batch trong vua co dong Batch co gia tri trong POSTING; HOLD toan bo ma de tranh post sai", "Product " & prod
End If
LastAllocationHasBlockingIssues = True
GoTo ContinueRow
End If

If sourceBin = "" Then
ModuleLog.WriteLog "LOI", "Source Bin trong o dong FT co Available > 0; khong the xac dinh Area/Zone", _
"POSTING dong " & (i + 1) & " - Product " & prod & " - Batch " & batch, prod
LastAllocationHasBlockingIssues = True
GoTo ContinueRow
End If

Dim binArea As String, binZone As String
binArea = ModuleHelper.GetBinArea(sourceBin)
binZone = ModuleHelper.GetKhuVucBinDynamic(sourceBin)

If InStr(1, prod, "|", vbBinaryCompare) > 0 Or InStr(1, batch, "|", vbBinaryCompare) > 0 Or _
InStr(1, binArea, "|", vbBinaryCompare) > 0 Or InStr(1, binZone, "|", vbBinaryCompare) > 0 Then
ModuleLog.WriteLog "LOI", "Product/Batch/Zone khong duoc chua ky tu |", "POSTING dong " & (i + 1)
LastAllocationHasBlockingIssues = True
GoTo ContinueRow
End If

If binZone = "KHAC" Then
If binArea = "SPECIAL" Then binArea = "OTHER"

If Not dictUnknownBins.Exists(UCase$(sourceBin)) Then
dictUnknownBins.Add UCase$(sourceBin), True
ModuleLog.WriteLog "CANH BAO", _
"BIN NGOAI DANH MUC | Area " & binArea & _
" | Zone KHAC | Source Bin " & sourceBin, _
"POSTING dong " & (i + 1), prod
End If
End If

Dim grDateRaw As Variant, grTimeRaw As Variant
grDateRaw = postingData(i, cGRDate)
grTimeRaw = postingData(i, cGRTime)

If ModuleHelper.CleanStr(grDateRaw) = "" Then
ModuleLog.WriteLog "LOI", "GR Date trong", "POSTING dong " & (i + 1) & " - Batch " & batch
GoTo ContinueRow
End If
If ModuleHelper.CleanStr(grTimeRaw) = "" Then
ModuleLog.WriteLog "LOI", "GR Time trong", "POSTING dong " & (i + 1) & " - Batch " & batch
GoTo ContinueRow
End If

Dim gDate As Date, gTime As Double
gDate = ModuleHelper.ParseSapDate(grDateRaw)
If gDate = 0 Then
ModuleLog.WriteLog "LOI", "GR Date khong parse duoc: " & ModuleHelper.CleanStr(grDateRaw), "POSTING dong " & (i + 1) & " - Batch " & batch
GoTo ContinueRow
End If

gTime = ModuleHelper.ParseSapTime(grTimeRaw)
If gTime < 0 Then
ModuleLog.WriteLog "LOI", "GR Time khong parse duoc: " & ModuleHelper.CleanStr(grTimeRaw), "POSTING dong " & (i + 1) & " - Batch " & batch
GoTo ContinueRow
End If

rowAvail(i) = avail
rowGRDateNum(i) = CDbl(gDate)
rowGRTimeNum(i) = gTime

Dim key As String
key = prod & "|" & batch & "|" & binArea & "|" & binZone

If Not dictGroups.Exists(key) Then
Dim col As Collection
Set col = New Collection
dictGroups.Add key, col
End If
dictGroups(key).Add i

ContinueRow:
Next i

If dictGroups.count = 0 Then
ModuleLog.WriteLog "LOI", "Khong co dong POSTING hop le de phan bo"
Exit Sub
End If

Dim nGroups As Long
nGroups = dictGroups.count

Dim gKeys As Variant
gKeys = dictGroups.keys

Dim gProduct() As String, gBatch() As String, gArea() As String, gZone() As String, gAvail() As Double
Dim gBucket() As Long, gSortVal() As Double, gGRDate() As Double, gGRTime() As Double, gFirstRow() As Long
ReDim gProduct(0 To nGroups - 1)
ReDim gBatch(0 To nGroups - 1)
ReDim gArea(0 To nGroups - 1)
ReDim gZone(0 To nGroups - 1)
ReDim gAvail(0 To nGroups - 1)
ReDim gBucket(0 To nGroups - 1)
ReDim gSortVal(0 To nGroups - 1)
ReDim gGRDate(0 To nGroups - 1)
ReDim gGRTime(0 To nGroups - 1)
ReDim gFirstRow(0 To nGroups - 1)

Dim g As Long, k As String, colG As Collection, r As Variant
Dim sumAvail As Double, firstIdx As Long, parts() As String

For g = 0 To nGroups - 1
k = CStr(gKeys(g))
Set colG = SortPostingRowsByGR(dictGroups(k), rowGRDateNum, rowGRTimeNum)
dictGroups.Remove k
dictGroups.Add k, colG
sumAvail = 0
firstIdx = colG(1)
For Each r In colG
sumAvail = sumAvail + rowAvail(CLng(r))
Next r

parts = Split(k, "|")
gProduct(g) = parts(0)
gBatch(g) = parts(1)
gArea(g) = parts(2)
gZone(g) = parts(3)
gAvail(g) = sumAvail
gFirstRow(g) = firstIdx

Dim minDate As Double, minTime As Double, isFirst As Boolean
isFirst = True
For Each r In colG
Dim rIdx As Long, rDate As Double, rTime As Double
rIdx = CLng(r)
rDate = rowGRDateNum(rIdx)
rTime = rowGRTimeNum(rIdx)
If isFirst Then
minDate = rDate: minTime = rTime: isFirst = False
ElseIf (rDate < minDate) Or (rDate = minDate And rTime < minTime) Then
minDate = rDate: minTime = rTime
End If
Next r
gGRDate(g) = minDate
gGRTime(g) = minTime

Dim isHsdBySuffix As Boolean, isHsdByList As Boolean
isHsdBySuffix = (Right$(UCase$(gBatch(g)), 4) = "_HSD")
isHsdByList = False
If specialDict.Exists(UCase$(gProduct(g))) Then
isHsdByList = (UCase$(ModuleHelper.CleanStr(specialDict(UCase$(gProduct(g))))) = "HSD")
End If

If isHsdBySuffix Or isHsdByList Then
gBucket(g) = 0
If isHsdBySuffix Then
Dim expDate As Date
expDate = ModuleHelper.ParseHsdExpiry(gBatch(g))
If expDate = 0 Then
ModuleLog.WriteLog "CANH BAO", "Batch HSD sai format, khong doc duoc han dung: " & gBatch(g), "Product " & gProduct(g)
gAvail(g) = 0 ' khong tu cap khi khong xac dinh duoc han dung
gSortVal(g) = 0
Else
gSortVal(g) = CDbl(expDate)
End If
Else
Dim listedExpDate As Date
listedExpDate = ModuleHelper.ParseListedHsdExpiry(gBatch(g))
If listedExpDate = 0 Then
ModuleLog.WriteLog "CANH BAO", "Ma " & gProduct(g) & _
" thuoc danh sach HSD nhung khong doc duoc YYMMDD tu Batch: " & gBatch(g)
gAvail(g) = 0
gSortVal(g) = 0
Else
gSortVal(g) = CDbl(listedExpDate)
End If
End If
Else
gBucket(g) = 1
gSortVal(g) = CDbl(GetBatchPriority(gBatch(g)))
End If
Next g

Dim idxArr() As Long
ReDim idxArr(0 To nGroups - 1)
For g = 0 To nGroups - 1
idxArr(g) = g
Next g
ModuleHelper.QuickSortGroups idxArr, gProduct, gBucket, gSortVal, gGRDate, gGRTime, gFirstRow, 0, nGroups - 1

Dim gRemaining() As Double
ReDim gRemaining(0 To nGroups - 1)
For g = 0 To nGroups - 1
gRemaining(g) = gAvail(g)
Next g

Dim dictResultQty As Object, dictResultMeta As Object
Dim dictDominantArea As Object, dictDominantZone As Object, dictClassQty As Object
Dim dictResultNotes As Object
Dim resultOrder As Collection
Set dictResultQty = CreateObject("Scripting.Dictionary")
Set dictResultMeta = CreateObject("Scripting.Dictionary")
Set dictDominantArea = CreateObject("Scripting.Dictionary")
Set dictDominantZone = CreateObject("Scripting.Dictionary")
Set dictClassQty = CreateObject("Scripting.Dictionary")
Set dictResultNotes = CreateObject("Scripting.Dictionary")
Set resultOrder = New Collection

Dim cOProduct As Long, cOQty As Long, cOTram As Long
cOProduct = orderCol("PRODUCT")
cOQty = orderCol("QTY")
cOTram = orderCol("TRAM")

Dim nOrder As Long
nOrder = UBound(orderData, 1)

Dim o As Long, oProduct As String, oTram As String, oQty As Double, needQty As Double
Dim validOrderCount As Long
For o = 1 To nOrder
oProduct = ModuleHelper.CleanStr(orderData(o, cOProduct))
oQty = ModuleHelper.CleanDouble(orderData(o, cOQty))
oTram = ModuleHelper.CleanStr(orderData(o, cOTram))

If oProduct = "" And ModuleHelper.CleanStr(orderData(o, cOQty)) = "" And oTram = "" Then GoTo NextOrder

If oProduct = "" Then
ModuleLog.WriteLog "LOI", "ORDER thieu Product", "ORDER dong " & (o + 1)
LastAllocationHasBlockingIssues = True
GoTo NextOrder
End If
If oQty <= 0 Then
ModuleLog.WriteLog "LOI", "ORDER Qty khong hop le hoac <= 0", "ORDER dong " & (o + 1) & " - Product " & oProduct
LastAllocationHasBlockingIssues = True
GoTo NextOrder
End If
If oTram = "" Then
ModuleLog.WriteLog "LOI", "ORDER thieu Tram/Text; dong nay khong duoc cap de tranh post sai Text len SAP", "ORDER dong " & (o + 1) & " - Product " & oProduct
LastAllocationHasBlockingIssues = True
GoTo NextOrder
End If

validOrderCount = validOrderCount + 1
needQty = oQty

Dim gi As Long, realIdx As Long, allocQty As Double, actualAlloc As Double, keyMark As String
For gi = 0 To nGroups - 1
If needQty <= 0 Then Exit For
realIdx = idxArr(gi)
If StrComp(gProduct(realIdx), oProduct, vbTextCompare) = 0 And gRemaining(realIdx) > 0 Then
allocQty = ModuleHelper.MinD(needQty, gRemaining(realIdx))

Dim allocatedBinsCsv As String, allocationNote As String
allocatedBinsCsv = ""
allocationNote = ""

keyMark = gProduct(realIdx) & "|" & gBatch(realIdx) & "|" & gArea(realIdx) & "|" & gZone(realIdx)
actualAlloc = MarkPostingRows(dictGroups(keyMark), rowAvail, rowAlloc, allocQty, _
oTram, oProduct, gBatch(realIdx), gZone(realIdx), _
postingData, cSourceBin, allocatedBinsCsv)

If UCase$(Trim$(gZone(realIdx))) = "KHAC" Then
allocationNote = "CANH BAO: BIN NGOAI DANH MUC"
If allocatedBinsCsv <> "" Then allocationNote = allocationNote & " [" & allocatedBinsCsv & "]"
End If

If actualAlloc > 0 Then
AddConsolidatedResult dictResultQty, dictResultMeta, dictDominantArea, dictDominantZone, _
dictClassQty, dictResultNotes, resultOrder, _
oTram, oProduct, gBatch(realIdx), gArea(realIdx), gZone(realIdx), _
actualAlloc, defaultPlant, defaultSLoc, defaultSLocTransfer, allocationNote
gRemaining(realIdx) = gRemaining(realIdx) - actualAlloc
needQty = needQty - actualAlloc
End If

If Abs(actualAlloc - allocQty) > 0.000001 Then
ModuleLog.WriteLog "LOI", "Sai lech noi bo giua Qty pool va Qty phan bo: yeu cau " & allocQty & ", cap duoc " & actualAlloc, _
"Product " & oProduct & " - Batch " & gBatch(realIdx) & " - Zone " & gZone(realIdx)
LastAllocationHasBlockingIssues = True
End If
End If
Next gi

If needQty > 0 Then
Dim allocatedOrderQty As Double
allocatedOrderQty = oQty - needQty
If allocatedOrderQty < 0 Then allocatedOrderQty = 0

LastShortageCount = LastShortageCount + 1

If allocatedOrderQty <= 0 Then
ModuleLog.WriteLog "THIEU TON", "HET TON KHO", "Tram " & oTram, oProduct

If LastShortageCount <= 10 Then
If LastShortageSummary <> "" Then LastShortageSummary = LastShortageSummary & vbCrLf
LastShortageSummary = LastShortageSummary & "- " & oProduct & " | Tram " & oTram & " | HET TON KHO"
End If
Else
ModuleLog.WriteLog "THIEU TON", _
"Yeu cau " & CStr(oQty) & " | Da cap " & CStr(allocatedOrderQty) & _
" | Con thieu " & CStr(needQty), _
"Tram " & oTram, oProduct

If LastShortageCount <= 10 Then
If LastShortageSummary <> "" Then LastShortageSummary = LastShortageSummary & vbCrLf
LastShortageSummary = LastShortageSummary & "- " & oProduct & " | Tram " & oTram & _
" | Yeu cau " & CStr(oQty) & " | Da cap " & CStr(allocatedOrderQty) & _
" | Con thieu " & CStr(needQty)
End If
End If
End If
NextOrder:
Next o

If validOrderCount = 0 Then
ModuleLog.WriteLog "LOI", "Khong co dong ORDER hop le de cap hang"
LastAllocationHasBlockingIssues = True
End If

Dim nResult As Long
nResult = resultOrder.count

If nResult > 0 Then
Dim resBuf() As Variant
ReDim resBuf(1 To nResult, 1 To 10)

Dim resultKey As Variant, meta As Variant
Dim dominantArea As String, dominantZone As String, resultLog As String
Dim bufRow As Long
bufRow = 0

For Each resultKey In resultOrder
meta = dictResultMeta(CStr(resultKey))
dominantArea = CStr(dictDominantArea(CStr(resultKey)))
dominantZone = CStr(dictDominantZone(CStr(resultKey)))
resultLog = ""
If ModuleHelper.CleanStr(meta(2)) = "" Then resultLog = "NO_BATCH_OK"
If dictResultNotes.Exists(CStr(resultKey)) Then
resultLog = MergeResultNote(resultLog, CStr(dictResultNotes(CStr(resultKey))))
End If

bufRow = bufRow + 1
FillResultBufferRow resBuf, bufRow, CStr(meta(1)), CDbl(dictResultQty(CStr(resultKey))), CStr(meta(0)), _
CStr(meta(3)), CStr(meta(4)), CStr(meta(5)), CStr(meta(2)), _
dominantArea, dominantZone, resultLog
Next resultKey

wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(1 + nResult, 10)).value = resBuf
End If

HighlightPosting wsPosting, nPosting, rowAlloc
LastAllocationSucceeded = True
End Sub

Private Sub FillResultBufferRow(ByRef buf() As Variant, ByVal bufRow As Long, ByVal product As String, ByVal qty As Double, _
ByVal tramText As String, ByVal plant As String, ByVal sloc As String, _
ByVal slocTransfer As String, ByVal batch As String, ByVal area As String, _
ByVal zone As String, ByVal logText As String)
buf(bufRow, 1) = product
buf(bufRow, 2) = qty
buf(bufRow, 3) = tramText
buf(bufRow, 4) = "'" & plant
buf(bufRow, 5) = "'" & sloc
buf(bufRow, 6) = "'" & slocTransfer
If Trim$(batch) = "" Then
buf(bufRow, 7) = Empty
Else
buf(bufRow, 7) = "'" & batch
End If
buf(bufRow, 8) = area
buf(bufRow, 9) = zone
buf(bufRow, 10) = logText
End Sub

Private Sub AddConsolidatedResult(ByVal dictQty As Object, ByVal dictMeta As Object, _
ByVal dictArea As Object, ByVal dictZone As Object, _
ByVal dictClassQty As Object, ByVal dictNotes As Object, _
ByVal resultOrder As Collection, _
ByVal tram As String, ByVal product As String, ByVal batch As String, _
ByVal area As String, ByVal zone As String, ByVal qty As Double, _
ByVal plant As String, ByVal sloc As String, ByVal slocTransfer As String, _
ByVal noteText As String)
Const SEP As String = "|~|"

Dim resultKey As String
resultKey = UCase$(Trim$(tram)) & SEP & UCase$(Trim$(product)) & SEP & UCase$(Trim$(batch)) & SEP & _
UCase$(Trim$(plant)) & SEP & UCase$(Trim$(sloc)) & SEP & UCase$(Trim$(slocTransfer))

If Not dictQty.Exists(resultKey) Then
dictQty.Add resultKey, 0#
dictMeta.Add resultKey, Array(tram, product, batch, plant, sloc, slocTransfer)
dictArea.Add resultKey, area
dictZone.Add resultKey, zone
dictNotes.Add resultKey, ""
resultOrder.Add resultKey
End If

dictQty(resultKey) = CDbl(dictQty(resultKey)) + qty

If Trim$(noteText) <> "" Then
dictNotes(resultKey) = MergeResultNote(CStr(dictNotes(resultKey)), noteText)
End If

Dim classKey As String, currentClassKey As String
classKey = resultKey & SEP & UCase$(Trim$(area)) & SEP & UCase$(Trim$(zone))
If Not dictClassQty.Exists(classKey) Then dictClassQty.Add classKey, 0#
dictClassQty(classKey) = CDbl(dictClassQty(classKey)) + qty

currentClassKey = resultKey & SEP & UCase$(Trim$(CStr(dictArea(resultKey)))) & SEP & _
UCase$(Trim$(CStr(dictZone(resultKey))))
If Not dictClassQty.Exists(currentClassKey) Then dictClassQty.Add currentClassKey, 0#

If CDbl(dictClassQty(classKey)) > CDbl(dictClassQty(currentClassKey)) Then
dictArea(resultKey) = area
dictZone(resultKey) = zone
End If
End Sub

Private Sub PrepareResultSheet(ByVal ws As Worksheet)
Dim lastRow As Long
lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row
If lastRow < 2 Then lastRow = ws.Cells(ws.rows.count, 10).End(xlUp).Row
If lastRow >= 2 Then ws.Range("A2:J" & CStr(lastRow)).ClearContents

ws.Range("A1:J1").value = Array("Product", "Qty", "Text", "Plant", "SLoc", "SLoc Transfer", "Batch", "Area", "Zone", "LOG")
ws.Columns(4).NumberFormat = "@"
ws.Columns(5).NumberFormat = "@"
ws.Columns(6).NumberFormat = "@"
ws.Columns(7).NumberFormat = "@"
End Sub

Private Sub MarkExistingOBGroupsStale()
Dim ws As Worksheet
On Error Resume Next
Set ws = ThisWorkbook.Worksheets("OB_GROUPS")
On Error GoTo 0
If ws Is Nothing Then Exit Sub

Dim lastRow As Long, r As Long, statusText As String
lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row
For r = 2 To lastRow
statusText = UCase$(ModuleHelper.CleanStr(ws.Cells(r, 9).value))
If statusText = "READY" And ModuleHelper.CleanStr(ws.Cells(r, 11).value) = "" Then
ws.Cells(r, 9).value = "Stale - Regenerate"
ws.Cells(r, 10).value = "FIFO da chay lai; can GOM OB lai truoc khi Post SAP."
End If
Next r
End Sub

Private Function SortPostingRowsByGR(ByVal rowsCol As Collection, ByRef rowGRDateNum() As Double, _
ByRef rowGRTimeNum() As Double) As Collection
Dim n As Long, i As Long, j As Long, tmp As Long
n = rowsCol.count

Dim arr() As Long
ReDim arr(1 To n)
For i = 1 To n
arr(i) = CLng(rowsCol(i))
Next i

For i = 2 To n
tmp = arr(i)
j = i - 1
Do While j >= 1
Dim dLeft As Double, dTmp As Double, tLeft As Double, tTmp As Double
dLeft = rowGRDateNum(arr(j))
dTmp = rowGRDateNum(tmp)
tLeft = rowGRTimeNum(arr(j))
tTmp = rowGRTimeNum(tmp)

If dLeft < dTmp Then Exit Do
If dLeft = dTmp And tLeft < tTmp Then Exit Do
If dLeft = dTmp And tLeft = tTmp And arr(j) < tmp Then Exit Do

arr(j + 1) = arr(j)
j = j - 1
Loop
arr(j + 1) = tmp
Next i

Dim sortedCol As Collection
Set sortedCol = New Collection
For i = 1 To n
sortedCol.Add arr(i)
Next i
Set SortPostingRowsByGR = sortedCol
End Function

Private Function BuildPostingBatchState(ByVal postingData As Variant, ByVal cProduct As Long, _
ByVal cBatch As Long, ByVal cAQty As Long, _
ByVal cTst As Long, ByVal cSourceBin As Long) As Object
Dim d As Object
Set d = CreateObject("Scripting.Dictionary")

Dim i As Long
Dim prod As String, batch As String, key As String
Dim sourceBin As String, stockStatus As String
Dim avail As Double, rowState As Long, oldState As Long

For i = 1 To UBound(postingData, 1)
prod = ModuleHelper.CleanStr(postingData(i, cProduct))

If prod <> "" Then
avail = ModuleHelper.CleanDouble(postingData(i, cAQty))
stockStatus = UCase$(ModuleHelper.CleanStr(postingData(i, cTst)))
sourceBin = ModuleHelper.CleanStr(postingData(i, cSourceBin))

If avail > 0 And IsWarehouseAvailableStatus(stockStatus) And _
sourceBin <> "" And Not IsBlockedSourceBin(sourceBin) Then

batch = ModuleHelper.CleanStr(postingData(i, cBatch))

If batch = "" Then
rowState = 1 ' NO_BATCH
Else
rowState = 2 ' BATCH
End If

key = UCase$(prod)

If Not d.Exists(key) Then
d.Add key, rowState
Else
oldState = CLng(d(key))
If oldState <> rowState Then d(key) = 3 ' MIXED -> HOLD
End If
End If
End If
Next i

Set BuildPostingBatchState = d
End Function

Private Function IsWarehouseAvailableStatus(ByVal stockStatus As String) As Boolean
IsWarehouseAvailableStatus = (UCase$(Trim$(stockStatus)) = "FT")
End Function

Private Function IsQualityBlockedStatus(ByVal stockStatus As String) As Boolean
stockStatus = UCase$(Trim$(stockStatus))
IsQualityBlockedStatus = (stockStatus = "QT" Or stockStatus = "QI")
End Function

Private Function IsBlockedSourceBin(ByVal sourceBin As String) As Boolean
sourceBin = UCase$(Trim$(sourceBin))

Select Case sourceBin
Case "APFL-BL01-01-01"
IsBlockedSourceBin = True
Case Else
IsBlockedSourceBin = False
End Select
End Function

Public Function GetBatchPriority(ByVal batch As String) As Long
Dim b As String
b = UCase$(batch)
If Right$(b, 3) = "-SC" Then
GetBatchPriority = 1
ElseIf Right$(b, 4) = "-UTX" Then
GetBatchPriority = 2
Else
GetBatchPriority = 3
End If
End Function

Public Function MarkPostingRows(rowsCol As Collection, rowAvail() As Double, rowAlloc() As Double, ByVal qty As Double, _
ByVal tram As String, ByVal product As String, ByVal batch As String, ByVal zone As String, _
ByVal postingData As Variant, ByVal cSourceBin As Long, _
ByRef allocatedBinsCsv As String) As Double
Dim r As Variant, idx As Long, available As Double, take As Double
Dim remain As Double, allocatedTotal As Double
remain = qty
For Each r In rowsCol
If remain <= 0 Then Exit For
idx = CLng(r)
available = rowAvail(idx) - rowAlloc(idx)
If available > 0 Then
take = ModuleHelper.MinD(remain, available)
rowAlloc(idx) = rowAlloc(idx) + take
allocatedTotal = allocatedTotal + take
remain = remain - take
AppendUniqueCsv allocatedBinsCsv, ModuleHelper.CleanStr(postingData(idx, cSourceBin))

End If
Next r
MarkPostingRows = allocatedTotal
End Function

Private Function MergeResultNote(ByVal existingText As String, ByVal newText As String) As String
existingText = Trim$(existingText)
newText = Trim$(newText)

If newText = "" Then
MergeResultNote = existingText
ElseIf existingText = "" Then
MergeResultNote = newText
ElseIf InStr(1, existingText, newText, vbTextCompare) > 0 Then
MergeResultNote = existingText
Else
MergeResultNote = existingText & " | " & newText
End If
End Function

Private Sub AppendUniqueCsv(ByRef csvText As String, ByVal itemText As String)
itemText = Trim$(itemText)
If itemText = "" Then Exit Sub

Dim items As Variant
Dim i As Long

If Trim$(csvText) <> "" Then
items = Split(csvText, ",")
For i = LBound(items) To UBound(items)
If StrComp(Trim$(CStr(items(i))), itemText, vbTextCompare) = 0 Then Exit Sub
Next i
csvText = csvText & ", " & itemText
Else
csvText = itemText
End If
End Sub

Public Sub HighlightPosting(ws As Worksheet, ByVal nPosting As Long, rowAlloc() As Double)
Dim lastCol As Long, allocCol As Long, c As Long
lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

allocCol = 0
For c = 1 To lastCol
If ModuleSAP.NormalizeHeader(ModuleHelper.CleanStr(ws.Cells(1, c).value)) = "ALLOCATEDQTY" Then
allocCol = c
Exit For
End If
Next c
If allocCol = 0 Then allocCol = lastCol + 1

ws.Cells(1, allocCol).value = "Allocated Qty"
ws.Cells(1, allocCol).Font.Bold = True

If nPosting <= 0 Then Exit Sub

ws.Range(ws.Cells(2, 1), ws.Cells(1 + nPosting, allocCol)).Interior.ColorIndex = xlColorIndexNone

Dim allocBuf() As Variant
ReDim allocBuf(1 To nPosting, 1 To 1)
Dim i As Long
For i = 1 To nPosting
allocBuf(i, 1) = rowAlloc(i)
Next i
ws.Range(ws.Cells(2, allocCol), ws.Cells(1 + nPosting, allocCol)).value = allocBuf

Dim runStart As Long
runStart = 0
For i = 1 To nPosting
If rowAlloc(i) > 0 Then
If runStart = 0 Then runStart = i
Else
If runStart > 0 Then
ws.Range(ws.Cells(1 + runStart, 1), ws.Cells(i, allocCol)).Interior.Color = RGB(198, 239, 206)
runStart = 0
End If
End If
Next i
If runStart > 0 Then
ws.Range(ws.Cells(1 + runStart, 1), ws.Cells(1 + nPosting, allocCol)).Interior.Color = RGB(198, 239, 206)
End If
End Sub

Private Function JoinCollection(c As Collection) As String
Dim s As String, itm As Variant
For Each itm In c
If s <> "" Then s = s & ", "
s = s & CStr(itm)
Next itm
JoinCollection = s
End Function

Public Function OneClickSelectMode(ByVal scopeText As String, ByVal hasOldData As Boolean, ByVal readyCount As Long, ByVal postedCount As Long, ByVal staleCount As Long, ByVal confirmBeforeSap As Boolean) As Long
Dim c As VbMsgBoxResult, t As String
If scopeText = "BUILD_ONLY" Then
If hasOldData Then
If MsgBox("TAO LAI FIFO/OB?" & vbCrLf & "YES = xoa RESULT + OB_GROUPS cu; NO = huy.", vbQuestion + vbYesNo + vbDefaultButton2, "LUOT MOI") <> vbYes Then Exit Function
End If
OneClickSelectMode = 1: Exit Function
End If
If hasOldData Then
t = "CHON CHE DO CHAY:" & vbCrLf & vbCrLf & "YES = LUOT MOI: xoa RESULT + OB_GROUPS, tao lai tu ORDER hien tai." & vbCrLf & "NO = TIEP TUC: giu du lieu cu, chi chay OB Ready." & vbCrLf & "CANCEL = thoat." & vbCrLf & vbCrLf & "Ready=" & readyCount & "; Posted=" & postedCount & "; Stale=" & staleCount
If postedCount > 0 Then t = t & vbCrLf & "LUU Y: xoa Excel KHONG hoan tac Delivery da tao tren SAP."
c = MsgBox(t, vbQuestion + vbYesNoCancel + vbDefaultButton1, "CHAY TOAN BO")
If c = vbYes Then OneClickSelectMode = 1: Exit Function
If c = vbNo Then
If readyCount > 0 Or postedCount > 0 Then OneClickSelectMode = 2 Else MsgBox "Khong co OB de tiep tuc. Chon YES de tao luot moi.", vbInformation, "Khong co luot dang do"
End If
Exit Function
End If
If confirmBeforeSap Then
If MsgBox("BAT DAU LUOT MOI?" & vbCrLf & "FIFO -> GOM OB -> MIGO 311", vbQuestion + vbYesNo + vbDefaultButton1, "CHAY TOAN BO") <> vbYes Then Exit Function
End If
OneClickSelectMode = 1
End Function

Public Sub OneClickClearRunData(Optional ByVal clearPosting As Boolean = True)
Dim ws As Worksheet, lr As Long
Set ws = ThisWorkbook.Worksheets("RESULT"): lr = ws.Cells(ws.rows.count, 1).End(xlUp).Row
If lr >= 2 Then ws.Range("A2:J" & CStr(lr)).ClearContents
Set ws = ThisWorkbook.Worksheets("OB_GROUPS"): lr = ws.Cells(ws.rows.count, 1).End(xlUp).Row
If lr >= 2 Then ws.Range("A2:Q" & CStr(lr)).ClearContents
If clearPosting Then OneClickClearPostingData
End Sub

Public Sub OneClickClearPostingData()
Dim ws As Worksheet, c As Long, candidate As Long, lr As Long
Set ws = ThisWorkbook.Worksheets("POSTING")
lr = 1
For c = 1 To 17
    candidate = ws.Cells(ws.rows.count, c).End(xlUp).Row
    If candidate > lr Then lr = candidate
Next c
If lr >= 2 Then
    ws.Range("A2:Q" & CStr(lr)).ClearContents
    ws.Range("A2:Q" & CStr(lr)).Interior.ColorIndex = xlColorIndexNone
End If
End Sub
