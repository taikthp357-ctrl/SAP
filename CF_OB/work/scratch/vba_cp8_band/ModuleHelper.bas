Attribute VB_Name = "ModuleHelper"
Option Explicit
Private mToolConfigLoaded As Boolean
Private mToolConfigCache As Object

Private mBinLoaded As Boolean
Private mBinCount As Long
Private mBinPatterns() As String
Private mBinZones() As String

' Hang thung cache (must stay in module declaration section)
Private mBoxLoaded As Boolean
Private mBoxCatalog As Object
Private mBoxRules As Collection
Private mBoxDuplicateLogged As Object

Public Function GetLastRow(ws As Worksheet, ByVal col As Long) As Long
Dim lr As Long
lr = ws.Cells(ws.rows.count, col).End(xlUp).row
If lr = 1 And ws.Cells(1, col).Value = "" Then
GetLastRow = 0
Else
GetLastRow = lr
End If
End Function

' === HANG THUNG: catalogue, packing rules and ORDER preflight ===

Public Sub ResetBoxCache()
    mBoxLoaded = False
    Set mBoxCatalog = Nothing
    Set mBoxRules = Nothing
    Set mBoxDuplicateLogged = Nothing
End Sub

Private Function BoxKey(ByVal v As Variant) As String
    BoxKey = UCase$(Trim$(CStr(v & "")))
End Function

Private Function BoxIsTrueFalse(ByVal v As Variant) As Boolean
    Dim s As String: s = UCase$(Trim$(CStr(v & "")))
    BoxIsTrueFalse = (s = "TRUE" Or s = "FALSE")
End Function

Private Function LoadBoxSetup(ByRef errOut As String) As Boolean
    Dim ws As Worksheet, lr As Long, data As Variant, i As Long, k As String, rec As Variant
    Dim wsc As Worksheet, rlr As Long, rd As Variant, a As Long, b As Long, n As Long, active As String
    errOut = "": Set mBoxCatalog = CreateObject("Scripting.Dictionary"): mBoxCatalog.CompareMode = vbTextCompare
    Set mBoxRules = New Collection: Set mBoxDuplicateLogged = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("DANH_MUC")
    Set wsc = ThisWorkbook.Worksheets("SYSTEM")
    On Error GoTo 0
    If ws Is Nothing Then errOut = "Khong tim thay sheet DANH_MUC.": Exit Function
    If wsc Is Nothing Then errOut = "Khong tim thay sheet CONFIG.": Exit Function

    ' DANH_MUC: D=Model, E=Material, F=Loai thung, G=Packing; header row 4, data row 5+
    lr = ws.Cells(ws.rows.count, 5).End(xlUp).row
    If lr >= 5 Then
        data = ws.Range("D5:G" & CStr(lr)).Value2
        For i = 1 To UBound(data, 1)
            k = BoxKey(data(i, 2))
            If k <> "" Then
                If mBoxCatalog.Exists(k) Then
                    If Not mBoxDuplicateLogged.Exists(k) Then
                        ModuleLog.WriteLog "CANH BAO", "Material trung trong DANH_MUC - HANG THUNG; chi dung dong dau tien.", "DANH_MUC dong " & (i + 4), k
                        mBoxDuplicateLogged.Add k, True
                    End If
                Else
                    rec = Array(Trim$(CStr(data(i, 1) & "")), data(i, 4), i + 4)
                    mBoxCatalog.Add k, rec
                End If
            End If
        Next i
    End If

    ' Packing config table: CONFIG G1:J5; title row 1, header row 2, data row 3+
    rlr = wsc.Cells(wsc.rows.count, 23).End(xlUp).row
    If rlr < 3 Then errOut = "CONFIG chua co bang gioi han thung tai cot G:J.": Exit Function
    rd = wsc.Range("W3:Z" & CStr(rlr)).Value2
    For i = 1 To UBound(rd, 1)
        If Trim$(CStr(rd(i, 1) & "")) <> "" Or Trim$(CStr(rd(i, 2) & "")) <> "" Or Trim$(CStr(rd(i, 3) & "")) <> "" Or Trim$(CStr(rd(i, 4) & "")) <> "" Then
            If Not IsNumeric(rd(i, 1)) Or Not IsNumeric(rd(i, 2)) Then errOut = errOut & "Packing Min/Max khong hop le: CONFIG dong " & (i + 2) & "." & vbCrLf: GoTo NextRule
            a = CLng(rd(i, 1)): b = CLng(rd(i, 2))
            If a > b Then errOut = errOut & "Packing Min lon hon Max: CONFIG dong " & (i + 2) & "." & vbCrLf
            If Not IsNumeric(rd(i, 3)) Or CLng(rd(i, 3)) <= 0 Or CDbl(rd(i, 3)) <> CLng(rd(i, 3)) Then errOut = errOut & "So thung phai la so nguyen duong: CONFIG dong " & (i + 2) & "." & vbCrLf
            If Not BoxIsTrueFalse(rd(i, 4)) Then errOut = errOut & "Kich hoat phai TRUE/FALSE: CONFIG dong " & (i + 2) & "." & vbCrLf
            active = UCase$(Trim$(CStr(rd(i, 4) & "")))
            If active = "TRUE" And IsNumeric(rd(i, 1)) And IsNumeric(rd(i, 2)) And IsNumeric(rd(i, 3)) Then mBoxRules.Add Array(a, b, CLng(rd(i, 3)), i + 2)
        End If
NextRule:
    Next i
    For i = 1 To mBoxRules.count
        For n = i + 1 To mBoxRules.count
            If CLng(mBoxRules(i)(0)) <= CLng(mBoxRules(n)(1)) And CLng(mBoxRules(n)(0)) <= CLng(mBoxRules(i)(1)) Then errOut = errOut & "Hai khoang Packing active bi chong len: CONFIG dong " & mBoxRules(i)(3) & " va " & mBoxRules(n)(3) & "." & vbCrLf
        Next n
    Next i
    ' Preserve the stable baseline safety rule: no active valid packing rule = STOP preflight.
    If mBoxRules.count = 0 Then errOut = errOut & "Khong co rule Packing active hop le." & vbCrLf
    If errOut <> "" Then Exit Function
    mBoxLoaded = True: LoadBoxSetup = True
End Function

Private Function BoxCountForPacking(ByVal packing As Double, ByRef boxCount As Long) As Boolean
    Dim i As Long, r As Variant
    For i = 1 To mBoxRules.count
        r = mBoxRules(i)
        If packing >= CDbl(r(0)) And packing <= CDbl(r(1)) Then boxCount = CLng(r(2)): BoxCountForPacking = True: Exit Function
    Next i
End Function

Public Function PrepareBoxOrders(ByRef errOut As String) As Boolean
    Dim ws As Worksheet, lr As Long, data As Variant, out() As Variant, i As Long, k As String, rec As Variant
    Dim packing As Double, boxes As Long, newQty As Double, oldQty As String
    ResetBoxCache
    If Not LoadBoxSetup(errOut) Then Exit Function
    Set ws = ThisWorkbook.Worksheets("ORDER"): lr = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lr < 2 Then PrepareBoxOrders = True: Exit Function
    data = ws.Range("A2:C" & CStr(lr)).Value2: ReDim out(1 To UBound(data, 1), 1 To 1)
    For i = 1 To UBound(data, 1)
        out(i, 1) = data(i, 2): k = BoxKey(data(i, 1))
        If k <> "" And mBoxCatalog.Exists(k) Then
            rec = mBoxCatalog(k)
            If Trim$(CStr(rec(0))) = "" Then errOut = errOut & "Hang thung dong dau tien thieu Model: Material " & k & ", dong " & rec(2) & "." & vbCrLf: GoTo NextOrder
            If Not IsNumeric(rec(1)) Or CDbl(rec(1)) <= 0 Then errOut = errOut & "Packing khong hop le: Material " & k & ", dong " & rec(2) & "." & vbCrLf: GoTo NextOrder
            packing = CDbl(rec(1))
            If Not BoxCountForPacking(packing, boxes) Then errOut = errOut & "Khong tim thay rule Packing cho Material " & k & " (Packing " & packing & ")." & vbCrLf: GoTo NextOrder
            newQty = packing * boxes: out(i, 1) = newQty
NextOrder:
        End If
    Next i
    If errOut <> "" Then Exit Function
    ws.Range("B2:B" & CStr(lr)).Value = out
    For i = 1 To UBound(data, 1)
        k = BoxKey(data(i, 1))
        If k <> "" And mBoxCatalog.Exists(k) Then
            rec = mBoxCatalog(k): packing = CDbl(rec(1)): Call BoxCountForPacking(packing, boxes)
            ModuleLog.WriteLog "HANG THUNG", "Model=" & rec(0) & "; Packing=" & packing & "; So thung=" & boxes & "; Qty cu=" & data(i, 2) & "; Qty moi=" & out(i, 1), "ORDER dong " & (i + 1), k
        End If
    Next i
    PrepareBoxOrders = True
End Function

Public Function GetBoxModel(ByVal product As String) As String
    Dim e As String, rec As Variant
    If Not mBoxLoaded Then Exit Function
    If mBoxCatalog.Exists(BoxKey(product)) Then rec = mBoxCatalog(BoxKey(product)): GetBoxModel = CStr(rec(0))
End Function

Public Function OrderSignature() As String
    Dim ws As Worksheet, lr As Long, d As Variant, i As Long, s As String, h As Double
    Set ws = ThisWorkbook.Worksheets("ORDER"): lr = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lr < 2 Then OrderSignature = "0": Exit Function
    d = ws.Range("A2:C" & CStr(lr)).Value2
    For i = 1 To UBound(d, 1): s = s & CStr(d(i, 1)) & Chr$(30) & CStr(d(i, 2)) & Chr$(30) & CStr(d(i, 3)) & Chr$(31): Next i
    For i = 1 To Len(s): h = (h * 31# + AscW(Mid$(s, i, 1))) Mod 68999981#: Next i
    OrderSignature = CStr(lr) & ":" & Format$(h, "0")
End Function

Public Function StoredOrderSignature() As String
    On Error Resume Next: StoredOrderSignature = Replace(ThisWorkbook.names("BOX_ORDER_SIGNATURE").RefersTo, "=" & Chr$(34), ""): StoredOrderSignature = Replace(StoredOrderSignature, Chr$(34), ""): On Error GoTo 0
End Function

Public Sub SaveOrderSignature()
    Dim v As String: v = OrderSignature()
    On Error Resume Next: ThisWorkbook.names("BOX_ORDER_SIGNATURE").Delete: On Error GoTo 0
    ThisWorkbook.names.Add "BOX_ORDER_SIGNATURE", "=" & Chr$(34) & v & Chr$(34)
End Sub

Public Function OrderDataIsStale() As Boolean
    Dim oldSig As String: oldSig = StoredOrderSignature()
    If oldSig = "" Then OrderDataIsStale = True Else OrderDataIsStale = (oldSig <> OrderSignature())
End Function


Public Function CleanStr(ByVal v As Variant) As String
If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
CleanStr = ""
Else
CleanStr = Trim$(CStr(v))
End If
End Function

Public Function CleanDouble(ByVal v As Variant) As Double
On Error GoTo Invalid

If IsNumeric(v) And VarType(v) <> vbString Then
CleanDouble = CDbl(v)
Exit Function
End If

Dim s As String
s = Trim$(CStr(v & ""))
If s = "" Then GoTo Invalid

s = Replace(s, " ", "")

Dim pDot As Long, pComma As Long
pDot = InStrRev(s, ".")
pComma = InStrRev(s, ",")

If pDot > 0 And pComma > 0 Then
If pDot > pComma Then
s = Replace(s, ",", "")
Else
s = Replace(s, ".", "")
s = Replace(s, ",", ".")
End If
ElseIf pComma > 0 Then
s = Replace(s, ",", ".")
End If

CleanDouble = val(s)
Exit Function

Invalid:
CleanDouble = 0
End Function

Public Function MinD(ByVal a As Double, ByVal b As Double) As Double
If a < b Then
MinD = a
Else
MinD = b
End If
End Function

Public Function ParseSapDate(ByVal v As Variant) As Date
On Error GoTo Invalid

If IsDate(v) And Not (VarType(v) = vbString) Then
ParseSapDate = CDate(v)
Exit Function
End If

Dim s As String
s = Trim$(CStr(v & ""))
If s = "" Then GoTo Invalid
If InStr(s, " ") > 0 Then s = Split(s, " ")(0)

Dim SEP As String
If InStr(s, ".") > 0 Then
SEP = "."
ElseIf InStr(s, "/") > 0 Then
SEP = "/"
ElseIf InStr(s, "-") > 0 Then
SEP = "-"
Else
If IsDate(s) Then
ParseSapDate = CDate(s)
Exit Function
End If
GoTo Invalid
End If

Dim p() As String
p = Split(s, SEP)
If UBound(p) <> 2 Then GoTo Invalid

Dim d As Integer, m As Integer, y As Integer
If Len(Trim$(p(0))) = 4 Then
y = CInt(p(0)): m = CInt(p(1)): d = CInt(p(2))
Else
d = CInt(p(0)): m = CInt(p(1)): y = CInt(p(2))
End If

Dim parsedDate As Date
parsedDate = DateSerial(y, m, d)
If Year(parsedDate) <> y Or Month(parsedDate) <> m Or Day(parsedDate) <> d Then GoTo Invalid
ParseSapDate = parsedDate
Exit Function

Invalid:
ParseSapDate = 0
End Function

Public Function ParseSapTime(ByVal v As Variant) As Double
On Error GoTo Invalid

If IsDate(v) And Not (VarType(v) = vbString) Then
ParseSapTime = CDbl(CDate(v)) - Int(CDbl(CDate(v)))
Exit Function
End If

Dim s As String
s = Trim$(CStr(v & ""))
If s = "" Then GoTo Invalid

If InStr(s, ":") > 0 Then
Dim p() As String
p = Split(s, ":")
If UBound(p) < 1 Or UBound(p) > 2 Then GoTo Invalid

Dim h As Integer, mi As Integer, se As Integer
h = CInt(p(0))
mi = CInt(p(1))
If UBound(p) >= 2 Then
se = CInt(p(2))
Else
se = 0
End If

If h < 0 Or h > 23 Or mi < 0 Or mi > 59 Or se < 0 Or se > 59 Then GoTo Invalid
ParseSapTime = CDbl(TimeSerial(h, mi, se))
Exit Function
End If

If IsNumeric(s) Then
Dim numVal As Double, fractionVal As Double
numVal = CDbl(s)
If numVal < 0 Then GoTo Invalid
fractionVal = numVal - Int(numVal)
If fractionVal = 0 And numVal <> 0 Then GoTo Invalid
ParseSapTime = fractionVal
Exit Function
End If

GoTo Invalid

Invalid:
ParseSapTime = -1 ' -1 = khong hop le (thoi gian that luon >= 0)
End Function

Public Function ParseHsdExpiry(ByVal batch As String) As Date
Dim b As String, bUpper As String
b = Trim$(batch)
bUpper = UCase$(b)

If Right$(bUpper, 4) <> "_HSD" Then
ParseHsdExpiry = 0
Exit Function
End If

Dim prefix As String
prefix = Left$(b, Len(b) - 4)
If Len(prefix) <> 6 Then
ParseHsdExpiry = 0
Exit Function
End If
If Not IsNumeric(prefix) Then
ParseHsdExpiry = 0
Exit Function
End If

On Error GoTo Invalid
Dim yy As Integer, mm As Integer, dd As Integer
yy = CInt(Left$(prefix, 2))
mm = CInt(Mid$(prefix, 3, 2))
dd = CInt(Right$(prefix, 2))
If mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Then GoTo Invalid

Dim result As Date
result = DateSerial(2000 + yy, mm, dd)
If Month(result) <> mm Or Day(result) <> dd Then GoTo Invalid

ParseHsdExpiry = result
Exit Function

Invalid:
ParseHsdExpiry = 0
End Function

Public Function ParseListedHsdExpiry(ByVal batch As String) As Date
Dim prefix As String
batch = Trim$(batch)
If Len(batch) < 6 Then Exit Function

prefix = Left$(batch, 6)
If Not IsNumeric(prefix) Then Exit Function

ParseListedHsdExpiry = ParseHsdExpiry(prefix & "_HSD")
End Function

Public Sub QuickSortGroups(idx() As Long, prod() As String, bucket() As Long, sortVal() As Double, grDate() As Double, _
grTime() As Double, origRow() As Long, ByVal lo As Long, ByVal hi As Long)
Dim i As Long, j As Long, tmp As Long
Dim pProd As String, pBucket As Long, pVal As Double, pDate As Double, pTime As Double, pRow As Long, pivotIdx As Long

If lo >= hi Then Exit Sub

i = lo: j = hi
pivotIdx = idx((lo + hi) \ 2)
pProd = prod(pivotIdx): pBucket = bucket(pivotIdx): pVal = sortVal(pivotIdx)
pDate = grDate(pivotIdx): pTime = grTime(pivotIdx): pRow = origRow(pivotIdx)

Do While i <= j
Do While CompareGroupKey(prod(idx(i)), bucket(idx(i)), sortVal(idx(i)), grDate(idx(i)), grTime(idx(i)), origRow(idx(i)), _
pProd, pBucket, pVal, pDate, pTime, pRow) < 0
i = i + 1
Loop
Do While CompareGroupKey(prod(idx(j)), bucket(idx(j)), sortVal(idx(j)), grDate(idx(j)), grTime(idx(j)), origRow(idx(j)), _
pProd, pBucket, pVal, pDate, pTime, pRow) > 0
j = j - 1
Loop
If i <= j Then
tmp = idx(i): idx(i) = idx(j): idx(j) = tmp
i = i + 1
j = j - 1
End If
Loop

If lo < j Then QuickSortGroups idx, prod, bucket, sortVal, grDate, grTime, origRow, lo, j
If i < hi Then QuickSortGroups idx, prod, bucket, sortVal, grDate, grTime, origRow, i, hi
End Sub

Private Function CompareGroupKey(ByVal prod1 As String, ByVal b1 As Long, ByVal v1 As Double, ByVal d1 As Double, ByVal t1 As Double, ByVal row1 As Long, _
ByVal prod2 As String, ByVal b2 As Long, ByVal v2 As Double, ByVal d2 As Double, ByVal t2 As Double, ByVal row2 As Long) As Integer
If prod1 <> prod2 Then
CompareGroupKey = IIf(prod1 < prod2, -1, 1)
ElseIf b1 <> b2 Then
CompareGroupKey = IIf(b1 < b2, -1, 1)
ElseIf v1 <> v2 Then
CompareGroupKey = IIf(v1 < v2, -1, 1)
ElseIf d1 <> d2 Then
CompareGroupKey = IIf(d1 < d2, -1, 1)
ElseIf t1 <> t2 Then
CompareGroupKey = IIf(t1 < t2, -1, 1)
ElseIf row1 <> row2 Then
CompareGroupKey = IIf(row1 < row2, -1, 1)
Else
CompareGroupKey = 0
End If
End Function

Public Sub ResetBinConfigCache()
mBinLoaded = False
mBinCount = 0
Erase mBinPatterns
Erase mBinZones
End Sub

Private Sub LoadBinConfigCache()
    mBinLoaded = True
    mBinCount = 0

    Dim wsConfig As Worksheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("DANH_MUC")
    On Error GoTo 0
    If wsConfig Is Nothing Then Exit Sub

    ' DANH_MUC: I=Bin Pattern, J=Zone; header row 4, data row 5+
    Dim lastRow As Long
    lastRow = wsConfig.Cells(wsConfig.rows.count, 9).End(xlUp).row
    If lastRow < 5 Then Exit Sub

    Dim raw As Variant
    raw = wsConfig.Range(wsConfig.Cells(5, 9), wsConfig.Cells(lastRow, 10)).Value2

    ReDim mBinPatterns(1 To lastRow - 4)
    ReDim mBinZones(1 To lastRow - 4)

    Dim i As Long, pat As String, zn As String
    For i = 1 To UBound(raw, 1)
        pat = UCase$(Trim$(CStr(raw(i, 1) & "")))
        zn = Trim$(CStr(raw(i, 2) & ""))
        If pat <> "" And zn <> "" Then
            mBinCount = mBinCount + 1
            mBinPatterns(mBinCount) = pat
            mBinZones(mBinCount) = zn
        End If
    Next i
End Sub

Public Function ValidateBinConfig() As Boolean
    Dim wsConfig As Worksheet
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("DANH_MUC")
    On Error GoTo 0

    If wsConfig Is Nothing Then
        ModuleLog.WriteLog "CANH BAO", "Khong co sheet DANH_MUC; tat ca Source Bin se vao Zone KHAC"
        ValidateBinConfig = True
        Exit Function
    End If

    Dim lastRow As Long, i As Long, activeRuleCount As Long
    Dim pattern As String, zone As String, dummyMatch As Boolean
    Dim hasError As Boolean
    lastRow = wsConfig.Cells(wsConfig.rows.count, 9).End(xlUp).row

    For i = 5 To lastRow
        pattern = UCase$(CleanStr(wsConfig.Cells(i, 9).Value))
        zone = UCase$(CleanStr(wsConfig.Cells(i, 10).Value))

        If pattern <> "" Or zone <> "" Then
            activeRuleCount = activeRuleCount + 1
            If pattern = "" Or zone = "" Then
                ModuleLog.WriteLog "LOI", "DANH_MUC BIN/ZONE phai co du ca Bin Pattern va Zone", "DANH_MUC dong " & i
                hasError = True
            ElseIf InStr(1, zone, "|", vbBinaryCompare) > 0 Then
                ModuleLog.WriteLog "LOI", "Zone khong duoc chua ky tu |", "DANH_MUC dong " & i
                hasError = True
            Else
                On Error Resume Next
                Err.Clear
                dummyMatch = ("TEST_BIN" Like pattern)
                If Err.Number <> 0 Then
                    ModuleLog.WriteLog "LOI", "Bin Pattern khong hop le: " & pattern, "DANH_MUC dong " & i
                    hasError = True
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    Next i

    If activeRuleCount = 0 Then
        ModuleLog.WriteLog "CANH BAO", "DANH_MUC chua co quy tac BIN/ZONE; tat ca Source Bin se vao Zone KHAC"
    End If

    ValidateBinConfig = Not hasError
End Function

Public Function GetBinArea(ByVal binName As String) As String
Dim normalizedBin As String
normalizedBin = UCase$(Trim$(binName))

If normalizedBin Like "APFL-P*" Then
GetBinArea = "P"
ElseIf normalizedBin Like "APFL-ZF*" Then
GetBinArea = "ZF"
Else
GetBinArea = "SPECIAL"
End If
End Function

Public Function GetKhuVucBinDynamic(ByVal binName As String) As String
Const DEFAULT_ZONE As String = "KHAC"

If Not mBinLoaded Then LoadBinConfigCache

Dim normalizedBin As String
normalizedBin = UCase$(Trim$(binName))
If normalizedBin = "" Then
GetKhuVucBinDynamic = DEFAULT_ZONE
Exit Function
End If

Dim i As Long
For i = 1 To mBinCount
If normalizedBin Like mBinPatterns(i) Then
GetKhuVucBinDynamic = mBinZones(i)
Exit Function
End If
Next i

GetKhuVucBinDynamic = DEFAULT_ZONE
End Function

Public Function BuildTramZoneText(ByVal tramText As String, ByVal zoneText As String) As String
tramText = Trim$(tramText)
zoneText = UCase$(Trim$(zoneText))
If zoneText = "" Then zoneText = "KHAC"

If tramText = "" Then
BuildTramZoneText = zoneText
Else
BuildTramZoneText = tramText & " - " & zoneText
End If
End Function

Public Function TramTLKey(ByVal tramText As String) As String
Dim re As Object, tlText As String
tramText = Trim$(tramText)
If tramText = "" Then Exit Function
Set re = CreateObject("VBScript.RegExp")
re.Global = False
re.IgnoreCase = True
re.pattern = "^KV[ ]*-?[ ]*[0-9]+([/,-][0-9]+)*(SV)?[ ]*[/\-_:]*[ ]*"
tlText = Trim$(re.Replace(tramText, ""))
If tlText = "" Then tlText = tramText
TramTLKey = UCase$(tlText)
End Function

Public Sub ResetToolConfigCache()
mToolConfigLoaded = False
Set mToolConfigCache = Nothing
End Sub

Private Sub LoadToolConfigCache()
Dim ws As Worksheet, lastRow As Long, data As Variant, i As Long
Dim keyText As String, valueText As String

Set mToolConfigCache = CreateObject("Scripting.Dictionary")
mToolConfigCache.CompareMode = vbTextCompare

On Error Resume Next
Set ws = ThisWorkbook.Worksheets("SYSTEM")
On Error GoTo 0
If ws Is Nothing Then
mToolConfigLoaded = True
Exit Sub
End If

lastRow = ws.Cells(ws.rows.count, 17).End(xlUp).row
If lastRow >= 3 Then
data = ws.Range("Q3:R" & CStr(lastRow)).Value2
If IsArray(data) Then
For i = 1 To UBound(data, 1)
keyText = UCase$(Trim$(CStr(data(i, 1) & "")))
valueText = Trim$(CStr(data(i, 2) & ""))
If keyText <> "" Then mToolConfigCache(keyText) = valueText
Next i
End If
End If
mToolConfigLoaded = True
End Sub

Public Function ToolConfig(ByVal keyName As String, ByVal defaultValue As String) As String
If Not mToolConfigLoaded Then LoadToolConfigCache
keyName = UCase$(Trim$(keyName))
If Not mToolConfigCache Is Nothing Then
If mToolConfigCache.Exists(keyName) Then
If Trim$(CStr(mToolConfigCache(keyName))) <> "" Then
ToolConfig = Trim$(CStr(mToolConfigCache(keyName)))
Exit Function
End If
End If
End If
ToolConfig = defaultValue
End Function

Public Function ToolConfigBool(ByVal keyName As String, ByVal defaultValue As Boolean) As Boolean
Dim raw As String
raw = UCase$(Trim$(ToolConfig(keyName, IIf(defaultValue, "TRUE", "FALSE"))))
Select Case raw
Case "TRUE", "YES", "Y", "1", "ON"
ToolConfigBool = True
Case "FALSE", "NO", "N", "0", "OFF"
ToolConfigBool = False
Case Else
ToolConfigBool = defaultValue
End Select
End Function

Public Function ToolConfigLong(ByVal keyName As String, ByVal defaultValue As Long, _
Optional ByVal minValue As Long = -2147483647, _
Optional ByVal maxValue As Long = 2147483647) As Long
Dim raw As String, resultValue As Long
raw = ToolConfig(keyName, CStr(defaultValue))
If IsNumeric(raw) Then resultValue = CLng(val(raw)) Else resultValue = defaultValue
If resultValue < minValue Then resultValue = minValue
If resultValue > maxValue Then resultValue = maxValue
ToolConfigLong = resultValue
End Function

Public Function ToolConfigDouble(ByVal keyName As String, ByVal defaultValue As Double, _
Optional ByVal minValue As Double = -1E+100, _
Optional ByVal maxValue As Double = 1E+100) As Double
Dim raw As String, resultValue As Double
raw = ToolConfig(keyName, CStr(defaultValue))
If IsNumeric(raw) Then resultValue = CDbl(raw) Else resultValue = defaultValue
If resultValue < minValue Then resultValue = minValue
If resultValue > maxValue Then resultValue = maxValue
ToolConfigDouble = resultValue
End Function

Public Function ValidateToolConfig(ByRef errorOut As String) As Boolean
Dim maxProducts As Long, minProducts As Long, maxRows As Long
Dim scopeText As String, v As Double
Dim defaultPlant As String, defaultSLoc As String, defaultSLocTransfer As String
errorOut = ""

maxProducts = ToolConfigLong("MAX_PRODUCTS_PER_OB", 4)
minProducts = ToolConfigLong("MIN_TARGET_PRODUCTS_PER_OB", 3)
maxRows = ToolConfigLong("MAX_ROWS_PER_OB", 20)
scopeText = UCase$(Trim$(ToolConfig("ONE_CLICK_POST_SCOPE", "ALL")))
defaultPlant = Trim$(ToolConfig("DEFAULT_PLANT", "3093"))
defaultSLoc = Trim$(ToolConfig("DEFAULT_SLOC", "1005"))
defaultSLocTransfer = Trim$(ToolConfig("DEFAULT_SLOC_TRANSFER", "3001"))

If maxProducts < 1 Or maxProducts > 20 Then errorOut = errorOut & "MAX_PRODUCTS_PER_OB phai 1..20." & vbCrLf
If minProducts < 1 Or minProducts > maxProducts Then errorOut = errorOut & "MIN_TARGET_PRODUCTS_PER_OB phai 1..MAX_PRODUCTS_PER_OB." & vbCrLf
If maxRows < 1 Or maxRows > 50 Then errorOut = errorOut & "MAX_ROWS_PER_OB phai 1..50." & vbCrLf
If scopeText <> "ALL" And scopeText <> "FIRST" And scopeText <> "BUILD_ONLY" Then _
errorOut = errorOut & "ONE_CLICK_POST_SCOPE chi nhan ALL, FIRST hoac BUILD_ONLY." & vbCrLf

If defaultPlant = "" Then errorOut = errorOut & "DEFAULT_PLANT khong duoc trong." & vbCrLf
If defaultSLoc = "" Then errorOut = errorOut & "DEFAULT_SLOC khong duoc trong." & vbCrLf
If defaultSLocTransfer = "" Then errorOut = errorOut & "DEFAULT_SLOC_TRANSFER khong duoc trong." & vbCrLf
If defaultSLoc <> "" And defaultSLoc = defaultSLocTransfer Then _
errorOut = errorOut & "DEFAULT_SLOC va DEFAULT_SLOC_TRANSFER khong duoc giong nhau." & vbCrLf

If Not IsValidBooleanConfig("ONE_CLICK_CONFIRM_BEFORE_SAP") Then _
errorOut = errorOut & "ONE_CLICK_CONFIRM_BEFORE_SAP phai TRUE/FALSE." & vbCrLf
If Not IsValidBooleanConfig("SAP_AUTO_ACCEPT_KNOWN_POPUPS") Then _
errorOut = errorOut & "SAP_AUTO_ACCEPT_KNOWN_POPUPS phai TRUE/FALSE." & vbCrLf
If Not IsValidBooleanConfig("SAP_STOP_ON_UNKNOWN_POPUP") Then _
errorOut = errorOut & "SAP_STOP_ON_UNKNOWN_POPUP phai TRUE/FALSE." & vbCrLf
If Not IsValidBooleanConfig("GROUP_BY_TL") Then _
errorOut = errorOut & "GROUP_BY_TL phai TRUE/FALSE." & vbCrLf

v = ToolConfigDouble("MIGO_BUSY_TIMEOUT_SEC", 60#)
If v < 10# Or v > 300# Then errorOut = errorOut & "MIGO_BUSY_TIMEOUT_SEC phai 10..300." & vbCrLf
v = ToolConfigDouble("MIGO_DELIVERY_GRACE_SEC", 5#)
If v < 0# Or v > 30# Then errorOut = errorOut & "MIGO_DELIVERY_GRACE_SEC phai 0..30." & vbCrLf
v = ToolConfigDouble("MIGO_OPEN_TIMEOUT_SEC", 20#)
If v < 5# Or v > 120# Then errorOut = errorOut & "MIGO_OPEN_TIMEOUT_SEC phai 5..120." & vbCrLf
v = ToolConfigLong("MIGO_POST_POLL_MS", 80)
If v < 30 Or v > 1000 Then errorOut = errorOut & "MIGO_POST_POLL_MS phai 30..1000." & vbCrLf
v = ToolConfigLong("MIGO_INPUT_IDLE_STABLE_SAMPLES", 1)
If v < 1 Or v > 10 Then errorOut = errorOut & "MIGO_INPUT_IDLE_STABLE_SAMPLES phai 1..10." & vbCrLf
v = ToolConfigDouble("MIGO_INPUT_IDLE_TIMEOUT_SEC", 4#)
If v < 1# Or v > 30# Then errorOut = errorOut & "MIGO_INPUT_IDLE_TIMEOUT_SEC phai 1..30." & vbCrLf
v = ToolConfigLong("MIGO_UI_POLL_MS", 100)
If v < 30 Or v > 1000 Then errorOut = errorOut & "MIGO_UI_POLL_MS phai 30..1000." & vbCrLf
v = ToolConfigLong("MIGO_INPUT_IDLE_POLL_MS", 80)
If v < 20 Or v > 1000 Then errorOut = errorOut & "MIGO_INPUT_IDLE_POLL_MS phai 20..1000." & vbCrLf
v = ToolConfigLong("MIGO_MAX_OB_PER_RUN", 0)
If v < 0 Or v > 1000 Then errorOut = errorOut & "MIGO_MAX_OB_PER_RUN phai 0..1000." & vbCrLf

ValidateUnmergedDataArea "ORDER", 2, 3, errorOut
ValidateUnmergedDataArea modSourceStockRepo.SSR_BackingName(), 2, modSourceStockRepo.SSR_COLS, errorOut   ' V2.8 Phase2: name via repo (SHEET mode = POSTING/17)
ValidateUnmergedDataArea "RESULT", 2, 10, errorOut
ValidateUnmergedDataArea "OB_GROUPS", 2, 17, errorOut
ValidateUnmergedDataArea "LOG", 2, 6, errorOut

ValidateToolConfig = (errorOut = "")
End Function

Private Sub ValidateUnmergedDataArea(ByVal sheetName As String, ByVal firstDataRow As Long, _
ByVal maxCol As Long, ByRef errorOut As String)
Dim ws As Worksheet, c As Long, candidate As Long, lastRow As Long
Dim rng As Range, mergedState As Variant
On Error Resume Next
Set ws = ThisWorkbook.Worksheets(sheetName)
On Error GoTo Failed
If ws Is Nothing Then
    errorOut = errorOut & "Thieu sheet " & sheetName & "." & vbCrLf
    Exit Sub
End If
lastRow = firstDataRow
For c = 1 To maxCol
    candidate = ws.Cells(ws.rows.count, c).End(xlUp).row
    If candidate > lastRow Then lastRow = candidate
Next c
Set rng = ws.Range(ws.Cells(firstDataRow, 1), ws.Cells(lastRow, maxCol))
mergedState = rng.MergeCells
If IsNull(mergedState) Or CBool(mergedState) Then
    errorOut = errorOut & sheetName & " co o gop trong vung du lieu " & _
        rng.Address(False, False) & "; hay bo Merge truoc khi chay." & vbCrLf
End If
Exit Sub
Failed:
errorOut = errorOut & "Khong kiem tra duoc layout sheet " & sheetName & ": " & _
    Err.Description & vbCrLf
End Sub

Public Function SaveResultAsNew(ByRef outPath As String, ByRef msg As String) As Boolean
Dim n As Long, d As String, fileAttr As Long
Dim oldEvents As Boolean, oldAlerts As Boolean
outPath = "": msg = "": SaveResultAsNew = False

If ThisWorkbook.Path = "" Then
msg = "Bo qua tu dong luu: file chua co duong dan. Du lieu van con trong workbook dang mo."
Exit Function
End If
If LCase$(Right$(ThisWorkbook.Name, 5)) <> ".xlsm" Then
msg = "Bo qua tu dong luu: phai dung file XLSM, khong dung XLTM."
Exit Function
End If
If ThisWorkbook.ReadOnly Then
msg = "Bo qua tu dong luu: file dang mo Read-only. Hay giai nen file ra thu muc co quyen ghi roi mo lai."
Exit Function
End If

On Error Resume Next
fileAttr = GetAttr(ThisWorkbook.FullName)
n = Err.Number: d = Err.Description
Err.Clear
On Error GoTo 0
If n = 0 Then
If (fileAttr And vbReadOnly) <> 0 Then
msg = "Bo qua tu dong luu: file tren dia dang co thuoc tinh Read-only."
Exit Function
End If
End If

oldEvents = Application.EnableEvents
oldAlerts = Application.DisplayAlerts
On Error GoTo f
Application.EnableEvents = False
Application.DisplayAlerts = False
DoEvents
ThisWorkbook.Save
Application.EnableEvents = oldEvents
Application.DisplayAlerts = oldAlerts
If Not ThisWorkbook.saved Then Err.Raise vbObjectError + 2672, , "Excel chua xac nhan workbook da duoc luu."
If Dir$(ThisWorkbook.FullName) = "" Then Err.Raise vbObjectError + 2673, , "Khong tim thay file sau khi Save."
outPath = ThisWorkbook.FullName
SaveResultAsNew = True
Exit Function
f:
n = Err.Number: d = Err.Description
On Error Resume Next
Application.EnableEvents = oldEvents
Application.DisplayAlerts = oldAlerts
On Error GoTo 0
msg = "Khong tu dong luu duoc file hien tai (Loi " & CStr(n) & ": " & d & "). " & _
"Tool KHONG dung; du lieu van con trong workbook dang mo. Hay bam Ctrl+S sau khi tool ket thuc."
End Function
Private Function IsValidBooleanConfig(ByVal keyName As String) As Boolean
Dim raw As String
raw = UCase$(Trim$(ToolConfig(keyName, "")))
IsValidBooleanConfig = (raw = "TRUE" Or raw = "FALSE" Or raw = "YES" Or raw = "NO" Or _
raw = "Y" Or raw = "N" Or raw = "1" Or raw = "0" Or _
raw = "ON" Or raw = "OFF")
End Function



