# v28_posting_migration_stage.ps1 — Stage POSTING→REGION consumer migrations
$ErrorActionPreference = 'Stop'
$vbaBase = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$vbaStage = 'g:\SAP\CF_OB\work\scratch\vba_posting_region'

if (Test-Path $vbaStage) { Remove-Item $vbaStage -Recurse -Force }
New-Item -ItemType Directory -Path $vbaStage | Out-Null
Copy-Item "$vbaBase\*.bas" $vbaStage
Copy-Item "$vbaBase\*.cls" $vbaStage
Write-Host "STAGE_CREATED"

# ═══════════════════════════════════════════════════════════════════════════════
# 1. modSourceStockRepo: Flip to REGION mode
# ═══════════════════════════════════════════════════════════════════════════════
$ssrPath = Join-Path $vbaStage 'modSourceStockRepo.bas'
$ssrContent = Get-Content $ssrPath -Raw -Encoding UTF8

$ssrContent = $ssrContent -replace 'Private Const SSR_BACKING As String = "SHEET"', 'Private Const SSR_BACKING As String = "REGION"'
$ssrContent = $ssrContent -replace 'Private Const SSR_REGION_COL0 As Long = 0', 'Private Const SSR_REGION_COL0 As Long = 30'

Set-Content $ssrPath -Value $ssrContent -Encoding UTF8 -NoNewline
Write-Host "SSR_FLIPPED_TO_REGION|col_offset=30"

# ═══════════════════════════════════════════════════════════════════════════════
# 2. ModuleSAP.ReadPosting: raw coords → SSR_ReadBlock
# ═══════════════════════════════════════════════════════════════════════════════
$sapPath = Join-Path $vbaStage 'ModuleSAP.bas'
$sapContent = Get-Content $sapPath -Raw -Encoding UTF8

# ReadPosting: lines 102-115
# OLD: lastRow = ModuleHelper.GetLastRow(ws, colMap("PRODUCT"))
#      dataArr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Value
# NEW: dataArr = modSourceStockRepo.SSR_ReadBlock()
$oldReadPosting = @'
Public Sub ReadPosting(ByRef dataArr As Variant, ByRef colMap As Object, ByRef missingRequired As Collection)
    Dim ws As Worksheet
    Set ws = modSourceStockRepo.SSR_BackingSheet()   ' V2.8 Phase2: source-stock repo owns WHERE (SHEET mode = POSTING, byte-identical)
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
'@

$newReadPosting = @'
Public Sub ReadPosting(ByRef dataArr As Variant, ByRef colMap As Object, ByRef missingRequired As Collection)
    Dim ws As Worksheet
    Set ws = modSourceStockRepo.SSR_BackingSheet()   ' V2.8 Phase4: REGION mode, band on DATA_CORE
    Set colMap = CreateObject("Scripting.Dictionary")
    Set missingRequired = New Collection
    MapPostingColumns ws, colMap, missingRequired
    dataArr = modSourceStockRepo.SSR_ReadBlock()     ' V2.8 Phase4: SSR band accessor replaces raw coords
End Sub
'@

$sapContent = $sapContent.Replace($oldReadPosting, $newReadPosting)
Write-Host "MIGRATED_ReadPosting"

# ═══════════════════════════════════════════════════════════════════════════════
# 3. ModuleSAP.DownloadPostingForOrder: raw Range("A2:Q") → SSR_CommitBlock
# ═══════════════════════════════════════════════════════════════════════════════
# Lines 273-297: entire commit block
$oldDownload = @'
Set ws = modSourceStockRepo.SSR_BackingSheet()   ' V2.8 Phase2: commit target via repo (SHEET mode = POSTING, byte-identical)
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
'@

$newDownload = @'
' V2.8 Phase4: REGION mode commit via SSR_CommitBlock (band on DATA_CORE)
If rows.count = 0 Then
    modSourceStockRepo.SSR_ClearData
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
modSourceStockRepo.SSR_CommitBlock outArr, rows.count
'@

$sapContent = $sapContent.Replace($oldDownload, $newDownload)
Write-Host "MIGRATED_DownloadPostingForOrder"

Set-Content $sapPath -Value $sapContent -Encoding UTF8 -NoNewline

Write-Host "ALL_CONSUMERS_MIGRATED"
