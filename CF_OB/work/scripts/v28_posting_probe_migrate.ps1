# v28_posting_probe_migrate.ps1 — Migrate modSAPReadOnlyProbe POSTING access to SSR
$ErrorActionPreference = 'Stop'
$vbaStage = 'g:\SAP\CF_OB\work\scratch\vba_posting_region'

$probePath = Join-Path $vbaStage 'modSAPReadOnlyProbe.bas'
$probeContent = Get-Content $probePath -Raw -Encoding UTF8

# SAPProbe_FirstPostingSource: lines 254-266
# OLD: lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
#      ws.Cells(r, 1/3/4/5).Value2
# NEW: lastRow = modSourceStockRepo.SSR_LastDataRow()
#      modSourceStockRepo.SSR_Cell(r, N).Value2

$oldFirstPosting = @'
Private Function SAPProbe_FirstPostingSource() As String
    Dim ws As Worksheet, r As Long, lastRow As Long
    On Error GoTo Done
    Set ws = modSourceStockRepo.SSR_BackingSheet()   ' V2.8 Phase2 redirect (SHEET mode = POSTING)
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If UCase$(SAPProbe_Text(ws.Cells(r, 5).Value2)) = "FT" Then
            SAPProbe_FirstPostingSource = SAPProbe_SourceKey(SAPProbe_Text(ws.Cells(r, 1).Value2), _
                                                               SAPProbe_Text(ws.Cells(r, 4).Value2), _
                                                               SAPProbe_Text(ws.Cells(r, 3).Value2))
            Exit Function
        End If
    Next r
Done:
End Function
'@

$newFirstPosting = @'
Private Function SAPProbe_FirstPostingSource() As String
    Dim r As Long, lastRow As Long
    On Error GoTo Done
    lastRow = modSourceStockRepo.SSR_LastDataRow()  ' V2.8 Phase4: REGION mode SSR band accessor
    For r = 2 To lastRow
        If UCase$(SAPProbe_Text(modSourceStockRepo.SSR_Cell(r, 5).Value2)) = "FT" Then
            SAPProbe_FirstPostingSource = SAPProbe_SourceKey(SAPProbe_Text(modSourceStockRepo.SSR_Cell(r, 1).Value2), _
                                                               SAPProbe_Text(modSourceStockRepo.SSR_Cell(r, 4).Value2), _
                                                               SAPProbe_Text(modSourceStockRepo.SSR_Cell(r, 3).Value2))
            Exit Function
        End If
    Next r
Done:
End Function
'@

$probeContent = $probeContent.Replace($oldFirstPosting, $newFirstPosting)
Write-Host "MIGRATED_SAPProbe_FirstPostingSource"

# SAPProbe_GetFirstSource: lines 268-299 (second POSTING block at 284-294)
$oldGetFirst = @'
    Set ws = modSourceStockRepo.SSR_BackingSheet()   ' V2.8 Phase2 redirect (SHEET mode = POSTING)
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 2 To lastRow
        If UCase$(SAPProbe_Text(ws.Cells(r, 5).Value2)) = "FT" Then
            material = SAPProbe_Text(ws.Cells(r, 1).Value2)
            binText = SAPProbe_Text(ws.Cells(r, 4).Value2)
            batchText = SAPProbe_Text(ws.Cells(r, 3).Value2)
            SAPProbe_GetFirstSource = (material <> "" And binText <> "" And batchText <> "")
            If SAPProbe_GetFirstSource Then Exit Function
        End If
    Next r
    messageText = "STOP_POST_PROBE_SOURCE_REQUIRED: build PLAN_SOURCE_LEGS or load a FT row in POSTING first."
'@

$newGetFirst = @'
    lastRow = modSourceStockRepo.SSR_LastDataRow()  ' V2.8 Phase4: REGION mode SSR band accessor
    For r = 2 To lastRow
        If UCase$(SAPProbe_Text(modSourceStockRepo.SSR_Cell(r, 5).Value2)) = "FT" Then
            material = SAPProbe_Text(modSourceStockRepo.SSR_Cell(r, 1).Value2)
            binText = SAPProbe_Text(modSourceStockRepo.SSR_Cell(r, 4).Value2)
            batchText = SAPProbe_Text(modSourceStockRepo.SSR_Cell(r, 3).Value2)
            SAPProbe_GetFirstSource = (material <> "" And binText <> "" And batchText <> "")
            If SAPProbe_GetFirstSource Then Exit Function
        End If
    Next r
    messageText = "STOP_POST_PROBE_SOURCE_REQUIRED: build PLAN_SOURCE_LEGS or load a FT row in POSTING first."
'@

$probeContent = $probeContent.Replace($oldGetFirst, $newGetFirst)
Write-Host "MIGRATED_SAPProbe_GetFirstSource"

Set-Content $probePath -Value $probeContent -Encoding UTF8 -NoNewline
Write-Host "PROBE_MIGRATION_COMPLETE"
