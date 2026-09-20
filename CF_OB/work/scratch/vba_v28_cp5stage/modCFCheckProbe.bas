Attribute VB_Name = "modCFCheckProbe"
Option Explicit

' ============================================================================
' modCFCheckProbe  (SCRATCH -- NEVER shipped into V2.8)
' CP5a golden/candidate equivalence probe for CF_CheckActualPosting.
' Seeds a deterministic multi-branch scenario into DATA_CORE + POSTING, runs
' CHECK headless (CF_CheckSilent=True), and returns a stable signature string
' of every write coordinate (DATA_CORE cols 15-19 + POSTING cols 7-23).
' The harness runs this on the seam-only (raw-coord) build and on the SSR
' refactor build; identical signatures prove the refactor is byte-identical in
' SHEET mode. Runs in a THROWAWAY copy of CP4 that is never saved -> no restore.
' ============================================================================

Public Function CFCheck_Probe() As String
    Dim wsWH As Worksheet, wsPost As Worksheet
    Dim r As Long, c As Long, sb As String

    Set wsWH = ThisWorkbook.Worksheets("DATA_CORE")
    Set wsPost = modSourceStockRepo.SSR_BackingSheet()

    Probe_SeedWH wsWH
    Probe_SeedPost wsPost

    modCF_Posting.CF_CheckSilent = True
    On Error GoTo Fail
    modCF_Posting.CF_CheckActualPosting
    modCF_Posting.CF_CheckSilent = False
    On Error GoTo 0

    sb = "SIG_V1"
    sb = sb & vbLf & "WH(15-19):"
    For r = 5 To 12
        sb = sb & vbLf & "wr" & r & "|"
        For c = 15 To 19
            sb = sb & Probe_CellText(wsWH, r, c) & "|"
        Next c
    Next r
    sb = sb & vbLf & "POST(7-23):"
    For r = 2 To 9
        sb = sb & vbLf & "pr" & r & "|"
        For c = 7 To 23
            sb = sb & Probe_CellText(wsPost, r, c) & "|"
        Next c
    Next r

    CFCheck_Probe = sb
    Exit Function
Fail:
    modCF_Posting.CF_CheckSilent = False
    CFCheck_Probe = "PROBE_ERR:" & CStr(Err.Number) & ":" & Err.Description
End Function

' Read a cell as a stable text token. Numbers via CStr; blanks as empty.
Private Function Probe_CellText(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Dim v As Variant
    v = ws.Cells(r, c).Value2
    If IsError(v) Then
        Probe_CellText = "#ERR"
    ElseIf IsEmpty(v) Then
        Probe_CellText = ""
    Else
        Probe_CellText = CStr(v)
    End If
End Function

' Seed DATA_CORE rows 5-12 to exercise many CHECK branches deterministically.
' Cols read by CHECK: 1 ob,2 doc,4 whNo,5 wtNo,6 material,7 sysQty,8 sysBin,
' 9 sysBatch,10 status,12 actualQty,13 actualBin,21 area,22 actualBatch,
' 25 obState,26 wtState. Everything else irrelevant to the write signature.
Private Sub Probe_SeedWH(ByVal ws As Worksheet)
    ws.Range("A5:AF20000").ClearContents
    '        ob    doc    wh    wt    mat   sQty sBin  sBatch stat  aQty aBin   area  aBatch  obState              wtState
    SeedWH ws, 5,  "OB1", "D1", "W1", "T1", "M1", 100, "B1", "BT1", "", 100, "B1",  "",  "BT1", "OB_OPEN_OR_PARTIAL", "WT_ACTIVE"   ' READY, no post
    SeedWH ws, 6,  "OB1", "D1", "W1", "T2", "M2", 100, "B2", "BT2", "", 80,  "B2",  "",  "BT2", "OB_OPEN_OR_PARTIAL", "WT_ACTIVE"   ' CHANGE_QTY
    SeedWH ws, 7,  "OB1", "D1", "W1", "T3", "M3", 100, "B3", "BT3", "", 100, "BX",  "",  "BTX", "OB_OPEN_OR_PARTIAL", "WT_ACTIVE"   ' POST_CHANGE_BATCH_BIN
    SeedWH ws, 8,  "OB1", "D1", "W1", "T4", "M4", 150, "B4", "BT4", "", 150, "B4",  "",  "BT4", "OB_OPEN_OR_PARTIAL", "WT_ACTIVE"   ' NEED_MON
    SeedWH ws, 9,  "OB1", "D1", "W1", "T5", "M5", 300, "B5", "BT5", "", 300, "B5",  "",  "BT5", "OB_OPEN_OR_PARTIAL", "WT_ACTIVE"   ' NOT_ENOUGH_STOCK
    SeedWH ws, 10, "OB1", "D1", "W1", "T6", "M6", 100, "B6", "BT6b", "", 100, "B6", "",  "BTZ", "OB_OPEN_OR_PARTIAL", "WT_ACTIVE"   ' BATCH_NOT_FOUND (matCount>0)
    SeedWH ws, 11, "OB1", "D1", "W1", "T7", "M7", 100, "B7", "BT7", "", 100, "",   "P", "BT7", "OB_OPEN_OR_PARTIAL", "WT_ACTIVE"   ' MULTI_BIN (area P)
    SeedWH ws, 12, "OB1", "D1", "W1", "T8", "M8", 100, "B8", "BT8", "", 100, "B8", "",  "BT8", "OB_DONE",            "WT_ACTIVE"   ' DONE (WH-only)
End Sub

Private Sub SeedWH(ByVal ws As Worksheet, ByVal r As Long, _
                   ByVal ob As String, ByVal doc As String, ByVal wh As String, ByVal wt As String, _
                   ByVal mat As String, ByVal sysQty As Variant, ByVal sysBin As String, ByVal sysBatch As String, _
                   ByVal stat As String, ByVal aQty As Variant, ByVal aBin As String, _
                   ByVal area As String, ByVal aBatch As String, _
                   ByVal obState As String, ByVal wtState As String)
    ws.Cells(r, 1).Value2 = ob
    ws.Cells(r, 2).Value2 = doc
    ws.Cells(r, 4).Value2 = wh
    ws.Cells(r, 5).Value2 = wt
    ws.Cells(r, 6).Value2 = mat
    ws.Cells(r, 7).Value2 = sysQty
    ws.Cells(r, 8).Value2 = sysBin
    ws.Cells(r, 9).Value2 = sysBatch
    ws.Cells(r, 10).Value2 = stat
    ws.Cells(r, 12).Value2 = aQty
    ws.Cells(r, 13).Value2 = aBin
    ws.Cells(r, 21).Value2 = area
    ws.Cells(r, 22).Value2 = aBatch
    ws.Cells(r, 25).Value2 = obState
    ws.Cells(r, 26).Value2 = wtState
End Sub

' Seed POSTING (SHEET-mode backing) rows 2-9. Cols read by CHECK: 1 product,
' 2 available,3 batch,4 bin,6 totalStock. Two M7 rows -> MULTI_BIN.
Private Sub Probe_SeedPost(ByVal ws As Worksheet)
    ws.Range("A2:W20000").ClearContents
    '         row  product avail batch  bin         total
    SeedPost ws, 2, "M1", 200, "BT1", "B1",        500   ' READY
    SeedPost ws, 3, "M2", 200, "BT2", "B2",        500   ' CHANGE_QTY
    SeedPost ws, 4, "M3", 200, "BTX", "BX",        500   ' POST_CHANGE_BATCH_BIN
    SeedPost ws, 5, "M4", 100, "BT4", "B4",        200   ' NEED_MON (avail<need<=total)
    SeedPost ws, 6, "M5", 200, "BT5", "B5",        200   ' NOT_ENOUGH_STOCK (need>total)
    SeedPost ws, 7, "M6", 200, "BT6", "B6",        500   ' M6 present but batch BT6<>BTZ
    SeedPost ws, 8, "M7", 100, "BT7", "APFL-P01",  300   ' MULTI_BIN row A
    SeedPost ws, 9, "M7", 100, "BT7", "APFL-P02",  300   ' MULTI_BIN row B
End Sub

Private Sub SeedPost(ByVal ws As Worksheet, ByVal r As Long, _
                     ByVal product As String, ByVal avail As Variant, ByVal batch As String, _
                     ByVal bin As String, ByVal total As Variant)
    ws.Cells(r, 1).Value2 = product
    ws.Cells(r, 2).Value2 = avail
    ws.Cells(r, 3).Value2 = batch
    ws.Cells(r, 4).Value2 = bin
    ws.Cells(r, 6).Value2 = total
End Sub
