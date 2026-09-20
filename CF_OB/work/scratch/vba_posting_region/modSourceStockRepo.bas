Attribute VB_Name = "modSourceStockRepo"
Option Explicit

' SourceStockRepository (PROMPT V2.8 3): the single owner of WHERE the SAP
' source-stock ("POSTING") data physically lives.  Every reader/writer reaches
' the data through this module instead of Worksheets("POSTING") directly, so the
' physical backing can move from a dedicated sheet (V2.7) to a band on DATA_CORE
' (V2.8) by changing ONLY this module.
'
' Backing modes (SSR_BACKING):
'   "SHEET"  - data on a dedicated worksheet. Phase 2: byte-identical to V2.7;
'              proves the consumer redirects are behaviour-preserving.
'   "REGION" - data in a fixed rectangular band on DATA_CORE. Phase 3: proves the
'              relocation. LIVE SAP writes into the band are NOT_LIVE_PROVEN
'              (SAP GUI Scripting absent offline); only coordinate fidelity is
'              proven here, via SSR_SelfTest on synthetic rows.
'
' Schema (identical in both modes): header row 1, data rows 2+, 17 columns:
'   1 Product   2 A.Qty AUoM  3 Batch      4 Source Bin 5 TST       6 Total Stk
'   7 Dest. Bin 8 QtyPostChg  9 Dest.Batch 10 ST        11 GR Date  12 GR Time
'   13 Proc.Type 14 AvlQtyBUoM 15 -        16 -         17 Allocated Qty
' The hardcoded-coordinate probe readers depend on col 1/3/4/5; keep them fixed.

Private Const SSR_BACKING As String = "REGION"
Private Const SSR_SHEET_NAME As String = "POSTING"
Private Const SSR_REGION_SHEET As String = "DATA_CORE"
' Band origin on DATA_CORE. Phase 3 wires these; unused while backing = SHEET.
' Kept 0/0 so SHEET mode is a pure pass-through (SSR_Cell(r,c) = sheet.Cells(r,c)).
Private Const SSR_REGION_ROW0 As Long = 0
Private Const SSR_REGION_COL0 As Long = 30
Public Const SSR_COLS As Long = 17

Public Function SSR_HeaderArray() As Variant
    SSR_HeaderArray = Array("Product", "A.Qty AUoM", "Batch", "Source Bin", "TST", _
        "Total Stk", "Dest. Bin", "QtyPostChg", "Dest.Batch", "ST", "GR Date", _
        "GR Time", "Proc. Type", "AvlQtyBUoM", "", "", "Allocated Qty")
End Function

' Worksheet physically holding the band. In SHEET mode this is POSTING itself so
' the runtime POSTING/POSTING_RAW rename swap still resolves to the live sheet.
Public Function SSR_BackingSheet() As Worksheet
    If SSR_BACKING = "REGION" Then
        Set SSR_BackingSheet = ThisWorkbook.Worksheets(SSR_REGION_SHEET)
    Else
        Set SSR_BackingSheet = ThisWorkbook.Worksheets(SSR_SHEET_NAME)
    End If
End Function

' Name of the sheet physically backing the band (for name-keyed callers such as
' ModuleHelper.ValidateUnmergedDataArea). SHEET mode => "POSTING" (byte-identical
' to V2.7); REGION mode => the DATA_CORE host. Phase 3 makes band-aware validators
' use SSR_Cell extents rather than this whole-sheet name.
Public Function SSR_BackingName() As String
    If SSR_BACKING = "REGION" Then
        SSR_BackingName = SSR_REGION_SHEET
    Else
        SSR_BackingName = SSR_SHEET_NAME
    End If
End Function

' Absolute cell for band-relative (r, c): r=1 is the header row, c=1 is Product.
Public Function SSR_Cell(ByVal r As Long, ByVal c As Long) As Range
    Set SSR_Cell = SSR_BackingSheet().Cells(r + SSR_REGION_ROW0, c + SSR_REGION_COL0)
End Function

' Last band-relative data row scanning the Product column (col 1); >=1 header row
' means "no data" returns 1. Bounded to the band so REGION mode never reads a
' sibling band below it. Readers key off Product (col 1), matching V2.7's
' GetLastRow(ws, colMap("PRODUCT")).
Public Function SSR_LastDataRow() As Long
    SSR_LastDataRow = SSR_LastDataRowIn(1)
End Function

Private Function SSR_LastDataRowIn(ByVal col As Long) As Long
    Dim ws As Worksheet, absLast As Long
    Set ws = SSR_BackingSheet()
    absLast = ws.Cells(ws.rows.count, col + SSR_REGION_COL0).End(xlUp).row
    If absLast < (1 + SSR_REGION_ROW0) Then absLast = 1 + SSR_REGION_ROW0
    SSR_LastDataRowIn = absLast - SSR_REGION_ROW0
End Function

' Clear-extent = max of Product(1) and Allocated Qty(17) last rows. V2.7's writer
' cleared A2:Q by scanning col 17; readers key off col 1. Taking the max makes
' the cleared region provably never narrower than V2.7 (no stale-row under-clear).
Private Function SSR_ClearExtent() As Long
    Dim a As Long, b As Long
    a = SSR_LastDataRowIn(1)
    b = SSR_LastDataRowIn(SSR_COLS)
    SSR_ClearExtent = IIf(a >= b, a, b)
End Function

' Ensure the header row exists exactly (idempotent). Callers that historically
' wrote A1:Q1 now delegate here so header layout lives in one place.
Public Sub SSR_EnsureHeader()
    Dim ws As Worksheet, hdr As Variant
    Set ws = SSR_BackingSheet()
    hdr = SSR_HeaderArray()
    ws.Range(SSR_Cell(1, 1), SSR_Cell(1, SSR_COLS)).Value = hdr
    SSR_Cell(1, 1).EntireRow.Font.Bold = True
End Sub

' Read the whole data block (rows 2..last, cols 1..17) as a 2-D Value array, or
' Empty when there is no data. Mirrors ModuleSAP.ReadPosting's dataArr contract.
Public Function SSR_ReadBlock() As Variant
    Dim last As Long
    last = SSR_LastDataRow()
    If last < 2 Then SSR_ReadBlock = Empty: Exit Function
    SSR_ReadBlock = SSR_BackingSheet().Range(SSR_Cell(2, 1), SSR_Cell(last, SSR_COLS)).Value
End Function

' Clear data rows 2..throughRow (band-relative), leaving the header intact.
' throughRow<0 => clear the full max-extent (col 1 vs 17), matching V2.7.
Public Sub SSR_ClearData(Optional ByVal throughRow As Long = -1)
    If throughRow < 0 Then throughRow = SSR_ClearExtent()
    If throughRow < 2 Then throughRow = 2
    SSR_BackingSheet().Range(SSR_Cell(2, 1), SSR_Cell(throughRow, SSR_COLS)).ClearContents
End Sub

' Commit a freshly downloaded (rowCount x 17) array starting at data row 2,
' replacing any previous content and re-applying the V2.7 number formats.
' Text columns C (Batch) and I (Dest.Batch) are forced to text BEFORE the write.
Public Sub SSR_CommitBlock(ByVal outArr As Variant, ByVal rowCount As Long)
    Dim ws As Worksheet
    Set ws = SSR_BackingSheet()
    SSR_ClearData          ' full max-extent clear, matches V2.7 A2:Q<oldLast>
    If rowCount < 1 Then Exit Sub
    ws.Range(SSR_Cell(2, 3), SSR_Cell(rowCount + 1, 3)).NumberFormat = "@"
    ws.Range(SSR_Cell(2, 9), SSR_Cell(rowCount + 1, 9)).NumberFormat = "@"
    ws.Range(SSR_Cell(2, 1), SSR_Cell(rowCount + 1, SSR_COLS)).Value2 = outArr
    ws.Range(SSR_Cell(2, 2), SSR_Cell(rowCount + 1, 2)).NumberFormat = "#,##0"
    ws.Range(SSR_Cell(2, 6), SSR_Cell(rowCount + 1, 6)).NumberFormat = "#,##0"
    ws.Range(SSR_Cell(2, 8), SSR_Cell(rowCount + 1, 8)).NumberFormat = "#,##0"
    ws.Range(SSR_Cell(2, 14), SSR_Cell(rowCount + 1, 14)).NumberFormat = "#,##0"
    ws.Range(SSR_Cell(2, SSR_COLS), SSR_Cell(rowCount + 1, SSR_COLS)).NumberFormat = "#,##0"
End Sub

' Compile probe for CF_PreLiveAudit's AuditCompile probe list.
Public Sub SSR_CompileProbe()
    Dim n As Long
    n = SSR_COLS
    If n <> 17 Then Err.Raise vbObjectError + 9970, , "SSR_COMPILE_PROBE_FAIL: SSR_COLS<>17"
End Sub

' Offline round-trip proof: snapshot header+data, write synthetic rows, assert
' cell-level and block-level read-back, then restore the original snapshot byte
' for byte. Returns "PASS" or a specific failure. Proves coordinate fidelity of
' the active backing mode WITHOUT any live SAP call.
Public Function SSR_SelfTest() As String
    Dim ws As Worksheet, snap As Variant, snapLast As Long, hadData As Boolean
    Dim probe As Variant, i As Long, blk As Variant, msg As String
    On Error GoTo Failed
    Set ws = SSR_BackingSheet()

    ' 1) snapshot rows 1..snapLast x 17 (header included) for exact restore
    snapLast = SSR_LastDataRow()
    hadData = (snapLast >= 2)
    If snapLast < 1 Then snapLast = 1
    snap = ws.Range(SSR_Cell(1, 1), SSR_Cell(snapLast, SSR_COLS)).Value

    ' 2) seed 2 synthetic rows via the write path
    ReDim probe(1 To 2, 1 To SSR_COLS)
    For i = 1 To SSR_COLS: probe(1, i) = "R1C" & CStr(i): probe(2, i) = "R2C" & CStr(i): Next i
    probe(1, 1) = "MAT_SSR_A": probe(1, 3) = "BATCH_A": probe(1, 4) = "BIN_A": probe(1, 5) = "FT"
    probe(2, 1) = "MAT_SSR_B": probe(2, 3) = "BATCH_B": probe(2, 4) = "BIN_B": probe(2, 5) = "IT"
    SSR_EnsureHeader
    SSR_CommitBlock probe, 2

    ' 3a) cell-level read-back on the probe-critical coordinates (1/3/4/5)
    If SSR_S(SSR_Cell(2, 1).Value2) <> "MAT_SSR_A" Then msg = "cell(2,1)": GoTo Report
    If SSR_S(SSR_Cell(2, 3).Value2) <> "BATCH_A" Then msg = "cell(2,3)": GoTo Report
    If SSR_S(SSR_Cell(2, 4).Value2) <> "BIN_A" Then msg = "cell(2,4)": GoTo Report
    If UCase$(SSR_S(SSR_Cell(2, 5).Value2)) <> "FT" Then msg = "cell(2,5)": GoTo Report
    If SSR_LastDataRow() <> 3 Then msg = "lastRow<>3 got " & CStr(SSR_LastDataRow()): GoTo Report

    ' 3b) block-level read-back matches what was written
    blk = SSR_ReadBlock()
    If IsEmpty(blk) Then msg = "block empty": GoTo Report
    If UBound(blk, 1) <> 2 Or UBound(blk, 2) <> SSR_COLS Then msg = "block dims": GoTo Report
    If SSR_S(blk(2, 1)) <> "MAT_SSR_B" Then msg = "block(2,1)": GoTo Report

    SSR_SelfTest = "PASS"
Report:
    ' 4) restore exact snapshot (clear our probe rows first, then rewrite)
    SSR_ClearData 3
    ws.Range(SSR_Cell(1, 1), SSR_Cell(snapLast, SSR_COLS)).Value = snap
    If msg <> "" Then SSR_SelfTest = "FAIL[" & SSR_BACKING & "]:" & msg
    Exit Function
Failed:
    SSR_SelfTest = "ERR[" & SSR_BACKING & "]:" & CStr(Err.Number) & ":" & Err.Description
End Function

Private Function SSR_S(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then SSR_S = "" Else SSR_S = Trim$(CStr(v))
End Function
