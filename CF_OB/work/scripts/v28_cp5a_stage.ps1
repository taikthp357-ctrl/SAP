# v28_cp5a_stage.ps1
# Produce two staged variants of modCF_Posting.bas from post-CP4 ground truth:
#   golden\ = SEAM ONLY (Public CF_CheckSilent + 3 MsgBox guards) -- raw coords.
#   ssr\    = golden + SCOPED coordinate refactor (wsPost.Cells/Range/LastDataRow
#             -> SSR band accessors; writer body ws.Cells(rowNo,-> SSR_Cell).
# ssr is derived FROM golden so the golden-vs-ssr diff is EXACTLY the coord swap.
# Fail-closed: every edit is count-gated; any mismatch throws and writes nothing.
# Pure ASCII / CRLF file -> cp1252 round-trip is byte-preserving.
$ErrorActionPreference = 'Stop'
$src   = 'g:\SAP\CF_OB\work\scratch\vba_v28\modCF_Posting.bas'
$gold  = 'g:\SAP\CF_OB\work\scratch\vba_v28_cp5stage\golden\modCF_Posting.bas'
$ssr   = 'g:\SAP\CF_OB\work\scratch\vba_v28_cp5stage\ssr\modCF_Posting.bas'
$outLog= 'g:\SAP\CF_OB\work\REPORTS\v28_cp5a_stage.log'
$enc   = [Text.Encoding]::GetEncoding(1252)
$nl    = "`r`n"
$sq    = [char]39
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Host $m }   # Write-Host, NOT Write-Output: never contaminate function return values
function J([string[]]$a){ [string]::Join($nl, $a) }
function Cnt([string]$hay,[string]$needle){ ([regex]::Matches($hay,[regex]::Escape($needle))).Count }
function ReplOnce([string]$hay,[string]$o,[string]$n,[string]$tag){
    $c = Cnt $hay $o
    if ($c -ne 1) { throw ("SEAM_ANCHOR_$tag expected 1 got $c") }
    Log ("  edit $tag : 1 site"); return $hay.Replace($o,$n)
}
function ReplN([string]$hay,[string]$o,[string]$n,[int]$expect,[string]$tag){
    $c = Cnt $hay $o
    if ($c -ne $expect) { throw ("COUNT_$tag expected $expect got $c") }
    Log ("  edit $tag : $c sites"); return $hay.Replace($o,$n)
}

$t0 = [System.IO.File]::ReadAllText($src, $enc)
Log ("SRC bytes=" + $t0.Length + " sha256=" + (Get-FileHash -Path $src -Algorithm SHA256).Hash)
# ==== LAYER 1: SEAM (golden) ==================================================
$g = $t0

# 1a) declare Public CF_CheckSilent after the CF_QUERY_ORDER const, before the Sub.
$declOld = J @('Private Const CF_QUERY_ORDER As String = "ORDER_QUERY"','','Public Sub CF_LoadPostingFromWH()')
$declNew = J @(
    'Private Const CF_QUERY_ORDER As String = "ORDER_QUERY"',
    '',
    ($sq + ' V2.8 CP5a: headless CHECK seam. Default False = unchanged (MsgBox shown).'),
    ($sq + ' Set True only by the scratch equivalence probe so CHECK runs without UI.'),
    'Public CF_CheckSilent As Boolean',
    '',
    'Public Sub CF_LoadPostingFromWH()'
)
$g = ReplOnce $g $declOld $declNew 'decl'

# 1b) guard the success/empty MsgBox block. Open guard before "If checked = 0",
#     close guard with an extra End If before "Exit Sub".
$g = ReplOnce $g ('    If checked = 0 Then') (J @('    If Not CF_CheckSilent Then','    If checked = 0 Then')) 'msg_open'
$closeOld = J @(
    '            vbInformation, "CHECK"',
    '    End If',
    '    Exit Sub'
)
$closeNew = J @(
    '            vbInformation, "CHECK"',
    '    End If',
    '    End If',
    '    Exit Sub'
)
$g = ReplOnce $g $closeOld $closeNew 'msg_close'

# 1c) guard the Failed-handler MsgBox.
$failOld = J @(
    '    MsgBox "CHECK failed" & vbCrLf & _',
    '        "Err " & CStr(Err.Number) & ": " & Err.Description, _',
    '        vbExclamation, "CHECK"',
    'End Sub'
)
$failNew = J @(
    '    If Not CF_CheckSilent Then',
    '    MsgBox "CHECK failed" & vbCrLf & _',
    '        "Err " & CStr(Err.Number) & ": " & Err.Description, _',
    '        vbExclamation, "CHECK"',
    '    End If',
    'End Sub'
)
$g = ReplOnce $g $failOld $failNew 'msg_fail'

# seam invariants: CF_CheckSilent now appears exactly 3x (1 decl + 2 guards).
$seamRefs = Cnt $g 'CF_CheckSilent'
if ($seamRefs -ne 3) { throw ("SEAM_CFCHECKSILENT_REFS expected 3 got $seamRefs") }
[System.IO.File]::WriteAllText($gold, $g, $enc)
Log ("GOLDEN written bytes=" + $g.Length + " CF_CheckSilent_refs=" + $seamRefs)
Log ("GOLDEN sha256=" + (Get-FileHash -Path $gold -Algorithm SHA256).Hash)
# ==== LAYER 2: SSR coordinate refactor (ssr = golden + coord swap) ============
$s = $g

# SSR-1: lastPost via band accessor (SHEET-mode identical to CF_LastDataRow).
$s = ReplOnce $s `
    'lastPost = CF_LastDataRow(wsPost, 1, 2)' `
    'lastPost = modSourceStockRepo.SSR_LastDataRow()' 'ssr_lastrow'

# SSR-2: clear the analysis band via SSR corners (SHEET-mode == "G2:W2000").
$s = ReplOnce $s `
    'wsPost.Range("G2:W2000").ClearContents' `
    'wsPost.Range(modSourceStockRepo.SSR_Cell(2, 7), modSourceStockRepo.SSR_Cell(2000, 23)).ClearContents' 'ssr_clear'

# SSR-3: the 10 CHECK-loop reads wsPost.Cells( -> SSR_Cell(.
$s = ReplN $s 'wsPost.Cells(' 'modSourceStockRepo.SSR_Cell(' 10 'ssr_reads'

# SSR-4: writer body only. Extract CF_WritePostingCheck (start anchor .. its End Sub),
#         swap its 25 ws.Cells(rowNo, -> SSR_Cell(rowNo,, splice back. The 6
#         CF_UpdatePostingResult sites (live MIGO writer) are OUTSIDE this window
#         and stay raw.
$startAnchor = 'Private Sub CF_WritePostingCheck(ByVal ws As Worksheet, ByVal rowNo As Long, _'
$si = $s.IndexOf($startAnchor)
if ($si -lt 0) { throw 'SSR_WRITER_START_NOT_FOUND' }
$endMarker = $nl + 'End Sub' + $nl
$ei = $s.IndexOf($endMarker, $si)
if ($ei -lt 0) { throw 'SSR_WRITER_END_NOT_FOUND' }
$ei = $ei + $endMarker.Length
$writer = $s.Substring($si, $ei - $si)
$writer = ReplN $writer 'ws.Cells(rowNo,' 'modSourceStockRepo.SSR_Cell(rowNo,' 25 'ssr_writer'
$s = $s.Substring(0, $si) + $writer + $s.Substring($ei)

# ---- SSR invariants (post-transform) ----
$rem_wsPostCells = Cnt $s 'wsPost.Cells('
$rem_lastRowRaw  = Cnt $s 'CF_LastDataRow(wsPost'
$rem_rangeRaw    = Cnt $s 'wsPost.Range("G2:W2000")'
$rem_wsRowNo     = Cnt $s 'ws.Cells(rowNo,'                # 6 survivors in CF_UpdatePostingResult
$ssrCellRefs     = Cnt $s 'modSourceStockRepo.SSR_Cell('   # 2 (clear) + 10 (reads) + 25 (writer) = 37
$ssrLastRefs     = Cnt $s 'modSourceStockRepo.SSR_LastDataRow()'
Log ("SSR_INVARIANTS wsPostCells=$rem_wsPostCells lastRowRaw=$rem_lastRowRaw rangeRaw=$rem_rangeRaw wsRowNoSurvivors=$rem_wsRowNo SSR_Cell=$ssrCellRefs SSR_LastDataRow=$ssrLastRefs")
if ($rem_wsPostCells -ne 0) { throw "SSR_LEFTOVER_wsPostCells=$rem_wsPostCells" }
if ($rem_lastRowRaw  -ne 0) { throw "SSR_LEFTOVER_lastRowRaw=$rem_lastRowRaw" }
if ($rem_rangeRaw    -ne 0) { throw "SSR_LEFTOVER_rangeRaw=$rem_rangeRaw" }
if ($rem_wsRowNo     -ne 6) { throw "SSR_wsRowNo_survivors expected 6 got $rem_wsRowNo" }
if ($ssrCellRefs     -ne 37){ throw "SSR_Cell refs expected 37 got $ssrCellRefs" }
if ($ssrLastRefs     -ne 1) { throw "SSR_LastDataRow refs expected 1 got $ssrLastRefs" }
# CF_CheckSilent seam must survive unchanged in ssr too.
$ssrSeam = Cnt $s 'CF_CheckSilent'
if ($ssrSeam -ne 3) { throw "SSR_SEAM refs expected 3 got $ssrSeam" }

[System.IO.File]::WriteAllText($ssr, $s, $enc)
Log ("SSR written bytes=" + $s.Length)
Log ("SSR sha256=" + (Get-FileHash -Path $ssr -Algorithm SHA256).Hash)

# ---- golden-vs-ssr diff surface (sanity: only coord lines differ) ----
$gLines = $g.Split([char]10); $sLines = $s.Split([char]10)
if ($gLines.Count -ne $sLines.Count) { throw ("LINECOUNT_DRIFT golden=" + $gLines.Count + " ssr=" + $sLines.Count) }
$diff = 0
for ($i=0; $i -lt $gLines.Count; $i++){ if ($gLines[$i] -ne $sLines[$i]) { $diff++ } }
Log ("GOLDEN_VS_SSR differing_lines=" + $diff + " (expect 2 clear/lastrow + 10 reads + 25 writer = 37)")
if ($diff -ne 37) { throw ("DIFF_LINES expected 37 got " + $diff) }

Set-Content -Path $outLog -Value (J $L) -Encoding UTF8
Log ('LOG_WRITTEN=' + $outLog)
Log 'STAGE_OK'


