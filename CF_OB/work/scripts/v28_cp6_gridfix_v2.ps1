# v28_cp6_gridfix_v2.ps1 — CP6: grid lazy-load fix + Vietnamese messages
# Exports clean base from CP5a; inserts helper + calls at ALL rowCount assignment sites.
$ErrorActionPreference = 'Stop'
$cp5a       = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP5a_CHECKBAND.xlsm'
$wip        = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP6_WIP.xlsm'
$vbaBase    = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$logPath    = 'g:\SAP\CF_OB\work\REPORTS\v28_cp6_gridfix.log'
$cleanExport = Join-Path $vbaBase '_cp6_clean_readonly.bas'
$enc        = [Text.Encoding]::GetEncoding(1252)
$preExcel   = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m) { [void]$L.Add($m); Write-Host $m }

if (-not (Test-Path $cp5a)) { throw "NO_CP5A=$cp5a" }

# ── PHASE 1: Export clean modCodexSAPReadOnly.bas from CP5a (ReadOnly) ──────
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wbRO = $xl.Workbooks.Open($cp5a, 0, $true)
    $found = $false
    foreach ($comp in $wbRO.VBProject.VBComponents) {
        if ($comp.Name -eq 'modCodexSAPReadOnly') {
            $comp.Export($cleanExport)
            $found = $true
            Log "EXPORTED_CLEAN_modCodexSAPReadOnly"
            break
        }
    }
    if (-not $found) { throw "COMP_NOT_FOUND=modCodexSAPReadOnly" }
    $wbRO.Close($false)
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# ── PHASE 2: Patch modCodexSAPReadOnly.bas ───────────────────────────────────
$roContent = [System.IO.File]::ReadAllText($cleanExport, $enc)

# 2a. Build helper as CRLF array to avoid here-string encoding issues
$helperLines = @(
    '',
    "' Grid lazy-load fix: force SAP GUI to render all rows before reading.",
    "' Issue #6/#2: weak machine + many rows causes unrendered cells to return blank.",
    'Private Sub Codex_EnsureGridRendered(ByVal grid As Object, ByVal ses As Object)',
    '    Dim rowCount As Long',
    '    On Error Resume Next',
    '    rowCount = CLng(grid.rowCount)',
    '    On Error GoTo 0',
    '    If rowCount > 20 Then',
    '        grid.SetCurrentCell rowCount - 1, 0',
    '        CWR_WaitIdle ses, 0.5',
    '        grid.SetCurrentCell 0, 0',
    '    End If',
    'End Sub',
    ''
)
$helper = $helperLines -join "`r`n"

# Insert after Option Explicit (CRLF-aware)
$roContent = $roContent -replace '(?m)^(Option Explicit)\r?\n', "`$1`r`n${helper}"
if ($roContent -notmatch 'Codex_EnsureGridRendered') { throw 'HELPER_INSERT_FAILED' }
Log 'INSERTED_HELPER_AFTER_OPTION_EXPLICIT'

# 2b. Insert Codex_EnsureGridRendered before every grid-rowCount read.
# Session var: scan backward from site to find "Set <gridVar> = <sesVar>.FindById".
# Skip lines inside the helper itself.
function Get-SesVar($lines, $upToIdx, $gridVar) {
    $pat = "Set\s+$([regex]::Escape($gridVar))\s*=\s*(\w+)\.FindById"
    for ($j = $upToIdx; $j -ge 0; $j--) {
        if ($lines[$j] -match $pat) { return $Matches[1] }
    }
    return 'ses'  # safe fallback
}

$lines = $roContent -split "`r`n"
$output = [System.Collections.ArrayList]::new()
$sitesFixed = 0
$inHelper   = $false

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]

    # Track whether we are inside the helper Sub (don't re-patch it)
    if ($line -match 'Private Sub Codex_EnsureGridRendered') { $inHelper = $true }
    if ($inHelper -and $line -match '^End Sub') { $inHelper = $false; [void]$output.Add($line); continue }

    # Pattern A: pre-stored  "X = CLng(Y.rowCount)" assignment
    if (-not $inHelper -and $line -match '^(\s+)\w+ = CLng\((\w+)\.rowCount\)') {
        $indent  = $Matches[1]
        $gridVar = $Matches[2]
        $sesVar  = Get-SesVar $lines $i $gridVar
        $prevLine = if ($output.Count -gt 0) { $output[$output.Count - 1] } else { '' }
        if ($prevLine -notmatch 'Codex_EnsureGridRendered') {
            [void]$output.Add("${indent}Codex_EnsureGridRendered ${gridVar}, ${sesVar}")
            $sitesFixed++
        }
    }
    # Pattern B: inline "For r = 0 To CLng(Y.rowCount) - 1" (sites 11-14)
    elseif (-not $inHelper -and $line -match '^(\s+)For \w+ = 0 To CLng\((\w+)\.rowCount\) - 1') {
        $indent  = $Matches[1]
        $gridVar = $Matches[2]
        $sesVar  = Get-SesVar $lines $i $gridVar
        $prevLine = if ($output.Count -gt 0) { $output[$output.Count - 1] } else { '' }
        if ($prevLine -notmatch 'Codex_EnsureGridRendered') {
            [void]$output.Add("${indent}Codex_EnsureGridRendered ${gridVar}, ${sesVar}")
            $sitesFixed++
        }
    }
    [void]$output.Add($line)
}
$roContent = $output -join "`r`n"
Log "INSERTED_CALLS_AT_${sitesFixed}_SITES"

# Verify no helper-internal call double-patched
$callCount = ([regex]::Matches($roContent, 'Codex_EnsureGridRendered')).Count
Log "TOTAL_CALL_REFS=$callCount"

# Write patched file to ground truth
[System.IO.File]::WriteAllText((Join-Path $vbaBase 'modCodexSAPReadOnly.bas'), $roContent, $enc)
Log 'WROTE_GROUND_TRUTH_modCodexSAPReadOnly.bas'

# ── PHASE 3: Create modCF_VietnameseMessages.bas ─────────────────────────────
$vnLines = @(
    'Attribute VB_Name = "modCF_VietnameseMessages"',
    'Option Explicit',
    '',
    "' Vietnamese message mapper for user-facing feedback.",
    "' Maps Result enum (English technical) + Detail context to Vietnamese string.",
    "' Used by Phase 1b worklist to display CHECK feedback in Vietnamese.",
    '',
    'Public Function CF_VN_GetMessage(ByVal result As String, ByVal detail As String) As String',
    '    Dim msg As String',
    '    result = UCase$(Trim$(result))',
    '    detail = Trim$(detail)',
    '    Select Case result',
    '        Case "PASS"',
    '            msg = "S" & ChrW(7855) & "n s" & ChrW(224) & "ng x" & ChrW(7917) & " l" & ChrW(253)',
    '        Case "BLOCKED_SOURCE_PLAN_INCOMPLETE"',
    '            If InStr(1, detail, "Multi-batch", vbTextCompare) > 0 Then',
    '                msg = "V" & ChrW(7883) & " tr" & ChrW(237) & " c" & ChrW(243) & " nhi" & ChrW(7873) & "u l" & ChrW(244) & ", c" & ChrW(7847) & "n ch" & ChrW(7885) & "n l" & ChrW(244) & " c" & ChrW(7909) & " th" & ChrW(7875)',
    '            ElseIf InStr(1, detail, "PICK_BIN_CANDIDATE_COUNT", vbTextCompare) > 0 Then',
    '                msg = "C" & ChrW(243) & " nhi" & ChrW(7873) & "u v" & ChrW(7883) & " tr" & ChrW(237) & " ngu" & ChrW(7891) & "n, c" & ChrW(7847) & "n ch" & ChrW(7885) & "n v" & ChrW(7883) & " tr" & ChrW(237) & " c" & ChrW(7909) & " th" & ChrW(7875)',
    '            Else',
    '                msg = "D" & ChrW(7919) & " li" & ChrW(7879) & "u ngu" & ChrW(7891) & "n ch" & ChrW(432) & "a " & ChrW(273) & ChrW(7847) & "y " & ChrW(273) & ChrW(7911)',
    '            End If',
    '        Case "BLOCKED_QTY_GROUP_BASELINE"',
    '            msg = "Nh" & ChrW(243) & "m s" & ChrW(7889) & " l" & ChrW(432) & ChrW(7907) & "ng ch" & ChrW(432) & "a x" & ChrW(225) & "c " & ChrW(273) & ChrW(7883) & "nh " & ChrW(273) & ChrW(432) & ChrW(7907) & "c baseline"',
    '        Case "BLOCKED_INCOMPLETE_EXECUTION"',
    '            msg = "H" & ChrW(224) & "nh " & ChrW(273) & ChrW(7897) & "ng tr" & ChrW(432) & ChrW(7899) & "c ch" & ChrW(432) & "a ho" & ChrW(224) & "n t" & ChrW(7845) & "t"',
    '        Case "BLOCKED_CREATE_OPEN"',
    '            msg = "Ch" & ChrW(7913) & "ng t" & ChrW(7915) & " PRDO " & ChrW(273) & "ang m" & ChrW(7903) & ", c" & ChrW(7847) & "n " & ChrW(273) & ChrW(243) & "ng tr" & ChrW(432) & ChrW(7899) & "c khi ti" & ChrW(7871) & "p t" & ChrW(7909) & "c"',
    '        Case "STOP"',
    '            If InStr(1, detail, "SOURCE_PLAN_INCOMPLETE", vbTextCompare) > 0 Then',
    '                msg = "D" & ChrW(7915) & "ng: d" & ChrW(7919) & " li" & ChrW(7879) & "u ngu" & ChrW(7891) & "n ch" & ChrW(432) & "a " & ChrW(273) & ChrW(7911)',
    '            ElseIf InStr(1, detail, "REGRESSION_GATE", vbTextCompare) > 0 Then',
    '                msg = "D" & ChrW(7915) & "ng: kh" & ChrW(244) & "ng th" & ChrW(7887) & "a " & ChrW(273) & "i" & ChrW(7873) & "u ki" & ChrW(7879) & "n an to" & ChrW(224) & "n regression"',
    '            Else',
    '                msg = "D" & ChrW(7915) & "ng: " & detail',
    '            End If',
    '        Case "READY_FOR_CF"',
    '            msg = "S" & ChrW(7855) & "n s" & ChrW(224) & "ng " & ChrW(273) & ChrW(7875) & " user x" & ChrW(225) & "c nh" & ChrW(7853) & "n WT"',
    '        Case "DONE_BY_USER_CF"',
    '            msg = ChrW(272) & ChrW(227) & " ho" & ChrW(224) & "n t" & ChrW(7845) & "t b" & ChrW(7903) & "i user CF"',
    '        Case "DONE"',
    '            msg = "Ho" & ChrW(224) & "n t" & ChrW(7845) & "t"',
    '        Case "PLAN_NEEDS_TAI_POSTING"',
    '            msg = "C" & ChrW(7847) & "n TAI POSTING " & ChrW(273) & ChrW(7875) & " ti" & ChrW(7871) & "p t" & ChrW(7909) & "c"',
    '        Case Else',
    '            msg = result',
    '            If detail <> "" Then msg = msg & " (" & Left$(detail, 100) & ")"',
    '    End Select',
    '    CF_VN_GetMessage = msg',
    'End Function',
    '',
    'Public Function CF_VN_GetStageLabel(ByVal stage As String) As String',
    '    stage = UCase$(Trim$(stage))',
    '    Select Case stage',
    '        Case "CHECK PREPARE":            CF_VN_GetStageLabel = "Chu" & ChrW(7849) & "n b" & ChrW(7883) & " ki" & ChrW(7875) & "m tra"',
    '        Case "CHECK DONE":               CF_VN_GetStageLabel = "Ki" & ChrW(7875) & "m tra ho" & ChrW(224) & "n t" & ChrW(7845) & "t"',
    '        Case "EXECUTION_CONTEXT_REUSE":  CF_VN_GetStageLabel = "T" & ChrW(225) & "i s" & ChrW(7917) & " d" & ChrW(7909) & "ng context"',
    '        Case "CANCEL_QUEUE":             CF_VN_GetStageLabel = "H" & ChrW(224) & "ng " & ChrW(273) & ChrW(7907) & "i cancel"',
    '        Case "CANCEL_GROUP":             CF_VN_GetStageLabel = "Cancel nh" & ChrW(243) & "m"',
    '        Case "CANCEL_BATCH_READBACK":    CF_VN_GetStageLabel = ChrW(272) & ChrW(7885) & "c l" & ChrW(7841) & "i sau cancel"',
    '        Case "LOCATION_GROUP":           CF_VN_GetStageLabel = "Thay " & ChrW(273) & ChrW(7893) & "i v" & ChrW(7883) & " tr" & ChrW(237)',
    '        Case "CREATE_GROUP":             CF_VN_GetStageLabel = "T" & ChrW(7841) & "o WT m" & ChrW(7899) & "i"',
    '        Case "FINAL_VERIFY":             CF_VN_GetStageLabel = "X" & ChrW(225) & "c minh cu" & ChrW(7889) & "i"',
    '        Case "WAIT_USER_CF":             CF_VN_GetStageLabel = "Ch" & ChrW(7901) & " user CF"',
    '        Case Else:                       CF_VN_GetStageLabel = stage',
    '    End Select',
    'End Function'
)
$vnContent = $vnLines -join "`r`n"
[System.IO.File]::WriteAllText((Join-Path $vbaBase 'modCF_VietnameseMessages.bas'), $vnContent, $enc)
Log 'CREATED_modCF_VietnameseMessages.bas'

# ── PHASE 4: Fresh WIP + import via COM ──────────────────────────────────────
Copy-Item $cp5a $wip -Force
Log "WIP_CREATED=$((Get-FileHash $wip -Algorithm SHA256).Hash)"

$xl2 = New-Object -ComObject Excel.Application
$xl2.Visible = $false; $xl2.DisplayAlerts = $false
$wb2 = $null
try {
    $wb2 = $xl2.Workbooks.Open($wip, 0, $false)
    Log 'OPENED_WIP'

    # Remove old versions of both modules if present
    foreach ($name in @('modCodexSAPReadOnly', 'modCF_VietnameseMessages')) {
        foreach ($comp in @($wb2.VBProject.VBComponents)) {
            if ($comp.Name -eq $name) {
                $wb2.VBProject.VBComponents.Remove($comp)
                Log "REMOVED=$name"
                break
            }
        }
    }

    # Import patched modules
    $wb2.VBProject.VBComponents.Import((Join-Path $vbaBase 'modCodexSAPReadOnly.bas'))
    Log 'IMPORTED_modCodexSAPReadOnly'
    $wb2.VBProject.VBComponents.Import((Join-Path $vbaBase 'modCF_VietnameseMessages.bas'))
    Log 'IMPORTED_modCF_VietnameseMessages'

    $wb2.Save()
    Log 'SAVED_WIP'
    $wb2.Close($false)
    $wb2 = $null
    Log "WIP_HASH=$((Get-FileHash $wip -Algorithm SHA256).Hash)"
} catch {
    Log "FATAL=$($_.Exception.Message)"
    if ($null -ne $wb2) { try { $wb2.Close($false) } catch {} }
    throw
} finally {
    $xl2.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl2) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # Kill Excel processes started by this script (skip if < 5 min)
    $postPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $newPids  = $postPids | Where-Object { $_ -notin $preExcel }
    if ($newPids -and ((Get-Date) - $scriptStart -gt [TimeSpan]::FromMinutes(5))) {
        $newPids | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue; Log "KILLED_HUNG=$_" }
    }
}

Set-Content -Path $logPath -Value ($L -join "`r`n") -Encoding UTF8
Log 'DONE'
