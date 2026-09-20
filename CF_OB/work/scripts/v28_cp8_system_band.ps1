# v28_cp8_system_band.ps1 — CP8: Merge CONFIG/SAP_ACTION_MAP/SAP_PROBE_LOG into SYSTEM
# Band layout (CORRECTED after full column audit):
#   SAP_MAP anchor : cols  1-16  (A:P)
#   CONFIG         : cols 17-26  (Q:Z)   offset +16  (key=1→17 val=2→18 ... packing G:J→W:Z)
#   SAP_ACTION_MAP : cols 27-36  (AA:AJ) offset +26  (COL_KEY=1→27 COL_NOTES=10→36)
#   SAP_PROBE_LOG  : cols 37-53  (AK:BA) offset +36  (Time=1→37 ... TotalMs=17→53)
#
# NOT_LIVE_PROVEN: All 3 siblings have live consumers — final structure correct;
# live paths (modCF_AutoRun_TwoPilot, modCF_Executor, modCodexWTRework,
#              modSAPActionRegistry, modSAPReadOnlyProbe, modSAPRunGuard)
# cannot be verified offline.

$ErrorActionPreference = 'Stop'
$cp7        = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP7_WORKLIST.xlsm'
$wip        = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP8_WIP.xlsm'
$vbaBase    = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$vbaBand    = 'g:\SAP\CF_OB\work\scratch\vba_cp8_band'
$regModule  = 'g:\SAP\CF_OB\work\scratch\modCFAllocRegression.bas'
$logPath    = 'g:\SAP\CF_OB\work\REPORTS\v28_cp8_system_band.log'
$preExcel   = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m) { [void]$L.Add($m); Write-Host $m }

if (-not (Test-Path $cp7)) { throw "NO_CP7=$cp7" }

# ─── helper: scope-limited column rewrite ─────────────────────────────────────
# Rewrites .Cells(anyVar, oldCol) → .Cells(anyVar, newCol) ONLY on lines
# inside a function that has set varName = Worksheets("SYSTEM").
# Also rewrites Range("A3:B → "Q3:R" and Range("G3:J" → "W3:Z" within scope.
function Rewrite-BandCols {
    param(
        [string]$fileContent,
        [string]$varName,          # 'ws', 'wsc' etc.
        [hashtable]$colMap,        # @{1=17; 2=18; ...}
        [hashtable]$rangeMap       # @{'Range("A3:B' = 'Range("Q3:R'; ...}
    )
    $lines = $fileContent -split "`r?`n"
    $inScope  = $false
    $scopeVar = $varName
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        # Enter scope: Set varName = ThisWorkbook.Worksheets("SYSTEM")
        if ($line -match "Set\s+$scopeVar\s*=.*Worksheets\(`"SYSTEM`"\)") {
            $inScope = $true
        }
        # Exit scope: next Set of same var to anything else, or End Sub/Function/Property
        if ($inScope -and $line -match "^\s*End\s+(Sub|Function|Property)") {
            $inScope = $false
        }
        if ($inScope -and $line -match "Set\s+$scopeVar\s*=" -and
            $line -notmatch "Worksheets\(`"SYSTEM`"\)" -and $line -notmatch "Nothing") {
            $inScope = $false
        }

        if ($inScope) {
            # Range replacements first (before cell-level replacements)
            foreach ($oldR in $rangeMap.Keys) {
                if ($line -match [Regex]::Escape($oldR)) {
                    $line = $line -replace [Regex]::Escape($oldR), $rangeMap[$oldR]
                }
            }
            # Column replacements (descending to avoid e.g. 1→17 then 17→???):
            foreach ($oldC in ($colMap.Keys | Sort-Object -Descending)) {
                $newC = $colMap[$oldC]
                # Pattern: .Cells(anyToken, oldCol) where oldCol is a standalone number
                $line = $line -replace "(\.$scopeVar\.Cells\s*\([^,]+,\s*)$oldC(\s*\))", "`${1}$newC`${2}"
                $line = $line -replace "(\.$scopeVar\.Cells\s*\([^,]+,\s*)$oldC(\s*\)\s*\.)", "`${1}$newC`${2}"
                # Also handle rows.count pattern: .Cells(ws.rows.count, oldCol)
                $line = $line -replace "(\.$scopeVar\.Cells\s*\($scopeVar\.[Rr]ows\.[Cc]ount\s*,\s*)$oldC(\s*\))", "`${1}$newC`${2}"
            }
        }
        $lines[$i] = $line
    }
    return ($lines -join "`r`n")
}

# Create band workspace
if (Test-Path $vbaBand) { Remove-Item $vbaBand -Recurse -Force }
New-Item -ItemType Directory -Path $vbaBand | Out-Null
Copy-Item "$vbaBase\*.bas" $vbaBand
Copy-Item "$vbaBase\*.cls" $vbaBand
Log "BAND_WORKSPACE_CREATED"

# CONFIG offset map (cols 1-10 each + 16)
$cfgColMap = @{1=17;2=18;3=19;4=20;5=21;6=22;7=23;8=24;9=25;10=26}
$cfgRangeMap = @{'Range("A3:B'='Range("Q3:R'; 'Range("G3:J'='Range("W3:Z'}

# ─── CONFIG consumers ─────────────────────────────────────────────────────────
# All 9 modules: change sheet name then rewrite columns within CONFIG scope

$configModules = @(
    @{file='modCF_FastContract.bas';    var='ws' },
    @{file='modCF_AutoRun_TwoPilot.bas'; var='ws' },
    @{file='modCF_Executor.bas';         var='ws' },
    @{file='modCodexWTRework.bas';       var='ws' },
    @{file='modCF_ApplyUserHandoff.bas'; var='ws' },
    @{file='modCF_ActualPicking.bas';    var='ws' },
    @{file='modCF_PreLiveAudit.bas';     var='ws' },
    @{file='modSAPRunGuard.bas';         var='ws' }
)

foreach ($m in $configModules) {
    $path = Join-Path $vbaBand $m.file
    $c = Get-Content $path -Raw -Encoding UTF8
    # 1. Rename sheet
    $c = $c -replace 'Worksheets\("CONFIG"\)', 'Worksheets("SYSTEM")'
    # 2. Rewrite band cols within scope
    $c = Rewrite-BandCols -fileContent $c -varName $m.var -colMap $cfgColMap -rangeMap $cfgRangeMap
    Set-Content $path -Value $c -Encoding UTF8 -NoNewline
    Log "CFG_REWRITTEN|$($m.file)"
}

# ModuleHelper special: uses 'wsc' for CONFIG and 'ws' for DANH_MUC — rewrite wsc only
$mhPath = Join-Path $vbaBand 'ModuleHelper.bas'
$mhContent = Get-Content $mhPath -Raw -Encoding UTF8
$mhContent = $mhContent -replace 'Worksheets\("CONFIG"\)', 'Worksheets("SYSTEM")'
$mhContent = Rewrite-BandCols -fileContent $mhContent -varName 'wsc' -colMap $cfgColMap -rangeMap $cfgRangeMap
Set-Content $mhPath -Value $mhContent -Encoding UTF8 -NoNewline
Log "CFG_REWRITTEN|ModuleHelper.bas"

# ─── SAP_ACTION_MAP consumers ──────────────────────────────────────────────────
# modSAPActionRegistry: update Private Const values (already declared, just change numbers)
$samRegPath = Join-Path $vbaBand 'modSAPActionRegistry.bas'
$samRegC = Get-Content $samRegPath -Raw -Encoding UTF8
$samRegC = $samRegC -replace 'Private Const SHEET_ACTION As String = "SAP_ACTION_MAP"', 'Private Const SHEET_ACTION As String = "SYSTEM"'
$samRegC = $samRegC -replace 'Private Const COL_KEY As Long = 1\b',      'Private Const COL_KEY As Long = 27'
$samRegC = $samRegC -replace 'Private Const COL_CONTEXT As Long = 3\b',  'Private Const COL_CONTEXT As Long = 29'
$samRegC = $samRegC -replace 'Private Const COL_COMMAND As Long = 5\b',  'Private Const COL_COMMAND As Long = 31'
$samRegC = $samRegC -replace 'Private Const COL_STATUS As Long = 6\b',   'Private Const COL_STATUS As Long = 32'
$samRegC = $samRegC -replace 'Private Const COL_EVIDENCE As Long = 7\b', 'Private Const COL_EVIDENCE As Long = 33'
$samRegC = $samRegC -replace 'Private Const COL_REVIEW As Long = 9\b',   'Private Const COL_REVIEW As Long = 35'
$samRegC = $samRegC -replace 'Private Const COL_NOTES As Long = 10\b',   'Private Const COL_NOTES As Long = 36'
Set-Content $samRegPath -Value $samRegC -Encoding UTF8 -NoNewline
Log "SAM_REWRITTEN|modSAPActionRegistry.bas"

# modSAPActionMap: direct ws access; cols used: 1(key), 5(command), 6(status) → 27, 31, 32
$samMapPath = Join-Path $vbaBand 'modSAPActionMap.bas'
$samMapC = Get-Content $samMapPath -Raw -Encoding UTF8
$samMapC = $samMapC -replace 'Worksheets\("SAP_ACTION_MAP"\)', 'Worksheets("SYSTEM")'
$samColMap = @{10=36;9=35;7=33;6=32;5=31;3=29;1=27}  # descending for safe replacement
$samMapC = Rewrite-BandCols -fileContent $samMapC -varName 'ws' -colMap $samColMap -rangeMap @{}
Set-Content $samMapPath -Value $samMapC -Encoding UTF8 -NoNewline
Log "SAM_REWRITTEN|modSAPActionMap.bas"

# modCF_PreLiveAudit: SAP_ACTION_MAP at lines 241 (col1 check), 396-406 (cols 1,5,6)
$preAuditPath = Join-Path $vbaBand 'modCF_PreLiveAudit.bas'
$preAuditC = Get-Content $preAuditPath -Raw -Encoding UTF8  # already CONFIG-rewritten
$preAuditC = $preAuditC -replace 'Worksheets\("SAP_ACTION_MAP"\)', 'Worksheets("SYSTEM")'
$preAuditC = Rewrite-BandCols -fileContent $preAuditC -varName 'ws' -colMap $samColMap -rangeMap @{}
Set-Content $preAuditPath -Value $preAuditC -Encoding UTF8 -NoNewline
Log "SAM_REWRITTEN|modCF_PreLiveAudit.bas"

# ─── SAP_PROBE_LOG consumer: modSAPReadOnlyProbe ──────────────────────────────
# Offset +36: cols 1-17 → 37-53
$splPath = Join-Path $vbaBand 'modSAPReadOnlyProbe.bas'
$splC = Get-Content $splPath -Raw -Encoding UTF8

# Change Const (SP_LOG points to SYSTEM now; EnsureLogSheet will return SYSTEM always)
$splC = $splC -replace 'Private Const SP_LOG As String = "SAP_PROBE_LOG"', 'Private Const SP_LOG As String = "SYSTEM"'

# Rewrite SAPProbe_EnsureLogSheet: remove Add-sheet logic, always return SYSTEM, update header cols
$splLines = $splC -split "`r?`n"
$ensureStart = -1; $ensureEnd = -1
for ($j = 0; $j -lt $splLines.Count; $j++) {
    if ($splLines[$j] -match '^Private Function SAPProbe_EnsureLogSheet') { $ensureStart = $j }
    if ($ensureStart -ge 0 -and $splLines[$j] -match '^End Function') { $ensureEnd = $j; break }
}
if ($ensureStart -ge 0 -and $ensureEnd -gt $ensureStart) {
    $newBody = @(
        $splLines[$ensureStart],
        '    Set SAPProbe_EnsureLogSheet = ThisWorkbook.Worksheets("SYSTEM")',
        '    If SAPProbe_Text(SAPProbe_EnsureLogSheet.Cells(1, 37).Value2) <> "Time" Then',
        '        SAPProbe_EnsureLogSheet.Cells(1, 37).Resize(1, 17).Value = Array("Time", "RunID", "Probe", "Action", "SessionID", "ExpectedKey", "OB", "Document", "WT", "Expected", "Live", "Target", "FindMs", "ReadMs", "Result", "ErrorStage", "TotalMs")',
        '        SAPProbe_EnsureLogSheet.rows(1).Font.Bold = True',
        '    End If',
        $splLines[$ensureEnd]
    )
    $splLines = $splLines[0..($ensureStart-1)] + $newBody + $splLines[($ensureEnd+1)..($splLines.Count-1)]
    $splC = $splLines -join "`r`n"
}

# Rewrite SAPProbe_Log: ws.Cells(r, N) where N=1..17 → N+36; End(xlUp) col1→37
$splColMap = @{}; for ($n=17; $n -ge 1; $n--) { $splColMap[$n] = $n+36 }
$splC = Rewrite-BandCols -fileContent $splC -varName 'ws' -colMap $splColMap -rangeMap @{}
Set-Content $splPath -Value $splC -Encoding UTF8 -NoNewline
Log "SPL_REWRITTEN|modSAPReadOnlyProbe.bas"

# modSAPRunGuard SAP_PROBE_LOG: ws at line 422, cols 1→37, 3→39, 15→51
$srgPath = Join-Path $vbaBand 'modSAPRunGuard.bas'
$srgC = Get-Content $srgPath -Raw -Encoding UTF8  # already CONFIG-rewritten
$srgC = $srgC -replace 'Worksheets\("SAP_PROBE_LOG"\)', 'Worksheets("SYSTEM")'
$srgColMap = @{15=51;3=39;1=37}  # only the probe-log columns used
$srgC = Rewrite-BandCols -fileContent $srgC -varName 'ws' -colMap $srgColMap -rangeMap @{}
Set-Content $srgPath -Value $srgC -Encoding UTF8 -NoNewline
Log "SPL_REWRITTEN|modSAPRunGuard.bas"

Log "ALL_MODULES_REWRITTEN"

# ─── PHASE 2: Copy baseline + data migration via COM ─────────────────────────
Copy-Item $cp7 $wip -Force
Log "WIP_CREATED"

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $null

try {
    $wb = $xl.Workbooks.Open($wip, 0, $false)
    Log 'OPENED_WIP'

    $wsSystem = $wb.Worksheets.Item('SYSTEM')
    $wsConfig  = $wb.Worksheets.Item('CONFIG')
    $wsSAM     = $wb.Worksheets.Item('SAP_ACTION_MAP')
    $wsSPL     = $wb.Worksheets.Item('SAP_PROBE_LOG')

    # CONFIG: up to col 10 → SYSTEM cols 17-26 (Q:Z)
    $cfgLR = $wsConfig.Cells.Item($wsConfig.Rows.Count, 1).End(-4162).Row
    if ($cfgLR -ge 1) {
        $cfgData = $wsConfig.Range("A1:J$cfgLR").Value2
        $wsSystem.Range("Q1").Resize($cfgLR, 10).Value2 = $cfgData
        Log "COPIED_CONFIG|rows=$cfgLR"
    }

    # SAP_ACTION_MAP: 10 cols → SYSTEM cols 27-36 (AA:AJ)
    $samLR = $wsSAM.Cells.Item($wsSAM.Rows.Count, 1).End(-4162).Row
    if ($samLR -ge 1) {
        $samData = $wsSAM.Range("A1:J$samLR").Value2
        $wsSystem.Range("AA1").Resize($samLR, 10).Value2 = $samData
        Log "COPIED_SAP_ACTION_MAP|rows=$samLR"
    }

    # SAP_PROBE_LOG: 17 cols → SYSTEM cols 37-53 (AK:BA)
    $splLR = $wsSPL.Cells.Item($wsSPL.Rows.Count, 1).End(-4162).Row
    if ($splLR -ge 1) {
        $splData = $wsSPL.Range("A1:Q$splLR").Value2
        $wsSystem.Range("AK1").Resize($splLR, 17).Value2 = $splData
        Log "COPIED_SAP_PROBE_LOG|rows=$splLR"
    }

    $wb.Save()
    Log 'DATA_MIGRATED'

    # ─── PHASE 3: Import rewritten modules ────────────────────────────────────
    $modsToImport = @(
        'modCF_FastContract.bas',
        'modCF_AutoRun_TwoPilot.bas',
        'modCF_Executor.bas',
        'modCodexWTRework.bas',
        'modCF_ApplyUserHandoff.bas',
        'modCF_ActualPicking.bas',
        'modCF_PreLiveAudit.bas',
        'modSAPRunGuard.bas',
        'ModuleHelper.bas',
        'modSAPActionRegistry.bas',
        'modSAPActionMap.bas',
        'modSAPReadOnlyProbe.bas'
    )

    foreach ($mod in $modsToImport) {
        $modName = $mod -replace '\.(bas|cls)$', ''
        foreach ($comp in @($wb.VBProject.VBComponents)) {
            if ($comp.Name -eq $modName) {
                $wb.VBProject.VBComponents.Remove($comp); break
            }
        }
        $wb.VBProject.VBComponents.Import((Join-Path $vbaBand $mod))
        Log "IMPORTED=$modName"
    }

    $wb.Save()
    Log 'MODULES_IMPORTED'

    # ─── PHASE 4: VeryHide sibling sheets ────────────────────────────────────
    $wsConfig.Visible = 2; $wsSAM.Visible = 2; $wsSPL.Visible = 2
    $wb.Save()
    Log 'SIBLINGS_VERYHIDDEN'

    # ─── PHASE 5: Offline probe (CONFIG read + SAP_ACTION_MAP read) ───────────
    # Seed: ensure CF_ConfigBool("FAST_FLOW_ENABLED", False) returns False
    # (reads from SYSTEM band col17, meaning the key-value is intact)
    $cfgProbe = $xl.Run('CF_ConfigValue', 'SAP_MODE')
    Log "PROBE_CONFIG_SAP_MODE=[$cfgProbe]"

    # SAPActionMap_Review reads col27(key)/31(cmd)/32(status)
    Dim $samReviewMsg = ''
    $samResult = $xl.Run('SAPActionMap_Review', [ref]$samReviewMsg)
    Log "PROBE_SAM_REVIEW|result=$samResult|msg=$samReviewMsg"

    $wb.Save()

    # ─── PHASE 6: Regression gates ────────────────────────────────────────────
    $r1 = $xl.Run('CFAlloc_SelfTest')
    if ($r1 -ne 'PASS') { throw "GATE_CFAlloc_FAIL|$r1" }
    Log "GATE_CFAlloc=PASS"

    $xl.Run('SSR_CompileProbe')
    Log "GATE_SSR_CompileProbe=RAN"

    $r3 = $xl.Run('SSR_SelfTest')
    if ($r3 -ne 'PASS') { throw "GATE_SSR_SelfTest_FAIL|$r3" }
    Log "GATE_SSR_SelfTest=PASS"

    if (Test-Path $regModule) {
        foreach ($comp in @($wb.VBProject.VBComponents)) {
            if ($comp.Name -eq 'modCFAllocRegression') { $wb.VBProject.VBComponents.Remove($comp); break }
        }
        $wb.VBProject.VBComponents.Import($regModule)
        Log 'IMPORTED_modCFAllocRegression'
    }
    $r4 = $xl.Run('CFReg_RunAll')
    if ($r4 -notmatch 'REG_ALL_PASS.*FAIL=0') { throw "GATE_CFReg_FAIL|$r4" }
    Log "GATE_CFReg=PASS|$r4"
    foreach ($comp in @($wb.VBProject.VBComponents)) {
        if ($comp.Name -eq 'modCFAllocRegression') { $wb.VBProject.VBComponents.Remove($comp); Log 'REMOVED_modCFAllocRegression'; break }
    }

    $r5 = $xl.Run('CF_FastModuleRegression')
    if (-not ($r5 -like 'PASS*')) { throw "GATE_FastReg_FAIL|$r5" }
    Log "GATE_FastReg=PASS"

    $wsTest = $wb.Worksheets.Item('PHASE3_TESTS')
    $r6 = $wsTest.Range('H1').Value2
    if ($r6 -notmatch '11/11.*SAP Write=0') { throw "GATE_Phase3_FAIL|$r6" }
    Log "GATE_Phase3=PASS|$r6"

    Log 'ALL_GATES_GREEN'

    $wb.Save()
    $wb.Close($false); $wb = $null
    Log "WIP_HASH=$((Get-FileHash $wip -Algorithm SHA256).Hash)"

    $cp8 = 'g:\SAP\CF_OB\work\WORKING\260821_CF_OB_V2.8_CP8_SYSTEM_BAND.xlsm'
    Move-Item $wip $cp8 -Force
    Log "CP8_CHECKPOINT=$cp8"
    Log "CP8_HASH=$((Get-FileHash $cp8 -Algorithm SHA256).Hash)"

} catch {
    Log "FATAL=$($_.Exception.Message)"
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    throw
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $postPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $newPids  = $postPids | Where-Object { $_ -notin $preExcel }
    if ($newPids -and ((Get-Date) - $scriptStart -gt [TimeSpan]::FromMinutes(5))) {
        $newPids | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue; Log "KILLED_HUNG=$_" }
    }
}

Set-Content -Path $logPath -Value ($L -join "`r`n") -Encoding UTF8
Log 'DONE'
