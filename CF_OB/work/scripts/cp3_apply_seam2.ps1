$ErrorActionPreference = 'Stop'
$wbPath  = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$outLog  = 'g:\SAP\CF_OB\work\REPORTS\cp3_seam2_patch.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

function Find-LineIndex {
    param($cm, [string]$needle, [int]$startAt = 1)
    $total = $cm.CountOfLines
    for ($i = $startAt; $i -le $total; $i++) {
        if ($cm.Lines($i, 1).Contains($needle)) { return $i }
    }
    return -1
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open($wbPath)   # writable
    $proj = $wb.VBProject
    $cm = $proj.VBComponents.Item('modCodexWTRework').CodeModule

    # Idempotency guard: bail if the FAST lock is already present.
    if ((Find-LineIndex $cm 'FAST_PATH_PLAN_LOCK') -ge 0) {
        Log "ALREADY_PATCHED_FAST_LOCK"
        $wb.Close($false)
    } else {
        # Anchor: the line inside CWR_RunFastToCFCore that loads the execution context.
        $anchor = 'If Not CWR_LoadExecutionContext(messageText) Then GoTo Stopped'
        $i = Find-LineIndex $cm $anchor 1290
        if ($i -lt 0) { throw "FAST anchor (CWR_LoadExecutionContext) not found near 1296" }
        Log ("FAST_ANCHOR_LINE=" + $i)

        # Verify we are inside CWR_RunFastToCFCore (anchor must be after its Function line).
        $fnLine = Find-LineIndex $cm 'Public Function CWR_RunFastToCFCore()'
        if ($fnLine -lt 0 -or $i -lt $fnLine -or ($i - $fnLine) -gt 15) {
            throw "Anchor not within CWR_RunFastToCFCore (fn=$fnLine anchor=$i)"
        }

        # Insert the SAME plan-lock guard the QUEUE path uses (CWR_AssertQueuePlanLock),
        # reading the same PLAN + WH_ACTUAL sheets. Additive; no existing line changed.
        $ins = @(
            '    '' FAST_PATH_PLAN_LOCK (CP3 Seam 2): FAST must verify the CHECK-locked',
            '    '' plan hash before any SAP write, exactly like the QUEUE path does.',
            '    stageText = "VERIFY_PLAN_LOCK"',
            '    If Not CWR_AssertQueuePlanLock(ThisWorkbook.Worksheets("PLAN"), _',
            '            ThisWorkbook.Worksheets("WH_ACTUAL"), messageText) Then GoTo Stopped'
        ) -join "`r`n"
        $cm.InsertLines($i + 1, $ins)
        Log ("FAST_LOCK_INSERTED_AFTER_LINE=" + $i)

        $wb.Save()
        Log "SAVED"
        $wb.Close($true)
    }
    Log "APPLY_DONE"
} catch {
    Log ("APPLY_FATAL=" + $_.Exception.Message)
    try { $wb.Close($false) } catch {}
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if ($preExcel -notcontains $p.Id) {
            try { if ($p.StartTime -ge $scriptStart) { Stop-Process -Id $p.Id -Force; Log ("KILLED_EXCEL_PID=" + $p.Id) } } catch {}
        }
    }
    Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
}
