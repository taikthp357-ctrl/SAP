$ErrorActionPreference = 'Stop'
$wbPath   = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$regMod   = 'g:\SAP\CF_OB\work\scratch\modCFAllocRegression.bas'
$outLog   = 'g:\SAP\CF_OB\work\REPORTS\cp3_verify.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open($wbPath)
    $proj = $wb.VBProject
    Log ("SHEETS=" + $wb.Worksheets.Count)
    Log ("VBA_COMPONENTS=" + $proj.VBComponents.Count)

    # 1) Prove the FAST path now contains the plan lock + still calls create.
    $cm = $proj.VBComponents.Item('modCodexWTRework').CodeModule
    $body = $cm.Lines(1, $cm.CountOfLines)
    Log ("FAST_LOCK_MARKERS=" + ([regex]::Matches($body,'FAST_PATH_PLAN_LOCK')).Count)
    Log ("ASSERTQUEUEPLANLOCK_CALLS=" + ([regex]::Matches($body,'CWR_AssertQueuePlanLock')).Count)

    # Extract the CWR_RunFastToCFCore body to confirm ordering (lock BEFORE create).
    $fnStart = -1; $fnEnd = -1
    for ($i=1; $i -le $cm.CountOfLines; $i++) {
        $ln = $cm.Lines($i,1)
        if ($fnStart -lt 0 -and $ln.Contains('Public Function CWR_RunFastToCFCore()')) { $fnStart = $i }
        elseif ($fnStart -ge 0 -and $ln -match '^End Function') { $fnEnd = $i; break }
    }
    if ($fnStart -ge 0 -and $fnEnd -ge 0) {
        $lockAt = -1; $createAt = -1
        for ($i=$fnStart; $i -le $fnEnd; $i++) {
            $ln = $cm.Lines($i,1)
            if ($lockAt -lt 0 -and $ln.Contains('CWR_AssertQueuePlanLock')) { $lockAt = $i }
            if ($createAt -lt 0 -and $ln.Contains('CWR_CreateOutboundWTFromExactCandidateFast()')) { $createAt = $i }
        }
        Log ("FAST_FN_RANGE=" + $fnStart + "-" + $fnEnd + " LOCK_AT=" + $lockAt + " CREATE_AT=" + $createAt)
        if ($lockAt -gt 0 -and $createAt -gt 0 -and $lockAt -lt $createAt) {
            Log "FAST_ORDER_OK=LOCK_BEFORE_CREATE"
        } else {
            Log "FAST_ORDER_FAIL"
        }
    }

    # 2) Compile proof: allocator self-test (VBA compiles all modules before run).
    $st = "NOT_RUN"; try { $st = $xl.Run('CFAlloc_SelfTest') } catch { $st = "RUN_ERROR: " + $_.Exception.Message }
    Log ("ALLOC_SELFTEST=" + $st)

    # 3) Native V2.6 harnesses must still pass (patch didn't break executor chain).
    $fast = "NOT_RUN"; try { $fast = $xl.Run('CF_FastModuleRegression') } catch { $fast = "RUN_ERROR: " + $_.Exception.Message }
    Log ("FAST_MODULE_REGRESSION=" + $fast)
    $p3 = "NOT_RUN"
    try {
        $xl.Run('CF_Phase3_RunUnitTests') | Out-Null
        foreach ($ws in $wb.Worksheets) {
            $h1 = $null; try { $h1 = $ws.Range('H1').Value2 } catch {}
            if ($h1 -ne $null -and ("$h1" -match 'UNIT TESTS')) { $p3 = "$h1"; break }
        }
    } catch { $p3 = "RUN_ERROR: " + $_.Exception.Message }
    Log ("PHASE3_UNIT_TESTS=" + $p3)

    # 4) Allocator Qty regression (import temp, run, remove).
    foreach ($c in @($proj.VBComponents)) { if ($c.Name -eq 'modCFAllocRegression') { $proj.VBComponents.Remove($c) } }
    $proj.VBComponents.Import($regMod) | Out-Null
    $reg = "NOT_RUN"; try { $reg = $xl.Run('CFReg_RunAll') } catch { $reg = "RUN_ERROR: " + $_.Exception.Message }
    Log ("ALLOC_REGRESSION=" + $reg)
    foreach ($c in @($proj.VBComponents)) { if ($c.Name -eq 'modCFAllocRegression') { $proj.VBComponents.Remove($c) } }

    $wb.Close($false)   # never save the temp reg module
    Log "CLOSED_NO_SAVE"
} catch {
    Log ("FATAL=" + $_.Exception.Message)
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
