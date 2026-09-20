$ErrorActionPreference = 'Stop'
$wbPath  = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$outLog  = 'g:\SAP\CF_OB\work\REPORTS\cp4_harden_mon.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }
function Find-LineIndex { param($cm,[string]$needle,[int]$startAt=1)
    for ($i=$startAt; $i -le $cm.CountOfLines; $i++){ if ($cm.Lines($i,1).Contains($needle)){ return $i } }
    return -1 }

$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    $wb = $xl.Workbooks.Open($wbPath)
    $cm = $wb.VBProject.VBComponents.Item('modCodexWTRework').CodeModule

    if ((Find-LineIndex $cm 'CP4_LEGACY_MON_REUSE_DISABLED') -ge 0) {
        Log "ALREADY_HARDENED"
        $wb.Close($false)
    } else {
        # Fail-closed the dead legacy same-tab MON reuse so it can never regress
        # the Single-MON-Broker invariant (never reuse PRDO tab as MON).
        $fn = Find-LineIndex $cm 'Private Function CWR_ReuseVerifiedPRDOTabForMON('
        if ($fn -lt 0) { throw "legacy reuse fn not found" }
        # Insert guard as the first executable line after the On Error line.
        $onErr = Find-LineIndex $cm 'On Error GoTo Failed' $fn
        if ($onErr -lt 0 -or ($onErr - $fn) -gt 6) { throw "On Error anchor not inside legacy fn" }
        $ins = @(
            '    '' CP4_LEGACY_MON_REUSE_DISABLED: Single MON Broker (CWR_GetMONSession) is the',
            '    '' only sanctioned resolver. This legacy same-tab reuse must never run.',
            '    messageText = "STOP_LEGACY_MON_REUSE_DISABLED|use CWR_GetMONSession dedicated reader; PRDO tab is never reused as MON."',
            '    Set ses = Nothing',
            '    Exit Function'
        ) -join "`r`n"
        $cm.InsertLines($onErr + 1, $ins)
        Log ("LEGACY_MON_REUSE_DISABLED_AT=" + ($onErr+1))
        $wb.Save()
        Log "SAVED"
        $wb.Close($true)
    }
    Log "DONE"
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
