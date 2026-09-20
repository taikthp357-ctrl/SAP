$ErrorActionPreference = 'Stop'
$wbPath  = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$modFile = 'g:\SAP\CF_OB\work\scratch\modCFGroupAllocator.bas'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date

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

    # --- Remove any prior copy of the allocator, then import fresh ---
    foreach ($c in @($proj.VBComponents)) {
        if ($c.Name -eq 'modCFGroupAllocator') { $proj.VBComponents.Remove($c) }
    }
    $proj.VBComponents.Import($modFile) | Out-Null
    Write-Output "IMPORTED modCFGroupAllocator"

    # --- Patch modCodexWTRework ---
    $wt = $proj.VBComponents.Item('modCodexWTRework')
    $cm = $wt.CodeModule

    # PATCH 1: per-member CONTRACT_B gate in CWR_SchedulerVerifyPRDOGroup
    $p1old = 'If Abs(liveTotal - expectedTarget) <= 0.000001 Then rec("QtyReady") = True'
    $i1 = Find-LineIndex $cm $p1old
    if ($i1 -lt 0) { throw "PATCH1 anchor not found" }
    $indent = ($cm.Lines($i1,1) -replace '\S.*$','')
    $cm.ReplaceLine($i1, $indent + 'If CFAlloc_IsContractBPass(rec) Then rec("QtyReady") = True')
    Write-Output "PATCH1_APPLIED line=$i1"

    # PATCH 2: CONTRACT_B gate condition in CWR_SchedulerVerifyQtySources
    $p2old = 'Abs(CDbl(rec("GroupLiveQty")) - CDbl(rec("GroupTarget"))) <= 0.000001 Then'
    $i2 = Find-LineIndex $cm $p2old
    if ($i2 -lt 0) { throw "PATCH2 anchor not found" }
    $indent2 = ($cm.Lines($i2,1) -replace '\S.*$','')
    $cm.ReplaceLine($i2, $indent2 + 'CFAlloc_IsContractBPass(rec) Then')
    Write-Output "PATCH2_APPLIED line=$i2"

    # PATCH 2b: correct the audit token (PRDO_QTY_ALREADY_TARGET / NO_QTY_WRITE)
    $p2bold = '"PRDO_QTY_ALREADY_TARGET|LIVE_TOTAL=" & CStr(rec("GroupLiveQty")) & _'
    $i2b = Find-LineIndex $cm $p2bold $i2
    if ($i2b -ge 0) {
        $ind = ($cm.Lines($i2b,1) -replace '\S.*$','')
        $cm.ReplaceLine($i2b, $ind + '"MEMBER_ALLOCATION_COMPLETE|REQUIRED_CREATE=0|LIVE_TOTAL=" & CStr(rec("GroupLiveQty")) & _')
        Write-Output "PATCH2b_APPLIED line=$i2b"
    }
    $p2cold = '"|TARGET=" & CStr(rec("GroupTarget")) & "|SOURCE_VERIFY=SKIP|NO_QTY_WRITE"'
    $i2c = Find-LineIndex $cm $p2cold $i2
    if ($i2c -ge 0) {
        $ind = ($cm.Lines($i2c,1) -replace '\S.*$','')
        $cm.ReplaceLine($i2c, $ind + '"|TARGET=" & CStr(rec("GroupTarget")) & "|SOURCE_VERIFY=SKIP|CONTRACT_B_PASS"')
        Write-Output "PATCH2c_APPLIED line=$i2c"
    }

    # PATCH 3: additive group-allocation-status log after PRDO baseline verify
    $p3anchor = 'If Not CWR_SchedulerVerifyQtySources(members, stopAll, messageText) Then'
    $i3 = Find-LineIndex $cm $p3anchor
    if ($i3 -ge 0) {
        $logLines = @(
            '        CWR_Log "QTY_GROUP_ALLOCATION_STATUS", CStr(members.item(1)("OB")), _',
            '            CStr(members.item(1)("Doc")), "PASS", CFAlloc_GroupStatusText(members)'
        ) -join "`r`n"
        $cm.InsertLines($i3, $logLines)
        Write-Output "PATCH3_APPLIED line=$i3"
    }

    $wb.Save()
    Write-Output "SAVED"
    $wb.Close($true)
    Write-Output "APPLY_DONE"
} catch {
    Write-Output ("APPLY_FATAL=" + $_.Exception.Message)
    try { $wb.Close($false) } catch {}
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if ($preExcel -notcontains $p.Id) {
            try { if ($p.StartTime -ge $scriptStart) { Stop-Process -Id $p.Id -Force } } catch {}
        }
    }
}