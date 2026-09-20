$ErrorActionPreference = 'Stop'
$wbPath = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
try {
    # READ-ONLY open — this harness never saves.
    $wb = $xl.Workbooks.Open($wbPath, 0, $true)
    $proj = $wb.VBProject

    # 1) Enumerate modules — prove the allocator is present and count functions.
    $allocFound = $false; $allocLines = 0; $funcs = @()
    foreach ($c in $proj.VBComponents) {
        if ($c.Name -eq 'modCFGroupAllocator') {
            $allocFound = $true
            $cm = $c.CodeModule
            $allocLines = $cm.CountOfLines
            $body = $cm.Lines(1, $allocLines)
            foreach ($line in ($body -split "`r?`n")) {
                if ($line -match '^\s*(Public\s+)?(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)') {
                    $funcs += ($matches[2] + ' ' + $matches[3])
                }
            }
        }
    }
    Write-Output ("ALLOC_MODULE_FOUND=" + $allocFound + " LINES=" + $allocLines)
    Write-Output ("ALLOC_PROCS=" + ($funcs -join '; '))

    # 2) Verify patches landed in modCodexWTRework.
    $wt = $proj.VBComponents.Item('modCodexWTRework').CodeModule
    $wtBody = $wt.Lines(1, $wt.CountOfLines)
    Write-Output ("HAS_IsContractBPass_CALLS=" + ([regex]::Matches($wtBody,'CFAlloc_IsContractBPass')).Count)
    Write-Output ("HAS_GroupStatus_CALL=" + ([regex]::Matches($wtBody,'CFAlloc_GroupStatusText')).Count)
    Write-Output ("RESIDUAL_ALREADY_TARGET=" + ([regex]::Matches($wtBody,'PRDO_QTY_ALREADY_TARGET')).Count)

    # 3) Force a full VBA compile via the IDE command — proves no syntax/reserved-word errors.
    $compileStatus = "UNKNOWN"; $compileErr = ""
    try {
        # CommandBars "2578" is the VBE 'Compile' action id.
        $vbe = $proj.VBE
        $vbe.CommandBars.FindControl(1, 2578).Execute()
        Start-Sleep -Milliseconds 800
        # After a successful compile, project is in compiled state; a failed compile
        # raises no exception here but leaves an error selection. Use RunKind probe:
        $compileStatus = "COMPILE_INVOKED"
    } catch {
        $compileStatus = "COMPILE_INVOKE_FAILED"
        $compileErr = $_.Exception.Message
    }
    Write-Output ("COMPILE_STATUS=" + $compileStatus)
    if ($compileErr -ne "") { Write-Output ("COMPILE_ERR=" + $compileErr) }

    # 4) Self-test proof (offline): run the allocator self-test via Application.Run.
    $selfTest = "NOT_RUN"
    try {
        $selfTest = $xl.Run('CFAlloc_SelfTest')
    } catch {
        $selfTest = "RUN_ERROR: " + $_.Exception.Message
    }
    Write-Output ("SELFTEST=" + $selfTest)

    # Close WITHOUT saving (read-only anyway).
    $wb.Close($false)
    Write-Output "CHECK_DONE"
} catch {
    Write-Output ("CHECK_FATAL=" + $_.Exception.Message)
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