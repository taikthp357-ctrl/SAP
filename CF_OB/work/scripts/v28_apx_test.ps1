# v28_apx_test.ps1
# Run the appendix P0 offline regression suite (modCFApxRegression.bas) on CP5a_CHECKBAND.
# READ-ONLY workbook open. The harness module lives in work\scratch\ (NOT vba_v28\)
# and is remove-then-imported at runtime (same pattern as modCFAllocRegression).
# Returns APX_P0_ALL_PASS / APX_P0_FAIL + counters + log path.
$ErrorActionPreference = 'Stop'
$cp5a   = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP5a_CHECKBAND.xlsm'
$apxMod = 'g:\SAP\CF_OB\work\scratch\modCFApxRegression.bas'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\v28_apx_test.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Host $m }

if (-not (Test-Path $cp5a))   { throw "NO_CP5a=$cp5a" }
if (-not (Test-Path $apxMod)) { throw "NO_APX_MODULE=$apxMod" }
Log ("CP5a_SHA256=" + (Get-FileHash -Path $cp5a -Algorithm SHA256).Hash)
Log ("APX_MODULE_SHA256=" + (Get-FileHash -Path $apxMod -Algorithm SHA256).Hash)

function ImportModule($vbp, $name, $path){
    $comp = $null
    foreach ($c in $vbp.VBComponents) { if ($c.Name -eq $name) { $comp = $c; break } }
    if ($comp -ne $null) { $vbp.VBComponents.Remove($comp) | Out-Null }
    $vbp.VBComponents.Import($path) | Out-Null
}

$result = 'NOT_RUN'
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    $wb = $xl.Workbooks.Open($cp5a, 0, $true)  # ReadOnly=TRUE (no Save)
    $vbp = $wb.VBProject
    ImportModule $vbp 'modCFApxRegression' $apxMod
    Log 'Imported modCFApxRegression'

    try {
        $result = [string]$xl.Run('CFApx_RunP0')
    } catch {
        $result = 'ERR:' + $_.Exception.Message
    }
    Log ("RESULT=" + $result)

    $wb.Close($false)  # NEVER save (ReadOnly test run)
} catch {
    Log ('FATAL=' + $_.Exception.Message)
    try { if ($wb -ne $null) { $wb.Close($false) } } catch {}
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if ($preExcel -notcontains $p.Id) {
            try { if ($p.StartTime -ge $scriptStart) { Stop-Process -Id $p.Id -Force; Log ('KILLED_EXCEL_PID=' + $p.Id) } } catch {}
        }
    }
}

Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
Log ('LOG_WRITTEN=' + $outLog)

if ($result -like 'APX_P0_ALL_PASS*') {
    Log 'APX_P0_OK'
    exit 0
} else {
    Log 'APX_P0_FAILED_OR_ERROR'
    exit 1
}
