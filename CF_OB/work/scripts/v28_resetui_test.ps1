$ErrorActionPreference = 'Stop'
$cp2 = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP2_REPO.xlsm'
$outLog = 'g:\SAP\CF_OB\work\REPORTS\v28_resetui_test.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

# Read-only verification. CFResetUI_SelfTest intentionally mutates PLAN!AJ2 and
# INPUT!H1:N3 in-memory; it restores AJ2 and this workbook is never saved.
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    $wb = $xl.Workbooks.Open($cp2, 0, $true)
    $result = 'NOT_RUN'; try { $result = [string]$xl.Run('CFResetUI_SelfTest') } catch { $result = 'ERR:' + $_.Exception.Message }
    Log ('CFResetUI_SelfTest=' + $result)
    $ok = ($result -match '^PASS\|')
    Log ('VERDICT resetUiOk=' + $ok)
    $wb.Close($false)
} catch {
    Log ('COM_FATAL=' + $_.Exception.Message)
    try { $wb.Close($false) } catch {}
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    foreach ($p in @(Get-Process EXCEL -ErrorAction SilentlyContinue)) {
        if ($preExcel -notcontains $p.Id) {
            try { if ($p.StartTime -ge $scriptStart) { Stop-Process -Id $p.Id -Force; Log ('KILLED_EXCEL_PID=' + $p.Id) } } catch {}
        }
    }
    Set-Content -Path $outLog -Value ($L -join "`r`n") -Encoding UTF8
}
