$ErrorActionPreference = 'Stop'
$wip     = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_WIP.xlsm'
$cpFix   = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP1b_LIVEFIX.xlsm'
$expDir  = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$regBas  = 'g:\SAP\CF_OB\work\scratch\modCFAllocRegression.bas'
$outLog  = 'g:\SAP\CF_OB\work\REPORTS\v28_livefix_verify.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

# Live-fix verify: re-import the edited executor (Bug A/B/C fixes) then prove the
# whole VBProject compiles + the offline allocation regression stays green.
# NOTE: this proves COMPILE + ALLOCATION regression only; the 3 live SAP paths
# remain NOT_LIVE_PROVEN (SAP GUI Scripting absent offline).
$reimport = @('modCodexWTRework')

$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    $wb = $xl.Workbooks.Open($wip, 0, $false)
    $proj = $wb.VBProject
    Log ("OPENED=" + $wb.Name + " components=" + $proj.VBComponents.Count)

    foreach ($name in $reimport) {
        $existing = $null
        foreach ($c in $proj.VBComponents) { if ($c.Name -eq $name) { $existing = $c; break } }
        if ($null -ne $existing) {
            if ($existing.Type -ne 1) { Log ("SKIP_NONSTD=" + $name + " type=" + $existing.Type); continue }
            $proj.VBComponents.Remove($existing)
            Log ("REMOVED=" + $name)
        } else { Log ("NEW_MODULE=" + $name) }
        $path = Join-Path $expDir ($name + '.bas')
        if (-not (Test-Path $path)) { Log ("MISSING_BAS=" + $path); continue }
        $proj.VBComponents.Import($path) | Out-Null
        Log ("IMPORTED=" + $name)
    }
    Log ("COMPONENTS_AFTER=" + $proj.VBComponents.Count)

    # 1) compile proof (VBA compiles the whole project on first Run)
    $st = "NOT_RUN"; try { $st = $xl.Run('CFAlloc_SelfTest') } catch { $st = "ERR:" + $_.Exception.Message }
    Log ("COMPILE_PROOF_CFAlloc_SelfTest=" + $st)

    # 2) allocation regression (runtime-import the external suite module)
    $regName = 'modCFAllocRegression'
    $er = $null; foreach ($c in $proj.VBComponents) { if ($c.Name -eq $regName) { $er = $c; break } }
    if ($null -ne $er) { $proj.VBComponents.Remove($er) }
    $proj.VBComponents.Import($regBas) | Out-Null
    Log ("IMPORTED_RUNTIME=" + $regName)
    $reg = "NOT_RUN"; try { $reg = $xl.Run('CFReg_RunAll') } catch { $reg = "ERR:" + $_.Exception.Message }
    Log ("CFReg_RunAll=" + $reg)

    # 3) fast module regression
    $fast = "NOT_RUN"; try { $fast = $xl.Run('CF_FastModuleRegression') } catch { $fast = "ERR:" + $_.Exception.Message }
    Log ("CF_FastModuleRegression=" + $fast)

    # 4) phase3 unit tests (Sub -> read PHASE3_TESTS!H1)
    $p3run = "NOT_RUN"; try { $xl.Run('CF_Phase3_RunUnitTests'); $p3run = "OK" } catch { $p3run = "ERR:" + $_.Exception.Message }
    Log ("CF_Phase3_RunUnitTests_call=" + $p3run)
    $p3 = "NOT_READ"; try { $p3 = [string]$wb.Worksheets('PHASE3_TESTS').Range('H1').Value2 } catch { $p3 = "ERR:" + $_.Exception.Message }
    Log ("PHASE3_TESTS_H1=" + $p3)

    # remove the runtime-only regression module before saving (keep WIP clean)
    $er2 = $null; foreach ($c in $proj.VBComponents) { if ($c.Name -eq $regName) { $er2 = $c; break } }
    if ($null -ne $er2) { $proj.VBComponents.Remove($er2); Log ("REMOVED_RUNTIME=" + $regName) }

    # verdict
    $stOk   = ($st -eq 'PASS')
    $regOk  = (("$reg" -match '^REG_ALL_PASS') -and ("$reg" -match 'FAIL=0'))
    $fastOk = ("$fast" -match '^PASS\|')
    $p3Ok   = $false
    if ("$p3" -match 'UNIT TESTS:\s*(\d+)/(\d+)\s*PASS') { $p3Ok = ($matches[1] -eq $matches[2]) -and ([int]$matches[1] -gt 0) }
    Log ("VERDICT stOk=" + $stOk + " regOk=" + $regOk + " fastOk=" + $fastOk + " p3Ok=" + $p3Ok)

    if ($stOk -and $regOk -and $fastOk -and $p3Ok) {
        $wb.Save()
        Log "SAVED_WIP"
        $wb.Close($false)
        Copy-Item -Path $wip -Destination $cpFix -Force
        Log ("CP1b_SAVED sha256=" + (Get-FileHash $cpFix -Algorithm SHA256).Hash)
    } else {
        $wb.Close($false)
        Log "NOT_SAVED_VERIFY_FAILED"
    }
} catch {
    Log ("COM_FATAL=" + $_.Exception.Message)
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
