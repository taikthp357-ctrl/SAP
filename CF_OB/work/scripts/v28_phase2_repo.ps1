$ErrorActionPreference = 'Stop'
$wip     = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_WIP.xlsm'
$cp2     = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP2_REPO.xlsm'
$expDir  = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$regBas  = 'g:\SAP\CF_OB\work\scratch\modCFAllocRegression.bas'
$outLog  = 'g:\SAP\CF_OB\work\REPORTS\v28_phase2_repo.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

# Phase 2: introduce modSourceStockRepo (NEW) + redirect the 5 POSTING data-binds
# (ModuleSAP reader+writer, modSAPReadOnlyProbe x2, modCF_Posting CHECK reader) and
# the name-keyed validator (ModuleHelper) through the repo. Backing = SHEET so
# behaviour is byte-identical to V2.7. modSourceStockRepo imported first so its
# Public procs exist before dependents compile.
$reimport = @('modSourceStockRepo','ModuleSAP','modSAPReadOnlyProbe','modCF_Posting','ModuleHelper')

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
        } else {
            Log ("NEW_MODULE=" + $name)
        }
        $path = Join-Path $expDir ($name + '.bas')
        if (-not (Test-Path $path)) { Log ("MISSING_BAS=" + $path); continue }
        $proj.VBComponents.Import($path) | Out-Null
        Log ("IMPORTED=" + $name)
    }
    Log ("COMPONENTS_AFTER=" + $proj.VBComponents.Count)

    # 1) whole-project compile proof (compiles every redirected dependent)
    $st = "NOT_RUN"; try { $st = $xl.Run('CFAlloc_SelfTest') } catch { $st = "ERR:" + $_.Exception.Message }
    Log ("COMPILE_PROOF_CFAlloc_SelfTest=" + $st)

    # 2) repo compile probe + offline coordinate round-trip (synthetic rows, backing=SHEET)
    $cp = "NOT_RUN"; try { $xl.Run('SSR_CompileProbe'); $cp = "OK" } catch { $cp = "ERR:" + $_.Exception.Message }
    Log ("SSR_CompileProbe=" + $cp)
    $ssr = "NOT_RUN"; try { $ssr = $xl.Run('SSR_SelfTest') } catch { $ssr = "ERR:" + $_.Exception.Message }
    Log ("SSR_SelfTest=" + $ssr)

    # 3) proven-green regression: runtime-import the external allocation suite first
    $regName = 'modCFAllocRegression'
    $er = $null; foreach ($c in $proj.VBComponents) { if ($c.Name -eq $regName) { $er = $c; break } }
    if ($null -ne $er) { $proj.VBComponents.Remove($er) }
    $proj.VBComponents.Import($regBas) | Out-Null
    Log ("IMPORTED_RUNTIME=" + $regName)
    $reg = "NOT_RUN"; try { $reg = $xl.Run('CFReg_RunAll') } catch { $reg = "ERR:" + $_.Exception.Message }
    Log ("CFReg_RunAll=" + $reg)
    $fast = "NOT_RUN"; try { $fast = $xl.Run('CF_FastModuleRegression') } catch { $fast = "ERR:" + $_.Exception.Message }
    Log ("CF_FastModuleRegression=" + $fast)
    $p3run = "NOT_RUN"; try { $xl.Run('CF_Phase3_RunUnitTests'); $p3run = "OK" } catch { $p3run = "ERR:" + $_.Exception.Message }
    Log ("CF_Phase3_RunUnitTests_call=" + $p3run)
    $p3 = "NOT_READ"; try { $p3 = [string]$wb.Worksheets('PHASE3_TESTS').Range('H1').Value2 } catch { $p3 = "ERR:" + $_.Exception.Message }
    Log ("PHASE3_TESTS_H1=" + $p3)

    # remove the runtime-only regression module before saving (keep WIP clean)
    $er2 = $null; foreach ($c in $proj.VBComponents) { if ($c.Name -eq $regName) { $er2 = $c; break } }
    if ($null -ne $er2) { $proj.VBComponents.Remove($er2); Log ("REMOVED_RUNTIME=" + $regName) }

    # verdict: compile PASS + SSR probe/round-trip PASS + full regression green
    $stOk   = ($st -eq 'PASS')
    $cpOk   = ($cp -eq 'OK')
    $ssrOk  = ($ssr -eq 'PASS')
    $regOk  = (("$reg" -match '^REG_ALL_PASS') -and ("$reg" -match 'FAIL=0'))
    $fastOk = ("$fast" -match '^PASS\|')
    $p3Ok   = $false
    if ("$p3" -match 'UNIT TESTS:\s*(\d+)/(\d+)\s*PASS') { $p3Ok = ($matches[1] -eq $matches[2]) -and ([int]$matches[1] -gt 0) }
    Log ("VERDICT stOk=" + $stOk + " cpOk=" + $cpOk + " ssrOk=" + $ssrOk + " regOk=" + $regOk + " fastOk=" + $fastOk + " p3Ok=" + $p3Ok)

    if ($stOk -and $cpOk -and $ssrOk -and $regOk -and $fastOk -and $p3Ok) {
        $wb.Save()
        Log "SAVED_WIP"
        $wb.Close($false)
        Copy-Item -Path $wip -Destination $cp2 -Force
        Log ("CP2_SAVED sha256=" + (Get-FileHash $cp2 -Algorithm SHA256).Hash)
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
