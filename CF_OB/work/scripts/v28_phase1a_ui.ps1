$ErrorActionPreference = 'Stop'
$wip     = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_WIP.xlsm'
$cp1     = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.8_CP1_UI.xlsm'
$expDir  = 'g:\SAP\CF_OB\work\scratch\vba_v28'
$outLog  = 'g:\SAP\CF_OB\work\REPORTS\v28_phase1a_ui.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

# Modules to re-import for phase 1a (standard modules only)
$reimport = @('modCF_InputLayout')

$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    $wb = $xl.Workbooks.Open($wip, 0, $false)
    $proj = $wb.VBProject
    Log ("OPENED=" + $wb.Name + " components=" + $proj.VBComponents.Count)

    foreach ($name in $reimport) {
        # remove existing standard module then import edited version
        $existing = $null
        foreach ($c in $proj.VBComponents) { if ($c.Name -eq $name) { $existing = $c; break } }
        if ($existing -ne $null) {
            if ($existing.Type -ne 1) { Log ("SKIP_NONSTD=" + $name + " type=" + $existing.Type); continue }
            $proj.VBComponents.Remove($existing)
            Log ("REMOVED=" + $name)
        }
        $path = Join-Path $expDir ($name + '.bas')
        $proj.VBComponents.Import($path) | Out-Null
        Log ("IMPORTED=" + $name)
    }
    Log ("COMPONENTS_AFTER=" + $proj.VBComponents.Count)

    # compile proof (whole project compiles before any proc runs)
    $st = "NOT_RUN"; try { $st = $xl.Run('CFAlloc_SelfTest') } catch { $st = "ERR:" + $_.Exception.Message }
    Log ("COMPILE_PROOF_CFAlloc_SelfTest=" + $st)

    # apply the layout (Workbook_Open is suppressed by EnableEvents=false, so call directly)
    $ap = "NOT_RUN"; try { $xl.Run('CF_ApplyInputLayout'); $ap = "OK" } catch { $ap = "ERR:" + $_.Exception.Message }
    Log ("CF_ApplyInputLayout=" + $ap)

    # enumerate INPUT buttons: name | visible | OnAction | text
    $wsIn = $wb.Worksheets('INPUT')
    # Required buttons (must exist). btnCF_LoadPosting is optional (no physical shape in V2.7 base).
    $expected = @{
        'btnCF_LoadWH'      = @($true,  'CF_LoadWH')
        'btnCF_CheckActual' = @($true,  'CF_CheckAndLoadPosting')
        'btnCF_ExecuteToCF' = @($true,  'CF_ExecutePlanToCFHandoff')
        'btnCF_Reset'       = @($true,  'CF_ResetUI')
        'btnCF_CheckSAP'    = @($false, 'CF_CheckSAP')
        'btnCF_OpenLog'     = @($false, 'CF_OpenLog')
        'btnCF_MapPOST'     = @($false, 'CF_MapSAP_POST')
    }
    $allOk = $true
    foreach ($sh in $wsIn.Shapes) {
        $oa=''; try { $oa = $sh.OnAction } catch {}
        $vis = ($sh.Visible -eq -1)
        $tx=''; try { $tx = $sh.TextFrame2.TextRange.Text } catch {}
        Log ("SHAPE|" + $sh.Name + "|vis=" + $vis + "|OnAction=" + $oa + "|text=" + ($tx -replace "`r`n"," "))
    }
    foreach ($k in $expected.Keys) {
        $want = $expected[$k]
        $found = $null
        foreach ($sh in $wsIn.Shapes) { if ($sh.Name -eq $k) { $found = $sh; break } }
        if ($found -eq $null) { Log ("MISSING_BUTTON=" + $k); $allOk=$false; continue }
        $vis = ($found.Visible -eq -1)
        $oa=''; try { $oa = $found.OnAction } catch {}
        $oaOk = $oa -match ([regex]::Escape($want[1]) + '\s*$')
        $visOk = ($vis -eq $want[0])
        if (-not ($oaOk -and $visOk)) { Log ("BUTTON_MISMATCH=" + $k + " wantVis=" + $want[0] + " gotVis=" + $vis + " wantMacro=" + $want[1] + " gotOnAction=" + $oa); $allOk=$false }
    }
    Log ("BUTTON_CONTRACT=" + $(if($allOk){'PASS'}else{'FAIL'}))

    if ($allOk -and $st -eq 'PASS') {
        $wb.Save()
        Log "SAVED_WIP"
        $wb.Close($false)
        Copy-Item -Path $wip -Destination $cp1 -Force
        Log ("CP1_SAVED sha256=" + (Get-FileHash $cp1 -Algorithm SHA256).Hash)
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
