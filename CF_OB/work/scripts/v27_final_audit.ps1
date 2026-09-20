$ErrorActionPreference = 'Stop'
$wbPath  = 'g:\SAP\CF_OB\work\OUTPUT\260820_CF_OB_V2.7.xlsm'
$outLog  = 'g:\SAP\CF_OB\work\REPORTS\v27_final_audit.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m); Write-Output $m }

# --- Static ZIP/XML integrity (no Excel needed) ---
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($wbPath)
    $names = $zip.Entries | ForEach-Object { $_.FullName }
    Log ("ZIP_ENTRIES=" + $names.Count)
    Log ("HAS_VBAPROJECT=" + [bool]($names -contains 'xl/vbaProject.bin'))
    Log ("HAS_WORKBOOK_XML=" + [bool]($names -contains 'xl/workbook.xml'))
    Log ("HAS_CONTENT_TYPES=" + [bool]($names -contains '[Content_Types].xml'))
    # parse workbook.xml to count sheet elements
    $e = $zip.GetEntry('xl/workbook.xml')
    $sr = New-Object System.IO.StreamReader($e.Open())
    $xml = $sr.ReadToEnd(); $sr.Close()
    $sheetCount = ([regex]::Matches($xml, '<sheet ')).Count
    Log ("WORKBOOK_XML_SHEETS=" + $sheetCount)
    $hidden = ([regex]::Matches($xml, 'state="veryHidden"')).Count
    $hiddenS = ([regex]::Matches($xml, 'state="hidden"')).Count
    Log ("SHEET_STATE_VERYHIDDEN=" + $hidden + " HIDDEN=" + $hiddenS)
    $zip.Dispose()
    Log "XML_PARSE_OK"
} catch {
    Log ("XML_FATAL=" + $_.Exception.Message)
}

# --- Live compile proof via Excel COM ---
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    $wb = $xl.Workbooks.Open($wbPath, 0, $true)  # ReadOnly
    Log ("OPENED_OK=" + $wb.Name)
    $proj = $wb.VBProject
    Log ("VBA_COMPONENTS=" + $proj.VBComponents.Count)
    # entry-point presence checks
    $hasReset=$false; $hasAlloc=$false
    foreach ($c in $proj.VBComponents) {
        if ($c.Name -eq 'modCFResetUI') { $hasReset=$true }
        if ($c.Name -eq 'modCFGroupAllocator') { $hasAlloc=$true }
    }
    Log ("HAS_modCFResetUI=" + $hasReset)
    Log ("HAS_modCFGroupAllocator=" + $hasAlloc)
    # compile proof: run a self-test (VBA compiles whole project before running any proc)
    $st = "NOT_RUN"; try { $st = $xl.Run('CFAlloc_SelfTest') } catch { $st = "ERR:" + $_.Exception.Message }
    Log ("COMPILE_PROOF_CFAlloc_SelfTest=" + $st)
    $rs = "NOT_RUN"; try { $rs = $xl.Run('CFResetUI_SelfTest') } catch { $rs = "ERR:" + $_.Exception.Message }
    Log ("CFResetUI_SelfTest=" + $rs)
    # sheet count live
    Log ("LIVE_SHEETS=" + $wb.Worksheets.Count)
    # plan lock cells intact
    $wsPlan = $null; try { $wsPlan = $wb.Worksheets('PLAN') } catch {}
    if ($wsPlan -ne $null) {
        Log ("PLAN_AJ2=" + $wsPlan.Range('AJ2').Value2)
        Log ("PLAN_AT2=" + $wsPlan.Range('AT2').Value2)
    }
    $wb.Close($false)
    Log "CLOSED_NO_SAVE"
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
