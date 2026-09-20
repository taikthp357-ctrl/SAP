$ErrorActionPreference = 'Stop'
$wbPath  = 'g:\SAP\CF_OB\work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm'
$outLog  = 'g:\SAP\CF_OB\work\REPORTS\v28_inventory.log'
$refLog  = 'g:\SAP\CF_OB\work\REPORTS\v28_posting_refs.log'
$preExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$scriptStart = Get-Date
$L = [System.Collections.ArrayList]::new()
$R = [System.Collections.ArrayList]::new()
function Log($m){ [void]$L.Add($m) }
function Ref($m){ [void]$R.Add($m) }

$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.EnableEvents=$false; $xl.AutomationSecurity=1
try {
    $wb = $xl.Workbooks.Open($wbPath, 0, $true)  # ReadOnly
    Log ("OPENED=" + $wb.Name)

    # --- Sheets: name, visibility, used range ---
    Log "=== SHEETS ==="
    foreach ($ws in $wb.Worksheets) {
        $vis = switch ($ws.Visible) { -1 {'Visible'} 0 {'Hidden'} 2 {'VeryHidden'} default {$ws.Visible} }
        $ur = $ws.UsedRange
        Log ("SHEET|" + $ws.Name + "|" + $vis + "|rows=" + $ur.Rows.Count + "|cols=" + $ur.Columns.Count)
    }

    # --- VBA components: name, type, line count ---
    Log "=== VBA COMPONENTS ==="
    $proj = $wb.VBProject
    Log ("VBA_COMPONENT_COUNT=" + $proj.VBComponents.Count)
    # type map: 1=StdModule 2=ClassModule 3=MSForm 100=Document
    foreach ($c in ($proj.VBComponents | Sort-Object Name)) {
        $cm = $c.CodeModule
        $lines = $cm.CountOfLines
        Log ("COMP|" + $c.Name + "|type=" + $c.Type + "|lines=" + $lines)
    }

    # --- POSTING reference scan across ALL components ---
    Ref "component|line|text"
    $patterns = @('POSTING')  # broad catch; we tag the variant per line
    foreach ($c in ($proj.VBComponents | Sort-Object Name)) {
        $cm = $c.CodeModule
        $n = $cm.CountOfLines
        if ($n -lt 1) { continue }
        $txt = $cm.Lines(1, $n)
        $arr = $txt -split "`r`n|`n"
        for ($i=0; $i -lt $arr.Count; $i++) {
            if ($arr[$i] -match 'POSTING') {
                $variant = 'OTHER'
                if     ($arr[$i] -match '"POSTING_RAW"')     { $variant='RAW' }
                elseif ($arr[$i] -match '"POSTING_CACHE"')   { $variant='CACHE' }
                elseif ($arr[$i] -match '"POSTING_SUMMARY"') { $variant='SUMMARY' }
                elseif ($arr[$i] -match '"POSTING_UI"')      { $variant='UI' }
                elseif ($arr[$i] -match '"POSTING"')         { $variant='SHEET_POSTING' }
                Ref ($c.Name + "|" + ($i+1) + "|" + $variant + "|" + $arr[$i].Trim())
            }
        }
    }

    # --- INPUT shapes with OnAction ---
    Log "=== INPUT SHAPES ==="
    try {
        $wsIn = $wb.Worksheets('INPUT')
        foreach ($sh in $wsIn.Shapes) {
            $oa = ''
            try { $oa = $sh.OnAction } catch {}
            $tx = ''
            try { $tx = $sh.TextFrame2.TextRange.Text } catch { try { $tx = $sh.TextFrame.Characters.Text } catch {} }
            Log ("SHAPE|" + $sh.Name + "|OnAction=" + $oa + "|Text=" + ($tx -replace "`r`n"," ") )
        }
    } catch { Log ("INPUT_SHAPES_ERR=" + $_.Exception.Message) }

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
    Set-Content -Path $refLog -Value ($R -join "`r`n") -Encoding UTF8
    # summary counts to stdout
    $refByComp = $R | Select-Object -Skip 1 | ForEach-Object { ($_ -split '\|')[0] } | Group-Object | Sort-Object Count -Descending
    Write-Output ("POSTING_REF_TOTAL=" + ($R.Count - 1))
    Write-Output "POSTING_REF_BY_COMPONENT:"
    $refByComp | ForEach-Object { Write-Output ("  " + $_.Name + " = " + $_.Count) }
}
