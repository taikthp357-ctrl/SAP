# v28_cp8_fix.ps1 — Fix CONFIG column rewrites that failed in scope-based approach
# Use direct string replacement on known patterns instead
$ErrorActionPreference = 'Stop'
$vbaBand = 'g:\SAP\CF_OB\work\scratch\vba_cp8_band'

# CONFIG consumers: directly replace .Cells(ws.rows.count, 1) → 17, .Cells(r, 1) → 17, .Cells(r, 2) → 18
# Do this for all CONFIG modules, iterating through exact patterns

$configFixes = @(
    @{
        file = 'modCF_FastContract.bas'
        replacements = @(
            @{old='ws.Cells(ws.rows.count, 1).End(xlUp).row'; new='ws.Cells(ws.rows.count, 17).End(xlUp).row'},
            @{old='ws.Cells(r, 1).Value2'; new='ws.Cells(r, 17).Value2'},
            @{old='ws.Cells(r, 2).Value2'; new='ws.Cells(r, 18).Value2'}
        )
    },
    @{
        file = 'modCF_AutoRun_TwoPilot.bas'
        replacements = @(
            @{old='ws.Cells(ws.rows.count, 1).End(xlUp).row'; new='ws.Cells(ws.rows.count, 17).End(xlUp).row'},
            @{old='ws.Cells(r, 1).Value2'; new='ws.Cells(r, 17).Value2'},
            @{old='ws.Cells(r, 2).Value2'; new='ws.Cells(r, 18).Value2'}
        )
    },
    @{
        file = 'modCF_Executor.bas'
        replacements = @(
            @{old='ws.Cells(ws.rows.count, 1).End(xlUp).row'; new='ws.Cells(ws.rows.count, 17).End(xlUp).row'},
            @{old='ws.Cells(r, 1).Value2'; new='ws.Cells(r, 17).Value2'},
            @{old='ws.Cells(r, 2).Value2'; new='ws.Cells(r, 18).Value2'}
        )
    },
    @{
        file = 'modCodexWTRework.bas'
        replacements = @(
            @{old='ws.Cells(ws.rows.count, 1).End(xlUp).row'; new='ws.Cells(ws.rows.count, 17).End(xlUp).row'},
            @{old='ws.Cells(r, 1).Value2'; new='ws.Cells(r, 17).Value2'},
            @{old='ws.Cells(r, 2).Value2'; new='ws.Cells(r, 18).Value2'}
        )
    },
    @{
        file = 'modCF_ApplyUserHandoff.bas'
        replacements = @(
            @{old='ws.Cells(ws.rows.count, 1).End(xlUp).row'; new='ws.Cells(ws.rows.count, 17).End(xlUp).row'},
            @{old='ws.Cells(r, 1).Value2'; new='ws.Cells(r, 17).Value2'},
            @{old='ws.Cells(r, 2).Value2'; new='ws.Cells(r, 18).Value2'},
            @{old='ws.Cells(r, 3).Value2'; new='ws.Cells(r, 19).Value2'},
            @{old='ws.Cells(r, 4).Value2'; new='ws.Cells(r, 20).Value2'},
            @{old='ws.Cells(r, 5).Value2'; new='ws.Cells(r, 21).Value2'},
            @{old='ws.Cells(r, 6).Value2'; new='ws.Cells(r, 22).Value2'}
        )
    },
    @{
        file = 'modCF_ActualPicking.bas'
        replacements = @(
            @{old='ws.Cells(ws.rows.count, 1).End(xlUp).row'; new='ws.Cells(ws.rows.count, 17).End(xlUp).row'},
            @{old='ws.Cells(r, 1).Value2'; new='ws.Cells(r, 17).Value2'},
            @{old='ws.Cells(r, 2).Value2'; new='ws.Cells(r, 18).Value2'},
            @{old='ws.Cells(r, 3).Value2'; new='ws.Cells(r, 19).Value2'},
            @{old='ws.Cells(r, 4).Value2'; new='ws.Cells(r, 20).Value2'},
            @{old='ws.Cells(r, 5).Value2'; new='ws.Cells(r, 21).Value2'},
            @{old='ws.Cells(r, 6).Value2'; new='ws.Cells(r, 22).Value2'},
            @{old='ws.Cells(lastRow, 1).Value2'; new='ws.Cells(lastRow, 17).Value2'},
            @{old='ws.Cells(lastRow, 2).Value2'; new='ws.Cells(lastRow, 18).Value2'}
        )
    },
    @{
        file = 'modCF_PreLiveAudit.bas'
        replacements = @(
            @{old='ws.Cells(ws.rows.count, 1).End(xlUp).row'; new='ws.Cells(ws.rows.count, 17).End(xlUp).row'},
            @{old='ws.Cells(r, 1).Value2'; new='ws.Cells(r, 17).Value2'},
            @{old='ws.Cells(r, 2).Value2'; new='ws.Cells(r, 18).Value2'},
            @{old='ws.Cells(lastRow + 1, 1).Value2'; new='ws.Cells(lastRow + 1, 17).Value2'},
            @{old='ws.Cells(lastRow + 1, 2).Value2'; new='ws.Cells(lastRow + 1, 18).Value2'}
        )
    },
    @{
        file = 'modSAPRunGuard.bas'
        replacements = @(
            @{old='ws.Cells(ws.rows.count, 1).End(xlUp).row'; new='ws.Cells(ws.rows.count, 17).End(xlUp).row'},
            @{old='ws.Cells(r, 1).Value2'; new='ws.Cells(r, 17).Value2'},
            @{old='ws.Cells(r, 2).Value2'; new='ws.Cells(r, 18).Value2'},
            @{old='ws.Cells(lastRow, 1).Value2'; new='ws.Cells(lastRow, 17).Value2'},
            @{old='ws.Cells(lastRow, 2).Value2'; new='ws.Cells(lastRow, 18).Value2'},
            @{old='ws.Cells(lastRow, 3).Value2'; new='ws.Cells(lastRow, 19).Value2'},
            @{old='ws.Cells(lastRow, 4).Value2'; new='ws.Cells(lastRow, 20).Value2'},
            @{old='ws.Cells(lastRow, 5).Value2'; new='ws.Cells(lastRow, 21).Value2'},
            @{old='ws.Cells(lastRow, 6).Value2'; new='ws.Cells(lastRow, 22).Value2'}
        )
    },
    @{
        file = 'ModuleHelper.bas'
        replacements = @(
            @{old='ws.Cells(ws.rows.count, "A").End(xlUp).row'; new='ws.Cells(ws.rows.count, 17).End(xlUp).row'},
            @{old='ws.Range("A3:B"'; new='ws.Range("Q3:R"'},
            @{old='wsc.Cells(wsc.rows.count, 7).End(xlUp).row'; new='wsc.Cells(wsc.rows.count, 23).End(xlUp).row'},
            @{old='wsc.Range("G3:J"'; new='wsc.Range("W3:Z"'}
        )
    }
)

foreach ($fix in $configFixes) {
    $path = Join-Path $vbaBand $fix.file
    $content = Get-Content $path -Raw -Encoding UTF8

    foreach ($rep in $fix.replacements) {
        $content = $content.Replace($rep.old, $rep.new)
    }

    Set-Content $path -Value $content -Encoding UTF8 -NoNewline
    Write-Host "FIXED_CONFIG|$($fix.file)"
}

# SAP_ACTION_MAP fixes: modSAPActionMap and modCF_PreLiveAudit need direct cell replacements
$samFixes = @(
    @{
        file = 'modSAPActionMap.bas'
        replacements = @(
            @{old='ws.Cells(ws.rows.count, 1).End(xlUp).row'; new='ws.Cells(ws.rows.count, 27).End(xlUp).row'},
            @{old='ws.Cells(1, 1).Value2'; new='ws.Cells(1, 27).Value2'},
            @{old='ws.Cells(r, 1).Value2'; new='ws.Cells(r, 27).Value2'},
            @{old='ws.Cells(r, 5).Value2'; new='ws.Cells(r, 31).Value2'},
            @{old='ws.Cells(r, 6).Value2'; new='ws.Cells(r, 32).Value2'}
        )
    },
    @{
        file = 'modCF_PreLiveAudit.bas'
        replacements = @(
            @{old='Worksheets("SYSTEM").Cells(1, 1).Value2'; new='Worksheets("SYSTEM").Cells(1, 27).Value2'}
        )
    }
)

foreach ($fix in $samFixes) {
    $path = Join-Path $vbaBand $fix.file
    $content = Get-Content $path -Raw -Encoding UTF8

    foreach ($rep in $fix.replacements) {
        $content = $content.Replace($rep.old, $rep.new)
    }

    Set-Content $path -Value $content -Encoding UTF8 -NoNewline
    Write-Host "FIXED_SAM|$($fix.file)"
}

# SAP_PROBE_LOG fixes: SAPProbe_Log has ws.Cells(r, N) for N=1..17
$splPath = Join-Path $vbaBand 'modSAPReadOnlyProbe.bas'
$splContent = Get-Content $splPath -Raw -Encoding UTF8

# Direct replacements for SAPProbe_Log function (cols 1-17 → 37-53)
for ($col = 17; $col -ge 1; $col--) {
    $newCol = $col + 36
    $splContent = $splContent.Replace("ws.Cells(r, $col).Value2", "ws.Cells(r, $newCol).Value2")
}
$splContent = $splContent.Replace('ws.Cells(ws.rows.count, 1).End(xlUp).row', 'ws.Cells(ws.rows.count, 37).End(xlUp).row')

Set-Content $splPath -Value $splContent -Encoding UTF8 -NoNewline
Write-Host "FIXED_SPL|modSAPReadOnlyProbe.bas"

# modSAPRunGuard SAP_PROBE_LOG access
$srgPath = Join-Path $vbaBand 'modSAPRunGuard.bas'
$srgContent = Get-Content $srgPath -Raw -Encoding UTF8
$srgContent = $srgContent.Replace('ws.Cells(ws.rows.count, 1).End(xlUp).row', 'ws.Cells(ws.rows.count, 37).End(xlUp).row')
$srgContent = $srgContent.Replace('ws.Cells(r, 3).Value2', 'ws.Cells(r, 39).Value2')  # probe column
$srgContent = $srgContent.Replace('ws.Cells(r, 15).Value2', 'ws.Cells(r, 51).Value2')  # result column
Set-Content $srgPath -Value $srgContent -Encoding UTF8 -NoNewline
Write-Host "FIXED_SPL|modSAPRunGuard.bas"

Write-Host "ALL_FIXES_APPLIED"
