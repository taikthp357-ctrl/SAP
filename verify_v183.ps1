# Verify V1.8.3 workbook button assignments & module integrity
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wbPath = "g:\SAP\v183_output\260803_V1.8.3_BBBG_VAT_SMART_GOS_ATT_REBUILT.xlsm"
    $wb = $excel.Workbooks.Open($wbPath)
    
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "VERIFYING V1.8.3 WORKBOOK INTEGRITY" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    
    # 1. Verify Components
    $vbaProj = $wb.VBProject
    Write-Host "Total VBA Components: $($vbaProj.VBComponents.Count)" -ForegroundColor Green
    
    $migoComp = $null
    foreach ($comp in $vbaProj.VBComponents) {
        if ($comp.Name -eq "MIGO311_RunPost_V24_ErrorGuard") {
            $migoComp = $comp
            break
        }
    }
    
    if ($null -ne $migoComp) {
        Write-Host "[OK] Module 'MIGO311_RunPost_V24_ErrorGuard' is correctly named and loaded!" -ForegroundColor Green
    } else {
        Write-Host "[FAIL] Module 'MIGO311_RunPost_V24_ErrorGuard' NOT FOUND!" -ForegroundColor Red
    }
    
    # 2. Verify Button Assignments
    $expectedButtons = @{
        "btnChooseFolder" = "SelectSourceFolder_NoOCR"
        "btnCheck"        = "CheckFiles_NoOCR"
        "btnHeadText"     = "SAP_FillHeadTextOnly"
        "btnAttach"       = "SAP_AttachCheckedFiles"
    }
    
    $buttonResults = @{}
    foreach ($ws in $wb.Worksheets) {
        foreach ($shape in $ws.Shapes) {
            if ($expectedButtons.ContainsKey($shape.Name)) {
                $buttonResults[$shape.Name] = $shape.OnAction
            }
        }
    }
    
    foreach ($btnName in $expectedButtons.Keys) {
        $expectedMacro = $expectedButtons[$btnName]
        if ($buttonResults.ContainsKey($btnName)) {
            $actualMacro = $buttonResults[$btnName]
            if ($actualMacro -match $expectedMacro) {
                Write-Host "  [OK] $btnName -> $actualMacro" -ForegroundColor Green
            } else {
                Write-Host "  [WARN] $btnName -> $actualMacro (Expected: $expectedMacro)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "  [WARN] $btnName - Button shape not found on sheets" -ForegroundColor Yellow
        }
    }
    
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
