# Compare core entry point functions between V1.7.1 (FIX29) and V1.8.3
$v171File = "g:\SAP\fix29_exported\MIGO311_RunPost_V24_ErrorGuard.bas"
$v183File = "g:\SAP\v183_audit_exported\MIGO311_RunPost_V24_ErrorGuard.bas"

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  COMPARING CORE FLOW: V1.7.1 vs V1.8.3" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

$macros = @(
    "SelectSourceFolder_NoOCR",
    "CheckFiles_NoOCR",
    "SAP_FillHeadTextOnly",
    "SAP_AttachCheckedFiles"
)

foreach ($m in $macros) {
    $f171 = Select-String -Path $v171File -Pattern "Sub\s+$m|Function\s+$m"
    $f183 = Select-String -Path $v183File -Pattern "Sub\s+$m|Function\s+$m"
    if ($f171 -and $f183) {
        Write-Host "[OK] Entry point '$m' exists in BOTH V1.7.1 and V1.8.3" -ForegroundColor Green
    } else {
        Write-Host "[WARN] Entry point '$m' missing!" -ForegroundColor Red
    }
}
