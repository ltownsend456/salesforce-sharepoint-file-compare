<#
.SYNOPSIS
    Shows where PowerShell looks for modules and whether PnP.PowerShell exists in those paths.
#>
Write-Host "PSModulePath (where PowerShell looks for modules):" -ForegroundColor Cyan
$bases = $env:PSModulePath -split ';'
$bases | ForEach-Object { Write-Host "  $_" }
Write-Host ""
Write-Host "Checking each path for PnP.PowerShell:" -ForegroundColor Cyan
foreach ($base in $bases) {
    $p = Join-Path $base "PnP.PowerShell"
    $exists = Test-Path $p
    Write-Host "  $p : $(if ($exists) { 'FOUND' } else { 'not found' })" -ForegroundColor $(if ($exists) { 'Green' } else { 'Gray' })
}
