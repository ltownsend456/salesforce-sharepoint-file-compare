<#
.SYNOPSIS
    Shows where PowerShell looks for modules and whether PnP.PowerShell exists in those paths.
#>
Write-Host "PSModulePath (where PowerShell looks for modules):" -ForegroundColor Cyan
$env:PSModulePath -split ';' | ForEach-Object { Write-Host "  $_" }
Write-Host ""
Write-Host "Checking known CurrentUser paths for PnP.PowerShell:" -ForegroundColor Cyan
$paths = @(
    (Join-Path $env:USERPROFILE "Documents\WindowsPowerShell\Modules\PnP.PowerShell"),
    (Join-Path $env:USERPROFILE "Documents\PowerShell\Modules\PnP.PowerShell")
)
foreach ($p in $paths) {
    $exists = Test-Path $p
    Write-Host "  $p : $(if ($exists) { 'FOUND' } else { 'not found' })" -ForegroundColor $(if ($exists) { 'Green' } else { 'Gray' })
}
