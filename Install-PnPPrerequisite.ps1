<#
.SYNONOPSIS
    One-time setup: trust PSGallery and install PnP.PowerShell so Export-SharePointFileList.ps1 can run.
#>

$ErrorActionPreference = "Stop"

# PowerShell Gallery requires TLS 1.2; many VDIs use older defaults
Write-Host "Step 0: Enabling TLS 1.2 for PowerShell Gallery..." -ForegroundColor Cyan
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host "Step 1: Setting PSGallery to Trusted (so install won't prompt)..." -ForegroundColor Cyan
try {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
} catch {
    Write-Host "Could not trust PSGallery (you may need to run PowerShell as Administrator once). Continuing." -ForegroundColor Yellow
}

Write-Host "Step 2: Installing PnP.PowerShell for CurrentUser (this can take 1-2 minutes)..." -ForegroundColor Cyan
try {
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
    Write-Host "Install-Module completed." -ForegroundColor Gray
} catch {
    Write-Host "ERROR from Install-Module:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.Exception.InnerException) { Write-Host $_.Exception.InnerException.Message -ForegroundColor Red }
    Write-Host ""
    Write-Host "Common causes on VDI: proxy/firewall blocking psgallery.com, or missing .NET. Try from a browser: https://www.powershellgallery.com" -ForegroundColor Yellow
    exit 1
}

Write-Host "Step 3: Verifying..." -ForegroundColor Cyan
$mod = Get-Module -ListAvailable -Name PnP.PowerShell
if (-not $mod) {
    # On VDIs the module path may not refresh; try loading from known CurrentUser paths
    $paths = @(
        (Join-Path $HOME "Documents\WindowsPowerShell\Modules\PnP.PowerShell"),
        (Join-Path $HOME "Documents\PowerShell\Modules\PnP.PowerShell")
    )
    foreach ($p in $paths) {
        if (Test-Path $p) {
            Write-Host "Found module at: $p" -ForegroundColor Gray
            Import-Module $p -Force -ErrorAction SilentlyContinue
            $mod = Get-Module -Name PnP.PowerShell
            if ($mod) { break }
        }
    }
}
if ($mod) {
    Write-Host "SUCCESS. PnP.PowerShell is available (Version $($mod.Version))." -ForegroundColor Green
    Write-Host "Run the export script in THIS SAME terminal: .\Export-SharePointFileList.ps1 -SiteUrl '...' -LibraryName 'Shared Documents' -OutputPath '.\SharePoint_FileList.csv'" -ForegroundColor Green
} else {
    Write-Host "Module not found. Checking where PowerShell looks for modules..." -ForegroundColor Yellow
    Write-Host "PSModulePath: $($env:PSModulePath -replace ';', \"`n  \")" -ForegroundColor Gray
    Write-Host "If your Documents folder is redirected (e.g. OneDrive), install may have gone there. Run the export script anyway - it will try to load PnP from known paths." -ForegroundColor Yellow
    exit 1
}
