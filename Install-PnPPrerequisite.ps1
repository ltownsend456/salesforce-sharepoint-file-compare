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
if ($mod) {
    Write-Host "SUCCESS. PnP.PowerShell is installed (Version $($mod.Version))." -ForegroundColor Green
    Write-Host "You can now run: .\Export-SharePointFileList.ps1 -SiteUrl '...' -LibraryName 'Shared Documents' -OutputPath '.\SharePoint_FileList.csv'" -ForegroundColor Green
} else {
    Write-Host "Module not found after install. Try closing this terminal, opening a new PowerShell window, and run: Get-Module -ListAvailable PnP.PowerShell" -ForegroundColor Yellow
    Write-Host "If it still doesn't appear, the install failed (see any error from Step 2 above)." -ForegroundColor Red
    exit 1
}
