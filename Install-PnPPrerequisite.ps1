<#
.SYNONOPSIS
    One-time setup: trust PSGallery and install PnP.PowerShell so Export-SharePointFileList.ps1 can run.
#>

$ErrorActionPreference = "Stop"

Write-Host "Step 1: Setting PSGallery to Trusted (so install won't prompt)..." -ForegroundColor Cyan
try {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
} catch {
    Write-Host "Could not trust PSGallery (you may need to run PowerShell as Administrator once). Continuing - you may see a prompt; choose [A] Yes to All." -ForegroundColor Yellow
}

Write-Host "Step 2: Installing PnP.PowerShell for CurrentUser..." -ForegroundColor Cyan
Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber

Write-Host "Step 3: Verifying..." -ForegroundColor Cyan
$mod = Get-Module -ListAvailable -Name PnP.PowerShell
if ($mod) {
    Write-Host "SUCCESS. PnP.PowerShell is installed (Version $($mod.Version))." -ForegroundColor Green
    Write-Host "You can now run: .\Export-SharePointFileList.ps1 -SiteUrl '...' -LibraryName 'Shared Documents' -OutputPath '.\SharePoint_FileList.csv'" -ForegroundColor Green
} else {
    Write-Host "Install may have failed. Check errors above." -ForegroundColor Red
    exit 1
}
