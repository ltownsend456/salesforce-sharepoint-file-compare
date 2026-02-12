<#
.SYNONOPSIS
    One-time setup: trust PSGallery and install PnP so Export-SharePointFileList.ps1 can run.
    On Windows PowerShell 5.1 (VDI), installs legacy SharePointPnPPowerShellOnline. On PowerShell 7+, installs PnP.PowerShell.
#>

$ErrorActionPreference = "Stop"

$psMajor = $PSVersionTable.PSVersion.Major
Write-Host "Detected PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Gray

# PowerShell Gallery requires TLS 1.2; many VDIs use older defaults
Write-Host "Step 0: Enabling TLS 1.2 for PowerShell Gallery..." -ForegroundColor Cyan
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host "Step 1: Setting PSGallery to Trusted (so install won't prompt)..." -ForegroundColor Cyan
try {
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
} catch {
    Write-Host "Could not trust PSGallery (you may need to run PowerShell as Administrator once). Continuing." -ForegroundColor Yellow
}

if ($psMajor -lt 7) {
    Write-Host "Step 2: Installing SharePointPnPPowerShellOnline (legacy, for Windows PowerShell 5.1)..." -ForegroundColor Cyan
    $moduleName = "SharePointPnPPowerShellOnline"
    try {
        Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber
        Write-Host "Install-Module completed." -ForegroundColor Gray
    } catch {
        Write-Host "ERROR from Install-Module:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "Alternative: Install PowerShell 7 (pwsh) and run the script with: pwsh -File .\Export-SharePointFileList.ps1 ..." -ForegroundColor Yellow
        exit 1
    }
} else {
    Write-Host "Step 2: Installing PnP.PowerShell for CurrentUser (this can take 1-2 minutes)..." -ForegroundColor Cyan
    $moduleName = "PnP.PowerShell"
    try {
        Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
        Write-Host "Install-Module completed." -ForegroundColor Gray
    } catch {
        Write-Host "ERROR from Install-Module:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }
}

Write-Host "Step 3: Verifying..." -ForegroundColor Cyan
$mod = Get-Module -ListAvailable -Name $moduleName | Select-Object -First 1
if (-not $mod) {
    $modulePaths = $env:PSModulePath -split ';'
    foreach ($base in $modulePaths) {
        $p = Join-Path $base $moduleName
        if (Test-Path $p) {
            Write-Host "Found module at: $p" -ForegroundColor Gray
            Import-Module $p -Force -ErrorAction SilentlyContinue
            $mod = Get-Module -Name $moduleName
            if ($mod) { break }
        }
    }
}
if ($mod) {
    Write-Host "SUCCESS. $moduleName is available (Version $($mod.Version))." -ForegroundColor Green
    Write-Host "Run: .\Export-SharePointFileList.ps1 -SiteUrl '...' -LibraryName 'Shared Documents' -OutputPath '.\SharePoint_FileList.csv'" -ForegroundColor Green
} else {
    Write-Host "Module not found. Run the export script anyway; it will try to load the module." -ForegroundColor Yellow
    exit 1
}
