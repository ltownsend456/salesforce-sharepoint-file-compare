<#
.SYNOPSIS
    Exports all files from a SharePoint Online document library to a CSV for comparison with Salesforce.
.DESCRIPTION
    Uses PnP PowerShell to connect to SharePoint and list all files (with metadata).
    Run this first, then use compare_sf_sharepoint.py with the Salesforce CSV and this output CSV.
.NOTES
    Requires: Install-Module -Name PnP.PowerShell
    Docs: https://learn.microsoft.com/sharepoint/dev/solution-guidance/pnp-provisioning-engine-and-the-core-library
#>

param(
    [Parameter(Mandatory = $true, HelpMessage = "SharePoint site URL, e.g. https://tenant.sharepoint.com/sites/YourSite")]
    [string]$SiteUrl,

    [Parameter(HelpMessage = "Document library name (e.g. 'Documents', 'Shared Documents', or your library name)")]
    [string]$LibraryName = "Documents",

    [Parameter(HelpMessage = "Output CSV path")]
    [string]$OutputPath = ".\SharePoint_FileList.csv",

    [Parameter(HelpMessage = "Include files in subfolders")]
    [switch]$Recurse
)

# Ensure PnP is available. On PS 5.1 we use SharePointPnPPowerShellOnline; on PS 7+ we use PnP.PowerShell.
function Import-PnPModuleIfAvailable {
    param([string]$modName, [ref]$lastError)
    if (Get-Module -Name $modName) { return $true }
    $avail = Get-Module -ListAvailable -Name $modName | Select-Object -First 1
    if (-not $avail) { return $false }
    try {
        Import-Module -Name $modName -Force -ErrorAction Stop
        return $true
    }
    catch {
        $lastError.Value = $_
        return $false
    }
}

$pnpModuleList = if ($PSVersionTable.PSVersion.Major -lt 7) {
    @("SharePointPnPPowerShellOnline", "PnP.PowerShell")
} else {
    @("PnP.PowerShell", "SharePointPnPPowerShellOnline")
}
$pnpLoaded = $false
$lastLoadError = $null

foreach ($modName in $pnpModuleList) {
    if (Import-PnPModuleIfAvailable -modName $modName -lastError ([ref]$lastLoadError)) {
        $pnpLoaded = $true
        break
    }
}

if (-not $pnpLoaded) {
    $modulePaths = $env:PSModulePath -split ';'
    foreach ($base in $modulePaths) {
        foreach ($modName in $pnpModuleList) {
            $modulePath = Join-Path $base $modName
            if (Test-Path $modulePath) {
                $loaded = $false
                try {
                    $env:PSModulePath = "$base;$env:PSModulePath"
                    Import-Module -Name $modName -Force -ErrorAction Stop
                    $loaded = $true
                }
                catch {
                    $lastLoadError = $_
                }
                if ($loaded) {
                    $pnpLoaded = $true
                    break
                }
            }
        }
        if ($pnpLoaded) {
            break
        }
    }
}

if (-not $pnpLoaded) {
    if ($lastLoadError -and $PSVersionTable.PSVersion.Major -lt 7) {
        Write-Host "On Windows PowerShell 5.1, PnP.PowerShell 3.x requires PowerShell 7.4.6+." -ForegroundColor Yellow
        Write-Host "Run .\Install-PnPPrerequisite.ps1 to install the legacy module (SharePointPnPPowerShellOnline) for 5.1." -ForegroundColor Yellow
    }
    if ($lastLoadError) {
        Write-Host "Load error: $($lastLoadError.Exception.Message)" -ForegroundColor Red
    }
    Write-Error "PnP module is required. Run: .\Install-PnPPrerequisite.ps1"
    exit 1
}

# Legacy module (PS 5.1) uses -UseWebLogin; modern PnP uses -Interactive
$connectParams = @{ Url = $SiteUrl }
$cmd = Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue
if ($cmd.Parameters['Interactive']) { $connectParams['Interactive'] = $true }
elseif ($cmd.Parameters['UseWebLogin']) { $connectParams['UseWebLogin'] = $true }
else { $connectParams['Interactive'] = $true }

Write-Host "Connecting to SharePoint: $SiteUrl" -ForegroundColor Cyan
Connect-PnPOnline @connectParams

try {
    $null = Get-PnPList -Identity $LibraryName -ErrorAction Stop
}
catch {
    Write-Host "List '$LibraryName' not found. Available lists:" -ForegroundColor Yellow
    Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 } | ForEach-Object { Write-Host "  - $($_.Title)" }
    exit 1
}

Write-Host "Retrieving files from library '$LibraryName' (this may take a while for large libraries)..." -ForegroundColor Cyan

# Get all items; filter to files (have FileLeafRef and not a folder)
$query = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where></Query></View>"
$items = Get-PnPListItem -List $LibraryName -Query $query -PageSize 500

$results = @()
foreach ($item in $items) {
    $fileRef = $item["FileRef"]
    $fileName = $item["FileLeafRef"]
    if (-not $fileName) { continue }
    $authorVal = $item["Author"]
    $editorVal = $item["Editor"]
    $authorStr = if ($null -ne $authorVal -and $authorVal.LookupValue) { $authorVal.LookupValue } else { "" }
    $editorStr = if ($null -ne $editorVal -and $editorVal.LookupValue) { $editorVal.LookupValue } else { "" }
    $results += [PSCustomObject]@{
        Name              = $fileName
        FileRef           = $fileRef
        FileExtension     = [System.IO.Path]::GetExtension($fileName)
        Length            = $item["File_x0020_Size"]
        TimeCreated       = $item["Created"]
        TimeLastModified  = $item["Modified"]
        Author            = $authorStr
        Editor            = $editorStr
    }
}

Disconnect-PnPOnline

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Exported $($results.Count) files to: $OutputPath" -ForegroundColor Green
Write-Host "Next step: Run the Python comparison script with this CSV and your Salesforce CSV."
