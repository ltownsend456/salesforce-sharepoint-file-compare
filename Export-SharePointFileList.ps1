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

# Ensure PnP is available (install once: Install-Module -Name PnP.PowerShell -Scope CurrentUser)
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "PnP.PowerShell is not installed. Run this command first (if prompted about PSGallery, choose [A] Yes to All):" -ForegroundColor Yellow
    Write-Host "  Install-Module -Name PnP.PowerShell -Scope CurrentUser" -ForegroundColor Cyan
    Write-Error "PnP.PowerShell is required. Install it, then run this script again."
    exit 1
}

Write-Host "Connecting to SharePoint: $SiteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

try {
    $list = Get-PnPList -Identity $LibraryName -ErrorAction Stop
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
