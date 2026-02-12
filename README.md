# Salesforce vs SharePoint file comparison

Automate comparing files between **Salesforce** (CSV export) and **SharePoint** (your site). Find what’s missing on each side and produce a list of files to give your coworker (SharePoint has, Salesforce doesn’t).

## What you get

- **Only in SharePoint** → list (and metadata) of files to give your coworker (missing in Salesforce).
- **Only in Salesforce** → list of files you don’t have in SharePoint.
- **In both** → files present in both (with metadata from both CSVs).

## Prerequisites

1. **Python 3.9+** (no extra packages; uses only the standard library).
2. **PnP PowerShell** (for exporting the SharePoint file list):
   ```powershell
   Install-Module -Name PnP.PowerShell -Scope CurrentUser
   ```
   [Microsoft: PnP PowerShell](https://learn.microsoft.com/sharepoint/dev/solution-guidance/pnp-provisioning-engine-and-the-core-library)

## Step 1: Export SharePoint file list to CSV

On your machine, open PowerShell and run:

```powershell
cd "C:\Users\laine\.cursor\salesforce-mcp"

.\Export-SharePointFileList.ps1 `
  -SiteUrl "https://YOUR-TENANT.sharepoint.com/sites/YOUR-SITE" `
  -LibraryName "Documents" `
  -OutputPath ".\SharePoint_FileList.csv"
```

- **SiteUrl**: Your SharePoint site URL (e.g. the CIM/CIP library site).
- **LibraryName**: Document library name (e.g. `Documents`, `Shared Documents`, or the exact name).
- A browser sign-in may appear for `-Interactive` auth.

This creates `SharePoint_FileList.csv` with columns such as: Name, FileRef, FileExtension, Length, TimeCreated, TimeLastModified, Author, Editor.

## Step 2: Run the comparison

Use your **Salesforce CSV** (e.g. `sfdc_2025_uspe_cims.csv`) and the **SharePoint CSV** from Step 1:

```powershell
python compare_sf_sharepoint.py `
  --salesforce "C:\Users\laine\OneDrive\Downloads\sfdc_2025_uspe_cims.csv" `
  --sharepoint ".\SharePoint_FileList.csv" `
  --output-dir ".\comparison_results"
```

Optional:

- `--case-sensitive` — compare file names case-sensitively (default is case-insensitive).

## Step 3: Use the results

Outputs are in the folder you set with `--output-dir` (default: `.\comparison_results`):

| File | Meaning |
|------|--------|
| `only_in_sharepoint__give_to_coworker.csv` | Files you have in SharePoint that Salesforce doesn’t — **give this list (or these files) to your coworker** |
| `only_in_salesforce__we_are_missing.csv` | Files in Salesforce that you don’t have in SharePoint |
| `in_both_salesforce.csv` | Rows from Salesforce CSV for files that appear in both |
| `in_both_sharepoint.csv` | Rows from SharePoint CSV for files that appear in both |
| `comparison_summary.txt` | Counts and short explanation |

All CSVs keep the original columns from the source export so you can use metadata (e.g. Opp Name, Fund, TimeLastModified) for follow-up.

## Salesforce CSV format

Expected column: **Document Title** (file name). Your export already has: ContentDocumentId, Document Title, FileExtension, Opp ID, Opp Name, Fund, Opp Owner Name, DownloadUrl, FilePath.

## SharePoint CSV format

Expected column: **Name** (or FileLeafRef / Title). The PowerShell script writes: Name, FileRef, FileExtension, Length, TimeCreated, TimeLastModified, Author, Editor.

## Re-running later

1. Export a fresh Salesforce CSV (same columns).
2. Re-run the PowerShell script to refresh `SharePoint_FileList.csv`.
3. Re-run `compare_sf_sharepoint.py` with the two CSVs and a new or same `--output-dir`.

## Troubleshooting

- **“List 'Documents' not found”**  
  Run the script without `-LibraryName` to see available lists, or use the exact library name (e.g. the CIM/CIP library name).

- **PowerShell execution policy**  
  If scripts won’t run: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

- **Large libraries**  
  The PnP script uses paging; if you have 5,000+ items and see timeouts, run during off-peak or split by folder if you use subfolders.
