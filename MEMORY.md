# Agent memory – Salesforce vs SharePoint file comparison

## Project purpose

- Compare file lists between **Salesforce** (CSV export) and **SharePoint** (document library).
- Identify: in both, only in SharePoint, only in Salesforce.
- Produce **“files to give coworker”**: items in SharePoint that are missing in Salesforce (coworker’s side = Salesforce; our side = SharePoint).
- Use metadata from both sources (e.g. Document Title, FileExtension, Opp Name, Fund, TimeLastModified, Author).

## Reproducing this project (setup steps)

1. **Workspace**: Use a folder (e.g. `salesforce-mcp`) with:
   - `Export-SharePointFileList.ps1` – PnP PowerShell script to export SharePoint file list to CSV.
   - `compare_sf_sharepoint.py` – Python 3.9+ script (stdlib only) to compare the two CSVs.
   - `README.md` – User-facing instructions.

2. **Dependencies**:
   - Python 3.9+ (no pip packages).
   - PnP.PowerShell: `Install-Module -Name PnP.PowerShell -Scope CurrentUser`.

3. **Inputs**:
   - Salesforce CSV with column **Document Title** (e.g. `sfdc_2025_uspe_cims.csv`).
   - SharePoint CSV with column **Name** (or FileLeafRef/Title), produced by `Export-SharePointFileList.ps1`.

4. **Workflow**:
   - Run PowerShell script: connect to SharePoint site, export document library to CSV.
   - Run Python script with `--salesforce` and `--sharepoint` paths and `--output-dir`.
   - Use `comparison_results/only_in_sharepoint__give_to_coworker.csv` for files to give coworker.

5. **Outputs**: `only_in_sharepoint__give_to_coworker.csv`, `only_in_salesforce__we_are_missing.csv`, `in_both_*.csv`, `comparison_summary.txt`.

## Important details

- Comparison is **case-insensitive** by default; use `--case-sensitive` in Python to change.
- Matching is by **file name** (Document Title vs Name); metadata is carried through for reporting.
- PowerShell uses `Get-PnPListItem` with `Scope='RecursiveAll'` and `FSObjType=0` to get all files (including subfolders).
- Salesforce CSV path used in examples: `C:\Users\laine\OneDrive\Downloads\sfdc_2025_uspe_cims.csv`.

## What NOT to use for the SharePoint file list

- **SharePoint CIM/CIP MCP search**: Do not use for generating the full file list. It only returns a limited, search-ranked subset of files, is slow for bulk inventory, and is not reliable for complete SF vs SP comparison. Use the PnP PowerShell export on the user's VDI instead.

## References

- Microsoft: PnP PowerShell – https://learn.microsoft.com/sharepoint/dev/solution-guidance/pnp-provisioning-engine-and-the-core-library
- Power Automate: Get files / Get items – https://learn.microsoft.com/sharepoint/dev/business-apps/power-automate/guidance/working-with-get-items-and-get-files
- Power Query: SharePoint.Files – https://learn.microsoft.com/powerquery-m/sharepoint-files
