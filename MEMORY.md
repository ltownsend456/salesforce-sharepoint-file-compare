# Agent memory -- Salesforce vs SharePoint file comparison

## Project purpose

- Compare file lists between **Salesforce** (CSV export) and **SharePoint** (document library).
- Identify: in both, only in SharePoint, only in Salesforce.
- Produce **"files to give coworker"**: items in SharePoint that are missing in Salesforce (coworker's side = Salesforce; our side = SharePoint).
- Use metadata from both sources (e.g. Document Title, FileExtension, Opp Name, Fund, TimeLastModified, Author).

## Reproducing this project (setup steps)

1. **Workspace**: Use a folder (e.g. `salesforce-mcp`) with:
   - `Export-SharePointFileList.ps1` -- PnP PowerShell script to export SharePoint file list to CSV.
   - `compare_sf_sharepoint.py` -- Python 3.9+ script (stdlib only) to compare the two CSVs (exact + fuzzy).
   - `README.md` -- User-facing instructions.

2. **Dependencies**:
   - Python 3.9+ (no pip packages; uses `difflib.SequenceMatcher` for fuzzy matching).
   - PnP.PowerShell: `Install-Module -Name PnP.PowerShell -Scope CurrentUser`.

3. **Inputs**:
   - Salesforce CSV with column **Document Title** (e.g. `sfdc_2025_uspe_cims.csv`).
   - SharePoint CSV with column **Name** (or FileLeafRef/Title), produced by `Export-SharePointFileList.ps1`.

4. **Workflow**:
   - Run PowerShell script: connect to SharePoint site, export document library to CSV.
   - Run Python script with `--salesforce` and `--sharepoint` paths and `--output-dir`.
   - Add `--fuzzy` for fuzzy matching (recommended); optionally set `--fuzzy-threshold` (default 70).
   - Use `comparison_results/only_in_sharepoint__give_to_coworker.csv` for files to give coworker.

5. **Outputs**: `only_in_sharepoint__give_to_coworker.csv`, `only_in_salesforce__we_are_missing.csv`, `in_both_*.csv`, `fuzzy_matches.csv`, `comparison_summary.txt`.

## Fuzzy matching (added Feb 2026)

The `--fuzzy` flag enables three-pass fuzzy matching:
- **Pass 1 (project codename)**: Extract "Project X" from both filenames; match on codename. Handles `project\s`, `project_`, `project-`, `project.` separators. High confidence.
- **Pass 2 (token fuzzy)**: Build inverted index of SP name tokens; find candidates by shared token; score with `difflib.SequenceMatcher` (60%) + Jaccard token overlap (40%). Guard: rejects pairs where both have different "Project X" codenames.
- **Pass 3 (Opp Name cross-reference)**: For SF files without a project codename in the filename but with one in the "Opp Name" column (e.g. "ChenMed CIM.pdf" + Opp "ChenMed (Project Resurgence)"), use the Opp Name's project codename to find SP matches.

### Performance
- 2.3k SF x 7k SP with token pre-filtering takes ~2 minutes on a VDI.
- Threshold 70 is recommended (good balance of precision and recall).
- `fuzzy_matches.csv` includes Match_Score and Match_Method for manual review.

### Key findings from analysis (Feb 2026)
- **SF export is NOT CIM-only**: Contains CIM (752), CIP (443), Teaser (148), Presentation (201), Memorandum (196), Overview (93), Other (479). SP export is CIM-only (7,001 files). This is the primary reason for the large "unmatched" count.
- **Results at threshold 70**: 131 exact + 470 fuzzy = 601 accounted for (of 2,224 SF files).
- **CIM-to-CIM**: ~340 of 762 CIM-named SF files matched (exact + fuzzy). The remaining 400+ are likely deals not in that SP library, or naming differences beyond fuzzy reach.
- **False positive mitigation**: Project name guard rejects token_fuzzy matches where both sides have "Project X" but X differs. Regex handles spaces, hyphens, underscores, dots.

## Important details

- Comparison is **case-insensitive** by default; use `--case-sensitive` in Python to change.
- Matching is by **file name** (Document Title vs Name); metadata is carried through for reporting.
- PowerShell uses `Get-PnPListItem` with `FSObjType=0` to get all files (including subfolders) via `-PageSize 500` pagination.
- Salesforce CSV path used in examples: `C:\Users\laine\OneDrive\Downloads\sfdc_2025_uspe_cims.csv`.
- SharePoint site: `https://hignet.sharepoint.com/sites/DATA-OP-CIM_CIP_Documents`, library: `Shared Documents`.

## PowerShell version

- **PnP.PowerShell 3.x** requires **PowerShell 7.4.6+**. On **Windows PowerShell 5.1** (common on VDIs) it will not load.
- **Install-PnPPrerequisite.ps1** detects PS version: on 5.1 it installs **SharePointPnPPowerShellOnline** (legacy); on 7+ it installs PnP.PowerShell.
- **Export-SharePointFileList.ps1** tries the legacy module first on 5.1, then PnP.PowerShell. Connect-PnPOnline uses -UseWebLogin (legacy) or -Interactive (modern).

## VDI / OneDrive

- On VDIs, Documents is often **OneDrive-redirected** (e.g. `C:\Users\...\OneDrive - HIG.net\Documents\...`). Install-Module -Scope CurrentUser installs to the first writable path in PSModulePath (the OneDrive path). Scripts must look for PnP in **every** path in $env:PSModulePath, not only $HOME\Documents\..., so Export and Install now iterate PSModulePath.

## SharePoint scope: CIM-only vs full library

- **Full library export** (no filter) can yield ~16k rows; **Salesforce CIM export** is ~2k. Comparing those directly gives a huge "only in SharePoint" list and is not what stakeholders expect.
- **Recommended**: Re-export SharePoint with **-DocumentTypeFilter "CIM"** so both sides are CIM-only (apples-to-apples). Replace the existing `SharePoint_FileList.csv` and re-run the Python comparison.
- **Export-SharePointFileList.ps1** supports `-DocumentTypeFilter "CIM"` (or any value). It requests `Document_x0020_Type` and `DocumentType`; if the column has a different internal name, run `Get-PnPField -List "Shared Documents"` on the site to find it and adjust the script.
- Choice/Lookup: script handles both string values and .LookupValue for Document Type. PowerShell 5.1-compatible (no `??`).

## What NOT to use for the SharePoint file list

- **SharePoint CIM/CIP MCP search**: Do not use for generating the full file list. It only returns a limited, search-ranked subset of files, is slow for bulk inventory, and is not reliable for complete SF vs SP comparison. Use the PnP PowerShell export on the user's VDI instead.

## Code verification (Feb 2026)

All code verified against official PnP PowerShell docs and Python docs:

- **-Fields parameter**: Confirmed correct (`String[]`). Docs: https://pnp.github.io/powershell/cmdlets/Get-PnPListItem.html
- **-PageSize 500 without -Query**: Confirmed works for paginating large lists (>5000 items). Do NOT combine with -Query (hits threshold).
- **Document_x0020_Type**: Confirmed correct internal name for "Document Type" (space -> `_x0020_`). Non-existent fields return null without error.
- **Choice field**: Returns a plain string (not FieldLookupValue). `$docType -is [string]` check is correct.
- **User/Person field (Author, Editor)**: Returns FieldUserValue; `.LookupValue` gives the display name.
- **?? null-coalescing operator**: NOT available in PowerShell 5.1, only 7+. Replaced with `if ($null -ne ...) { ... } else { ... }`.
- **Connect-PnPOnline**: Legacy uses `-UseWebLogin`; modern PnP.PowerShell uses `-Interactive`.
- **FSObjType**: 0 = file, 1 = folder. Confirmed.
- **Python utf-8-sig encoding**: Correctly strips BOM from CSV. csv.DictReader reads first column name cleanly.
- **Python errors='replace'**: Works but may silently replace non-UTF-8 chars. Acceptable for file name comparison.
- **difflib.SequenceMatcher**: stdlib, no pip install. Slower than rapidfuzz but accurate; optimized with token-indexed candidate search.

## References

- Microsoft: PnP PowerShell -- https://learn.microsoft.com/sharepoint/dev/solution-guidance/pnp-provisioning-engine-and-the-core-library
- PnP Get-PnPListItem docs -- https://pnp.github.io/powershell/cmdlets/Get-PnPListItem.html
- Power Automate: Get files / Get items -- https://learn.microsoft.com/sharepoint/dev/business-apps/power-automate/guidance/working-with-get-items-and-get-files
- Power Query: SharePoint.Files -- https://learn.microsoft.com/powerquery-m/sharepoint-files
