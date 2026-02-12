#!/usr/bin/env python3
r"""
Compare files between Salesforce (CSV export) and SharePoint (CSV export).

Outputs:
- Files in BOTH systems
- Only in SharePoint (files to give your coworker - they don't have these in Salesforce)
- Only in Salesforce (files coworker has that you don't have in SharePoint)
- Optional metadata comparison report

Usage:
  python compare_sf_sharepoint.py --salesforce <path-to-sfdc-csv> --sharepoint <path-to-sp-csv> --output-dir .\comparison_results
"""

import argparse
import csv
import sys
from pathlib import Path
from collections import defaultdict


# Column names: Salesforce CSV uses "Document Title", SharePoint export uses "Name"
SF_FILE_COLUMN = "Document Title"
SP_FILE_COLUMN = "Name"


def normalize_name(name: str) -> str:
    """Normalize file name for comparison (strip, optional case-insensitive)."""
    if not name or not isinstance(name, str):
        return ""
    return name.strip()


def load_salesforce_csv(path: str) -> tuple[list[dict], set[str], dict[str, list[dict]]]:
    """Load Salesforce CSV. Returns (rows, set of normalized names, name -> rows for metadata)."""
    rows = []
    names = set()
    by_name = defaultdict(list)
    with open(path, newline="", encoding="utf-8", errors="replace") as f:
        reader = csv.DictReader(f)
        for row in reader:
            title = row.get(SF_FILE_COLUMN, "").strip()
            if not title:
                continue
            rows.append(row)
            key = normalize_name(title)
            names.add(key)
            by_name[key].append(row)
    return rows, names, dict(by_name)


def load_sharepoint_csv(path: str) -> tuple[list[dict], set[str], dict[str, list[dict]]]:
    """Load SharePoint CSV. Returns (rows, set of normalized names, name -> rows)."""
    rows = []
    names = set()
    by_name = defaultdict(list)
    with open(path, newline="", encoding="utf-8", errors="replace") as f:
        reader = csv.DictReader(f)
        for row in reader:
            # Accept "Name" or "FileLeafRef" or "Title"
            name = row.get(SP_FILE_COLUMN) or row.get("FileLeafRef") or row.get("Title", "")
            if not name or not str(name).strip():
                continue
            name = str(name).strip()
            rows.append(row)
            key = normalize_name(name)
            names.add(key)
            by_name[key].append(row)
    return rows, names, dict(by_name)


def write_csv(path: Path, rows: list[dict], fieldnames: list[str] | None = None):
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    if fieldnames is None:
        fieldnames = list(rows[0].keys())
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def main():
    parser = argparse.ArgumentParser(
        description="Compare Salesforce and SharePoint file lists; output what's missing where."
    )
    parser.add_argument(
        "--salesforce", "-sf",
        required=True,
        help="Path to Salesforce export CSV (must have 'Document Title' column)",
    )
    parser.add_argument(
        "--sharepoint", "-sp",
        required=True,
        help="Path to SharePoint export CSV (must have 'Name' or 'FileLeafRef' column)",
    )
    parser.add_argument(
        "--output-dir", "-o",
        default="./comparison_results",
        help="Directory for output CSVs and report (default: ./comparison_results)",
    )
    parser.add_argument(
        "--case-sensitive",
        action="store_true",
        help="Compare file names case-sensitively (default: case-insensitive)",
    )
    args = parser.parse_args()

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Load both CSVs
    sf_path = Path(args.salesforce)
    sp_path = Path(args.sharepoint)
    if not sf_path.exists():
        print(f"Error: Salesforce CSV not found: {sf_path}", file=sys.stderr)
        sys.exit(1)
    if not sp_path.exists():
        print(f"Error: SharePoint CSV not found: {sp_path}", file=sys.stderr)
        sys.exit(1)

    sf_rows, sf_names, sf_by_name = load_salesforce_csv(str(sf_path))
    sp_rows, sp_names, sp_by_name = load_sharepoint_csv(str(sp_path))

    if args.case_sensitive:
        sf_names = {normalize_name(n) for n in sf_names}
        sp_names = {normalize_name(n) for n in sp_names}
    else:
        sf_names = {normalize_name(n).lower() for n in sf_names}
        sp_names = {normalize_name(n).lower() for n in sp_names}

    # Build name -> original display name (first occurrence)
    sf_display = {}
    for r in sf_rows:
        t = r.get(SF_FILE_COLUMN, "").strip()
        k = normalize_name(t).lower() if not args.case_sensitive else normalize_name(t)
        if k and k not in sf_display:
            sf_display[k] = t
    sp_display = {}
    for r in sp_rows:
        n = (r.get(SP_FILE_COLUMN) or r.get("FileLeafRef") or r.get("Title", "") or "").strip()
        if isinstance(n, str) and n:
            k = normalize_name(n).lower() if not args.case_sensitive else normalize_name(n)
            if k not in sp_display:
                sp_display[k] = n

    in_both = sf_names & sp_names
    only_sharepoint = sp_names - sf_names
    only_salesforce = sf_names - sp_names

    def first_sf_row(key):
        for k, rows in sf_by_name.items():
            if (k.lower() if not args.case_sensitive else k) == key and rows:
                return rows[0]
        return None

    def first_sp_row(key):
        for k, rows in sp_by_name.items():
            if (k.lower() if not args.case_sensitive else k) == key and rows:
                return rows[0]
        return None

    # Build output rows for each category (use first row per name for metadata)
    only_sp_rows = []
    for k in sorted(only_sharepoint):
        display = sp_display.get(k, k)
        r = first_sp_row(k)
        if r:
            only_sp_rows.append({**r, "Name": display})
        else:
            only_sp_rows.append({"Name": display})

    only_sf_rows = []
    for k in sorted(only_salesforce):
        display = sf_display.get(k, k)
        r = first_sf_row(k)
        if r:
            only_sf_rows.append({**r, SF_FILE_COLUMN: display})
        else:
            only_sf_rows.append({SF_FILE_COLUMN: display})

    both_sf_rows = []
    both_sp_rows = []
    for k in sorted(in_both):
        r = first_sf_row(k)
        if r:
            both_sf_rows.append(r)
        r = first_sp_row(k)
        if r:
            both_sp_rows.append(r)

    # Write CSVs
    write_csv(out_dir / "only_in_sharepoint__give_to_coworker.csv", only_sp_rows)
    write_csv(out_dir / "only_in_salesforce__we_are_missing.csv", only_sf_rows)
    write_csv(out_dir / "in_both_salesforce.csv", both_sf_rows)
    write_csv(out_dir / "in_both_sharepoint.csv", both_sp_rows)

    # Summary report
    report_lines = [
        "Salesforce vs SharePoint file comparison",
        "=" * 50,
        f"Salesforce CSV: {sf_path}",
        f"SharePoint CSV: {sp_path}",
        f"Output directory: {out_dir}",
        "",
        "Counts:",
        f"  Total in Salesforce:    {len(sf_names)}",
        f"  Total in SharePoint:   {len(sp_names)}",
        f"  In BOTH:                {len(in_both)}",
        f"  Only in SharePoint (files to give coworker): {len(only_sharepoint)}",
        f"  Only in Salesforce (we are missing these):  {len(only_salesforce)}",
        "",
        "Output files:",
        "  - only_in_sharepoint__give_to_coworker.csv  -> Give this list/files to coworker (he doesn't have these in Salesforce)",
        "  - only_in_salesforce__we_are_missing.csv    -> We don't have these in SharePoint",
        "  - in_both_salesforce.csv / in_both_sharepoint.csv -> Present in both (with metadata)",
        "",
    ]
    report_path = out_dir / "comparison_summary.txt"
    report_path.write_text("\n".join(report_lines), encoding="utf-8")

    for line in report_lines:
        print(line)

    print("\nDone. Open the output CSVs in the folder above.")


if __name__ == "__main__":
    main()
