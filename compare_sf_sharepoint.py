#!/usr/bin/env python3
r"""
Compare files between Salesforce (CSV export) and SharePoint (CSV export).

Outputs:
- Files in BOTH systems (exact match)
- Fuzzy matches (likely the same document, different name)
- Only in SharePoint (files to give your coworker - they don't have these in Salesforce)
- Only in Salesforce (files coworker has that you don't have in SharePoint)

Usage:
  python compare_sf_sharepoint.py --salesforce <path> --sharepoint <path> --output-dir .\comparison_results
  python compare_sf_sharepoint.py --salesforce <path> --sharepoint <path> --output-dir .\comparison_results --fuzzy --fuzzy-threshold 70
"""

import argparse
import csv
import re
import sys
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path


# Column names: Salesforce CSV uses "Document Title", SharePoint export uses "Name"
SF_FILE_COLUMN = "Document Title"
SP_FILE_COLUMN = "Name"

# Noise words to strip when normalizing for fuzzy comparison
NOISE_WORDS = {
    "hig", "h.i.g.", "h.i.g", "capital", "management", "llc", "inc",
    "confidential", "information", "memorandum", "presentation",
    "vf", "vfinal", "final", "v1", "v2", "v3", "vs", "draft",
    "copy", "print", "redacted",
    "pdf", "docx", "xlsx", "pptx", "doc", "xls", "ppt", "msg", "xps",
    "middle", "market", "advantage", "fund", "infrastructure", "infra",
    "small", "cap", "lbo", "growth",
}

# Regex patterns for stripping dates, version suffixes, etc.
DATE_PATTERNS = [
    r"\(\w+\s+\d{4}\)",          # (April 2025), (Sep 2025)
    r"\(\w+\s+\d{4}\)",          # (Summer 2025), (Fall 2025)
    r"\d{1,2}[\.\-/]\d{1,2}[\.\-/]\d{2,4}",  # 12.18.24, 2025-01-01
    r"\d{4}[\.\-/]\d{1,2}[\.\-/]\d{1,2}",    # 2025.06.24
    r"(?:january|february|march|april|may|june|july|august|september|october|november|december)\s*\d{4}",
    r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\.?\s*\d{4}",
    r"(?:spring|summer|fall|winter|q[1-4])\s*[-_]?\s*\d{4}",
    r"\d{4}",                     # bare year (removed last to avoid stripping project numbers)
]


def normalize_name(name: str) -> str:
    """Normalize file name for comparison (strip whitespace)."""
    if not name or not isinstance(name, str):
        return ""
    return name.strip()


def normalize_for_fuzzy(name: str) -> str:
    """
    Aggressively normalize a file name for fuzzy comparison:
    - Lowercase
    - Remove file extension
    - Remove dates, version tags, noise words
    - Collapse separators and whitespace
    """
    if not name:
        return ""
    s = name.lower().strip()
    # Remove file extension
    s = re.sub(r"\.[a-z]{2,5}$", "", s)
    # Remove dates
    for pat in DATE_PATTERNS:
        s = re.sub(pat, " ", s, flags=re.IGNORECASE)
    # Replace separators with spaces
    s = re.sub(r"[_\-\.\(\)\[\],;:]+", " ", s)
    # Remove noise words
    tokens = s.split()
    tokens = [t for t in tokens if t not in NOISE_WORDS and len(t) > 1]
    return " ".join(tokens).strip()


def extract_project_name(name: str) -> str | None:
    """
    Extract the project codename from a filename.
    E.g. 'Project Eagle CIM (HIG).pdf' -> 'eagle'
         'Project TITAN_CIM_H.I.G. Capital.pdf' -> 'titan'
         'Project-ORANGE-CIM-1.pdf' -> 'orange'
         'Project_Quartz_CIM.pdf' -> 'quartz'
    Handles space, hyphen, underscore, and dot separators after 'Project'.
    """
    m = re.search(r"project[\s_\-\.]+([a-zA-Z]+)", name, re.IGNORECASE)
    if m:
        codename = m.group(1).lower()
        # Skip if the "codename" is actually a generic word (e.g. "Confidential")
        if codename in NOISE_WORDS or len(codename) < 3:
            return None
        return codename
    return None


def tokenize(name: str) -> set[str]:
    """
    Tokenize a normalized name into significant tokens for candidate lookup.
    """
    norm = normalize_for_fuzzy(name)
    tokens = set(norm.split())
    # Remove very short tokens (1 char) and pure numbers
    return {t for t in tokens if len(t) > 1 and not t.isdigit()}


def fuzzy_score(a: str, b: str) -> int:
    """
    Compute a fuzzy similarity score (0-100) between two normalized names.
    Uses SequenceMatcher for overall similarity and token overlap as a boost.
    """
    norm_a = normalize_for_fuzzy(a)
    norm_b = normalize_for_fuzzy(b)
    if not norm_a or not norm_b:
        return 0

    # SequenceMatcher ratio (0.0 - 1.0)
    seq_ratio = SequenceMatcher(None, norm_a, norm_b).ratio()

    # Token overlap (Jaccard)
    tokens_a = set(norm_a.split())
    tokens_b = set(norm_b.split())
    if tokens_a and tokens_b:
        jaccard = len(tokens_a & tokens_b) / len(tokens_a | tokens_b)
    else:
        jaccard = 0.0

    # Weighted combination: 60% sequence, 40% token overlap
    score = int((0.6 * seq_ratio + 0.4 * jaccard) * 100)
    return score


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
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = (
                row.get(SP_FILE_COLUMN)
                or row.get("FileLeafRef")
                or row.get("Title")
                or row.get("\ufeffName")
                or ""
            )
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


def find_fuzzy_matches(
    sf_unmatched: set[str],
    sp_unmatched: set[str],
    sf_by_name: dict[str, list[dict]],
    sp_by_name: dict[str, list[dict]],
    threshold: int = 70,
    case_sensitive: bool = False,
) -> tuple[list[dict], set[str], set[str]]:
    """
    Find fuzzy matches between unmatched SF and SP names.
    Uses a multi-pass approach for performance:
      Pass 1: Project codename match (fast, high confidence)
      Pass 2: Token-indexed candidate search + SequenceMatcher (medium)

    Returns: (fuzzy_match_rows, sf_matched_keys, sp_matched_keys)
    """
    print("  Running fuzzy matching...")

    # Build lookup structures for SP
    # Map: lowered original name -> original name
    sp_orig = {}
    for k in sp_by_name:
        key = k if case_sensitive else k.lower()
        if key not in sp_orig:
            sp_orig[key] = k

    sf_orig = {}
    for k in sf_by_name:
        key = k if case_sensitive else k.lower()
        if key not in sf_orig:
            sf_orig[key] = k

    # ---- Pass 1: Project codename matching ----
    print("    Pass 1: Project codename matching...")
    sp_by_project: dict[str, list[str]] = defaultdict(list)
    for sp_key in sp_unmatched:
        proj = extract_project_name(sp_key)
        if proj:
            sp_by_project[proj].append(sp_key)

    fuzzy_rows = []
    sf_matched = set()
    sp_matched = set()

    for sf_key in sorted(sf_unmatched):
        proj = extract_project_name(sf_key)
        if not proj or proj not in sp_by_project:
            continue
        # Find best match among SP names with same project codename
        best_score = 0
        best_sp = None
        for sp_key in sp_by_project[proj]:
            if sp_key in sp_matched:
                continue
            score = fuzzy_score(sf_key, sp_key)
            if score > best_score:
                best_score = score
                best_sp = sp_key
        if best_sp and best_score >= threshold:
            sf_matched.add(sf_key)
            sp_matched.add(best_sp)
            # Get original-case names for display
            sf_display = sf_orig.get(sf_key, sf_key)
            sp_display = sp_orig.get(best_sp, best_sp)
            sf_row = sf_by_name.get(sf_display, [{}])[0]
            sp_row = sp_by_name.get(sp_display, [{}])[0]
            fuzzy_rows.append({
                "SF_Name": sf_row.get(SF_FILE_COLUMN, sf_display),
                "SP_Name": sp_row.get(SP_FILE_COLUMN, sp_row.get("Name", sp_display)),
                "Match_Score": best_score,
                "Match_Method": "project_name",
                "SF_Opp_Name": sf_row.get("Opp Name", ""),
                "SF_Fund": sf_row.get("Fund", ""),
            })

    print(f"    Pass 1 matched: {len(sf_matched)} pairs")

    # ---- Pass 2: Token-indexed candidate search ----
    print("    Pass 2: Token-indexed fuzzy search...")
    sf_remaining = sf_unmatched - sf_matched
    sp_remaining = sp_unmatched - sp_matched

    # Build inverted index: token -> set of SP keys
    sp_token_index: dict[str, set[str]] = defaultdict(set)
    for sp_key in sp_remaining:
        for tok in tokenize(sp_key):
            sp_token_index[tok].add(sp_key)

    pass2_count = 0
    comparisons = 0
    for sf_key in sorted(sf_remaining):
        sf_tokens = tokenize(sf_key)
        if not sf_tokens:
            continue

        # Find candidate SP names: any SP name sharing >= 1 significant token
        candidates = set()
        for tok in sf_tokens:
            if tok in sp_token_index:
                candidates.update(sp_token_index[tok])
        candidates -= sp_matched  # skip already matched

        if not candidates:
            continue

        sf_proj = extract_project_name(sf_key)

        best_score = 0
        best_sp = None
        for sp_key in candidates:
            comparisons += 1
            # GUARD: if both have "Project X" but X differs, skip (prevents false positives)
            sp_proj = extract_project_name(sp_key)
            if sf_proj and sp_proj and sf_proj != sp_proj:
                continue
            score = fuzzy_score(sf_key, sp_key)
            if score > best_score:
                best_score = score
                best_sp = sp_key

        if best_sp and best_score >= threshold:
            sf_matched.add(sf_key)
            sp_matched.add(best_sp)
            # Remove from token index
            for tok in tokenize(best_sp):
                sp_token_index.get(tok, set()).discard(best_sp)

            sf_display = sf_orig.get(sf_key, sf_key)
            sp_display = sp_orig.get(best_sp, best_sp)
            sf_row = sf_by_name.get(sf_display, [{}])[0]
            sp_row = sp_by_name.get(sp_display, [{}])[0]
            fuzzy_rows.append({
                "SF_Name": sf_row.get(SF_FILE_COLUMN, sf_display),
                "SP_Name": sp_row.get(SP_FILE_COLUMN, sp_row.get("Name", sp_display)),
                "Match_Score": best_score,
                "Match_Method": "token_fuzzy",
                "SF_Opp_Name": sf_row.get("Opp Name", ""),
                "SF_Fund": sf_row.get("Fund", ""),
            })
            pass2_count += 1

    print(f"    Pass 2 matched: {pass2_count} pairs ({comparisons:,} comparisons)")

    # ---- Pass 3: Opp Name cross-reference ----
    # If SF has "Opp Name" with a project codename not in the filename,
    # use that codename to find SP matches (e.g. "ChenMed CIM.pdf" + Opp "ChenMed (Project Resurgence)")
    print("    Pass 3: Opp Name project codename cross-reference...")
    sf_remaining2 = (sf_unmatched - sf_matched)
    pass3_count = 0

    for sf_key in sorted(sf_remaining2):
        sf_display_name = sf_orig.get(sf_key, sf_key)
        sf_row_list = sf_by_name.get(sf_display_name, [])
        if not sf_row_list:
            continue
        sf_row = sf_row_list[0]
        opp_name = sf_row.get("Opp Name", "")
        opp_proj = extract_project_name(opp_name)
        file_proj = extract_project_name(sf_key)
        # Only use Opp Name if filename doesn't already have a project codename
        if not opp_proj or file_proj:
            continue
        if opp_proj not in sp_by_project:
            continue
        # Find best SP match with this project codename
        best_score = 0
        best_sp = None
        for sp_key in sp_by_project[opp_proj]:
            if sp_key in sp_matched:
                continue
            score = fuzzy_score(sf_key, sp_key)
            # Also try scoring with the opp project name prepended
            alt_name = f"project {opp_proj} {normalize_for_fuzzy(sf_key)}"
            alt_score = fuzzy_score(alt_name, sp_key)
            use_score = max(score, alt_score)
            if use_score > best_score:
                best_score = use_score
                best_sp = sp_key
        if best_sp and best_score >= max(threshold - 10, 60):  # slightly lower threshold for opp-name cross-ref
            sf_matched.add(sf_key)
            sp_matched.add(best_sp)
            sp_display_name = sp_orig.get(best_sp, best_sp)
            sp_row = sp_by_name.get(sp_display_name, [{}])[0]
            fuzzy_rows.append({
                "SF_Name": sf_row.get(SF_FILE_COLUMN, sf_display_name),
                "SP_Name": sp_row.get(SP_FILE_COLUMN, sp_row.get("Name", sp_display_name)),
                "Match_Score": best_score,
                "Match_Method": f"opp_name_xref (Project {opp_proj.title()})",
                "SF_Opp_Name": opp_name,
                "SF_Fund": sf_row.get("Fund", ""),
            })
            pass3_count += 1

    print(f"    Pass 3 matched: {pass3_count} pairs")
    print(f"    Total fuzzy matches: {len(fuzzy_rows)}")

    return fuzzy_rows, sf_matched, sp_matched


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
    parser.add_argument(
        "--fuzzy",
        action="store_true",
        help="Enable fuzzy matching for files with similar but not identical names",
    )
    parser.add_argument(
        "--fuzzy-threshold",
        type=int,
        default=70,
        help="Fuzzy match score threshold 0-100 (default: 70). Higher = stricter.",
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
        sf_keys = {normalize_name(n) for n in sf_names}
        sp_keys = {normalize_name(n) for n in sp_names}
    else:
        sf_keys = {normalize_name(n).lower() for n in sf_names}
        sp_keys = {normalize_name(n).lower() for n in sp_names}

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

    in_both_exact = sf_keys & sp_keys
    only_sharepoint = sp_keys - sf_keys
    only_salesforce = sf_keys - sp_keys

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

    # ---- Fuzzy matching ----
    fuzzy_match_rows = []
    sf_fuzzy_matched = set()
    sp_fuzzy_matched = set()

    if args.fuzzy:
        # Analyze SF doc type breakdown
        sf_breakdown = defaultdict(int)
        for name in sf_keys:
            nl = name.lower()
            if "cim" in nl:
                sf_breakdown["CIM"] += 1
            elif "cip" in nl:
                sf_breakdown["CIP"] += 1
            elif "teaser" in nl:
                sf_breakdown["Teaser"] += 1
            elif "presentation" in nl:
                sf_breakdown["Presentation"] += 1
            elif "memorandum" in nl:
                sf_breakdown["Memorandum"] += 1
            elif "overview" in nl:
                sf_breakdown["Overview"] += 1
            else:
                sf_breakdown["Other"] += 1

        print("\nSalesforce document type breakdown (by filename keywords):")
        for dtype, cnt in sorted(sf_breakdown.items(), key=lambda x: -x[1]):
            print(f"  {dtype}: {cnt}")
        print()

        fuzzy_match_rows, sf_fuzzy_matched, sp_fuzzy_matched = find_fuzzy_matches(
            sf_unmatched=only_salesforce,
            sp_unmatched=only_sharepoint,
            sf_by_name=sf_by_name,
            sp_by_name=sp_by_name,
            threshold=args.fuzzy_threshold,
            case_sensitive=args.case_sensitive,
        )

        # Update only_ sets to exclude fuzzy matches
        only_salesforce = only_salesforce - sf_fuzzy_matched
        only_sharepoint = only_sharepoint - sp_fuzzy_matched

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
    for k in sorted(in_both_exact):
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

    if args.fuzzy and fuzzy_match_rows:
        # Sort fuzzy matches by score descending
        fuzzy_match_rows.sort(key=lambda r: -r["Match_Score"])
        write_csv(
            out_dir / "fuzzy_matches.csv",
            fuzzy_match_rows,
            fieldnames=["SF_Name", "SP_Name", "Match_Score", "Match_Method", "SF_Opp_Name", "SF_Fund"],
        )

    # Summary report
    report_lines = [
        "Salesforce vs SharePoint file comparison",
        "=" * 60,
        f"Salesforce CSV: {sf_path}",
        f"SharePoint CSV: {sp_path}",
        f"Output directory: {out_dir}",
        "",
        "Counts:",
        f"  Total unique in Salesforce:    {len(sf_keys)}",
        f"  Total unique in SharePoint:    {len(sp_keys)}",
        "",
        f"  EXACT match (same file name):  {len(in_both_exact)}",
    ]
    if args.fuzzy:
        report_lines += [
            f"  FUZZY match (similar name):    {len(fuzzy_match_rows)}  (threshold: {args.fuzzy_threshold})",
            f"    -> via project codename:     {sum(1 for r in fuzzy_match_rows if r['Match_Method'] == 'project_name')}",
            f"    -> via token fuzzy:          {sum(1 for r in fuzzy_match_rows if r['Match_Method'] == 'token_fuzzy')}",
            f"    -> via Opp Name cross-ref:   {sum(1 for r in fuzzy_match_rows if r['Match_Method'].startswith('opp_name'))}",
            f"  TOTAL accounted for (SF side): {len(in_both_exact) + len(sf_fuzzy_matched)} of {len(sf_keys)}",
        ]
    report_lines += [
        "",
        f"  Only in SharePoint (give to coworker): {len(only_sharepoint)}",
        f"  Only in Salesforce (we are missing):   {len(only_salesforce)}",
        "",
        "Output files:",
        "  - only_in_sharepoint__give_to_coworker.csv  -> Files to give coworker (not in SF)",
        "  - only_in_salesforce__we_are_missing.csv    -> Files we don't have in SP",
        "  - in_both_salesforce.csv / in_both_sharepoint.csv -> Exact matches (with metadata)",
    ]
    if args.fuzzy:
        report_lines.append(
            "  - fuzzy_matches.csv -> Likely same doc, different name (review manually)"
        )
    report_lines.append("")

    report_path = out_dir / "comparison_summary.txt"
    report_path.write_text("\n".join(report_lines), encoding="utf-8")

    for line in report_lines:
        print(line)

    print("\nDone. Open the output CSVs in the folder above.")


if __name__ == "__main__":
    main()
