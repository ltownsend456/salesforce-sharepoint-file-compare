# Verification: Salesforce vs SharePoint comparison alignment

This document checks that the Salesforce CSV, SharePoint CSV, and comparison logic match what you defined (per your reference images).

---

## 1. Intended Salesforce source (from your spec)

| Requirement | Detail |
|-------------|--------|
| **Doc object** | `HIG_Document__c` — Id, OwnerId, Name, CreatedDate, LastModifiedDate, etc. |
| **Opp object** | `Opportunity` — Id, Name, Opportunity_Fund__c, Opportunity_Owner_Full_Name__c |
| **Join** | `HIG_Document__c.Opportunity__c` = `Opportunity.Id` |
| **Filters** | `Doc_Type__c = 'CIM'` and `Opportunity__c != NULL`; Opportunities with `CreatedDate >= 2021-01-01` |
| **Purpose** | “SFDC 2025 US PE CIM names here + opp name” (CIM documents tied to opportunities) |

---

## 2. What the Salesforce CSV actually contains

**File:** `sfdc_2025_uspe_cims.csv`

| Column in CSV | Maps to / Notes |
|---------------|------------------|
| **ContentDocumentId** | ContentDocument Id (Salesforce files); not `HIG_Document__c.Id` |
| **Document Title** | File name (used for comparison) ✓ |
| **FileExtension** | File type ✓ |
| **Opp ID** | Opportunity Id ✓ |
| **Opp Name** | Opportunity name ✓ |
| **Fund** | Maps to Opportunity_Fund__c ✓ |
| **Opp Owner Name** | Maps to Opportunity_Owner_Full_Name__c ✓ |
| **DownloadUrl** | Link to download from Salesforce |
| **FilePath** | Local path in export context |

**Not present in CSV:** `Doc_Type__c`, `Opportunity__c`, or `HIG_Document__c` Id. So we **cannot check inside this file** whether the filters `Doc_Type__c = 'CIM'` and `Opportunity__c != NULL` were applied. That must be guaranteed when the export is built (report/query/ETL). The filename “uspe_cims” and the “CIM” wording in many titles are consistent with it being a CIM-focused export.

**Conclusion:** The CSV has the right **shape** for “CIM names + opp name” and the comparison key (file name) is present. Scope (CIM-only, Opp non-null, Opp created ≥ 2021) is assumed to be enforced in the source of this export.

---

## 3. What the SharePoint CSV actually contains

**File:** `SharePoint_FileList.csv`  
**Source:** Site `DATA-OP-CIM_CIP_Documents` → “Shared Documents” (CIM/CIP library).

| Column in CSV | Purpose |
|---------------|--------|
| **Name** | File name (used for comparison) ✓ |
| **FileRef** | Full path in SharePoint |
| **FileExtension** | File type |
| **Length** | File size |
| **TimeCreated** / **TimeLastModified** | Dates |
| **Author** / **Editor** | People |

**Conclusion:** SharePoint side is “all files in the CIM/CIP document library.” That matches the intent of comparing CIM/CIP documents between Salesforce and SharePoint. The **Name** column is the correct comparison key.

---

## 4. Comparison logic (what we do)

| Aspect | Implementation | Correct for your spec? |
|--------|-----------------|---------------------------|
| **Match key** | File name: Salesforce **Document Title** ↔ SharePoint **Name** | ✓ Yes — “same document” = same file name. |
| **Case** | Case-insensitive by default | ✓ Reduces false “missing” due to casing. |
| **Output “Salesforce missing”** | Files in SharePoint that are not in Salesforce (**only_in_sharepoint__give_to_coworker.csv**) | ✓ These are the ones to give your coworker. |
| **Output “SharePoint missing”** | Files in Salesforce that are not in SharePoint | ✓ For your side to backfill or reconcile. |
| **In both** | Same file name in both CSVs | ✓ Present in both systems. |

So the comparison is **accurate** for:

- Identifying which **file names** exist in one system but not the other.
- Producing the list of **files SharePoint has that Salesforce is missing** (for your coworker).

---

## 5. What we cannot verify from the CSVs alone

- That the Salesforce export was built **only** from:
  - `HIG_Document__c` joined to `Opportunity` on `Opportunity__c`
  - With `Doc_Type__c = 'CIM'` and `Opportunity__c != NULL`
  - And `Opportunity.CreatedDate >= 2021-01-01`
- Whether the export is from **ContentDocument** vs **HIG_Document__c** (your spec shows HIG_Document__c; the CSV has ContentDocumentId). That’s a data-source choice; the comparison is still valid on **file name** as long as this export is the one you want to use as “Salesforce CIM list.”

---

## 6. Summary

| Question | Answer |
|----------|--------|
| Do both CSVs have the right columns for comparison? | **Yes** — Document Title (SF) and Name (SP). |
| Is the comparison key (file name) correct? | **Yes** — same file = same name in both. |
| Does the SF CSV align with “CIM names + opp name”? | **Yes** — it has document titles and Opp Name, Fund, Opp Owner. |
| Can we prove the CIM/Opp filters were applied in SF? | **No** — not from the CSV; ensure the export/report uses `Doc_Type__c = 'CIM'` and `Opportunity__c != NULL` (and Opp CreatedDate ≥ 2021 if required). |
| Is the SharePoint scope correct? | **Yes** — CIM/CIP document library. |

**Bottom line:** The Salesforce and SharePoint CSVs and the current comparison logic are **aligned with what you’re looking for**: we compare by file name and correctly identify what Salesforce is missing (SharePoint-only list) and what SharePoint is missing (Salesforce-only list). The only thing to confirm outside this repo is that the **source** of `sfdc_2025_uspe_cims.csv` applies the HIG_Document__c/Opportunity filters you specified.
