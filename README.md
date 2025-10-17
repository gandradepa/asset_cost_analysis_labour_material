# TSBC Website Data — README

This repository contains small data-processing scripts used to prepare TSBC unit and permit data and export a consolidated Excel file for the TSBC website.

## Files of interest

- `tsbc_data.py` — Script that:
  - reads unit data from: `S:\\MaintOpsPlan\\WorkMgt\\SharePoint Dashboards\\DATA\\Asset_Management\\asset_data\\tsbc_units\\full-table-data.xlsx` (sheet `data`) into `tsbc_unit`.
  - reads permit data from: `S:\\MaintOpsPlan\\WorkMgt\\SharePoint Dashboards\\DATA\\Asset_Management\\asset_data\\tsbc_permits\\full-table-data.xlsx` (sheet `data`) into `tsbc_permit`.
  - normalizes all `tsbc_unit` columns to string type and converts column names to UPPERCASE for both datasets.
  - renames specific `tsbc_unit` columns (examples: `Site Address ` → `SITE ADDRESS`, `Classification` → `UNIT CLASS`).
  - adds an empty `MCP START DATE` column to `tsbc_unit`.
  - inserts an empty `WORK CLASS` column into `tsbc_permit`, renames `STATUS` → `PERMIT STATUS`, and adds an empty `DESCRIPTION` column to `tsbc_permit`.
  - merges `tsbc_permit` (left) with `tsbc_unit` (right) using `PERMIT NUMBER` (left) and `OPERATING PERMIT NUMBER` (right) producing `TSBC_website_data`.
  - drops duplicate/unused columns (e.g., `SITE ADDRESS_y`, `DESCRIPTION`, `UNIT NAME`) and renames `SITE ADDRESS_x` → `SITE ADDRESS`.
  - reorders columns to a TSBC website-friendly sequence and filters out permits where `PERMIT STATUS` equals `Closed`.
  - writes the final dataset to Excel: `S:\\MaintOpsPlan\\WorkMgt\\SharePoint Dashboards\\DATA\\Asset_Management\\asset_data\\TSBC_website_data.xlsx`.

Notes:
- The script expects column names in the inputs to match the names used in the script (e.g., `WO #`, `Reported on`, `Asset Code`, `Amount`, `Cost Occurred on`, `Property`).


**Notes:**
- The script uppercases all column names so joins and lookups are case-insensitive with respect to original capitalization.

## Data contract (inputs)

- Units workbook: `tsbc_units/full-table-data.xlsx` (sheet `data`). Key columns used/created in the script include:
  - `SITE ADDRESS`, `SITE BUILDING NAME`, `UNIT NAME`, `UNIT CLASS`, `UNIT NUMBER`, `UNIT STATUS`, `OPERATING PERMIT NUMBER`, `LAST INSPECTION DATE`, `MCP START DATE` (added)
- Permits workbook: `tsbc_permits/full-table-data.xlsx` (sheet `data`). Key columns used/created include:
  - `SITE ADDRESS`, `SITE/BUILDING NAME`, `PERMIT TYPE`, `WORK CLASS` (added), `PERMIT STATUS` (renamed from `STATUS`), `ISSUED`, `EXPIRY`, `PERMIT NUMBER`, `DESCRIPTION` (added)

## Output

- `TSBC_website_data.xlsx` — Excel workbook written to `S:\\MaintOpsPlan\\WorkMgt\\SharePoint Dashboards\\DATA\\Asset_Management\\asset_data\\TSBC_website_data.xlsx` containing the consolidated dataset with columns ordered for TSBC website ingestion.

**Expected (example) column order in the output:**

1. SITE ADDRESS
2. SITE BUILDING NAME
3. UNIT CLASS
4. UNIT NUMBER
5. UNIT STATUS
6. OPERATING PERMIT NUMBER
7. LAST INSPECTION DATE
8. MCP START DATE
9. PERMIT TYPE
10. WORK CLASS
11. PERMIT STATUS
12. ISSUED
13. EXPIRY
14. PERMIT NUMBER

## Requirements

- Python 3.8+ (tested with 3.8–3.13)
- pandas
- openpyxl (for reading/writing .xlsx files)

**Install dependencies (Windows PowerShell):**

```powershell
python -m pip install --upgrade pip; \\
python -m pip install pandas openpyxl
```

## How to run `tsbc_data.py`

Open PowerShell on a machine with access to the `S:` network drive and run:

```powershell
python "s:\\MaintOpsPlan\\AssetMgt\\Asset Management Process\\Database\\1. Cost_Analisys\\01. Cost_files\\04. Billing files\\tsbc_data.py"
```

The script uses hard-coded network paths defined near the top of the file. If you want to run it from another machine or user, ensure the `S:` drive is mapped and accessible and that the file paths remain valid.

## Behavior, assumptions and edge cases

- All `tsbc_unit` columns are coerced to string using `astype(str)` to avoid dtype mismatches during the merge.
- Column names are uppercased; the script expects input column names that match known labels after uppercasing.
- Rows with `PERMIT STATUS == 'Closed'` are filtered out prior to export.

## Troubleshooting

- Permission errors: ensure read access to input files and write access to the output folder.
- Missing dependencies: `ImportError` indicates pandas or openpyxl are not installed.
- Column/key mismatches: run the script interactively and print `tsbc_unit.columns` and `tsbc_permit.columns` to debug any unexpected column names.

## Quick improvements / next steps

- Parameterize file paths via CLI args or a small config file (JSON/YAML).
- Add logging and error handling for missing/invalid input files.
- Add a small test harness that runs the script on a tiny sample dataset.

## Contact / author

Maintainers of this workspace (update with a responsible contact or team).

---

Generated based on the current `tsbc_data.py` script.
