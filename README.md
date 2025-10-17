## BOps Labour & Material Cost Analysis — README

This repository contains a small data-processing script that collects labour and material cost records from Excel files, enriches them with property metadata and BRM work-order types, and writes a consolidated CSV summary.

### Files

- `asset_cost_labour_material.py` — Main script. It:
  - finds Excel files matching: `UBC - BOps Labour & Material Cost Analysis*.xlsx` in the configured `network_path_bill` folder,
  - reads a property list Excel workbook (`UBC - All Properties List with GPS Coordinates.xlsx`) to attach property metadata,
  - reads BRM cost detail files matching `UBC - Work Orders - Cost Detail (ungrouped) (BRM)*.xlsx` to attach order types,
  - merges datasets, computes fiscal year, formats dates, adds metadata columns, and writes `BOps Labour & Material Cost Analysis_full_info.csv`.

### Data contract (inputs)

- Labour & Material Excel files: any files in `network_path_bill` that start with `UBC - BOps Labour & Material Cost Analysis`.
- Property metadata: `UBC - All Properties List with GPS Coordinates.xlsx` (expected columns: `Name`, `Zone`, `FM`, `Geo Zone`, `GPS Coordinates`, `Owner Rep`).
- BRM work orders: files in `network_path_brm` matching `UBC - Work Orders - Cost Detail (ungrouped) (BRM)*.xlsx` (expected columns: `Order #`, `Order Type`).

Notes:
- The script expects column names in the inputs to match the names used in the script (e.g., `WO #`, `Reported on`, `Asset Code`, `Amount`, `Cost Occurred on`, `Property`).

### Output

- `BOps Labour & Material Cost Analysis_full_info.csv` — consolidated CSV written to the `network_path_bill` folder. It contains merged fields from the bill files, property metadata, BRM order type, computed `Fiscal Year`, `Last_update` (script run date), and `Last_bill_date` (latest cost date).

### Requirements

- Python 3.8+ (tested with Python 3.8–3.11)
- pandas
- openpyxl (for reading .xlsx files)

Install dependencies (Windows PowerShell):

```powershell
python -m pip install --upgrade pip; \
python -m pip install pandas openpyxl
```

### How to run

Open PowerShell on a machine that has access to the `S:` network drive and run:

```powershell
python "s:\\MaintOpsPlan\\AssetMgt\\Asset Management Process\\Database\\1. Cost_Analisys\\01. Cost_files\\04. Billing files\\asset_cost_labour_material.py"
```

The script uses hard-coded network paths defined near the top of the file. If you want to run it from another machine or user, ensure the `S:` drive is mapped and accessible and that the file paths remain valid.

### Behavior, assumptions and edge cases

- Date parsing: the script coercively parses `Cost Occurred on` and formats to `MM/DD/YYYY`. Rows with unparseable dates will become NaT and may affect `Fiscal Year` and `Last_bill_date` calculations.
- The `WO #` column is normalized to string; the BRM `Order #` is converted to string (floats formatted with 2 decimals in the script). Mismatched formats between the invoice and BRM files may cause join misses.
- If no matching input files are found, the script may raise an error during concat or write an empty CSV. Confirm that the file patterns and network paths are correct.
- The script expects the property metadata file to contain a `Name` column to join on `Property`.

### Troubleshooting

- Permission errors: ensure the user running the script has read access to the Excel files and write access to the output folder.
- Missing dependencies: `ImportError` indicates pandas or openpyxl are not installed.
- Column/key mismatches: inspect input files for the exact column names used by the script (case-sensitive). Consider adding a small wrapper to log the columns found when debugging.

### Quick improvements / next steps

- Parameterize file paths and patterns via CLI args or a small config file (JSON or YAML).
- Add logging and error handling for better observability.
- Add unit tests for the merge logic and a small example CSV/XLSX fixture to validate behavior.

### Contact / author

Maintainers of this workspace (see repository owner) — update this section with a responsible contact or team.

---

Generated based on the contents of `asset_cost_labour_material.py`.
