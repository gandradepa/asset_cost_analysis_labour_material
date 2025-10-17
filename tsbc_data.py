import pandas as pd
import os

# --- INPUT PATHS (UNC) ---
file_path_unit   = r"\\files.ubc.ca\team\bops\MaintOpsPlan\WorkMgt\SharePoint Dashboards\DATA\Asset_Management\asset_data\tsbc_units\full-table-data.xlsx"
file_path_permit = r"\\files.ubc.ca\team\bops\MaintOpsPlan\WorkMgt\SharePoint Dashboards\DATA\Asset_Management\asset_data\tsbc_permits\full-table-data.xlsx"

# Optional: clearer error if a file can't be seen by Power BI/Python
for p in (file_path_unit, file_path_permit):
    if not os.path.exists(p):
        raise FileNotFoundError(f"Path not found (check UNC, access, and filename): {p}")

# --- READ EXCEL ---
tsbc_unit = pd.read_excel(file_path_unit, sheet_name='data', engine='openpyxl')
tsbc_permit = pd.read_excel(file_path_permit, sheet_name='data', engine='openpyxl')

# --- UNITS CLEANUP ---
tsbc_unit = tsbc_unit.astype(str)
tsbc_unit.rename(columns={
    "Site Address ": "SITE ADDRESS",
    "Site/Building Name": "SITE BUILDING NAME",
    "Classification": "UNIT CLASS",
    "Unit Number ": "UNIT NUMBER",
    "Status": "UNIT STATUS",
    "Operating Permit Number": "OPERATING PERMIT NUMBER",
    "Last Inspection Date": "LAST INSPECTION DATE"
}, inplace=True)
tsbc_unit.columns = tsbc_unit.columns.str.upper()
tsbc_unit['MCP START DATE'] = ''

if 'UNIT NAME' in tsbc_unit.columns:
    unit_name_col = tsbc_unit.pop('UNIT NAME')
    tsbc_unit['UNIT NAME'] = unit_name_col

# --- PERMITS CLEANUP ---
tsbc_permit.columns = tsbc_permit.columns.str.upper()
if 'PERMIT TYPE' in tsbc_permit.columns and 'ISSUED' in tsbc_permit.columns:
    issued_index = tsbc_permit.columns.get_loc('ISSUED')
    tsbc_permit.insert(issued_index, 'WORK CLASS', '')

tsbc_permit.rename(columns={'STATUS': 'PERMIT STATUS'}, inplace=True)

if 'PERMIT STATUS' in tsbc_permit.columns:
    permit_status_col = tsbc_permit.pop('PERMIT STATUS')
    if 'WORK CLASS' in tsbc_permit.columns and 'ISSUED' in tsbc_permit.columns:
        issued_index = tsbc_permit.columns.get_loc('ISSUED')
        tsbc_permit.insert(issued_index, 'PERMIT STATUS', permit_status_col)
    else:
        tsbc_permit['PERMIT STATUS'] = permit_status_col

tsbc_permit['DESCRIPTION'] = ''

# --- JOIN ---
TSBC_website_data = pd.merge(
    tsbc_permit,
    tsbc_unit,
    left_on='PERMIT NUMBER',
    right_on='OPERATING PERMIT NUMBER',
    how='left'
)

# --- POST-JOIN CLEANUP ---
if 'SITE ADDRESS_y' in TSBC_website_data.columns:
    TSBC_website_data.drop('SITE ADDRESS_y', axis=1, inplace=True)
TSBC_website_data.rename(columns={'SITE ADDRESS_x': 'SITE ADDRESS'}, inplace=True)
if 'SITE/BUILDING NAME' in TSBC_website_data.columns:
    TSBC_website_data.drop('SITE/BUILDING NAME', axis=1, inplace=True)
for col in ['DESCRIPTION', 'UNIT NAME']:
    if col in TSBC_website_data.columns:
        TSBC_website_data.drop(col, axis=1, inplace=True)

desired_order = [
    'SITE ADDRESS', 'SITE BUILDING NAME', 'UNIT CLASS', 'UNIT NUMBER', 'UNIT STATUS',
    'OPERATING PERMIT NUMBER', 'LAST INSPECTION DATE', 'MCP START DATE', 'PERMIT TYPE',
    'WORK CLASS', 'PERMIT STATUS', 'ISSUED', 'EXPIRY', 'PERMIT NUMBER'
]
existing_columns = [c for c in desired_order if c in TSBC_website_data.columns]
TSBC_website_data = TSBC_website_data[existing_columns]

if 'PERMIT STATUS' in TSBC_website_data.columns:
    TSBC_website_data = TSBC_website_data[TSBC_website_data['PERMIT STATUS'] != 'Closed']

# IMPORTANT: Power BI will expose this DataFrame in the Navigator
TSBC_website_data
