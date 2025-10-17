import os
import glob
import pandas as pd
from datetime import datetime

'''Table of contents:

	1. Part I: Search for all files starting with "UBC - BOps Labour & Material Cost Analysis"
	2. Part II: Load the property_info dataset
	3. Part III: Load the "UBC - Work Orders - Cost Detail (ungrouped) (BRM)*.xlsx" dataset
    4. Part IV: Create a bill date dataset: "BOps Labour & Material Cost Analysis_full_info.csv" '''

# Define the network paths
network_path_bill = r"S:\MaintOpsPlan\AssetMgt\Asset Management Process\Database\1. Cost_Analisys\03. Bill Type\bill_labor_material_date\Current_files"
network_path_property = r"S:\MaintOpsPlan\AssetMgt\Asset Management Process\Database\4. Asset all data\02. Up to date file\UBC - All Properties List with GPS Coordinates.xlsx"
network_path_brm = r"S:\MaintOpsPlan\AssetMgt\Asset Management Process\Database\1. Cost_Analisys\01. Cost_files\01. Cost by BRM"

# PART I: Search for all files starting with "UBC - BOps Labour & Material Cost Analysis"
file_pattern = os.path.join(network_path_bill, "UBC - BOps Labour & Material Cost Analysis*.xlsx")
excel_files = glob.glob(file_pattern)

dfs = []
for file in excel_files:
    df = pd.read_excel(file, dtype={'Asset Code': str})  # Ensure 'Asset Code' is read as string
    dfs.append(df)

bill = pd.concat(dfs, ignore_index=True)

# Convert relevant columns to appropriate formats
bill['WO #'] = bill['WO #'].astype(str)
bill['Reported on'] = pd.to_datetime(bill['Reported on'], errors='coerce').dt.strftime('%m/%d/%Y')
bill['Asset Code'] = bill['Asset Code']  # Keep the original formatting of 'Asset Code'
bill['Amount'] = pd.to_numeric(bill['Amount'], errors='coerce')
bill['Cost Occurred on'] = pd.to_datetime(bill['Cost Occurred on'], errors='coerce')

# Calculate Fiscal Year based on the given formula
bill['Fiscal Year'] = bill['Cost Occurred on'].apply(
    lambda x: str(x.year if x.month >= 4 else x.year - 1) if pd.notnull(x) else None
)

bill = bill[bill['WO #'] != 'WO #']

# PART II: Load the property_info dataset
property_info = pd.read_excel(network_path_property)
property_info = property_info[['Name', 'Zone', 'FM', 'Geo Zone', 'GPS Coordinates', 'Owner Rep']]
bill_asset_info = pd.merge(bill, property_info, left_on='Property', right_on='Name', how='left')

# PART III: Load the "UBC - Work Orders - Cost Detail (ungrouped) (BRM)*.xlsx" dataset
brm_file_pattern = os.path.join(network_path_brm, "UBC - Work Orders - Cost Detail (ungrouped) (BRM)*.xlsx")
brm_files = glob.glob(brm_file_pattern)

brm_dfs = []
for file in brm_files:
    brm_df = pd.read_excel(file)
    brm_dfs.append(brm_df)

brm_cost = pd.concat(brm_dfs, ignore_index=True)
# Ensure 'Order #' is properly converted to string
brm_cost['Order #'] = brm_cost['Order #'].apply(lambda x: f"{x:.2f}" if isinstance(x, (float, int)) else str(x))
brm_cost = brm_cost[['Order #', 'Order Type']]

# PART IV: Create a bill date dataset: "BOps Labour & Material Cost Analysis_full_info.csv"
merged_data = pd.merge(bill_asset_info, brm_cost, left_on='WO #', right_on='Order #', how='left')

# Ensure 'Cost Occurred on' in merged_data is date type and formatted as MM/DD/YYYY
merged_data['Cost Occurred on'] = pd.to_datetime(merged_data['Cost Occurred on'], errors='coerce').dt.strftime('%m/%d/%Y')

# Add the Last_update column with the current date
creation_date = datetime.now().strftime('%m-%d-%Y')
merged_data['Last_update'] = creation_date

# Add the Last_bill_date column with the maximum date from 'Cost Occurred on'
max_date = pd.to_datetime(merged_data['Cost Occurred on'], errors='coerce').max().strftime('%m-%d-%Y')
merged_data['Last_bill_date'] = max_date

# Change output file format to CSV
output_path = os.path.join(network_path_bill, "BOps Labour & Material Cost Analysis_full_info.csv")
merged_data.to_csv(output_path, index=False)

print(merged_data)
