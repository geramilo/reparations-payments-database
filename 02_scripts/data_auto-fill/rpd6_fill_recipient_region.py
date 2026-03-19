"""
Script name: rpd6_fill_recipient_region.py

Description:
This script assigns the recipient region for each recipient country in the RPD, using data from the dataset “Countries and Economies Data.xlsx”.

Author: Geraldine Ramilo  
Project: REPAIR  

Notes:
- Developed with AI-assisted support (ChatGPT)  
- Reviewed and tested by the author  
- Preserves existing Excel formatting  
- Creates a new output file and does not overwrite the original database  
- Please ensure that both the input file and the lookup file are located in the same directory as this script before execution  
"""


import openpyxl
import pandas as pd

# -----------------------------
# File names
# -----------------------------
input_file = "RPD (V1).xlsx"
lookup_file = "Countries and Economies Data.xlsx"
output_file = "RPD (V1) - updated.xlsx"

# -----------------------------
# Load main workbook (keeps formatting)
# -----------------------------
wb_main = openpyxl.load_workbook(input_file)
ws_main = wb_main["DATABASE"]

# -----------------------------
# Load lookup file with pandas
# -----------------------------
df_lookup = pd.read_excel(lookup_file, sheet_name="2024")

# Clean column names
df_lookup.columns = df_lookup.columns.str.strip()

# Create dictionary: Country -> Region
country_to_region = dict(
    zip(
        df_lookup["Entity (Country or economy)"].astype(str).str.strip(),
        df_lookup["Region"]
    )
)

# -----------------------------
# Find column indexes dynamically
# -----------------------------
def find_column_index(sheet, column_name):
    for cell in sheet[1]:
        if cell.value == column_name:
            return cell.column
    return None

main_country_col = find_column_index(ws_main, "Recipient country or economy")
main_region_col = find_column_index(ws_main, "Recipient region")

if not all([main_country_col, main_region_col]):
    raise ValueError("Required columns not found in main file.")

# -----------------------------
# Fill Region column
# -----------------------------
for row in ws_main.iter_rows(min_row=2):
    country_cell = row[main_country_col - 1]
    region_cell = row[main_region_col - 1]

    country_name = country_cell.value

    if country_name:
        key = str(country_name).strip()
        region_value = country_to_region.get(key, "Not applicable")
    else:
        region_value = "Not applicable"

    region_cell.value = region_value

# -----------------------------
# Save output (preserves formatting)
# -----------------------------
wb_main.save(output_file)

print("Done! File saved as:", output_file)
