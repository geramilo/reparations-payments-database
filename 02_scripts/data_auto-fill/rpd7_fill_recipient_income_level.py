"""
Script name: rpd7_fill_recipient_income_level.py

Description:
This script assigns the recipient income level for each recipient country in the RPD, using data from the dataset “Countries and Economies Data.xlsx”.

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

# Create dictionary: Country -> Income Level
country_to_income = dict(
    zip(
        df_lookup["Entity (Country or economy)"].astype(str).str.strip(),
        df_lookup["Income level classification"]
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
main_income_col = find_column_index(ws_main, "Recipient income level")

if not all([main_country_col, main_income_col]):
    raise ValueError("Required columns not found in main file.")

# -----------------------------
# Fill Income Level column
# -----------------------------
for row in ws_main.iter_rows(min_row=2):
    country_cell = row[main_country_col - 1]
    income_cell = row[main_income_col - 1]

    country_name = country_cell.value

    if country_name:
        key = str(country_name).strip()
        income_value = country_to_income.get(key, "Not applicable")
    else:
        income_value = "Not applicable"

    income_cell.value = income_value

# -----------------------------
# Save output (preserves formatting)
# -----------------------------
wb_main.save(output_file)

print("Done! File saved as:", output_file)
