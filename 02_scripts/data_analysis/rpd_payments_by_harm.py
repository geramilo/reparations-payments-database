"""
Script name: rpd_spayments_by_harm.py

Description:
This script analyzes the RPD dataset of reparations programs and visualizes annual payments by the type of harm. 
It expands multi-year payments to individual years, aggregates totals by "Main harm," identifies the top 3 harms 
per year, groups the rest as "Other," and produces a stacked bar chart with a line showing total payments per year.

Columns used in the analysis:
- "Start year of reparations payment"
- "End year of reparations payment"
- "Paid (Total USD 2024)"
- "Programme name"
- "Main harm"

Author: Geraldine Ramilo  
Project: REPAIR  

Notes:
- Developed with AI-assisted support (ChatGPT)  
- Reviewed and tested by the author  
- Requires the Excel file 'RPD (V1).xlsx' to be present in the working directory  
- Drops rows with missing or invalid numeric data in the required columns  
- Only includes payments greater than zero  
- Expands multi-year payments evenly across all years in the payment range  
- Aggregates payments per year by "Main harm" and identifies top 3 harms per year  
- All other harms are grouped under "Other"  
- Produces a stacked bar chart of top harms and a line chart of total payments  
- Uses a consistent color palette for harms (tab20)  
- Fully portable to any system with Python, pandas, matplotlib, and the dataset  
"""

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

# === STEP 1: Load your data ===
file_path = "RPD (V1).xlsx"
sheet_name = "DATABASE"
df = pd.read_excel(file_path, sheet_name=sheet_name)

# === STEP 2: Clean and prepare the data ===
cols = ["Start year of reparations payment", "End year of reparations payment", 
        "Paid (Total USD 2024)", "Programme name", "Main harm"]
for col in cols:
    if col not in df.columns:
        raise ValueError(f"Missing column: {col}")

# Convert to numeric
df["Start year of reparations payment"] = pd.to_numeric(df["Start year of reparations payment"], errors="coerce")
df["End year of reparations payment"] = pd.to_numeric(df["End year of reparations payment"], errors="coerce")
df["Paid (Total USD 2024)"] = pd.to_numeric(df["Paid (Total USD 2024)"], errors="coerce")

# Drop invalid rows
df = df.dropna(subset=cols)
df = df[df["Paid (Total USD 2024)"] > 0]

# === STEP 3: Expand payments across years ===
records = []
for _, row in df.iterrows():
    start = int(row["Start year of reparations payment"])
    end = int(row["End year of reparations payment"])
    total = float(row["Paid (Total USD 2024)"])
    programme = row["Programme name"]
    harm = row["Main harm"]
    
    years = list(range(start, end + 1))
    amount_per_year = total / len(years)
    
    for year in years:
        records.append({"Year": year, "Amount": amount_per_year, 
                        "Programme": programme, "Main harm": harm})

yearly_df = pd.DataFrame(records)

# === STEP 4: Prepare data for stacked bar by top 3 harms + Other ===
stacked_data = []

for year, group in yearly_df.groupby("Year"):
    harm_sum = group.groupby("Main harm")["Amount"].sum().sort_values(ascending=False)
    top_harms = harm_sum.head(3)
    other_sum = harm_sum[3:].sum()
    
    row = top_harms.to_dict()
    row["Other"] = other_sum
    row["Year"] = year
    stacked_data.append(row)

stacked_df = pd.DataFrame(stacked_data).fillna(0).set_index("Year").sort_index()

# === STEP 5: Plot stacked bar chart ===
plt.figure(figsize=(14, 7))
colors = plt.cm.tab20.colors  # nice color palette

# Keep consistent color mapping
harms = stacked_df.columns.tolist()
for i, harm in enumerate(harms):
    if i == 0:
        plt.bar(stacked_df.index, stacked_df[harm], color=colors[i], label=harm)
        bottom = stacked_df[harm].copy()
    else:
        plt.bar(stacked_df.index, stacked_df[harm], bottom=bottom, color=colors[i], label=harm)
        bottom += stacked_df[harm]

# === STEP 6: Add line for total payments ===
total_per_year = yearly_df.groupby("Year")["Amount"].sum()
plt.plot(total_per_year.index, total_per_year.values, color='black', linewidth=2, label="Total payments")

# === STEP 7: Y-axis formatting ===
def format_money(x):
    if x >= 1e9:
        return f"$ {x/1e9:.2f} B"
    elif x >= 1e6:
        return f"$ {x/1e6:.2f} M"
    elif x >= 1e3:
        return f"$ {x/1e3:.2f} K"
    else:
        return f"$ {x:.2f}"

formatter = FuncFormatter(lambda x, _: format_money(x))
plt.gca().yaxis.set_major_formatter(formatter)

# === STEP 8: Labels and legend ===
plt.title("")
plt.xlabel("Year")
plt.ylabel("Estimated annual reparations flows")
plt.grid(axis='y', linestyle='--', alpha=0.7)

# Legend below chart
plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.12), ncol=3, fontsize=11)
plt.tight_layout()
plt.show()
