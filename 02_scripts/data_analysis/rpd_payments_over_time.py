"""
Script name: rpd_payments_over_time.py

Description:
This script analyzes the RPD dataset, calculating total payments adjusted to 2024 USD. 
It expands multi-year payments to individual years, aggregates yearly totals, identifies the top 5 years by total 
payments, writes a summary to a text file, and produces a bar + line chart showing reparations paid per year.

Columns used in the analysis:
- "Start year of reparations payment"
- "End year of reparations payment"
- "Paid (Total USD 2024)"
- "Programme name"

Author: Geraldine Ramilo  
Project: REPAIR  

Notes:
- Developed with AI-assisted support (ChatGPT)  
- Reviewed and tested by the author  
- Requires the Excel file 'RPD (V1).xlsx' to be present in the working directory  
- Drops rows with missing or invalid numeric data in the required columns  
- Only includes payments greater than zero  
- Expands multi-year payments evenly across all years in the payment range  
- Aggregates yearly totals and identifies the top 5 years by total payments  
- Creates a text summary file 'top_years_note.txt' listing top years and contributing programmes  
- Generates a combined bar and line plot of total payments per year  
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
cols = ["Start year of reparations payment", "End year of reparations payment", "Paid (Total USD 2024)", "Programme name"]
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
    
    years = list(range(start, end + 1))
    amount_per_year = total / len(years)
    
    for year in years:
        records.append({"Year": year, "Amount": amount_per_year, "Programme": programme})

# === STEP 4: Aggregate per year ===
yearly_df = pd.DataFrame(records)
sum_per_year = yearly_df.groupby("Year", as_index=False)["Amount"].sum()

# === STEP 5: Save top years info to file, including programmes ===
top_years = sum_per_year.sort_values(by="Amount", ascending=False).head(5)

def format_money(x):
    if x >= 1e9:
        return f"$ {x/1e9:.2f} B"
    elif x >= 1e6:
        return f"$ {x/1e6:.2f} M"
    elif x >= 1e3:
        return f"$ {x/1e3:.2f} K"
    else:
        return f"$ {x:.2f}"

with open("top_years_note.txt", "w") as f:
    f.write("Top 5 Years by Total Reparations (Adjusted to 2024 USD)\n")
    f.write("="*65 + "\n")
    for _, row in top_years.iterrows():
        year = int(row['Year'])
        amount = row['Amount']
        formatted_amount = format_money(amount)
        # Get programmes contributing to this year
        programmes = yearly_df[yearly_df["Year"] == year]["Programme"].unique()
        f.write(f"{year}: {formatted_amount}\n")
        f.write(f"Programmes: {', '.join(programmes)}\n\n")

print("✅ File 'top_years_note.txt' created with top 5 years and their programmes.")

# === STEP 6: Plot the bar + line chart ===
plt.figure(figsize=(12, 6))
plt.bar(sum_per_year["Year"], sum_per_year["Amount"], color='skyblue', edgecolor='black')
plt.plot(sum_per_year["Year"], sum_per_year["Amount"], color='darkblue', linewidth=2)

plt.title("Reparations paid per year (Adjusted to 2024 USD)")
plt.xlabel("Year")

# Custom Y-axis formatting
formatter = FuncFormatter(lambda x, _: format_money(x))
plt.gca().yaxis.set_major_formatter(formatter)
plt.ylabel("Total payments")

plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show()
