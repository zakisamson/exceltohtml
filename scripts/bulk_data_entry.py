import pandas as pd
from openpyxl import load_workbook

# File paths
excel_file = "./data/data.xlsx"    # Existing sales orders file
csv_file = "./data/csv_example.csv"  # New data from an external source
sheet_name = "SalesOrders"  # The target sheet

# Load the new sales orders from CSV
new_orders_df = pd.read_csv(csv_file)

# Load the existing Excel file
wb = load_workbook(excel_file)
ws = wb[sheet_name]

# Append each new order as a row in the Excel sheet
for row in new_orders_df.itertuples(index=False, name=None):
    ws.append(row)

# Save the updated Excel file
wb.save(excel_file)
print("✅ New orders added successfully!")
