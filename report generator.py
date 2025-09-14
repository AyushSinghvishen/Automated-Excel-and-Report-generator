import os
import pandas as pd

# Folder with Excel files
folder = "excel_reports"

# List to store data from all files
all_data = []

# Read all Excel files from the folder
for file in os.listdir(folder):
    if file.endswith(".xlsx"):
        path = os.path.join(folder, file)
        df = pd.read_excel(path)
        all_data.append(df)

# Combine all data
data = pd.concat(all_data, ignore_index=True)

# Clean column names
data.columns = [col.strip().lower() for col in data.columns]

# Calculate summary
total_sales = data["sales"].sum()
total_expenses = data["expenses"].sum()
summary = pd.DataFrame({
    "Metric": ["Total Sales", "Total Expenses"],
    "Value": [total_sales, total_expenses]
})

# Save both full data and summary into one Excel file
with pd.ExcelWriter("final_report.xlsx", engine='openpyxl') as writer:
    data.to_excel(writer, sheet_name="Merged Data", index=False)
    summary.to_excel(writer, sheet_name="Summary", index=False)

print("✅ Excel report saved as 'final_report.xlsx'")
