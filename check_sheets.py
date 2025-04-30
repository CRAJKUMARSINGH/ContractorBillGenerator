import pandas as pd

# Check sheet names in one of the test files
file_path = "test_files/PRIYANKA SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx"

# Read the Excel file
with pd.ExcelFile(file_path) as xls:
    print("Available sheets:", xls.sheet_names)
