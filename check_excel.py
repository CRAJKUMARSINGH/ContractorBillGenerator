import pandas as pd

def check_excel_file(file_path):
    """Check the structure of the Excel file"""
    print(f"\nChecking Excel file: {file_path}")
    
    with pd.ExcelFile(file_path) as xls:
        print("\nAvailable sheets:", xls.sheet_names)
        
        for sheet_name in xls.sheet_names:
            print(f"\n=== Sheet: {sheet_name} ===")
            df = pd.read_excel(xls, sheet_name=sheet_name)
            print("\nFirst few rows:")
            print(df.head())
            print("\nColumns:", df.columns.tolist())
            print("\nData types:")
            print(df.dtypes)

if __name__ == '__main__':
    check_excel_file("test_files/SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx")
