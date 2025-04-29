import pandas as pd
import numpy as np
import os

# Create a directory for test files if it doesn't exist
os.makedirs('test_files', exist_ok=True)

# Create a sample Excel file with the expected structure
# First 21 rows will be headers (to be skipped)
# Data starts from row 22

# Create header data (first 21 rows)
header_data = [
    ["BILL OF QUANTITIES"],
    ["Name of Contractor or supplier :", "M/s Seema Electrical Udaipur"],
    ["Name of Work ;-", "Electric Repair and MTC work at Govt. Ambedkar hostel Ambamata, Govardhanvilas, Udaipur"],
    ["Serial No. of this bill :", "First & Final Bill"],
    ["No. and date of the last bill-", "Not Applicable"],
    ["Reference to work order or Agreement :", "1179 Dt. 09-01-2025"],
    ["Agreement No.", "48/2024-25"],
    ["Date of written order to commence work :", "09-01-2025"],
    ["St. date of Start :", "18-01-2025"],
    ["St. date of completion :", "17-04-2025"],
    ["Date of actual completion of work :", "01-03-2025"],
    ["Date of measurement :", "03-03-2025"],
    ["WORK ORDER AMOUNT RS.", "854678"],
    [""],  # Empty row
    [""],  # Empty row
    ["ITEM DETAILS"],
    [""],  # Empty row
    [""],  # Empty row
    [""],  # Empty row
    [""],  # Column headers will be handled in our code
    [""]   # Empty row before data starts
]

# Create bill data
bill_data = []

# Add some sample items
bill_data.append([1, "LED lights installation - 20W", "Nos", 10, 1200, 12, 1200, 12000, 14400])
bill_data.append([2, "Electrical wiring - 2.5 sq mm", "Meter", 500, 45, 520, 45, 22500, 23400])
bill_data.append([3, "Switch board installation", "Nos", 15, 850, 15, 850, 12750, 12750])
bill_data.append([4, "Earthing work", "Job", 2, 3500, 2, 3500, 7000, 7000])
bill_data.append([5, "Ceiling fan installation", "Nos", 8, 650, 10, 650, 5200, 6500])
bill_data.append([6, "MCB installation - 32A", "Nos", 6, 450, 5, 450, 2700, 2250])
bill_data.append([7, "Panel board installation", "Nos", 1, 15000, 1, 15000, 15000, 15000])
bill_data.append([8, "Conduit pipe installation", "Meter", 300, 35, 350, 35, 10500, 12250])
bill_data.append([9, "Cable laying - 4 sq mm", "Meter", 200, 65, 180, 65, 13000, 11700])
bill_data.append([10, "Light fixtures installation", "Nos", 25, 350, 28, 350, 8750, 9800])

# Create the full Excel file
with pd.ExcelWriter('test_files/sample_bill.xlsx', engine='openpyxl') as writer:
    # Write header (rows 1-21)
    header_df = pd.DataFrame(header_data)
    header_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
    
    # Write bill data starting from row 22
    columns = ['S_No', 'Item_Description', 'Unit', 'Work_Order_Qty', 'Work_Order_Rate', 
              'Bill_Qty', 'Bill_Rate', 'Work_Order_Amount', 'Bill_Amount']
    bill_df = pd.DataFrame(bill_data, columns=columns)
    bill_df.to_excel(writer, sheet_name='Sheet1', startrow=21, index=False)

print("Sample Excel file created at: test_files/sample_bill.xlsx")

# Create another sample with different data (more items and missing values)
bill_data2 = []
for i in range(1, 21):
    # Generate some random data with realistic values
    desc = f"Item {i} - Electrical work {i}"
    unit = np.random.choice(["Nos", "Meter", "Job", "Set", "Piece"])
    work_order_qty = np.random.randint(1, 100)
    work_order_rate = np.random.randint(50, 5000)
    
    # Some items have different bill quantities (to test deviations)
    if i % 3 == 0:
        bill_qty = work_order_qty + np.random.randint(-5, 10)
    else:
        bill_qty = work_order_qty
    
    # Make sure bill_qty is not negative
    bill_qty = max(0, bill_qty)
    
    bill_rate = work_order_rate
    work_order_amount = work_order_qty * work_order_rate
    bill_amount = bill_qty * bill_rate
    
    bill_data2.append([i, desc, unit, work_order_qty, work_order_rate, 
                       bill_qty, bill_rate, work_order_amount, bill_amount])

# Create the full Excel file with more items
with pd.ExcelWriter('test_files/sample_bill_large.xlsx', engine='openpyxl') as writer:
    # Write header (rows 1-21)
    header_df.to_excel(writer, sheet_name='Sheet1', index=False, header=False)
    
    # Write bill data starting from row 22
    bill_df2 = pd.DataFrame(bill_data2, columns=columns)
    bill_df2.to_excel(writer, sheet_name='Sheet1', startrow=21, index=False)

print("Larger sample Excel file created at: test_files/sample_bill_large.xlsx")