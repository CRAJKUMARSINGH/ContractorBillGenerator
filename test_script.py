import os
import pandas as pd
import subprocess
import platform
from app import process_bill, generate_pdf, create_word_doc

def check_wkhtmltopdf():
    """Check if wkhtmltopdf is installed and download if needed"""
    try:
        subprocess.run(['wkhtmltopdf', '--version'], capture_output=True, check=True)
        print("wkhtmltopdf is already installed")
        return True
    except subprocess.CalledProcessError:
        print("wkhtmltopdf not found. Downloading...")
        
        # Download appropriate version based on OS
        system = platform.system()
        if system == "Windows":
            # Download Windows installer
            import urllib.request
            url = "https://github.com/wkhtmltopdf/packaging/releases/download/0.12.6.1-2/wkhtmltox-0.12.6.1-2.msvc2015-win64.exe"
            urllib.request.urlretrieve(url, "wkhtmltopdf_installer.exe")
            print("Downloaded wkhtmltopdf installer")
            
            # Run installer silently
            subprocess.run(["wkhtmltopdf_installer.exe", "/S"])
            print("wkhtmltopdf installed successfully")
            return True
        else:
            print("Unsupported operating system")
            return False

# Check and install wkhtmltopdf before running tests
if not check_wkhtmltopdf():
    print("Failed to install wkhtmltopdf. Please install it manually from https://wkhtmltopdf.org/downloads.html")
    exit(1)

print("Starting tests...")

# Rest of the test script remains the same

def run_test(file_path, bill_type, bill_number):
    print(f"\nTesting file: {file_path} with {bill_type} {bill_number}")
    
    # Test data
    test_data = {
        "start_date": "01-04-2025",
        "completion_date": "30-04-2025",
        "actual_completion_date": "30-04-2025",
        "measurement_date": "",
        "order_date": "01-04-2025",
        "contractor_name": "Test Contractor",
        "work_name": "Test Work",
        "bill_serial": "12345",
        "agreement_no": "AG/2025/001",
        "work_order_ref": "WO/2025/001",
        "work_order_amount": 100000.0,
        "premium_percent": 5.0,
        "amount_paid_last_bill": 0.0 if bill_number == "First" else 50000.0,
        "bill_type": bill_type,
        "bill_number": bill_number,
        "last_bill_reference": "BILL/2025/001" if bill_number != "First" else "N/A",
        "measurement_officer": "",
        "officer_name": "",
        "officer_designation": "",
        "authorising_officer_name": "",
        "authorising_officer_designation": ""
    }
    
    # Read all sheets from the Excel file
    with pd.ExcelFile(file_path) as xls:
        ws_wo = pd.read_excel(xls, sheet_name="Work Order")
        ws_bq = pd.read_excel(xls, sheet_name="Bill Quantity")
        ws_extra = pd.read_excel(xls, sheet_name="Extra Items")
    
    print("Data being used:")
    print("-" * 50)
    print(f"Start Date: {test_data['start_date']}")
    print(f"Completion Date: {test_data['completion_date']}")
    print(f"Actual Completion Date: {test_data['actual_completion_date']}")
    print(f"Measurement Date: {test_data['measurement_date']}")
    print(f"Order Date: {test_data['order_date']}")
    print(f"Contractor Name: {test_data['contractor_name']}")
    print(f"Work Name: {test_data['work_name']}")
    print(f"Bill Serial: {test_data['bill_serial']}")
    print(f"Agreement No: {test_data['agreement_no']}")
    print(f"Work Order Ref: {test_data['work_order_ref']}")
    print(f"Work Order Amount: {test_data['work_order_amount']}")
    print(f"Premium Percent: {test_data['premium_percent']}")
    print(f"Amount Paid Last Bill: {test_data['amount_paid_last_bill']}")
    print(f"Last Bill Reference: {test_data['last_bill_reference']}")
    print("-" * 50)
    
    # Process bill
    try:
        print("Processing bill data...")
        first_page_data, last_page_data, deviation_data, extra_items_data, note_sheet_data = \
            process_bill(ws_wo, ws_bq, ws_extra, test_data["premium_percent"], "Fixed", 
                        test_data["amount_paid_last_bill"], bill_number == "First", test_data)
        
        if not all([first_page_data, last_page_data]):
            print("Failed to process bill data")
            return False
            
        print("Bill data processed successfully")
        
        # Create temporary directory for output
        temp_dir = "test_output"
        os.makedirs(temp_dir, exist_ok=True)
        
        # Generate PDF
        pdf_path = os.path.join(temp_dir, f"{os.path.basename(file_path)}_{bill_type}_{bill_number}.pdf")
        
        # Generate first page
        print("Generating first page...")
        if not generate_pdf("first_page", first_page_data, "portrait", pdf_path, is_first_page=True):
            print("Failed to generate first page")
            return False
        
        print("Generating certificate II...")
        if not generate_pdf("certificate_ii", last_page_data, "portrait", pdf_path):
            print("Failed to generate certificate II")
            return False
        
        print("Generating certificate III...")
        if not generate_pdf("certificate_iii", last_page_data, "portrait", pdf_path):
            print("Failed to generate certificate III")
            return False
        
        # Generate deviation statement for final bills
        if bill_type == "Final Bill" and deviation_data:
            print("Generating deviation statement...")
            if not generate_pdf("deviation_statement", deviation_data, "portrait", pdf_path):
                print("Failed to generate deviation statement")
                return False
        
        # Generate note sheet if present
        if note_sheet_data:
            print("Generating note sheet...")
            if not generate_pdf("note_sheet", note_sheet_data, "portrait", pdf_path):
                print("Failed to generate note sheet")
                return False
        
        # Generate extra items if present
        if extra_items_data:
            print("Generating extra items...")
            if not generate_pdf("extra_items", extra_items_data, "portrait", pdf_path):
                print("Failed to generate extra items")
                return False
        
        print(f"Successfully generated PDF for {bill_type} {bill_number}")
        return True
        
    except Exception as e:
        print(f"Error processing {bill_type} {bill_number}: {str(e)}")
        import traceback
        print("Full error traceback:")
        print(traceback.format_exc())
        return False

if __name__ == "__main__":
    # Test different combinations
    test_files = [
        "PRIYANKA SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx",
        "PRIYANKA SAMPLE BILL INPUT- WITH EXTRA ITEMS.xlsx",
        "SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx",
        "SAMPLE BILL INPUT- WITH EXTRA ITEMS.xlsx",
           ]
    
    bill_types = ["Running Bill", "Final Bill"]
    bill_numbers = ["First", "Second", "Third"]
    
    # Run tests
    for file_name in test_files:
        file_path = os.path.join("test_files", file_name)
        if os.path.exists(file_path):
            for bill_type in bill_types:
                for bill_number in bill_numbers:
                    run_test(file_path, bill_type, bill_number)
