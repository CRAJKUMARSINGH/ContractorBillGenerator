import sys
import time
from pathlib import Path
import pandas as pd
import streamlit as st
from app import process_bill
from utils import generate_pdf
import traceback
import webbrowser
import pdfkit
from jinja2 import Environment, FileSystemLoader
import os

# Add the project root to Python path
sys.path.append(str(Path(__file__).parent))

def process_test_file(file_path):
    # Set default values
    test_data = {
        "work_order_amount": 111111,
        "premium_percent": 4.11,
        "premium_type": "above",
        "is_first_bill": "no",  # Default to no
        "is_final_bill": "yes",  # Default to yes
        "amount_paid_last_bill": 11111,
        "start_date": "2025-01-01",
        "completion_date": "2025-04-25",
        "actual_completion_date": "2026-08-27",
        "work_name": "Madhuban Niwas Construction work",
        "contractor_name": "M/s Bhajan Lal",
        "work_order_ref": "786",
        "agreement_no": "12/2034-25",
        "bill_serial": "",
        "voucher_number": ""
    }

    try:
        # Read the Excel file
        print(f"\nProcessing file: {file_path.name}")
        with pd.ExcelFile(file_path) as xls:
            # Read sheets
            ws_bq = pd.read_excel(xls, "Bill Quantity", header=None)
            ws_wo = pd.read_excel(xls, "Work Order", header=None)
            ws_extra = pd.read_excel(xls, "Extra Items", header=None)

            # Process bill
            print("Processing bill...")
            first_page_data, bill_totals, deviation_data, extra_items_data, _ = process_bill(
                ws_wo=ws_wo,
                ws_bq=ws_bq,
                ws_extra=ws_extra,
                premium_percent=test_data["premium_percent"],
                premium_type=test_data["premium_type"],
                amount_paid_last_bill=test_data["amount_paid_last_bill"],
                is_first_bill=test_data["is_first_bill"] == "yes",
                user_inputs=test_data
            )

            # Ensure totals are properly initialized
            if 'totals' not in first_page_data:
                first_page_data['totals'] = {}

            # Calculate grand total
            work_order_total = first_page_data['totals'].get('work_order_total', 0)
            extra_items_total = first_page_data['totals'].get('extra_items_total', 0)
            premium_amount = first_page_data['totals'].get('premium_amount', 0)
            
            # Calculate original payable (before deductions)
            original_payable = work_order_total + extra_items_total + premium_amount
            
            # Calculate current payable (after deducting last bill amount)
            amount_paid_last_bill = test_data.get('amount_paid_last_bill', 0)
            payable = original_payable - amount_paid_last_bill
            
            # Set all required totals
            first_page_data['totals'] = {
                'work_order_total': work_order_total,
                'extra_items_total': extra_items_total,
                'premium_amount': premium_amount,
                'bill_amount': work_order_total + premium_amount,
                'grand_total': original_payable,
                'original_payable': original_payable,
                'amount_paid_last_bill': amount_paid_last_bill,
                'payable': payable
            }

            # Add premium data if not already present
            if 'premium' not in first_page_data['totals']:
                first_page_data['totals']['premium'] = {
                    'percent': test_data['premium_percent'] / 100,
                    'amount': premium_amount
                }

            # Ensure items are lists
            first_page_data['items'] = first_page_data.get('items', [])
            first_page_data['header'] = first_page_data.get('header', [])

            # Prepare template data
            template_data = {
                'data': {
                    'items': first_page_data['items'],
                    'header': first_page_data['header'],
                    'user_inputs': test_data,
                    'totals': first_page_data['totals']
                },
                'totals': first_page_data['totals']
            }

            # Add extra items data if it exists
            extra_items = first_page_data.get('extra_items', [])
            if extra_items:
                template_data['data']['extra_items'] = {
                    'items': extra_items,
                    'total': first_page_data['totals'].get('extra_items_total', 0)
                }

            # Get template and render HTML
            print("Rendering template...")
            template_dir = Path(__file__).parent / "templates"
            env = Environment(
                loader=FileSystemLoader(str(template_dir)),
                autoescape=True
            )
            template = env.get_template("first_page.html")
            
            # Try to render the template with error handling
            try:
                html_content = template.render(**template_data)
            except Exception as template_error:
                print(f"\nError rendering template:")
                print(f"Error: {str(template_error)}")
                print("Full traceback:")
                print(traceback.format_exc())
                print("\nTemplate data:")
                print(f"Work order total: {work_order_total}")
                print(f"Extra items total: {extra_items_total}")
                print(f"Premium amount: {premium_amount}")
                print(f"Grand total: {first_page_data['totals'].get('grand_total', 'Not set')}")
                print(f"Template data keys: {list(template_data.keys())}")
                print(f"Template data totals keys: {list(template_data['totals'].keys()) if 'totals' in template_data else 'No totals'}")
                return False

            # Generate PDF
            print("Generating PDF...")
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            output_dir = Path(__file__).parent / "test_output"
            output_dir.mkdir(exist_ok=True)
            
            pdf_path = output_dir / f"test_{file_path.name}_{timestamp}.pdf"
            
            # Generate PDF with proper configuration
            try:
                generate_pdf(
                    sheet_name="First Page",
                    data=template_data,
                    orientation="portrait",
                    output_path=str(pdf_path)
                )
                
                print(f"\nSuccess! PDF generated at: {pdf_path}")
                return True
                
            except Exception as pdf_error:
                print(f"\nError generating PDF:")
                print(f"Error: {str(pdf_error)}")
                print("Full traceback:")
                print(traceback.format_exc())
                print("\nPlease ensure wkhtmltopdf is installed and added to PATH.")
                print("You can download it from: https://wkhtmltopdf.org/downloads.html")
                return False

    except Exception as e:
        print(f"\nError processing {file_path.name}:")
        print(f"Error: {str(e)}")
        print("Full traceback:")
        print(traceback.format_exc())
        return False

def main():
    # Get test files directory
    test_dir = Path(__file__).parent / "test_files"
    test_files = [
        "PRIYANKA SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx",
        "PRIYANKA SAMPLE BILL INPUT- WITH EXTRA ITEMS.xlsx",
        "SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx",
        "SAMPLE BILL INPUT- WITH EXTRA ITEMS.xlsx"
    ]

    print("\nStarting bill generator tests...")
    print("This will process all files with default values.")
    print("Any errors will be shown in detail.")
    print("-" * 50)
    
    # Process each file
    for file_name in test_files:
        file_path = test_dir / file_name
        if file_path.exists():
            process_test_file(file_path)
        else:
            print(f"Warning: File not found: {file_name}")

    print("\nAll tests completed!")
    print("Check the test_output directory for generated PDFs.")

if __name__ == "__main__":
    main()
