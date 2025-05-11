import pandas as pd
from core_functions import process_bill, generate_bill_notes, generate_pdf
import os
from datetime import datetime

def read_excel_file(file_path, sheet_name):
    """Read Excel file and skip empty rows"""
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    # Skip empty rows
    df = df.dropna(how='all')
    return df

def main():
    # Sample test data
    user_inputs = {
        "start_date": "01-01-2025",
        "completion_date": "30-04-2025",
        "actual_completion_date": "30-04-2025",
        "measurement_date": "30-04-2025",
        "order_date": "01-01-2025",
        "contractor_name": "Test Contractor",
        "work_name": "Test Work",
        "bill_serial": "1",
        "agreement_no": "AG123",
        "work_order_ref": "WO123",
        "work_order_amount": 100000,
        "bill_type": "Running Bill",
        "bill_number": "First",
        "last_bill": "Not Applicable"
    }

    print("\n=== Sample Output - No Extra Items ===")
    print("\nReading Excel file...")
    try:
        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx") as xls:
            ws_wo = read_excel_file(xls, "Work Order")
            ws_bq = read_excel_file(xls, "Bill Quantity")
            ws_extra = read_excel_file(xls, "Extra Items")

        print("\nProcessing bill...")
        first_page, last_page, deviation, extra_items, note_sheet = process_bill(
            ws_wo, ws_bq, ws_extra,
            premium_percent=10,
            premium_type="Percentage",
            amount_paid_last_bill=0,
            is_first_bill=True,
            user_inputs=user_inputs
        )
        
        print("\n=== Bill Processing Results ===")
        print("\nFirst Page Totals:")
        print(f"Work Order Total: {first_page['totals']['work_order_total']}")
        print(f"Premium Amount: {first_page['totals']['premium_amount']}")
        print(f"Grand Total: {first_page['totals']['grand_total']}")
        
        print("\nBill Items:")
        for item in first_page['items']:
            print(f"Description: {item['description']}")
            print(f"Quantity: {item['quantity']}")
            print(f"Rate: {item['rate']}")
            print(f"Amount: {item['amount']}")
            print("-" * 40)
        
        if extra_items['items']:
            print("\nExtra Items:")
            for item in extra_items['items']:
                print(f"Description: {item['description']}")
                print(f"Quantity: {item['quantity']}")
                print(f"Rate: {item['rate']}")
                print(f"Amount: {item['amount']}")
                print("-" * 40)

        # Generate PDF
        pdf_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 
                               "output", 
                               f"bill_{user_inputs['bill_number']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
        
        print(f"\nGenerating PDF at: {pdf_path}")
        generate_pdf(pdf_path, user_inputs, {
            'first_page': first_page,
            'last_page': last_page,
            'deviation': deviation,
            'extra_items': extra_items,
            'note_sheet': note_sheet
        })
        print("PDF generated successfully!")

    except Exception as e:
        print("\nError processing bill:", str(e))

    print("\n\n=== Sample Bill Notes ===")
    print("\nGenerating sample bill notes...")
    try:
        notes = generate_bill_notes(user_inputs)
        print("\nBill Notes:")
        print(notes)

    except Exception as e:
        print("\nError generating bill notes:", str(e))

if __name__ == '__main__':
    main()
