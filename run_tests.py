import unittest
import pandas as pd
import os
import sys
import tempfile
import shutil
from unittest.mock import patch, MagicMock

# Add the parent directory to Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Mock Streamlit
sys.modules['streamlit'] = MagicMock()

from app import process_bill, generate_bill_notes, generate_pdf, create_word_doc, number_to_words

class TestContractorBillGenerator(unittest.TestCase):
    def setUp(self):
        # Sample test data
        self.test_data = {
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
            "premium_percent": 10,
            "premium_type": "Percentage",
            "amount_paid_last_bill": 0,
            "is_first_bill": True,
            "bill_type": "Running Bill",
            "bill_number": "1",
            "last_bill": False
        }

    def read_excel_file(self, file_path, sheet_name):
        """Read Excel file and skip empty rows"""
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        # Skip empty rows
        df = df.dropna(how='all')
        return df

    def test_process_bill_no_extra_items(self):
        # Read test Excel file
        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx") as xls:
            ws_wo = self.read_excel_file(xls, "Work Order")
            ws_bq = self.read_excel_file(xls, "Bill Quantity")
            ws_extra = self.read_excel_file(xls, "Extra Items")

        # Process the bill
        result = process_bill(ws_wo, ws_bq, ws_extra, self.test_data)
        
        # Verify the result
        self.assertIsNotNone(result)
        self.assertIn("total_amount", result)
        self.assertIn("bill_notes", result)

if __name__ == '__main__':
    unittest.main()
