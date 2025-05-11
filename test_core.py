import unittest
import pandas as pd
import os
import sys
import tempfile
import shutil
from unittest.mock import patch, MagicMock
from datetime import datetime

# Add the parent directory to Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Mock Streamlit components
mock_st = MagicMock()
mock_st.columns.return_value = [MagicMock(), MagicMock()]
mock_st.date_input.return_value = datetime(2025, 1, 1)
mock_st.text_input.return_value = "Test Value"
mock_st.selectbox.return_value = "Test Value"
mock_st.file_uploader.return_value = None

sys.modules['streamlit'] = mock_st

# Import functions directly from core_functions.py
from core_functions import process_bill, generate_bill_notes

class TestContractorBillGenerator(unittest.TestCase):
    def setUp(self):
        # Sample test data
        self.test_data = {
            "start_date": datetime(2025, 1, 1),
            "completion_date": datetime(2025, 4, 30),
            "actual_completion_date": datetime(2025, 4, 30),
            "measurement_date": datetime(2025, 4, 30),
            "order_date": datetime(2025, 1, 1),
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
        # Read Excel file with proper header detection
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        # Skip empty rows
        df = df.dropna(how='all')
        # Reset index to ensure proper DataFrame structure
        df = df.reset_index(drop=True)
        return df

    def test_process_bill_no_extra_items(self):
        """Test with Excel file containing no extra items"""
        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx") as xls:
            # Verify column names
            ws_bq = self.read_excel_file(xls, "Bill Quantity")
            print("\nTest file: SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx")
            print("Bill Quantity sheet columns:", list(ws_bq.columns))
            print("Sample data:")
            print(ws_bq.head())
            
            # Map unnamed columns
            column_mapping = {
                'Unnamed: 0': 'Description',
                'Unnamed: 1': 'Quantity',
                'Unnamed: 2': 'Rate'
            }
            ws_bq = ws_bq.rename(columns=column_mapping)
            
            required_columns = ['Quantity', 'Rate', 'Description']
            missing_columns = [col for col in required_columns if col not in ws_bq.columns]
            if missing_columns:
                print(f"Missing columns: {missing_columns}")
                print("Available columns:", list(ws_bq.columns))
                
            ws_wo = self.read_excel_file(xls, "Work Order")
            ws_extra = self.read_excel_file(xls, "Extra Items")
            
            result = process_bill(
                ws_wo, ws_bq, ws_extra,
                self.test_data["premium_percent"],
                self.test_data["premium_type"],
                self.test_data["amount_paid_last_bill"],
                self.test_data["is_first_bill"],
                self.test_data
            )
            
            self.assertIsNotNone(result)
            self.assertEqual(len(result), 5)

    def test_with_extra_items(self):
        """Test with Excel file containing extra items"""
        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- WITH EXTRA ITEMS.xlsx") as xls:
            ws_bq = self.read_excel_file(xls, "Bill Quantity")
            print("\nTest file: SAMPLE BILL INPUT- WITH EXTRA ITEMS.xlsx")
            print("Bill Quantity sheet columns:", list(ws_bq.columns))
            print("Sample data:")
            print(ws_bq.head())
            
            # Map unnamed columns
            column_mapping = {
                'Unnamed: 0': 'Description',
                'Unnamed: 1': 'Quantity',
                'Unnamed: 2': 'Rate'
            }
            ws_bq = ws_bq.rename(columns=column_mapping)
            
            required_columns = ['Quantity', 'Rate', 'Description']
            missing_columns = [col for col in required_columns if col not in ws_bq.columns]
            if missing_columns:
                print(f"Missing columns: {missing_columns}")
                print("Available columns:", list(ws_bq.columns))
                
            ws_wo = self.read_excel_file(xls, "Work Order")
            ws_extra = self.read_excel_file(xls, "Extra Items")
            
            result = process_bill(
                ws_bq, ws_bq, ws_extra,
                self.test_data["premium_percent"],
                self.test_data["premium_type"],
                self.test_data["amount_paid_last_bill"],
                self.test_data["is_first_bill"],
                self.test_data
            )
            
            self.assertIsNotNone(result)
            self.assertEqual(len(result), 5)

    def test_final_bill(self):
        """Test final bill with deviation data"""
        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- WITH EXTRA ITEMS.xlsx") as xls:
            ws_bq = self.read_excel_file(xls, "Bill Quantity")
            required_columns = ['Quantity', 'Rate', 'Description']
            missing_columns = [col for col in required_columns if col not in ws_bq.columns]
            if missing_columns:
                print(f"Missing columns: {missing_columns}")
                print("Available columns:", list(ws_bq.columns))
                
            ws_wo = self.read_excel_file(xls, "Work Order")
            ws_extra = self.read_excel_file(xls, "Extra Items")
            test_data = self.test_data.copy()
            test_data["bill_type"] = "Final Bill"
            
            result = process_bill(
                ws_wo, ws_bq, ws_extra,
                self.test_data["premium_percent"],
                self.test_data["premium_type"],
                self.test_data["amount_paid_last_bill"],
                self.test_data["is_first_bill"],
                test_data
            )
            
            self.assertIsNotNone(result)
            self.assertEqual(len(result), 5)

    def test_bill_notes_generation(self):
        """Test bill notes generation"""
        notes = generate_bill_notes(self.test_data)
        self.assertIsNotNone(notes)
        self.assertIn("Contractor Name:", notes)
        self.assertIn("Work Name:", notes)
        self.assertIn("Bill Type:", notes)

    def test_invalid_dates(self):
        """Test with invalid date order"""
        test_data = self.test_data.copy()
        test_data["start_date"] = datetime(2025, 4, 30)
        test_data["completion_date"] = datetime(2025, 1, 1)
        
        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx") as xls:
            ws_bq = self.read_excel_file(xls, "Bill Quantity")
            
            with self.assertRaises(ValueError) as context:
                process_bill(ws_bq, ws_bq, ws_bq, 
                            self.test_data["premium_percent"],
                            self.test_data["premium_type"],
                            self.test_data["amount_paid_last_bill"],
                            self.test_data["is_first_bill"],
                            test_data)
            self.assertTrue("Start date cannot be after completion date" in str(context.exception))

    def test_negative_values(self):
        """Test with negative quantities or rates"""
        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx") as xls:
            ws_bq = self.read_excel_file(xls, "Bill Quantity")
            ws_bq.loc[0, "Quantity"] = -1  # Make one quantity negative
            
            with self.assertRaises(ValueError) as context:
                process_bill(ws_bq, ws_bq, ws_bq, 
                            self.test_data["premium_percent"],
                            self.test_data["premium_type"],
                            self.test_data["amount_paid_last_bill"],
                            self.test_data["is_first_bill"],
                            self.test_data)
            self.assertTrue("Negative quantities are not allowed" in str(context.exception))

    def test_invalid_premium(self):
        """Test with invalid premium percentage"""
        test_data = self.test_data.copy()
        test_data["premium_percent"] = 101  # Invalid premium
        
        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx") as xls:
            ws_bq = self.read_excel_file(xls, "Bill Quantity")
            
            with self.assertRaises(ValueError) as context:
                process_bill(ws_bq, ws_bq, ws_bq, 
                            test_data["premium_percent"],
                            test_data["premium_type"],
                            test_data["amount_paid_last_bill"],
                            test_data["is_first_bill"],
                            test_data)
            self.assertTrue("Premium percentage must be between 0 and 100" in str(context.exception))

    def test_missing_required_fields(self):
        """Test with missing required fields"""
        test_data = self.test_data.copy()
        test_data["start_date"] = None
        
        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx") as xls:
            ws_bq = self.read_excel_file(xls, "Bill Quantity")
            
            with self.assertRaises(ValueError) as context:
                process_bill(ws_bq, ws_bq, ws_bq, 
                            self.test_data["premium_percent"],
                            self.test_data["premium_type"],
                            self.test_data["amount_paid_last_bill"],
                            self.test_data["is_first_bill"],
                            test_data)
            self.assertTrue("Start date is required" in str(context.exception))

    def test_missing_columns(self):
        """Test with missing required columns"""
        test_data = self.test_data.copy()
        
        # Create a DataFrame with missing columns
        data = pd.DataFrame({"Description": ["Test Item"]})
        
        with self.assertRaises(ValueError) as context:
            process_bill(data, data, data, 
                        self.test_data["premium_percent"],
                        self.test_data["premium_type"],
                        self.test_data["amount_paid_last_bill"],
                        self.test_data["is_first_bill"],
                        test_data)
        self.assertTrue("Missing required columns" in str(context.exception))

    def test_empty_excel_sheets(self):
        """Test with empty Excel sheets"""
        empty_df = pd.DataFrame()
        
        with self.assertRaises(ValueError) as context:
            process_bill(empty_df, empty_df, empty_df, 
                        self.test_data["premium_percent"],
                        self.test_data["premium_type"],
                        self.test_data["amount_paid_last_bill"],
                        self.test_data["is_first_bill"],
                        self.test_data)
            self.assertTrue("Empty Excel sheet" in str(context.exception))

if __name__ == '__main__':
    unittest.main()
