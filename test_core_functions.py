import unittest
import pandas as pd
from datetime import datetime
import os
from core_functions import process_bill, generate_bill_notes

class TestContractorBillGenerator(unittest.TestCase):
    def setUp(self):
        # Common test data
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

    def test_no_extra_items(self):
        """Test with Excel file containing no extra items"""
        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx") as xls:
            ws_wo = pd.read_excel(xls, "Work Order", header=None)
            ws_bq = pd.read_excel(xls, "Bill Quantity", header=None)
            ws_extra = pd.read_excel(xls, "Extra Items", header=None)

        result = process_bill(ws_wo, ws_bq, ws_extra, 
                             self.test_data["premium_percent"],
                             self.test_data["premium_type"],
                             self.test_data["amount_paid_last_bill"],
                             self.test_data["is_first_bill"],
                             self.test_data)

        self.assertIsNotNone(result)
        self.assertEqual(len(result), 5)  # Should return 5 data structures
        self.assertGreater(result[0]["totals"]["grand_total"], 0)

    def test_with_extra_items(self):
        """Test with Excel file containing extra items"""
        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- WITH EXTRA ITEMS.xlsx") as xls:
            ws_wo = pd.read_excel(xls, "Work Order", header=None)
            ws_bq = pd.read_excel(xls, "Bill Quantity", header=None)
            ws_extra = pd.read_excel(xls, "Extra Items", header=None)

        result = process_bill(ws_wo, ws_bq, ws_extra, 
                             self.test_data["premium_percent"],
                             self.test_data["premium_type"],
                             self.test_data["amount_paid_last_bill"],
                             self.test_data["is_first_bill"],
                             self.test_data)

        self.assertIsNotNone(result)
        self.assertEqual(len(result), 5)
        self.assertGreater(result[3]["totals"]["extra_items_total"], 0)

    def test_final_bill(self):
        """Test final bill with deviation data"""
        test_data = self.test_data.copy()
        test_data["bill_type"] = "Final Bill"
        test_data["last_bill"] = True

        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- WITH EXTRA ITEMS.xlsx") as xls:
            ws_wo = pd.read_excel(xls, "Work Order", header=None)
            ws_bq = pd.read_excel(xls, "Bill Quantity", header=None)
            ws_extra = pd.read_excel(xls, "Extra Items", header=None)

        result = process_bill(ws_wo, ws_bq, ws_extra, 
                             self.test_data["premium_percent"],
                             self.test_data["premium_type"],
                             self.test_data["amount_paid_last_bill"],
                             self.test_data["is_first_bill"],
                             test_data)

        self.assertIsNotNone(result)
        self.assertEqual(len(result), 5)
        self.assertIsNotNone(result[2])  # Deviation data should be present

    def test_bill_notes_generation(self):
        """Test bill notes generation"""
        notes = generate_bill_notes(self.test_data)
        self.assertIsNotNone(notes)
        self.assertIn("Contractor Name:", notes)
        self.assertIn("Work Name:", notes)
        self.assertIn("Bill Type:", notes)

    def test_premium_calculation(self):
        """Test premium calculation with different percentages"""
        test_data = self.test_data.copy()
        test_data["premium_percent"] = 15  # Change premium to 15%

        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- NO EXTRA ITEMS.xlsx") as xls:
            ws_wo = pd.read_excel(xls, "Work Order", header=None)
            ws_bq = pd.read_excel(xls, "Bill Quantity", header=None)
            ws_extra = pd.read_excel(xls, "Extra Items", header=None)

        result = process_bill(ws_wo, ws_bq, ws_extra, 
                             test_data["premium_percent"],
                             test_data["premium_type"],
                             test_data["amount_paid_last_bill"],
                             test_data["is_first_bill"],
                             test_data)

        self.assertIsNotNone(result)
        self.assertEqual(len(result), 5)
        premium_amount = result[0]["totals"]["premium_amount"]
        bill_amount = result[0]["totals"]["bill_amount"]
        self.assertAlmostEqual(premium_amount, bill_amount * 0.15, places=2)

if __name__ == '__main__':
    unittest.main()
