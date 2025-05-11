import os
import sys
import pandas as pd
import unittest
from unittest.mock import patch
import streamlit as st
from datetime import date
import tempfile
import shutil

# Add the parent directory to Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import process_bill, generate_pdf, create_word_doc

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

        # Process bill
        result = process_bill(ws_wo, ws_bq, ws_extra, 
                        self.test_data["premium_percent"], 
                        self.test_data["premium_type"],
                        self.test_data["amount_paid_last_bill"],
                        self.test_data["is_first_bill"],
                        self.test_data)

        # Verify results
        self.assertIsNotNone(result)
        first_page_data, last_page_data, deviation_data, extra_items_data, note_sheet_data = result
        
        # Verify data is not None
        self.assertIsNotNone(first_page_data)
        self.assertIsNotNone(last_page_data)
        self.assertIsNone(extra_items_data)  # No extra items in this test file
        
        # Verify basic data structure
        self.assertIn("header", first_page_data)
        self.assertIn("items", first_page_data)
        self.assertIn("totals", first_page_data)
        self.assertIn("payable_amount", last_page_data)
        self.assertIn("certificate_items", last_page_data)
        self.assertIn("bill_type", last_page_data)
        self.assertIn("bill_number", last_page_data)
        self.assertIn("last_bill", last_page_data)

    def test_process_bill_with_extra_items(self):
        # Read test Excel file
        with pd.ExcelFile("test_files/SAMPLE BILL INPUT- WITH EXTRA ITEMS.xlsx") as xls:
            ws_wo = self.read_excel_file(xls, "Work Order")
            ws_bq = self.read_excel_file(xls, "Bill Quantity")
            ws_extra = self.read_excel_file(xls, "Extra Items")

        # Process bill
        result = process_bill(ws_wo, ws_bq, ws_extra, 
                        self.test_data["premium_percent"], 
                        self.test_data["premium_type"],
                        self.test_data["amount_paid_last_bill"],
                        self.test_data["is_first_bill"],
                        self.test_data)

        # Verify results
        self.assertIsNotNone(result)
        first_page_data, last_page_data, deviation_data, extra_items_data, note_sheet_data = result
        
        # Verify data is not None
        self.assertIsNotNone(first_page_data)
        self.assertIsNotNone(last_page_data)
        self.assertIsNotNone(extra_items_data)  # Should have extra items
        
        # Verify extra items data
        self.assertIn("items", extra_items_data)
        self.assertIn("totals", extra_items_data)
        self.assertIn("work_order_total", extra_items_data["totals"])
        self.assertIn("extra_items_total", extra_items_data["totals"])
        self.assertIn("grand_total", extra_items_data["totals"])
        self.assertGreater(len(extra_items_data["items"]), 0)

    def test_pdf_generation(self):
        # Create temporary directory for testing
        temp_dir = tempfile.mkdtemp()
        try:
            # Test PDF generation with sample data
            test_pdf_path = os.path.join(temp_dir, "test.pdf")
            success = generate_pdf("first_page", self.test_data, "portrait", test_pdf_path, True)
            
            # Verify PDF was generated
            self.assertTrue(success)
            self.assertTrue(os.path.exists(test_pdf_path))
            
            # Verify PDF size
            self.assertGreater(os.path.getsize(test_pdf_path), 1024)  # Should be at least 1KB

        finally:
            # Clean up
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

    def test_word_document_generation(self):
        # Create temporary directory for testing
        temp_dir = tempfile.mkdtemp()
        try:
            # Test Word document generation with sample data
            test_doc_path = os.path.join(temp_dir, "test.docx")
            success = create_word_doc("first_page", self.test_data, test_doc_path)
            
            # Verify Word document was generated
            self.assertTrue(success)
            self.assertTrue(os.path.exists(test_doc_path))
            
            # Verify document size
            self.assertGreater(os.path.getsize(test_doc_path), 1024)  # Should be at least 1KB

        finally:
            # Clean up
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

if __name__ == '__main__':
    unittest.main()
