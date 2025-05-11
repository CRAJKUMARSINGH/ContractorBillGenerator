import unittest
import os
import pandas as pd
import subprocess
import shutil
from app import process_bill, generate_pdf
from datetime import datetime

def check_dependencies():
    """Check if required system dependencies are installed."""
    try:
        subprocess.run(['wkhtmltopdf', '--version'], check=True, capture_output=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("Warning: wkhtmltopdf not found. PDF generation tests will be skipped.")
        return False
    
    try:
        subprocess.run(['pdftk', '--version'], check=True, capture_output=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("Warning: pdftk not found. PDF merging tests will be skipped.")
        return False
    
    return True

class TestBillGenerator(unittest.TestCase):
    """Test suite for the bill generator application."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.has_dependencies = False

    @classmethod
    def setUpClass(cls):
        """Set up class-level dependencies."""
        cls.has_dependencies = check_dependencies()

    def setUp(self):
        """Set up test data for each test case."""
        self.test_data = {
            "start_date": "01-04-2025",
            "completion_date": "30-04-2025",
            "actual_completion_date": "30-04-2025",
            "measurement_date": "30-04-2025",
            "order_date": "01-04-2025",
            "contractor_name": "Test Contractor",
            "work_name": "Test Work",
            "bill_serial": "12345",
            "agreement_no": "AG/2025/001",
            "work_order_ref": "WO/2025/001",
            "work_order_amount": 100000.0,
            "premium_percent": 5.0,
            "amount_paid_last_bill": 0.0,
            "bill_type": "Regular Bill",
            "bill_number": "First",
            "last_bill_reference": "N/A",
            "measurement_officer": "",
            "officer_name": "",
            "officer_designation": "",
            "authorising_officer_name": "",
            "authorising_officer_designation": ""
        }

    def test_process_bill(self):
        # Create test dataframes
        ws_wo = pd.DataFrame({
            "Item": ["Test Item 1", "Test Item 2"],
            "Description": ["Description 1", "Description 2"],
            "Quantity": [100, 200]
        })

        ws_bq = pd.DataFrame({
            "Item": ["Test Item 1", "Test Item 2"],
            "Quantity": [50, 75],
            "Rate": [100, 200],
            "Amount": [5000, 15000]
        })

        ws_extra = pd.DataFrame({
            "Item": ["Extra Item 1"],
            "Quantity": [10],
            "Rate": [500],
            "Amount": [5000]
        })

        # Process bill
        first_page_data, last_page_data, deviation_data, extra_items_data, note_sheet_data = \
            process_bill(ws_wo, ws_bq, ws_extra, 5.0, "Fixed", 0.0, True, self.test_data)

        # Verify data structure
        self.assertIsNotNone(first_page_data)
        self.assertIsNotNone(last_page_data)
        self.assertIsNotNone(extra_items_data)
        self.assertIsNone(deviation_data)
        self.assertIsNone(note_sheet_data)

        # Verify totals
        self.assertEqual(first_page_data["totals"]["grand_total"], 20000)
        self.assertEqual(first_page_data["totals"]["premium"]["amount"], 1000)
        self.assertEqual(first_page_data["totals"]["original_payable"], 21000)

    @unittest.skipIf(not self.has_dependencies, "Skipping PDF generation tests due to missing dependencies")
    def test_generate_pdf(self):
        """Test PDF generation with valid data."""
        # Create test data
        test_data = {
            "header": [
                ["Test Header 1", "Value 1"],
                ["Test Header 2", "Value 2"]
            ],
            "items": [
                {
                    "unit": "Unit",
                    "quantity": 100,
                    "rate": 100,
                    "amount": 10000
                }
            ],
            "totals": {
                "grand_total": 10000,
                "premium": {
                    "percent": 0.05,
                    "amount": 500
                }
            }
        }

        # Create temporary output path
        temp_dir = "temp_test"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        
        output_path = os.path.join(temp_dir, "test_output.pdf")

        # Generate PDF
        result = generate_pdf("first_page", test_data, "portrait", output_path, True)
        
        # Verify PDF was created
        self.assertTrue(result)
        self.assertTrue(os.path.exists(output_path))
        self.assertGreater(os.path.getsize(output_path), 1024)

        # Clean up
        os.remove(output_path)
        os.rmdir(temp_dir)

    def test_template_rendering(self):
        """Test template rendering without PDF generation."""
        # Create test data
        test_data = {
            "header": [
                ["Test Header 1", "Value 1"],
                ["Test Header 2", "Value 2"]
            ],
            "items": [
                {
                    "unit": "Unit",
                    "quantity": 100,
                    "rate": 100,
                    "amount": 10000
                }
            ],
            "totals": {
                "grand_total": 10000,
                "premium": {
                    "percent": 0.05,
                    "amount": 500
                }
            }
        }

        try:
            env = Environment(loader=FileSystemLoader('templates'))
            template = env.get_template("first_page.html")
            rendered_html = template.render(test_data)
            self.assertTrue(rendered_html)
        except Exception as e:
            self.fail(f"Template rendering failed: {str(e)}")

    def test_invalid_data(self):
        # Test with missing required data
        with self.assertRaises(ValueError):
            process_bill(None, None, None, 5.0, "Fixed", 0.0, True, {})

        # Test with invalid dates
        invalid_data = self.test_data.copy()
        invalid_data["start_date"] = "invalid-date"
        with self.assertRaises(ValueError):
            process_bill(None, None, None, 5.0, "Fixed", 0.0, True, invalid_data)

        # Test with invalid premium type
        invalid_data = self.test_data.copy()
        with self.assertRaises(ValueError):
            process_bill(None, None, None, 5.0, "InvalidType", 0.0, True, invalid_data)

if __name__ == '__main__':
    unittest.main()
