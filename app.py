import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import os
import webbrowser
import time
import tempfile
import base64
from jinja2 import Environment, FileSystemLoader
from utils import (
    process_bill,
    generate_bill_notes,
    generate_pdf,
    create_word_doc,
    number_to_words
)
import traceback
import shutil
import subprocess
import zipfile
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Initialize Jinja2 environment
env = Environment(loader=FileSystemLoader("templates"), cache_size=0)
env.filters['strptime'] = lambda s, fmt: datetime.strptime(s, fmt) if s else None

st.title("Contractor Bill Generator")

# Application description
st.markdown("""
This application processes contractor bill data from Excel files and generates formatted bills.

### Instructions:
1. Fill in the required details in the sidebar
2. Upload an Excel file containing three sheets:
   - Work Order (ws_wo)
   - Bill Quantity (ws_bq)
   - Extra Items (ws_extra)
3. View the bill summary and download the processed documents
""")

# Sidebar for input parameters
st.sidebar.header("Bill Information")

# Required fields with validation
with st.sidebar.form("bill_info_form"):
    st.subheader("Mandatory Fields")
    
    # Bold labels for mandatory fields
    st.markdown("<strong style='color:red'>* Required fields</strong>", unsafe_allow_html=True)
    
    # Date fields
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input(
            "Start Date *",
            date.today(),
            help="The date when work started",
            key="start_date"
        )
        
        completion_date = st.date_input(
            "Scheduled Completion Date *",
            date.today(),
            help="The date when work was scheduled to complete",
            key="completion_date"
        )
        
        # Measurement Date is non-mandatory
        measurement_date = st.date_input(
            "Measurement Date",
            None,
            help="The date when measurements were taken (optional)",
            key="measurement_date"
        )
    
    with col2:
        actual_completion_date = st.date_input(
            "Actual Completion Date *",
            date.today(),
            help="The actual date when work was completed",
            key="actual_completion_date"
        )
        
        order_date = st.date_input(
            "Order Date *",
            date.today(),
            help="Date of written order to commence work",
            key="order_date"
        )
    
    # Contractor and work details (all optional)
    st.markdown("---")
    st.subheader("Contractor & Work Details")
    
    contractor_name = st.text_input(
        "Contractor Name",
        "",
        help="Name of the contractor (optional)"
    )
    
    work_name = st.text_input(
        "Work Name",
        "",
        help="Name of the work/project (optional)"
    )
    
    bill_serial = st.text_input(
        "Bill Serial Number",
        "",
        help="Serial number of this bill (optional)"
    )
    
    agreement_no = st.text_input(
        "Agreement Number",
        "",
        help="Number of the agreement (optional)"
    )
    
    work_order_ref = st.text_input(
        "Work Order Reference",
        "",
        help="Reference number of the work order (optional)"
    )
    
    # Financial details
    st.markdown("---")
    st.subheader("Financial Details")
    
    work_order_amount = st.number_input(
        "Work Order Amount *",
        min_value=0.0,
        help="Total amount in the work order"
    )
    
    # Premium details (always fixed)
    premium_percent = st.number_input(
        "Premium Percentage",
        min_value=-100.0,
        max_value=100.0,
        value=0.0,
        help="Premium percentage (positive for above, negative for below)"
    )
    premium_type = "Fixed"  # Always fixed
    
    # Bill type and number
    bill_type = st.selectbox(
        "Bill Type",
        ["Running Bill", "Final Bill"],
        help="Type of bill being generated"
    )
    
    bill_number = st.selectbox(
        "Bill Number",
        ["First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth", "Tenth"],
        help="The sequence number of this bill in the series"
    )
    
    # Amount details
    if bill_number != "First":
        amount_paid_last_bill = st.number_input(
            "Amount Paid in Last Bill",
            min_value=0.0,
            help="Amount paid in the previous bill"
        )
    else:
        amount_paid_last_bill = 0.0  # First bill has no previous payment
    
    # Non-mandatory fields
    st.markdown("---")
    st.subheader("Optional Fields")
    
    # Last Bill Reference
    last_bill_reference = st.text_input(
        "Last Bill Reference",
        "",
        help="Reference to the previous bill (optional)"
    )
    
    # Measurement Date
    measurement_date = st.date_input(
        "Measurement Date",
        None,
        help="The date when measurements were taken (optional)"
    )
    
    # Bill type selection
    bill_type = st.radio(
        "Bill Type",
        ["Running Bill", "Final Bill"],
        help="Select whether this is a running bill or the final bill"
    )
    
    # Bill number selection
    bill_number = st.selectbox(
        "Bill Number",
        ["First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth", "Tenth"],
        help="Select the sequential number of this bill"
    )
    
    # Last bill reference
    last_bill_options = ["Not Applicable"]
    if bill_number != "First":
        bill_index = ["First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth", "Tenth"].index(bill_number)
        previous_bills = [f"{prev} Running Bill" for prev in ["First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth"][:bill_index]]
        last_bill_options = previous_bills + last_bill_options
    
    last_bill = st.selectbox(
        "Last Bill Reference",
        last_bill_options,
        help="Reference to the previous bill if applicable"
    )
    
    submitted = st.form_submit_button("Submit Details")

# Check for required dependencies
import subprocess

def check_dependencies():
    def check_dependency(executable):
        try:
            subprocess.run([executable, "--version"], capture_output=True, check=True)
            return True
        except Exception as e:
            return False
    
    if not check_dependency("wkhtmltopdf"):
        st.error("wkhtmltopdf is not installed. Please install it from https://wkhtmltopdf.org/downloads.html")
        return False
    
    if not check_dependency("pdftk"):
        st.error("pdftk is not installed. Please install it from https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/")
        return False
    
    return True

def check_and_display_dependencies():
    try:
        if not check_dependencies():
            st.stop()  # Stop the Streamlit app if dependencies are missing
    except Exception as e:
        st.error(f"Error checking dependencies: {str(e)}")
        st.stop()  # Stop the Streamlit app if there's an error

check_and_display_dependencies()

# Main area for file upload and processing
st.subheader("Upload Bill Data")

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx"],
    help="Upload Excel file containing Work Order, Bill Quantity, and Extra Items sheets"
)

def generate_pdf(template_name, data, orientation, output_path, is_first_page=False):
    """
    Generate PDF from HTML template.
    
    Args:
        template_name: Name of the template file (without .html extension)
        data: Dictionary containing data for template
        orientation: Page orientation (portrait/landscape)
        output_path: Path where PDF should be saved
        is_first_page: Boolean indicating if this is the first page
    """
    try:
        # Validate template name
        if not template_name:
            raise ValueError("Template name cannot be empty")
        
        # Validate orientation
        if orientation not in ["portrait", "landscape"]:
            raise ValueError("Orientation must be either 'portrait' or 'landscape'")
        
        # Validate output path
        if not output_path:
            raise ValueError("Output path cannot be empty")
        
        # Load template
        template_path = os.path.join("templates", f"{template_name}.html")
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found: {template_path}")
            
        # Create temporary HTML file
        temp_html = os.path.join(os.path.dirname(output_path), f"{template_name}.html")
        
        # Load and render template using Jinja2
        try:
            env = Environment(loader=FileSystemLoader('templates'))
            template = env.get_template(f"{template_name}.html")
            rendered_html = template.render(data)
        except Exception as e:
            raise ValueError(f"Error rendering template {template_name}: {str(e)}")
            
        # Create temporary directory for PDFs
        temp_dir = tempfile.mkdtemp()
        
        try:
            # Generate PDF for this template
            temp_pdf = os.path.join(temp_dir, f"{template_name}.pdf")
            
            # Write rendered HTML to file
            try:
                with open(temp_html, "w", encoding="utf-8") as f:
                    f.write(rendered_html)
            except Exception as e:
                raise IOError(f"Error writing HTML file: {str(e)}")
            
            # Generate PDF using wkhtmltopdf
            cmd = [
                "wkhtmltopdf",
                "--orientation", orientation,
                "--page-size", "A4",
                "--margin-top", "10mm",
                "--margin-bottom", "10mm",
                "--margin-left", "15mm",
                "--margin-right", "15mm",
                "--quiet",
                temp_html,
                temp_pdf
            ]
            
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            except subprocess.CalledProcessError as e:
                raise RuntimeError(f"Error generating PDF: {e.stderr}\nCommand: {' '.join(cmd)}")
            except FileNotFoundError:
                raise RuntimeError("wkhtmltopdf executable not found. Please install it from https://wkhtmltopdf.org/downloads.html")
            
            # Verify PDF was created
            if not os.path.exists(temp_pdf):
                raise IOError(f"Failed to create PDF for {template_name}")
                
            # Verify PDF size (should be at least 1KB)
            if os.path.getsize(temp_pdf) < 1024:
                raise IOError(f"Generated PDF for {template_name} is too small")
                
            # If this is the first page, create the final PDF
            if is_first_page:
                try:
                    shutil.copy2(temp_pdf, output_path)
                except Exception as e:
                    raise IOError(f"Error copying PDF: {str(e)}")
            else:
                # For subsequent pages, merge with existing PDF
                try:
                    cmd = [
                        "pdftk",
                        output_path,
                        temp_pdf,
                        "cat",
                        "output",
                        output_path
                    ]
                    result = subprocess.run(cmd, capture_output=True, text=True, check=True)
                except subprocess.CalledProcessError as e:
                    raise RuntimeError(f"Error merging PDF pages: {e.stderr}\nCommand: {' '.join(cmd)}")
                except FileNotFoundError:
                    raise RuntimeError("pdftk executable not found. Please install it from https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/")
            
            return True
            
        finally:
            # Clean up temporary files
            if os.path.exists(temp_html):
                os.remove(temp_html)
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
            if os.path.exists(temp_dir):
                os.rmdir(temp_dir)
                
    except Exception as e:
        print(f"Error in generate_pdf: {str(e)}")
        print(f"Full traceback: {traceback.format_exc()}")
        return False
        # Load template
        template_path = os.path.join("templates", f"{template_name}.html")
        if not os.path.exists(template_path):
            st.error(f"Template file not found: {template_path}")
            return False
            
        # Create temporary HTML file
        temp_html = os.path.join(os.path.dirname(output_path), f"{template_name}.html")
        
        # Load and render template using Jinja2
        env = Environment(loader=FileSystemLoader('templates'))
        template = env.get_template(f"{template_name}.html")
        rendered_html = template.render(data)
        
        # Create temporary directory for PDFs
        temp_dir = tempfile.mkdtemp()
        
        try:
            # Generate PDF for this template
            temp_pdf = os.path.join(temp_dir, f"{template_name}.pdf")
            
            # Write rendered HTML to file
            with open(temp_html, "w", encoding="utf-8") as f:
                f.write(rendered_html)
            
            # Generate PDF using wkhtmltopdf
            cmd = [
                "wkhtmltopdf",
                "--orientation", orientation,
                "--page-size", "A4",
                "--margin-top", "10mm",
                "--margin-bottom", "10mm",
                "--margin-left", "15mm",
                "--margin-right", "15mm",
                "--quiet",
                temp_html,
                temp_pdf
            ]
            
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            except subprocess.CalledProcessError as e:
                st.error(f"Error generating PDF for {template_name}: {e.stderr}")
                st.error(f"Command: {' '.join(cmd)}")
                return False
            except FileNotFoundError:
                st.error("wkhtmltopdf executable not found. Please install it from https://wkhtmltopdf.org/downloads.html")
                return False
            
            # Verify PDF was created
            if not os.path.exists(temp_pdf):
                st.error(f"Failed to create PDF for {template_name}")
                return False
                
            # Verify PDF size (should be at least 1KB)
            if os.path.getsize(temp_pdf) < 1024:
                st.error(f"Generated PDF for {template_name} is too small")
                return False
                
            # If this is the first page, create the final PDF
            if is_first_page:
                try:
                    shutil.copy2(temp_pdf, output_path)
                except Exception as e:
                    st.error(f"Error copying PDF: {str(e)}")
                    return False
            else:
                # For subsequent pages, merge with existing PDF
                try:
                    cmd = [
                        "pdftk",
                        output_path,
                        temp_pdf,
                        "cat",
                        "output",
                        output_path
                    ]
                    result = subprocess.run(cmd, capture_output=True, text=True, check=True)
                except subprocess.CalledProcessError as e:
                    st.error(f"Error merging PDF pages: {e.stderr}")
                    st.error(f"Command: {' '.join(cmd)}")
                    return False
                except FileNotFoundError:
                    st.error("pdftk executable not found. Please install it from https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/")
                    return False
                    
            # Verify final PDF size
            if os.path.getsize(output_path) < 1024:
                st.error("Final PDF is too small")
                return False
                
            return True
                
        finally:
            # Clean up temporary files
            if os.path.exists(temp_html):
                os.remove(temp_html)
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            
    except Exception as e:
        st.error(f"Error in generate_pdf: {str(e)}")
        print(f"Full traceback: {traceback.format_exc()}")
        return False

def create_word_doc(template_name, data, output_path):
    """
    Create Word document from template.
    
    Args:
        template_name: Name of the template file (without .docx extension)
        data: Dictionary containing data for template
        output_path: Path where Word document should be saved
    """
    try:
        # Create a new Word document
        doc = Document()
        
        # Add heading
        heading = doc.add_heading('CONTRACTOR BILL', 0)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add main content based on template
        if template_name == "first_page":
            # Add header table
            header_table = doc.add_table(rows=len(data["header"]), cols=7)
            for i, row in enumerate(data["header"]):
                for j, cell in enumerate(row):
                    header_table.cell(i, j).text = str(cell)
            
            # Add items table
            items_table = doc.add_table(rows=1, cols=9)
            items_table.style = 'Table Grid'
            
            # Add table headers
            headers = [
                "Item No.",
                "Item of Work", 
                "Unit", 
                "Quantity", 
                "Rate", 
                "Amount", 
                "Remark"
            ]
            for i, header in enumerate(headers):
                cell = items_table.cell(0, i)
                cell.text = header
                cell.paragraphs[0].runs[0].font.bold = True
            
            # Add items
            for item in data["items"]:
                row = items_table.add_row()
                for i, key in enumerate(["serial_no", "description", "unit", "quantity", 
                                      "rate", "amount", "remark"]):
                    cell = row.cells[i]
                    cell.text = str(item.get(key, ""))
        
        elif template_name == "certificate_ii":
            # Add certificate text
            doc.add_paragraph("CERTIFICATE II").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("\n")
            
            # Add certificate items
            for item in data["certificate_items"]:
                doc.add_paragraph(f"{item['name']}: {item['value']}")
        
        elif template_name == "certificate_iii":
            # Add certificate text
            doc.add_paragraph("CERTIFICATE III").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("\n")
            
            # Add certificate details
            doc.add_paragraph(f"Payable Amount: {data['payable_amount']}")
            doc.add_paragraph(f"Amount in Words: {data['amount_words']}")
        
        elif template_name == "deviation_statement":
            # Add deviation text
            doc.add_paragraph("DEVIATION STATEMENT").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("\n")
            
            # Add deviation summary
            for key, value in data["summary"].items():
                doc.add_paragraph(f"{key}: {value}")
        
        elif template_name == "note_sheet":
            # Add note sheet text
            doc.add_paragraph("NOTE SHEET").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("\n")
            
            # Add notes
            for note in data["notes"]:
                doc.add_paragraph(note)
        
        elif template_name == "extra_items":
            # Add extra items text
            doc.add_paragraph("EXTRA ITEMS").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("\n")
            
            # Add items table
            items_table = doc.add_table(rows=1, cols=8)
            items_table.style = 'Table Grid'
            
            # Add table headers
            headers = [
                "Serial No.",
                "Description", 
                "Unit", 
                "Quantity", 
                "Rate", 
                "Amount", 
                "Remark"
            ]
            for i, header in enumerate(headers):
                cell = items_table.cell(0, i)
                cell.text = header
                cell.paragraphs[0].runs[0].font.bold = True
            
            # Add items
            for item in data["items"]:
                row = items_table.add_row()
                for i, key in enumerate(["serial_no", "description", "unit", "quantity", 
                                      "rate", "amount", "remark"]):
                    cell = row.cells[i]
                    cell.text = str(item.get(key, ""))
        
        # Save document
        doc.save(output_path)
        
        # Verify file was created
        if not os.path.exists(output_path):
            st.error(f"Failed to create Word document: {output_path}")
            return False
            
        # Verify file size (should be at least 1KB)
        if os.path.getsize(output_path) < 1024:
            st.error(f"Generated Word document is too small: {output_path}")
            return False
            
        return True
        
    except Exception as e:
        st.error(f"Error creating Word document: {str(e)}")
        print(f"Full traceback: {traceback.format_exc()}")
        return False

if submitted and uploaded_file:
    try:
        # Initialize success flag
        success = True
        
        # Read all sheets from the Excel file
        with pd.ExcelFile(uploaded_file) as xls:
            ws_wo = pd.read_excel(xls, sheet_name="Work Order")
            ws_bq = pd.read_excel(xls, sheet_name="Bill Quantity")
            ws_extra = pd.read_excel(xls, sheet_name="Extra Items")

        # Prepare header data for first page
        first_page_data.update({
            "header": [
                ["Start Date:", data.get("start_date", "")],
                ["Completion Date:", data.get("completion_date", "")],
                ["Actual Completion Date:", data.get("actual_completion_date", "")],
                ["Order Date:", data.get("order_date", "")],
                ["Contractor Name:", data.get("contractor_name", "")],
                ["Work Name:", data.get("work_name", "")],
                ["Bill Serial:", data.get("bill_serial", "")],
                ["Agreement No:", data.get("agreement_no", "")],
                ["Work Order Ref:", data.get("work_order_ref", "")],
                ["Work Order Amount:", data.get("work_order_amount", "")],
                ["Premium Percent:", data.get("premium_percent", "")],
                ["Amount Paid Last Bill:", data.get("amount_paid_last_bill", "")],
                ["Bill Type:", data.get("bill_type", "")],
                ["Bill Number:", data.get("bill_number", "")],
                ["Last Bill Reference:", data.get("last_bill_reference", "")]
            ],
            "items": [],  # Will be populated with bill items
            "totals": {
                "grand_total": 0,
                "premium": {
                    "percent": 0,
                    "amount": 0
                },
                "original_payable": 0,
                "extra_items_total": 0,
                "amount_paid_last_bill": data.get("amount_paid_last_bill", 0),
                "payable": 0
            }
        })
        
        # Process bill data
        first_page_data, last_page_data, deviation_data, extra_items_data, note_sheet_data = \
            process_bill(ws_wo, ws_bq, ws_extra, premium_percent, "Fixed", 
                       amount_paid_last_bill, bill_number == "First", {
                           "start_date": start_date.strftime("%d-%m-%Y"),
                           "completion_date": completion_date.strftime("%d-%m-%Y"),
                           "actual_completion_date": actual_completion_date.strftime("%d-%m-%Y"),
                           "measurement_date": measurement_date.strftime("%d-%m-%Y"),
                            "order_date": order_date.strftime("%d-%m-%Y"),
                            "contractor_name": contractor_name,
                            "work_name": work_name,
                            "bill_serial": bill_serial,
                            "agreement_no": agreement_no,
                            "work_order_ref": work_order_ref,
                            "work_order_amount": work_order_amount,
                            "premium_percent": premium_percent,
                            "premium_type": "Fixed",
                            "amount_paid_last_bill": amount_paid_last_bill,
                            "bill_type": bill_type,
                            "bill_number": bill_number,
                            "last_bill_reference": last_bill_reference,
                            "measurement_officer": measurement_officer,
                            "officer_name": officer_name,
                            "officer_designation": officer_designation,
                            "authorising_officer_name": authorising_officer_name,
                            "authorising_officer_designation": authorising_officer_designation
                        })
        if not all([first_page_data, last_page_data]):
            st.error("Failed to process bill data")
            success = False

        if success:
            # Create temporary directory
            temp_dir = tempfile.mkdtemp()
            
            try:
                # Generate PDF
                pdf_path = os.path.join(temp_dir, "bill_output.pdf")
                
                # Generate first page (is_first_page=True)
                if not generate_pdf("first_page", first_page_data, "portrait", pdf_path, is_first_page=True):
                    success = False
                    st.error("Failed to generate first page")
                    st.stop()
                
                # Generate certificate II and III
                if not generate_pdf("certificate_ii", last_page_data, "portrait", pdf_path):
                    success = False
                    st.error("Failed to generate certificate II")
                    st.stop()
                if not generate_pdf("certificate_iii", last_page_data, "portrait", pdf_path):
                    success = False
                    st.error("Failed to generate certificate III")
                    st.stop()
                
                # Generate deviation statement for final bills
                if bill_type == "Final Bill" and deviation_data:
                    if not generate_pdf("deviation_statement", deviation_data, "portrait", pdf_path):
                        success = False
                        st.error("Failed to generate deviation statement")
                        st.stop()
                
                # Generate note sheet
                if note_sheet_data:
                    if not generate_pdf("note_sheet", note_sheet_data, "portrait", pdf_path):
                        success = False
                        st.error("Failed to generate note sheet")
                        st.stop()
                
                # Generate extra items
                if extra_items_data:
                    if not generate_pdf("extra_items", extra_items_data, "portrait", pdf_path):
                        success = False
                        st.error("Failed to generate extra items")
                        st.stop()
                
                # Check if PDF was generated successfully
                if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) < 1024:
                    st.error("Failed to generate complete PDF")
                    success = False
                    st.stop()
                
                # Generate Word documents
                word_dir = os.path.join(temp_dir, "word_docs")
                os.makedirs(word_dir, exist_ok=True)
                
                # Generate individual Word documents
                word_files = []
                if first_page_data:
                    if not create_word_doc("first_page", first_page_data, 
                        os.path.join(word_dir, "First_Page.docx")):
                        success = False
                        st.error("Failed to generate first page Word document")
                        st.stop()
                if last_page_data:
                    if not create_word_doc("certificate_ii", last_page_data, 
                        os.path.join(word_dir, "Certificate_II.docx")):
                        success = False
                        st.error("Failed to generate Certificate II Word document")
                        st.stop()
                    if not create_word_doc("certificate_iii", last_page_data, 
                        os.path.join(word_dir, "Certificate_III.docx")):
                        success = False
                        st.error("Failed to generate Certificate III Word document")
                        st.stop()
                if deviation_data and bill_type == "Final Bill":
                    if not create_word_doc("deviation_statement", deviation_data, 
                        os.path.join(word_dir, "Deviation_Statement.docx")):
                        success = False
                        st.error("Failed to generate Deviation Statement Word document")
                        st.stop()
                if note_sheet_data:
                    if not create_word_doc("note_sheet", note_sheet_data, 
                        os.path.join(word_dir, "Note_Sheet.docx")):
                        success = False
                        st.error("Failed to generate Note Sheet Word document")
                        st.stop()
                if extra_items_data:
                    if not create_word_doc("extra_items", extra_items_data, 
                        os.path.join(word_dir, "Extra_Items.docx")):
                        success = False
                        st.error("Failed to generate Extra Items Word document")
                        st.stop()
                
                # Download options
                col1, col2 = st.columns(2)
                
                with col1:
                    # Download PDF
                    try:
                        with open(pdf_path, "rb") as f:
                            st.download_button(
                                label="Download PDF Document",
                                data=f,
                                file_name=f"Contractor_Bill_{contractor_name}_{date.today().strftime('%Y-%m-%d')}.pdf",
                                mime="application/pdf"
                            )
                    except Exception as e:
                        st.error(f"Error preparing PDF for download: {str(e)}")
                        success = False
                
                with col2:
                    # Download Word documents as zip
                    try:
                        word_zip_path = os.path.join(temp_dir, "word_docs.zip")
                        with zipfile.ZipFile(word_zip_path, 'w') as zipf:
                            for root, _, files in os.walk(word_dir):
                                for file in files:
                                    zipf.write(os.path.join(root, file), file)
                        
                        with open(word_zip_path, "rb") as f:
                            st.download_button(
                                label="Download Word Documents",
                                data=f,
                                file_name=f"Contractor_Bill_Docs_{contractor_name}_{date.today().strftime('%Y-%m-%d')}.zip",
                                mime="application/zip"
                            )
                    except Exception as e:
                        st.error(f"Error preparing Word documents for download: {str(e)}")
                        success = False

            except Exception as e:
                st.error(f"Error generating documents: {str(e)}")
                print(f"Full traceback: {traceback.format_exc()}")
                success = False
            
            finally:
                # Clean up temporary directory
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)

    except Exception as e:
        st.error(f"Error processing bill: {str(e)}")
        print(f"Full traceback: {traceback.format_exc()}")
