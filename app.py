import streamlit as st
import pandas as pd
from datetime import date
import os
import tempfile
from jinja2 import Environment, FileSystemLoader
import traceback
from utils import process_bill, generate_pdf, combine_pdfs

# Initialize form state at the very top
if 'form_state' not in st.session_state:
    st.session_state.form_state = {
        'uploaded_file': None,
        'premium_percent': 0.0,
        'premium_type': 'Above',
        'premium_position': 'Percentage',
        'amount_paid_last_bill': 0,
        'start_date': date.today(),
        'completion_date': date.today(),
        'bill_type': 'Running Bill',
        'work_order_amount': 0,
        'processing': False,
        'error': None,
        'bill_number': 'First'
    }

# Initialize Jinja2 environment
env = Environment(loader=FileSystemLoader(os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")))

# Title and description
st.title("Contractor Bill Generator")
st.markdown("""
1. Upload Excel file containing Work Order, Bill Quantity, and Extra Items sheets
2. Fill in the required fields
3. Click Process Bill to generate the PDF
""")

# File upload
uploaded_file = st.file_uploader("Upload Bill Input File", type=['xlsx'])

if uploaded_file:
    st.write(f"Uploaded file: {uploaded_file.name}")

# Main content
with st.form("bill_form", clear_on_submit=False):
    # Get form data
    st.session_state.form_state["start_date"] = st.date_input(
        "Start Date *",
        st.session_state.form_state.get("start_date", date.today()),
        help="Start date of the work"
    )
    
    st.session_state.form_state["completion_date"] = st.date_input(
        "Completion Date *",
        st.session_state.form_state.get("completion_date", date.today()),
        help="Expected completion date"
    )

    st.session_state.form_state["bill_type"] = st.radio(
        "Bill Type",
        ["Running Bill", "Final Bill"],
        index=0 if st.session_state.form_state.get("bill_type") == "Running Bill" else 1,
        help="Select whether this is a running bill or the final bill"
    )

    st.session_state.form_state["bill_number"] = st.selectbox(
        "Bill Number",
        ["First", "Second"],
        index=0 if st.session_state.form_state.get("bill_number") == "First" else 1,
        help="Select the bill number"
    )

    st.session_state.form_state["work_order_amount"] = st.number_input(
        "Work Order Amount *",
        min_value=0,
        value=st.session_state.form_state.get("work_order_amount", 0),
        help="Total work order amount"
    )

    # Premium settings
    st.session_state.form_state["premium_percent"] = st.number_input(
        "Premium Percentage",
        min_value=0.0,
        max_value=100.0,
        step=0.1,
        value=st.session_state.form_state.get("premium_percent", 0.0)
    )

    st.session_state.form_state["premium_type"] = st.selectbox(
        "Premium Type",
        ["Above", "Below"],
        index=0 if st.session_state.form_state.get("premium_type") == "Above" else 1
    )

    st.session_state.form_state["premium_position"] = st.selectbox(
        "Premium Position",
        ["Percentage", "Fixed"],
        index=0 if st.session_state.form_state.get("premium_position") == "Percentage" else 1
    )

    # Amount paid in last bill
    st.session_state.form_state["amount_paid_last_bill"] = st.number_input(
        "Amount Paid in Last Bill",
        min_value=0,
        value=st.session_state.form_state.get("amount_paid_last_bill", 0)
    )

    # Submit button
    submitted = st.form_submit_button("Process Bill")

    if submitted and uploaded_file is not None:
        try:
            # Read the uploaded file
            with pd.ExcelFile(uploaded_file) as xls:
                ws_wo = pd.read_excel(xls, "Work Order", header=None)
                ws_bq = pd.read_excel(xls, "Bill Quantity", header=None)
                ws_extra = pd.read_excel(xls, "Extra Items", header=None)

            # Prepare user inputs
            user_inputs = {
                "work_order_amount": st.session_state.form_state["work_order_amount"],
                "premium_percent": st.session_state.form_state["premium_percent"],
                "premium_type": st.session_state.form_state["premium_type"],
                "amount_paid_last_bill": st.session_state.form_state["amount_paid_last_bill"],
                "start_date": st.session_state.form_state["start_date"],
                "completion_date": st.session_state.form_state["completion_date"],
                "bill_type": st.session_state.form_state["bill_type"],
                "is_first_bill": st.session_state.form_state["bill_type"] == "Running Bill" and st.session_state.form_state["bill_number"] == "First",
                "premium_position": st.session_state.form_state["premium_position"]
            }

            # Process the bill
            first_page_data, last_page_data, deviation_data, extra_items_data, note_sheet_data = process_bill(
                ws_wo, ws_bq, ws_extra,
                premium_percent=st.session_state.form_state["premium_percent"],
                premium_type=st.session_state.form_state["premium_type"],
                amount_paid_last_bill=st.session_state.form_state["amount_paid_last_bill"],
                is_first_bill=user_inputs["is_first_bill"],
                user_inputs=user_inputs
            )

            # Create a temporary directory
            temp_dir = tempfile.mkdtemp()
            
            # Generate PDFs for each page
            pdf_paths = []
            for sheet_name, data in {
                "First Page": first_page_data,
                "Last Page": last_page_data,
                "Deviation Statement": deviation_data,
                "Extra Items": extra_items_data,
                "Note Sheet": note_sheet_data
            }.items():
                if data is None:
                    continue
                    
                # Prepare template data
                template_data = data.copy()
                template_data.update({
                    "premium_percent": st.session_state.form_state["premium_percent"],
                    "premium_type": st.session_state.form_state["premium_type"],
                    "amount_paid_last_bill": st.session_state.form_state["amount_paid_last_bill"],
                    "premium_position": st.session_state.form_state["premium_position"]
                })
                
                # Get template and render HTML
                template = env.get_template(f"{sheet_name.lower().replace(' ', '_')}.html")
                html_content = template.render(**template_data)
                
                # Generate PDF
                pdf_path = os.path.join(temp_dir, f"{sheet_name.lower().replace(' ', '_')}.pdf")
                generate_pdf(html_content, output_path=pdf_path)
                pdf_paths.append(pdf_path)
            
            # Combine all PDFs into one
            combined_pdf_path = os.path.join(temp_dir, "combined_bill.pdf")
            combine_pdfs(pdf_paths, combined_pdf_path)
            
            # Display success message and download link
            st.success("Bill processed successfully!")
            
            # Create download button
            with open(combined_pdf_path, "rb") as f:
                pdf_bytes = f.read()
            st.download_button(
                label="Download Bill",
                data=pdf_bytes,
                file_name="contractor_bill.pdf",
                mime="application/pdf"
            )

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.error(traceback.format_exc())

# Add clear form button
if st.button("Clear Form"):
    st.session_state.form_state = {
        'uploaded_file': None,
        'premium_percent': 0.0,
        'premium_type': 'Above',
        'premium_position': 'Percentage',
        'amount_paid_last_bill': 0,
        'start_date': date.today(),
        'completion_date': date.today(),
        'bill_type': 'Running Bill',
        'work_order_amount': 0,
        'processing': False,
        'error': None,
        'bill_number': 'First'
    }
    st.experimental_rerun()

def process_bill(ws_wo, ws_bq, ws_extra, premium_percent, premium_type, premium_position, amount_paid_last_bill, is_first_bill, user_inputs):
    # Process bill logic here
    pass

def generate_pdf(html_content, output_path):
    # Generate PDF logic here
    pass

def combine_pdfs(pdf_paths, output_path):
    # Combine PDFs logic here
    pass
