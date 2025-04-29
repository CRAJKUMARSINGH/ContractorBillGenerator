import streamlit as st
import pandas as pd
import datetime
from utils import (
    process_excel_file, 
    generate_bill_summary, 
    export_to_excel,
    export_to_word,
    export_to_pdf,
    generate_certificate_2,
    generate_certificate_3,
    generate_deviation_statement,
    generate_note_sheet,
    generate_extra_items_slip
)

st.set_page_config(page_title="Contractor Bill Generator", layout="wide")

def main():
    st.title("Contractor Bill Generator")
    
    # Application description
    st.markdown("""
    This application processes contractor bill data from Excel files and generates formatted bills with calculated deviations (for final bills only).
    
    ### Instructions:
    1. Fill in the required details in the sidebar
    2. Upload an Excel file containing bill data (rows from 22 onwards will be processed)
    3. View the bill summary and download the processed Excel file
    
    **Note:** Deviations will only be calculated for final bills.
    """)
    
    # Sidebar for input parameters
    st.sidebar.header("Bill Information")
    
    # Required fields with validation
    with st.sidebar.form("bill_info_form"):
        st.subheader("Mandatory Fields")
        
        # Bold labels for mandatory fields
        st.markdown("<strong style='color:red'>* Required fields</strong>", unsafe_allow_html=True)
        
        start_date = st.date_input(
            "Start Date *",
            datetime.date.today(),  # Default to today
            help="The date when work started (e.g., 18-01-2025)"
        )
        
        completion_date = st.date_input(
            "Scheduled Completion Date *",
            datetime.date.today(),  # Default to today
            help="The scheduled date for work completion (e.g., 17-04-2025)"
        )
        
        actual_completion_date = st.date_input(
            "Actual Completion Date *",
            datetime.date.today(),  # Default to today
            help="The date when work was actually completed (e.g., 01-03-2025)"
        )
        
        work_order_amount = st.number_input(
            "Work Order Amount (Rs.) *",
            min_value=0.0,
            value=0.0,  # Starting with zero
            step=1000.0,
            help="Total amount specified in the work order (e.g., 854678)"
        )
        
        # Optional fields
        st.subheader("Optional Fields")
        st.markdown("<small>These fields can be left blank and filled manually in the exported document</small>", unsafe_allow_html=True)
        
        contractor_name = st.text_input(
            "Contractor Name",
            value="",  # Empty by default
            placeholder="e.g., M/s Seema Electrical Udaipur",
            help="Name of the contractor or supplier"
        )
        
        work_name = st.text_input(
            "Work Name",
            value="",  # Empty by default
            placeholder="e.g., Electric Repair and MTC work at Govt. Ambedkar hostel Ambamata, Govardhanvilas, Udaipur",
            help="Description of the work being billed"
        )
        
        # Bill type selection - Running or Final
        bill_type = st.radio(
            "Bill Type",
            ["Running Bill", "Final Bill"],
            index=1,  # Default to Final Bill
            help="Select whether this is a running bill or the final bill"
        )
        
        is_final_bill = bill_type == "Final Bill"
        
        # Bill number selection
        bill_number = st.selectbox(
            "Bill Number",
            ["First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth", "Tenth"],
            index=0,  # Default to First
            help="Select the sequential number of this bill"
        )
        
        # Construct bill serial based on selections
        bill_serial = f"{bill_number} {bill_type}"
        
        # Last bill reference
        last_bill_options = ["Not Applicable"]
        if bill_number != "First":
            # If not the first bill, add previous bill numbers as options
            bill_index = ["First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth", "Tenth"].index(bill_number)
            previous_bills = [f"{prev} Running Bill" for prev in ["First", "Second", "Third", "Fourth", "Fifth", "Sixth", "Seventh", "Eighth", "Ninth"][:bill_index]]
            last_bill_options = previous_bills + last_bill_options
        
        last_bill = st.selectbox(
            "Last Bill Reference",
            last_bill_options,
            index=0,
            help="Reference to the previous bill if applicable"
        )
        
        # Amount paid in last running bill - Always ask for this field and make it prominent
        st.markdown("<strong style='color:blue'>Amount Paid in Last Running Bill</strong>", unsafe_allow_html=True)
        st.markdown("<small>Enter the total amount paid in previous running bills. This is important for calculating the net amount payable.</small>", unsafe_allow_html=True)
        st.markdown("<small>For first bills, you can leave this as 0.0</small>", unsafe_allow_html=True)
        last_bill_amount = st.number_input(
            "Amount Paid in Last Running Bill (Rs.) *",  # Added asterisk to indicate importance
            min_value=0.0,
            value=0.0,
            step=1000.0,
            help="Total amount paid in the previous running bill(s). Required for calculating net amount payable."
        )
        
        work_order_ref = st.text_input(
            "Work Order Reference",
            value="",  # Empty by default
            placeholder="e.g., 1179 Dt. 09-01-2025",
            help="Reference number and date of the work order"
        )
        
        agreement_no = st.text_input(
            "Agreement Number",
            value="",  # Empty by default
            placeholder="e.g., 48/2024-25",
            help="Contract agreement number"
        )
        
        order_date = st.date_input(
            "Order Date",
            datetime.date.today(),  # Default to today
            help="Date of written order to commence work (e.g., 09-01-2025)"
        )
        
        measurement_date = st.date_input(
            "Measurement Date",
            datetime.date.today(),  # Default to today
            help="Date when measurements were taken (e.g., 03-03-2025)"
        )
        
        submitted = st.form_submit_button("Submit Details")
    
    # Main area for file upload and processing
    st.subheader("Upload Bill Data")
    
    # File upload section with clear instructions
    st.info("Upload an Excel file containing bill data. The file should have at least 22 rows, with data starting from row 22.")
    
    # Provide example file download
    with open("test_files/sample_bill.xlsx", "rb") as file:
        st.download_button(
            label="Download Sample Excel Template",
            data=file,
            file_name="bill_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Download a sample Excel template to see the expected format"
        )
    
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"], help="Upload your bill data Excel file. Headers (first 21 rows) will be skipped.")
    
    if uploaded_file is not None:
        st.success("File uploaded successfully!")
        
        # Process the Excel file
        with st.spinner("Processing file..."):
            # Show details about the uploaded file for debugging
            file_details = {
                "Filename": uploaded_file.name,
                "File size": f"{uploaded_file.size / 1024:.2f} KB",
                "File type": uploaded_file.type
            }
            st.write("**File Details:**", file_details)
            
            # Process the file
            df = process_excel_file(uploaded_file, is_final_bill)
            
            if df is not None:
                # Store contractor details
                contractor_details = {
                    'contractor_name': contractor_name,
                    'work_name': work_name,
                    'bill_serial': bill_serial,
                    'last_bill': last_bill,
                    'work_order_ref': work_order_ref,
                    'agreement_no': agreement_no,
                    'order_date': order_date.strftime("%d-%m-%Y"),
                    'start_date': start_date.strftime("%d-%m-%Y"),
                    'completion_date': completion_date.strftime("%d-%m-%Y"),
                    'actual_completion_date': actual_completion_date.strftime("%d-%m-%Y"),
                    'measurement_date': measurement_date.strftime("%d-%m-%Y"),
                    'work_order_amount': f"{work_order_amount:,.2f}",
                    'is_final_bill': is_final_bill
                }
                
                # Generate bill summary
                summary = generate_bill_summary(df, contractor_details)
                
                # Create two columns for the summary display
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("Bill Summary")
                    st.metric("Total Work Order Amount", f"₹ {summary['total_work_order_amount']:,.2f}")
                    st.metric("Total Bill Amount", f"₹ {summary['total_bill_amount']:,.2f}")
                    
                    if is_final_bill:
                        # Add color based on deviation (red for negative, green for positive)
                        deviation_color = "normal"
                        if summary['total_deviation'] > 0:
                            deviation_color = "off"  # Green in Streamlit
                        elif summary['total_deviation'] < 0:
                            deviation_color = "inverse"  # Red in Streamlit
                            
                        st.metric(
                            "Total Deviation", 
                            f"₹ {summary['total_deviation']:,.2f}", 
                            f"{summary['deviation_percentage']}%",
                            delta_color=deviation_color
                        )
                
                with col2:
                    st.subheader("Contract Details")
                    st.write(f"**Contractor:** {contractor_name}")
                    st.write(f"**Work:** {work_name}")
                    st.write(f"**Bill Type:** {bill_serial}")
                    st.write(f"**Agreement No.:** {agreement_no}")
                    st.write(f"**Period:** {start_date.strftime('%d-%m-%Y')} to {actual_completion_date.strftime('%d-%m-%Y')}")
                
                # Display bill data
                st.subheader("Bill Data")
                # Format the dataframe for display
                display_df = df.copy()
                
                # Style the dataframe to highlight deviations
                if is_final_bill and 'Deviation_Qty' in display_df.columns:
                    # Highlight rows with deviations
                    st.dataframe(
                        display_df.style.apply(
                            lambda x: ['background-color: rgba(255,204,203,0.5)' if x['Deviation_Qty'] < 0 else
                                      'background-color: rgba(204,255,204,0.5)' if x['Deviation_Qty'] > 0 else 
                                      '' for i in range(len(x))], 
                            axis=1
                        )
                    )
                else:
                    st.dataframe(display_df)
                
                # Always add the last bill amount to the summary
                summary['last_bill_amount'] = last_bill_amount
                
                # Document download section
                st.subheader("Generate Documents")
                
                # Create download options in tabs
                download_tabs = st.tabs(["Excel", "Word", "PDF", "Individual Documents"])
                
                with download_tabs[0]:  # Excel Tab
                    # Generate and provide download link for complete Excel file with all sheets
                    excel_data = export_to_excel(df, summary)
                    st.download_button(
                        label="Download Complete Bill Package (Excel)",
                        data=excel_data,
                        file_name=f"Contractor_Bill_{contractor_name}_{datetime.date.today().strftime('%Y-%m-%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Download Excel file with bill data and all supporting documents"
                    )
                
                with download_tabs[1]:  # Word Tab
                    # Placeholder for Word export functionality
                    st.info("Word export will be implemented in a future update. Currently, you can download as Excel and open in Word.")
                    word_data = export_to_word(df, summary)
                    st.download_button(
                        label="Download as Word Document (Preview)",
                        data=word_data,
                        file_name=f"Contractor_Bill_{contractor_name}_{datetime.date.today().strftime('%Y-%m-%d')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        help="This is a preview feature - full Word export will be available in a future update",
                        disabled=True
                    )
                
                with download_tabs[2]:  # PDF Tab
                    # Placeholder for PDF export functionality
                    st.info("PDF export will be implemented in a future update. Currently, you can download as Excel and export to PDF.")
                    pdf_data = export_to_pdf(df, summary)
                    st.download_button(
                        label="Download as PDF Document (Preview)",
                        data=pdf_data,
                        file_name=f"Contractor_Bill_{contractor_name}_{datetime.date.today().strftime('%Y-%m-%d')}.pdf",
                        mime="application/pdf",
                        help="This is a preview feature - full PDF export will be available in a future update",
                        disabled=True
                    )
                
                with download_tabs[3]:  # Individual Documents Tab
                    st.write("Download individual documents:")
                    
                    # Certificate-II
                    st.subheader("Certificate-II")
                    st.text_area("Certificate-II Content", generate_certificate_2(summary), height=200)
                    
                    # Certificate-III
                    st.subheader("Certificate-III")
                    st.text_area("Certificate-III Content", generate_certificate_3(summary), height=200)
                    
                    # Deviation Statement (only for final bill)
                    if is_final_bill:
                        st.subheader("Deviation Statement")
                        st.text_area("Deviation Statement Content", generate_deviation_statement(df, summary), height=300)
                    
                    # Note Sheet
                    st.subheader("Note Sheet")
                    st.text_area("Note Sheet Content", generate_note_sheet(summary), height=300)
                    
                    # Extra Items Slip (always show, even if empty)
                    extra_items_df = st.session_state.extra_items_df if hasattr(st.session_state, 'extra_items_df') else pd.DataFrame()
                    st.subheader("Extra Items Slip")
                    extra_items_content = generate_extra_items_slip(extra_items_df, summary)
                    st.text_area("Extra Items Slip Content", extra_items_content, height=300)
                
                # Additional options for viewing specific data
                st.subheader("View Options")
                
                view_option = st.selectbox(
                    "Select data to view",
                    ["All Data", "Items with Deviations", "Summary Only"]
                )
                
                if view_option == "Items with Deviations" and is_final_bill:
                    if 'Deviation_Qty' in df.columns:
                        deviation_df = df[df['Deviation_Qty'] != 0]
                        st.dataframe(deviation_df)
                        st.write(f"Number of items with deviations: {len(deviation_df)}")
                    else:
                        st.warning("Deviation data not available")
                
                elif view_option == "Summary Only":
                    # Display only the summary in a more formatted way
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write("**Contractor Details**")
                        st.write(f"Contractor: {contractor_name}")
                        st.write(f"Work: {work_name}")
                        st.write(f"Bill Serial: {bill_serial}")
                        st.write(f"Work Order Ref: {work_order_ref}")
                        st.write(f"Agreement No.: {agreement_no}")
                    
                    with col2:
                        st.write("**Dates**")
                        st.write(f"Start Date: {start_date.strftime('%d-%m-%Y')}")
                        st.write(f"Scheduled Completion: {completion_date.strftime('%d-%m-%Y')}")
                        st.write(f"Actual Completion: {actual_completion_date.strftime('%d-%m-%Y')}")
                        st.write(f"Measurement Date: {measurement_date.strftime('%d-%m-%Y')}")
            else:
                st.error("Failed to process the Excel file. Please check the format and try again.")

if __name__ == "__main__":
    main()
