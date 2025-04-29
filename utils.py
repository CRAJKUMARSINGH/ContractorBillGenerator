import pandas as pd
import numpy as np
import streamlit as st
import io
from jinja2 import Environment, FileSystemLoader
from num2words import num2words
import os
import traceback
from datetime import datetime

# Initialize Jinja2 environment
env = Environment(loader=FileSystemLoader("templates"), cache_size=0)
env.filters['strptime'] = lambda s, fmt: datetime.strptime(s, fmt) if s else None

def process_excel_file(file, is_final_bill):
    """
    Process uploaded Excel file to extract bill data.
    
    Args:
        file: Uploaded Excel file object
        is_final_bill: Boolean indicating if this is a final bill (to calculate deviations)
        
    Returns:
        DataFrame with processed bill data and extra items DataFrame if any
    """
    try:
        # Read the Excel file starting from row 22 (skipping headers)
        df = pd.read_excel(file, header=None, skiprows=21)
        
        # Check if DataFrame is empty
        if df.empty:
            st.error("No data found in the Excel file after skipping header rows.")
            return None
        
        # Show first few rows and column info to console for debugging
        print(f"First few rows of data:\n{df.head()}")
        print(f"Column count: {len(df.columns)}")
            
        # Rename columns for better clarity
        # Assuming standard structure: S.No, Item Description, Unit, Work Order Qty, Work Order Rate, Bill Qty, etc.
        columns = ['S_No', 'Item_Description', 'Unit', 'Work_Order_Qty', 'Work_Order_Rate', 
                  'Bill_Qty', 'Bill_Rate', 'Work_Order_Amount', 'Bill_Amount']
        
        # Use first columns only if there are enough columns in the dataframe
        if len(df.columns) >= len(columns):
            df = df.iloc[:, :len(columns)]
            df.columns = columns
        else:
            # Adjust columns based on available data
            available_columns = columns[:len(df.columns)]
            df.columns = available_columns
            st.info(f"Excel file has fewer columns than expected. Using available columns: {', '.join(available_columns)}")
        
        # Check for extra items (items in the bill that weren't in the original work order)
        # Extra items are typically indicated by having quantities and rates in bill columns but zeros or NaN in work order columns
        extra_items_mask = ((df['Work_Order_Qty'].isna() | (df['Work_Order_Qty'] == 0)) & 
                           (df['Bill_Qty'].notna() & (df['Bill_Qty'] > 0))) if 'Bill_Qty' in df.columns and 'Work_Order_Qty' in df.columns else pd.Series(False, index=df.index)
        
        # Create a separate DataFrame for extra items if any are found
        extra_items_df = None
        if extra_items_mask.any():
            extra_items_df = df[extra_items_mask].copy()
            # Remove extra items from the main DataFrame
            df = df[~extra_items_mask].copy()
            st.info(f"Found {len(extra_items_df)} extra items in the bill. These will be processed separately.")
            # Initialize extra_items_df in session state even if empty
            st.session_state.extra_items_df = extra_items_df
        else:
            # Create an empty DataFrame for extra items
            extra_items_df = pd.DataFrame(columns=df.columns)
            st.session_state.extra_items_df = extra_items_df
        
        # Remove rows with NaN in critical columns or empty rows
        df = df.dropna(subset=['Item_Description'], how='all')
        
        # Convert numeric columns to float
        numeric_cols = ['Work_Order_Qty', 'Work_Order_Rate', 'Bill_Qty', 'Bill_Rate', 
                       'Work_Order_Amount', 'Bill_Amount']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Calculate amounts based on quantities and rates
        if 'Work_Order_Qty' in df.columns and 'Work_Order_Rate' in df.columns:
            df['Work_Order_Amount'] = df['Work_Order_Qty'] * df['Work_Order_Rate']
        
        if 'Bill_Qty' in df.columns and 'Bill_Rate' in df.columns:
            df['Bill_Amount'] = df['Bill_Qty'] * df['Bill_Rate']
            
        # For any missing Bill quantities or rates, use Work Order values (for non-final bills)
        if 'Bill_Qty' not in df.columns and 'Work_Order_Qty' in df.columns:
            df['Bill_Qty'] = df['Work_Order_Qty']
            st.info("Bill quantities not found in file. Using Work Order quantities by default.")
            
        if 'Bill_Rate' not in df.columns and 'Work_Order_Rate' in df.columns:
            df['Bill_Rate'] = df['Work_Order_Rate']
            st.info("Bill rates not found in file. Using Work Order rates by default.")
            
        if 'Bill_Amount' not in df.columns and 'Work_Order_Amount' in df.columns:
            df['Bill_Amount'] = df['Work_Order_Amount']
            st.info("Bill amounts not found in file. Calculated from quantities and rates.")
            
        # Also process extra items if found
        if extra_items_df is not None and not extra_items_df.empty:
            # Convert numeric columns in extra items
            for col in numeric_cols:
                if col in extra_items_df.columns:
                    extra_items_df[col] = pd.to_numeric(extra_items_df[col], errors='coerce')
            
            # Calculate bill amounts for extra items
            if 'Bill_Qty' in extra_items_df.columns and 'Bill_Rate' in extra_items_df.columns:
                extra_items_df['Bill_Amount'] = extra_items_df['Bill_Qty'] * extra_items_df['Bill_Rate']
                
            # Add extra items flag
            extra_items_df['Is_Extra_Item'] = True
            
            # Store extra items in the session state
            st.session_state.extra_items_df = extra_items_df
        
        # Calculate deviation if it's a final bill
        if is_final_bill:
            if 'Bill_Qty' in df.columns and 'Work_Order_Qty' in df.columns:
                df['Deviation_Qty'] = df['Bill_Qty'] - df['Work_Order_Qty']
                # Avoid division by zero for percentage calculation
                df['Deviation_Percentage'] = np.where(
                    df['Work_Order_Qty'] > 0,
                    (df['Deviation_Qty'] / df['Work_Order_Qty'] * 100).round(2),
                    0
                )
                
            if 'Bill_Amount' in df.columns and 'Work_Order_Amount' in df.columns:
                df['Deviation_Amount'] = df['Bill_Amount'] - df['Work_Order_Amount']
                
            # Add deviation indicators for easy identification
            if 'Deviation_Qty' in df.columns:
                df['Deviation_Status'] = np.where(
                    df['Deviation_Qty'] == 0, 
                    'No Change',
                    np.where(df['Deviation_Qty'] > 0, 'Increase', 'Decrease')
                )
                
            # Display a message to the user if deviations are found
            if 'Deviation_Qty' in df.columns:
                deviation_count = (df['Deviation_Qty'] != 0).sum()
                if deviation_count > 0:
                    st.info(f"Found {deviation_count} items with quantity deviations in the final bill")
                else:
                    st.success("No quantity deviations found in the final bill")
        
        return df
    
    except Exception as e:
        st.error(f"Error processing Excel file: {str(e)}")
        # Show more detailed error information for debugging
        import traceback
        st.error(f"Detailed error: {traceback.format_exc()}")
        return None

def generate_bill_summary(df, contractor_details):
    """
    Generate bill summary from processed data.
    
    Args:
        df: DataFrame with processed bill data
        contractor_details: Dictionary containing contractor and job details
        
    Returns:
        Dictionary with bill summary information
    """
    summary = {}
    
    # Total amounts
    summary['total_work_order_amount'] = df['Work_Order_Amount'].sum() if 'Work_Order_Amount' in df.columns else 0
    summary['total_bill_amount'] = df['Bill_Amount'].sum() if 'Bill_Amount' in df.columns else 0
    
    # Include extra items in the bill total
    extra_items_df = st.session_state.extra_items_df if hasattr(st.session_state, 'extra_items_df') else None
    if extra_items_df is not None and not extra_items_df.empty and 'Bill_Amount' in extra_items_df.columns:
        extra_items_total = extra_items_df['Bill_Amount'].sum()
        summary['extra_items_amount'] = extra_items_total
        summary['total_bill_amount'] += extra_items_total
        st.info(f"Added extra items worth ₹{extra_items_total:,.2f} to the bill total")
    else:
        summary['extra_items_amount'] = 0
    
    # Calculate deviation if final bill
    if contractor_details.get('is_final_bill'):
        summary['total_deviation'] = summary['total_bill_amount'] - summary['total_work_order_amount']
        summary['deviation_percentage'] = (summary['total_deviation'] / summary['total_work_order_amount'] * 100).round(2) if summary['total_work_order_amount'] > 0 else 0
    
    # Add contractor details
    summary.update(contractor_details)
    
    return summary

def generate_certificate_2(summary):
    """
    Generate Certificate-II document content using the HTML template
    
    Args:
        summary: Dictionary with bill summary data
    
    Returns:
        String containing Certificate-II document text
    """
    try:
        template = env.get_template('certificate_ii.html')
        return template.render(summary=summary)
    except Exception as e:
        st.error(f"Error generating Certificate-II: {str(e)}")
        # Fallback to original format
        return f"""
        II. CERTIFICATE AND SIGNATURES

        The measurements on which are based the entries in columns 1 to 6 of Account I, were made by ............. on ............., and are recorded at page .......... of Measurement Book No. ............

        *Certified that in addition to and quite apart from the quantities of work actually executed, as shown in column 4 of Account I, some work has actually been done in connection with several items and the value of such work (after deduction therefrom the proportionate amount of secured advances, if any, ultimately recoverable on account of the quantities of materials used therein) is in no case, less than the advance payments as per item 2 of the Memorandum, if payments made or proposed to be made, for the convenience of the contractor, in anticipation of and subject to the result of detailed measurements, which will be made as soon as possible.

                                                                               Dated signature of officer preparing the bill
                                                                               Designation

                                                                               +Dated signature of officer authorising payment
                                                                               Designation
        """

def generate_certificate_3(summary):
    """
    Generate Certificate-III document content using the HTML template
    
    Args:
        summary: Dictionary with bill summary data
    
    Returns:
        String containing Certificate-III document text
    """
    try:
        # Calculate deductions
        sd_amount = round(0.1 * summary.get('total_bill_amount', 0), 2)
        it_amount = round(0.02 * summary.get('total_bill_amount', 0), 2)
        gst_amount = round(0.02 * summary.get('total_bill_amount', 0), 2)
        lc_amount = round(0.01 * summary.get('total_bill_amount', 0), 2)
        total_deductions = sd_amount + it_amount + gst_amount + lc_amount
        by_cheque = summary.get('total_bill_amount', 0) - total_deductions
        
        # Format the data for the template
        data = {
            'payable_amount': f"{summary.get('total_bill_amount', 0):,.2f}",
            'total_123': f"{summary.get('total_bill_amount', 0):,.2f}",
            'balance_4_minus_5': f"{summary.get('total_bill_amount', 0):,.2f}",
            'amount_paid_last_bill': f"{summary.get('last_bill_amount', 0):,.2f}",
            'payment_now': f"{by_cheque:,.2f}",
            'by_cheque': f"{by_cheque:,.2f}",
            'cheque_amount_words': number_to_words(by_cheque),
            'total_recovery': f"{total_deductions:,.2f}",
            'certificate_items': [
                {'name': 'S.D.', 'percentage': 10, 'value': f"{sd_amount:,.2f}"},
                {'name': 'I.T.', 'percentage': 2, 'value': f"{it_amount:,.2f}"},
                {'name': 'GST', 'percentage': 2, 'value': f"{gst_amount:,.2f}"},
                {'name': 'L.C.', 'percentage': 1, 'value': f"{lc_amount:,.2f}"}
            ]
        }
        
        template = env.get_template('certificate_iii.html')
        return template.render(data=data)
    except Exception as e:
        st.error(f"Error generating Certificate-III: {str(e)}")
        # Fallback to original format
        return f"""
        III. MEMORANDUM OF PAYMENTS

        1. Total value of work actually measured as per Account I, column 9, entry (A) Rs. {summary.get('total_bill_amount', 0):,.2f}
        2. Total up-to-date advance payments for work not yet measured, as per Account I, column 3, entry (B) Rs. ________
        3. Total up-to-date secured advances on security of materials as per Account II, column 8, entry (C) Rs. ________
        4. Total (Items 1 + 2 + 3) Rs. {summary.get('total_bill_amount', 0):,.2f}
        5. Deduct amount withheld:
           (a) From previous bill as security deposit Rs. ________
           (b) From this bill as security deposit Rs. ________
           (c) Total security deposit [(a) + (b)] Rs. ________
        6. Deduct amount previously withheld:
           (a) Sales tax/VAT 18% on Rs. {summary.get('total_bill_amount', 0):,.2f} Rs. {round(0.18 * summary.get('total_bill_amount', 0), 2):,.2f}
           (b) Income tax 2% on Rs. {summary.get('total_bill_amount', 0):,.2f} Rs. {round(0.02 * summary.get('total_bill_amount', 0), 2):,.2f}
           (c) Other deductions (specify) Rs. ________
        7. Total deductions [(5) + (6)] Rs. {round(0.2 * summary.get('total_bill_amount', 0), 2):,.2f}
        8. Net payment (Item 4 - Item 7) Rs. {round(0.8 * summary.get('total_bill_amount', 0), 2):,.2f}
        9. Total amount of payments already made as per entry (K) of last bill No. ________ dated ________ Rs. {summary.get('last_bill_amount', 0):,.2f}
        10. Payment now to be made:
            (a) By recovery of amounts creditable to this work Rs. ________
            (b) By cheque Rs. {round(0.8 * summary.get('total_bill_amount', 0), 2):,.2f}
            
                                                                               Dated signature of officer preparing the bill
                                                                               Designation

                                                                               +Dated signature of officer authorising payment
                                                                               Designation
        """

def generate_deviation_statement(df, summary):
    """
    Generate deviation statement document using the HTML template
    
    Args:
        df: DataFrame with processed bill data
        summary: Dictionary with bill summary
    
    Returns:
        String containing deviation statement text
    """
    if 'Deviation_Qty' not in df.columns:
        return "Deviation statement is only available for final bills with deviation data."
    
    # Filter rows with deviations
    deviation_df = df[df['Deviation_Qty'] != 0].copy()
    
    if len(deviation_df) == 0:
        return "No deviations found in the bill data."
    
    # Format data for the template
    deviation_items = []
    for _, row in deviation_df.iterrows():
        item_no = row.get('S_No', '')
        description = row.get('Item_Description', '')
        unit = row.get('Unit', '')
        qty_wo = row.get('Work_Order_Qty', 0)
        rate = row.get('Work_Order_Rate', 0)
        amt_wo = row.get('Work_Order_Amount', 0)
        qty_bill = row.get('Bill_Qty', 0)
        amt_bill = row.get('Bill_Amount', 0)
        
        # Calculate deviation
        deviation_qty = row.get('Deviation_Qty', 0)
        
        excess_qty = max(0, deviation_qty)
        excess_amount = max(0, row.get('Deviation_Amount', 0))
        
        saving_qty = max(0, -deviation_qty)
        saving_amount = max(0, -row.get('Deviation_Amount', 0))
        
        deviation_items.append({
            'serial_no': item_no,
            'description': description,
            'unit': unit,
            'qty_wo': qty_wo,
            'rate': rate,
            'amt_wo': amt_wo,
            'qty_bill': qty_bill,
            'amt_bill': amt_bill,
            'excess_qty': excess_qty,
            'excess_amt': excess_amount,
            'saving_qty': saving_qty,
            'saving_amt': saving_amount,
            'remark': ""
        })
    
    # Calculate summary data
    work_order_total = sum(item['amt_wo'] for item in deviation_items)
    executed_total = sum(item['amt_bill'] for item in deviation_items)
    overall_excess = sum(item['excess_amt'] for item in deviation_items)
    overall_saving = sum(item['saving_amt'] for item in deviation_items)
    net_difference = executed_total - work_order_total
    deviation_percentage = (net_difference / work_order_total) * 100 if work_order_total > 0 else 0
    
    deviation_data = {
        'items': deviation_items,
        'summary': {
            'work_order_total': work_order_total,
            'executed_total': executed_total,
            'overall_excess': overall_excess,
            'overall_saving': overall_saving,
            'net_difference': net_difference,
            'deviation_percentage': deviation_percentage,
            'premium': {
                'percent': 0  # Placeholder
            },
            'tender_premium_f': 0,
            'tender_premium_h': 0,
            'tender_premium_j': 0,
            'tender_premium_l': 0,
            'grand_total_f': work_order_total,
            'grand_total_h': executed_total,
            'grand_total_j': overall_excess,
            'grand_total_l': overall_saving
        }
    }
    
    try:
        template = env.get_template('deviation_statement.html')
        return template.render(data=deviation_data, header_data={
            8: [0, summary.get('work_name', '')],
            12: [0, 0, 0, 0, summary.get('agreement_no', '')]
        })
    except Exception as e:
        st.error(f"Error generating Deviation Statement: {str(e)}")
        
        # Fallback to original format
        statement = f"""
        DEVIATION STATEMENT
        
        Name of Work: {summary.get('work_name', '')}
        Agreement No.: {summary.get('agreement_no', '')}
        Name of Contractor: {summary.get('contractor_name', '')}
        Bill: {summary.get('bill_serial', '')}
        Date: {summary.get('measurement_date', '')}
        
        Total Work Order Amount: Rs. {summary.get('total_work_order_amount', 0):,.2f}
        Total Bill Amount: Rs. {summary.get('total_bill_amount', 0):,.2f}
        Total Deviation: Rs. {summary.get('total_deviation', 0):,.2f}
        Deviation Percentage: {summary.get('deviation_percentage', 0)}%
        
        Items with Quantity Deviations:
        """
        
        for _, row in deviation_df.iterrows():
            statement += f"""
            Item: {row['Item_Description']}
            Work Order Quantity: {row['Work_Order_Qty']}
            Actual Quantity: {row['Bill_Qty']}
            Deviation: {row['Deviation_Qty']} ({row['Deviation_Percentage']}%)
            Financial Impact: Rs. {row['Deviation_Amount']:,.2f}
            Reason: <To be filled by Engineer>
            """
        
        statement += """
        
        Certified that the above deviations were necessary for the proper execution of the work.
        
        Junior Engineer / Assistant Engineer                Executive Engineer
        """
        
        return statement

def generate_note_sheet(summary):
    """
    Generate note sheet document content using the HTML template
    
    Args:
        summary: Dictionary with bill summary data
    
    Returns:
        String containing note sheet text
    """
    # Calculate some fields if not already present
    if 'sd_amount' not in summary:
        summary['sd_amount'] = round(0.1 * summary.get('total_bill_amount', 0), 2)
    
    if 'it_amount' not in summary:
        summary['it_amount'] = round(0.02 * summary.get('total_bill_amount', 0), 2)
    
    if 'gst_amount' not in summary:
        summary['gst_amount'] = round(0.02 * summary.get('total_bill_amount', 0), 2)
    
    if 'lc_amount' not in summary:
        summary['lc_amount'] = round(0.01 * summary.get('total_bill_amount', 0), 2)
    
    if 'balance_amount' not in summary:
        summary['balance_amount'] = max(0, summary.get('total_work_order_amount', 0) - summary.get('total_bill_amount', 0))
    
    if 'cheque_amount' not in summary:
        total_deductions = summary.get('sd_amount', 0) + summary.get('it_amount', 0) + summary.get('gst_amount', 0) + summary.get('lc_amount', 0)
        summary['cheque_amount'] = summary.get('total_bill_amount', 0) - total_deductions
    
    if 'payment_now' not in summary:
        summary['payment_now'] = summary.get('cheque_amount', 0)
    
    # Format dates correctly
    date_format = '%d/%m/%Y'
    for date_key in ['start_date', 'completion_date', 'actual_completion_date', 'measurement_date']:
        if date_key in summary:
            try:
                # Convert from DD-MM-YYYY to DD/MM/YYYY if needed
                if '-' in summary[date_key]:
                    date_obj = datetime.strptime(summary[date_key], '%d-%m-%Y')
                    summary[date_key] = date_obj.strftime(date_format)
            except Exception:
                pass
    
    # Format for note sheet template
    note_sheet_data = {
        'header': {
            'agreement_no': summary.get('agreement_no', ''),
            'name_of_work': summary.get('work_name', ''),
            'name_of_firm': summary.get('contractor_name', ''),
            'date_commencement': summary.get('start_date', ''),
            'date_completion': summary.get('completion_date', ''),
            'actual_completion': summary.get('actual_completion_date', '')
        },
        'work_order_amount': summary.get('total_work_order_amount', 0),
        'totals': {
            'original_payable': summary.get('total_bill_amount', 0)
        },
        'deductions': {
            'recovery_sd': summary.get('sd_amount', 0),
            'recovery_it': summary.get('it_amount', 0),
            'recovery_gst': summary.get('gst_amount', 0),
            'recovery_lc': summary.get('lc_amount', 0),
            'recovery_deposit_v': 0,
            'liquidated_damages': 0,
            'by_cheque': summary.get('cheque_amount', 0),
            'payment_now': summary.get('payment_now', 0)
        },
        'notes': [
            f"Submitted for approval of payment Rs. {summary.get('payment_now', 0)}/- against {summary.get('bill_serial', '')} for work {summary.get('work_name', '')}."
        ]
    }
    
    try:
        template = env.get_template('note_sheet.html')
        return template.render(note_sheet_data=note_sheet_data)
    except Exception as e:
        st.error(f"Error generating Note Sheet: {str(e)}")
        
        # Fallback to original format
        return f"""
        NOTE SHEET

        Bill for Agreement No. {summary.get('agreement_no', '')}

        1. Chargeable Head: 8443-00-108-00-00
        2. Agreement No.: {summary.get('agreement_no', '')}
        3. Adm. Section: 
        4. Tech. Section: 
        5. M.B No.: 887/Pg. No. 04-20
        6. Name of Sub Dn: Rajsamand
        7. Name of Work: {summary.get('work_name', '')}
        8. Name of Firm: {summary.get('contractor_name', '')}
        9. Original/Deposit: Deposit
        11. Date of Commencement: {summary.get('start_date', '')}
        12. Date of Completion: {summary.get('completion_date', '')}
        13. Actual Date of Completion: {summary.get('actual_completion_date', '')}
        14. In case of delay whether Provisional Extension Granted: {("Yes" if summary.get('delay_days', 0) > 0 else "No delay")}
        15. Whether any notice issued: 
        16. Amount of Work Order Rs.: {summary.get('total_work_order_amount', 0)}
        17. Actual Expenditure up to this Bill Rs.: {summary.get('total_bill_amount', 0)}
        18. Balance to be done Rs.: {summary.get('balance_amount', 0)}

        Net Amount of This Bill Rs.: {summary.get('total_bill_amount', 0)}
        Deductions:
            S.D. @ 10%: {summary.get('sd_amount', 0)}
            I.T. @ 2%: {summary.get('it_amount', 0)}
            GST @ 2%: {summary.get('gst_amount', 0)}
            L.C. @ 1%: {summary.get('lc_amount', 0)}
            Dep-V: 0
            Liquidated Damages: 0
            Cheque: {summary.get('cheque_amount', 0)}
            Total: {summary.get('payment_now', 0)}

        Submitted for approval of payment Rs. {summary.get('payment_now', 0)}/- against {summary.get('bill_serial', '')} for work {summary.get('work_name', '')}.
        """

def generate_extra_items_slip(extra_items_df, summary):
    """
    Generate extra items slip document content using the HTML template
    
    Args:
        extra_items_df: DataFrame with extra items data
        summary: Dictionary with bill summary data
    
    Returns:
        String containing extra items slip text
    """
    if extra_items_df is None or len(extra_items_df) == 0:
        return "No extra items found in the bill data."
    
    # Format data for the template
    extra_items = []
    for _, row in extra_items_df.iterrows():
        item_no = row.get('S_No', '')
        description = row.get('Item_Description', '')
        unit = row.get('Unit', '')
        quantity = row.get('Bill_Qty', 0)
        rate = row.get('Bill_Rate', 0)
        amount = row.get('Bill_Amount', quantity * rate)
        
        extra_items.append({
            'item_no': item_no,
            'description': description,
            'unit': unit,
            'quantity': quantity,
            'rate': rate,
            'amount': amount,
            'remarks': ""
        })
    
    extra_items_data = {
        'items': extra_items,
        'header': {
            'agreement_no': summary.get('agreement_no', ''),
            'work_name': summary.get('work_name', ''),
            'contractor_name': summary.get('contractor_name', '')
        },
        'total_amount': sum(item['amount'] for item in extra_items)
    }
    
    try:
        template = env.get_template('extra_items.html')
        return template.render(data=extra_items_data)
    except Exception as e:
        st.error(f"Error generating Extra Items Slip: {str(e)}")
        
        # Fallback to original format
        slip = f"""
        EXTRA ITEMS SLIP
        
        Name of Work: {summary.get('work_name', '')}
        Agreement No.: {summary.get('agreement_no', '')}
        Name of Contractor: {summary.get('contractor_name', '')}
        Bill: {summary.get('bill_serial', '')}
        Date: {summary.get('measurement_date', '')}
        
        The following extra items were executed as per site requirements:
        """
        
        total_extra_amount = 0
        
        for item in extra_items:
            item_amount = item['amount']
            total_extra_amount += item_amount
            
            slip += f"""
            Item: {item['description']}
            Quantity: {item['quantity']}
            Rate: Rs. {item['rate']:,.2f}
            Amount: Rs. {item_amount:,.2f}
            Reason: <To be filled by Engineer>
            """
        
        slip += f"""
        
        Total Extra Items Amount: Rs. {total_extra_amount:,.2f}
        
        Certified that the above extra items were necessary for the proper execution of the work and were executed as per the approval of the competent authority.
        
        Junior Engineer / Assistant Engineer                Executive Engineer
        """
        
        return slip

def export_to_excel(df, summary):
    """
    Export processed data and summary to Excel.
    
    Args:
        df: DataFrame with processed bill data
        summary: Dictionary with bill summary
        
    Returns:
        Excel file as bytes object
    """
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write the header information
        header_data = [
            ["Name of Contractor or supplier:", summary.get('contractor_name', '')],
            ["Name of Work:", summary.get('work_name', '')],
            ["Serial No. of this bill:", summary.get('bill_serial', '')],
            ["No. and date of the last bill-", summary.get('last_bill', 'Not Applicable')],
            ["Reference to work order or Agreement:", summary.get('work_order_ref', '')],
            ["Agreement No.", summary.get('agreement_no', '')],
            ["Date of written order to commence work:", summary.get('order_date', '')],
            ["St. date of Start:", summary.get('start_date', '')],
            ["Scheduled Date of Completion:", summary.get('completion_date', '')],
            ["Actual date of Completion:", summary.get('actual_completion_date', '')],
            ["Date of MB measurement:", summary.get('measurement_date', '')],
            ["Work Order Amount (Rs.):", summary.get('work_order_amount', '')],
            ["Bill Amount (Rs.):", summary.get('total_bill_amount', 0)]
        ]
        
        # Add deviation information if final bill
        if summary.get('is_final_bill'):
            header_data.extend([
                ["Deviation (Rs.):", summary.get('total_deviation', 0)],
                ["Deviation (%):", summary.get('deviation_percentage', 0)]
            ])
        
        header_df = pd.DataFrame(header_data)
        header_df.to_excel(writer, sheet_name='Bill', index=False, header=False)
        
        # Add empty rows after header
        empty_rows = 2
        empty_df = pd.DataFrame([[""] * len(df.columns)] * empty_rows)
        empty_df.to_excel(writer, sheet_name='Bill', startrow=len(header_data) + 1, index=False, header=False)
        
        # Write bill data
        df.to_excel(writer, sheet_name='Bill', startrow=len(header_data) + empty_rows + 1, index=False)
        
        # Add other sheets for reports and certificates
        # Certificate-II
        cert2_text = generate_certificate_2(summary).replace('<br>', '\n')
        cert2_df = pd.DataFrame({"Certificate-II": [cert2_text]})
        cert2_df.to_excel(writer, sheet_name='Certificate-II', index=False)
        
        # Certificate-III
        cert3_text = generate_certificate_3(summary).replace('<br>', '\n')
        cert3_df = pd.DataFrame({"Certificate-III": [cert3_text]})
        cert3_df.to_excel(writer, sheet_name='Certificate-III', index=False)
        
        # Note Sheet
        note_text = generate_note_sheet(summary).replace('<br>', '\n')
        note_df = pd.DataFrame({"Note Sheet": [note_text]})
        note_df.to_excel(writer, sheet_name='Note Sheet', index=False)
        
        # Deviation Statement (if final bill)
        if summary.get('is_final_bill'):
            dev_text = generate_deviation_statement(df, summary).replace('<br>', '\n')
            dev_df = pd.DataFrame({"Deviation Statement": [dev_text]})
            dev_df.to_excel(writer, sheet_name='Deviation Statement', index=False)
            
        # Extra Items (if any)
        if hasattr(st.session_state, 'extra_items_df') and not st.session_state.extra_items_df.empty:
            extra_items_df = st.session_state.extra_items_df
            extra_items_df.to_excel(writer, sheet_name='Extra Items', index=False)
            
            # Also add extra items slip
            extra_text = generate_extra_items_slip(extra_items_df, summary).replace('<br>', '\n')
            extra_slip_df = pd.DataFrame({"Extra Items Slip": [extra_text]})
            extra_slip_df.to_excel(writer, sheet_name='Extra Items Slip', index=False)
    
    # Get the value of the Excel file
    output.seek(0)
    return output.getvalue()

def export_to_pdf(df, summary):
    """
    Placeholder for PDF export functionality.
    In a real implementation, this would use a PDF library to generate a PDF document.
    
    Args:
        df: DataFrame with processed bill data
        summary: Dictionary with bill summary
        
    Returns:
        PDF file as bytes object (currently returns a message)
    """
    # This is a placeholder - in a real implementation, you'd use a library like ReportLab, WeasyPrint, or pdfkit
    return "PDF export functionality will be implemented in a future update."

def export_to_word(df, summary):
    """
    Placeholder for Word document export functionality.
    In a real implementation, this would use a library to generate a Word document.
    
    Args:
        df: DataFrame with processed bill data
        summary: Dictionary with bill summary
        
    Returns:
        Word document as bytes object (currently returns a message)
    """
    # This is a placeholder - in a real implementation, you'd use a library like python-docx
    return "Word export functionality will be implemented in a future update."