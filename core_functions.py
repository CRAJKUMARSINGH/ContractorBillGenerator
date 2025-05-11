import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import traceback
import os

def read_excel_file(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    """
    Read Excel sheet with proper column names.
    
    Args:
        xls: Excel file object
        sheet_name: Name of the sheet to read
        
    Returns:
        DataFrame containing the sheet data
    
    Raises:
        ValueError: If required columns are missing or sheet is invalid
    """
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        
        # Validate sheet has enough rows
        min_rows = 20 if sheet_name in ["Bill Quantity", "Work Order"] else 6
        if len(df) < min_rows:
            raise ValueError(f"{sheet_name} sheet has insufficient rows: {len(df)}")

        # Skip rows based on sheet name
        if sheet_name in ["Bill Quantity", "Work Order"]:
            df = df.iloc[19:]  # Skip first 19 rows
        elif sheet_name == "Extra Items":
            df = df.iloc[5:]   # Skip first 5 rows
        
        # Reset index and assign generic column names
        df = df.reset_index(drop=True)
        df.columns = [f"Col_{i}" for i in range(len(df.columns))]
        
        # Convert necessary columns to appropriate types
        if sheet_name == "Bill Quantity":
            df['Col_3'] = pd.to_numeric(df['Col_3'], errors='coerce')  # Quantity
            df['Col_4'] = pd.to_numeric(df['Col_4'], errors='coerce')  # Rate
            
        return df
        
    except Exception as e:
        raise ValueError(f"Error reading {sheet_name} sheet: {str(e)}")

def process_bill(ws_wo: pd.DataFrame, ws_bq: pd.DataFrame, ws_extra: pd.DataFrame,
                 premium_percent: float, premium_type: str, amount_paid_last_bill: float,
                 is_first_bill: bool, user_inputs: dict) -> tuple:
    """
    Process bill data from Excel sheets and prepare data for templates.
    
    Args:
        ws_wo: Work Order sheet DataFrame
        ws_bq: Bill Quantity sheet DataFrame
        ws_extra: Extra Items sheet DataFrame
        premium_percent: Premium percentage
        premium_type: Premium type ("Fixed" or "Percentage")
        amount_paid_last_bill: Amount paid in previous bill
        is_first_bill: Boolean indicating if this is the first bill
        user_inputs: Dictionary of user inputs from the form
        
    Returns:
        Tuple of data dictionaries for each document
    
    Raises:
        ValueError: If validation fails
    """
    try:
        # Validate input types
        for df, name in [(ws_bq, "Bill Quantity"), (ws_wo, "Work Order"), (ws_extra, "Extra Items")]:
            if not isinstance(df, pd.DataFrame):
                raise ValueError(f"{name} sheet must be a DataFrame")
        
        # Validate required columns
        if ws_bq.shape[1] < 5:
            raise ValueError(f"Bill Quantity sheet has insufficient columns: {ws_bq.shape[1]}")
        
        # Validate bill type
        bill_type = user_inputs.get("bill_type", "").lower()
        if bill_type not in ["final bill", "running bill"]:
            raise ValueError("Invalid bill type. Must be either 'Final Bill' or 'Running Bill'")
        
        # Validate dates
        for date_field in ["start_date", "completion_date", "actual_completion_date", "order_date"]:
            if not isinstance(user_inputs.get(date_field), (datetime, date)):
                raise ValueError(f"{date_field.replace('_', ' ').title()} is required")
        
        if user_inputs["start_date"] > user_inputs["completion_date"]:
            raise ValueError("Start date cannot be after completion date")
        
        # Validate quantities and rates
        if ws_bq["Col_3"].lt(0).any():
            raise ValueError("Negative quantities are not allowed")
        if ws_bq["Col_4"].lt(0).any():
            raise ValueError("Negative rates are not allowed")
        
        # Validate premium
        if not isinstance(premium_percent, (int, float)):
            raise ValueError("Premium percentage must be a number")
        if premium_percent < -100 or premium_percent > 100:
            raise ValueError("Premium percentage must be between -100 and 100")
        
        # Validate amount paid for non-first bills
        if not is_first_bill and amount_paid_last_bill <= 0:
            raise ValueError("Amount paid via last bill is mandatory for non-first bills")

        # Initialize data structures
        first_page_data = {
            "contractor_name": user_inputs.get("contractor_name", ""),
            "work_name": user_inputs.get("work_name", ""),
            "bill_serial": user_inputs.get("bill_serial", ""),
            "agreement_no": user_inputs.get("agreement_no", ""),
            "work_order_ref": user_inputs.get("work_order_ref", ""),
            "work_order_amount": float(user_inputs.get("work_order_amount", 0)),
            "start_date": user_inputs["start_date"],
            "completion_date": user_inputs["completion_date"],
            "actual_completion_date": user_inputs["actual_completion_date"],
            "measurement_date": user_inputs.get("measurement_date"),
            "order_date": user_inputs["order_date"],
            "premium_percent": premium_percent,
            "premium_type": premium_type,
            "amount_paid_last_bill": amount_paid_last_bill,
            "is_first_bill": is_first_bill,
            "bill_type": bill_type,
            "bill_number": user_inputs.get("bill_number", ""),
            "items": []
        }

        # Process bill quantities
        bill_totals = {
            "total_quantity": 0,
            "total_amount": 0,
            "premium_amount": 0,
            "grand_total": 0
        }

        for idx, row in ws_bq.iterrows():
            try:
                qty = float(row["Col_3"]) if pd.notnull(row["Col_3"]) else 0
                rate = float(row["Col_4"]) if pd.notnull(row["Col_4"]) else 0
                desc = str(row["Col_1"]) if pd.notnull(row["Col_1"]) else "Item"
                amount = qty * rate
                
                first_page_data["items"].append({
                    "description": desc,
                    "quantity": qty,
                    "rate": rate,
                    "amount": amount
                })
                
                bill_totals["total_quantity"] += qty
                bill_totals["total_amount"] += amount
            except (ValueError, TypeError):
                continue

        # Calculate premium
        bill_totals["premium_amount"] = bill_totals["total_amount"] * (premium_percent / 100)
        bill_totals["grand_total"] = bill_totals["total_amount"] + bill_totals["premium_amount"]

        # Process extra items
        extra_items_data = {"items": [], "total": 0}
        if not ws_extra.empty and ws_extra.shape[1] >= 6:
            for idx, row in ws_extra.iterrows():
                try:
                    qty = float(row["Col_3"]) if pd.notnull(row["Col_3"]) else 0
                    rate = float(row["Col_5"]) if pd.notnull(row["Col_5"]) else 0
                    desc = str(row["Col_2"]) if pd.notnull(row["Col_2"]) else "Extra Item"
                    amount = qty * rate
                    
                    extra_items_data["items"].append({
                        "description": desc,
                        "quantity": qty,
                        "rate": rate,
                        "amount": amount
                    })
                    
                    extra_items_data["total"] += amount
                except (ValueError, TypeError):
                    continue

        # Prepare deviation data for final bill
        deviation_data = None
        if bill_type == "final bill":
            deviation_data = {
                "deviations": [],
                "total_deviation": bill_totals["grand_total"] - first_page_data["work_order_amount"]
            }

        return (
            first_page_data,
            bill_totals,
            deviation_data,
            extra_items_data,
            user_inputs
        )

    except Exception as e:
        error_msg = f"Error processing bill: {str(e)}\n{traceback.format_exc()}"
        raise ValueError(error_msg)

def generate_bill_notes(bill_data):
    """
    Generate bill notes based on bill data.
    
    Args:
        bill_data: Dictionary containing bill metadata
        
    Returns:
        str: Formatted bill notes
    """
    try:
        notes = [
            f"Contractor Name: {bill_data['contractor_name']}",
            f"Work Name: {bill_data['work_name']}",
            f"Bill Serial: {bill_data['bill_serial']}",
            f"Agreement No: {bill_data['agreement_no']}",
            f"Work Order Ref: {bill_data['work_order_ref']}",
            f"Work Order Amount: {bill_data['work_order_amount']:,.2f}",
            f"Premium: {bill_data['premium_percent']}%",
            f"Bill Type: {bill_data['bill_type']}",
            f"Bill Number: {bill_data['bill_number']}"
        ]
        
        if bill_data['is_first_bill']:
            notes.append("This is the first bill")
        else:
            notes.append(f"Amount Paid in Last Bill: {bill_data['amount_paid_last_bill']:,.2f}")
        
        return "\n".join(notes)
    except Exception as e:
        return f"Error generating notes: {str(e)}"

def generate_pdf(output_path, bill_data, bill_result):
    """
    Generate a PDF bill using ReportLab.
    
    Args:
        output_path: Path where PDF should be saved
        bill_data: Dictionary containing bill metadata
        bill_result: Processed bill result tuple
    """
    try:
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        elements = []
        styles = getSampleStyleSheet()
        
        # Title
        title = Paragraph("Contractor Bill", styles['Title'])
        elements.append(title)
        
        # Basic Information
        info = [
            ["Contractor Name:", bill_data['contractor_name']],
            ["Work Name:", bill_data['work_name']],
            ["Agreement No:", bill_data['agreement_no']],
            ["Work Order Ref:", bill_data['work_order_ref']],
            ["Bill Number:", bill_data['bill_number']],
            ["Bill Date:", datetime.now().strftime("%d-%m-%Y")]
        ]
        
        info_table = Table(info, colWidths=[150, 350])
        info_table.setStyle(TableStyle([
            ('FONT', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        elements.append(info_table)
        
        # Bill Quantities
        elements.append(Paragraph("\nBill Quantities", styles['Heading2']))
        
        qty_headers = ["Description", "Quantity", "Rate", "Amount"]
        qty_data = [qty_headers]
        for item in bill_result[0]['items']:
            qty_data.append([
                item['description'],
                f"{item['quantity']:.2f}",
                f"{item['rate']:.2f}",
                f"{item['amount']:.2f}"
            ])
        
        qty_table = Table(qty_data, colWidths=[250, 80, 80, 80])
        qty_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONT', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(qty_table)
        
        # Extra Items
        if bill_result[3] and bill_result[3]["items"]:
            elements.append(Paragraph("\nExtra Items", styles['Heading2']))
            
            extra_headers = ["Description", "Quantity", "Rate", "Amount"]
            extra_data = [extra_headers]
            for item in bill_result[3]['items']:
                extra_data.append([
                    item['description'],
                    f"{item['quantity']:.2f}",
                    f"{item['rate']:.2f}",
                    f"{item['amount']:.2f}"
                ])
            
            extra_table = Table(extra_data, colWidths=[250, 80, 80, 80])
            extra_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONT', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            elements.append(extra_table)
        
        # Total Amount
        elements.append(Paragraph("\nTotal Amount", styles['Heading2']))
        total = Paragraph(f"₹ {bill_result[1]['grand_total']:,.2f}", styles['Heading3'])
        elements.append(total)
        
        # Bill Notes
        notes = generate_bill_notes(bill_data).split('\n')
        notes_table = Table([[note] for note in notes], colWidths=[500])
        notes_table.setStyle(TableStyle([
            ('FONT', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('LEFTPADDING', (0, 0), (-1, -1), 20),
        ]))
        elements.append(Paragraph("\nBill Notes", styles['Heading2']))
        elements.append(notes_table)
        
        # Build PDF
        doc.build(elements)
        
    except Exception as e:
        raise ValueError(f"Error generating PDF: {str(e)}\n{traceback.format_exc()}")