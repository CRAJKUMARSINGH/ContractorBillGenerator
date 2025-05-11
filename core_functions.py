import pandas as pd
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

def read_excel_file(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    """
    Read Excel sheet with proper column names
    
    Args:
        xls: Excel file object
        sheet_name: Name of the sheet to read
        
    Returns:
        DataFrame containing the sheet data
    
    Raises:
        ValueError: If required columns are missing
    """
    try:
        # Read the sheet with optimized memory usage
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
        
        # Skip rows based on sheet name
        if sheet_name in ["Bill Quantity", "Work Order"]:
            df = df.iloc[19:]  # Skip first 19 rows
        elif sheet_name == "Extra Items":
            df = df.iloc[5:]   # Skip first 5 rows
        
        # Reset index
        df = df.reset_index(drop=True)
        
        # Convert necessary columns to appropriate types
        if sheet_name == "Bill Quantity":
            df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
            df['Rate'] = pd.to_numeric(df['Rate'], errors='coerce')
            
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
        premium_type: Premium type ("Percentage" or "Fixed")
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
        if not isinstance(ws_bq, pd.DataFrame):
            raise ValueError("Bill Quantity sheet must be a DataFrame")
        if not isinstance(ws_wo, pd.DataFrame):
            raise ValueError("Work Order sheet must be a DataFrame")
        if not isinstance(ws_extra, pd.DataFrame):
            raise ValueError("Extra Items sheet must be a DataFrame")
        
        # Map unnamed columns to our required columns
        column_mapping = {
            'Unnamed: 0': 'Description',
            'Unnamed: 1': 'Quantity',
            'Unnamed: 2': 'Rate'
        }
        
        # Rename columns in ws_bq
        ws_bq = ws_bq.rename(columns=column_mapping)
        
        # Validate required columns
        required_columns = ['Quantity', 'Rate', 'Description']
        missing_columns = [col for col in required_columns if col not in ws_bq.columns]
        if missing_columns:
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
            
        # Validate bill type
        bill_type = user_inputs.get("bill_type", "").lower()
        if bill_type not in ["final bill", "running bill"]:
            raise ValueError("Invalid bill type. Must be either 'Final Bill' or 'Running Bill'")
            
        # Validate dates
        start_date = user_inputs.get("start_date")
        completion_date = user_inputs.get("completion_date")
        if not isinstance(start_date, datetime):
            raise ValueError("Start date is required")
        if not isinstance(completion_date, datetime):
            raise ValueError("Completion date is required")
        
        if start_date > completion_date:
            raise ValueError("Start date cannot be after completion date")
            
        # Validate quantities and rates
        if any(ws_bq["Quantity"] < 0):
            raise ValueError("Negative quantities are not allowed")
        if any(ws_bq["Rate"] < 0):
            raise ValueError("Negative rates are not allowed")
            
        # Validate premium
        if not isinstance(premium_percent, (int, float)):
            raise ValueError("Premium percentage must be a number")
        if premium_percent < 0 or premium_percent > 100:
            raise ValueError("Premium percentage must be between 0 and 100")
            
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
            "work_order_amount": user_inputs.get("work_order_amount", 0),
            "start_date": start_date,
            "completion_date": completion_date,
            "actual_completion_date": user_inputs.get("actual_completion_date"),
            "measurement_date": user_inputs.get("measurement_date"),
            "order_date": user_inputs.get("order_date"),
            "premium_percent": premium_percent,
            "premium_type": premium_type,
            "amount_paid_last_bill": amount_paid_last_bill,
            "is_first_bill": is_first_bill,
            "bill_type": bill_type,
            "bill_number": user_inputs.get("bill_number", "")
        }

        # Calculate totals
        bill_totals = {
            "total_quantity": ws_bq["Quantity"].sum(),
            "total_rate": ws_bq["Rate"].sum(),
            "total_amount": (ws_bq["Quantity"] * ws_bq["Rate"]).sum(),
            "premium_amount": (ws_bq["Quantity"] * ws_bq["Rate"]).sum() * (premium_percent / 100),
            "grand_total": (ws_bq["Quantity"] * ws_bq["Rate"]).sum() + ((ws_bq["Quantity"] * ws_bq["Rate"]).sum() * (premium_percent / 100))
        }

        # Calculate extra items if present
        extra_items_data = None
        if not ws_extra.empty:
            extra_items_data = {
                "items": ws_extra.to_dict('records'),
                "total": (ws_extra["Quantity"] * ws_extra["Rate"]).sum()
            }

        # Prepare deviation data for final bill
        deviation_data = None
        if bill_type == "final bill":
            deviation_data = {
                "deviations": [],
                "total_deviation": 0
            }

        return (
            first_page_data,
            bill_totals,
            deviation_data,
            extra_items_data,
            user_inputs
        )

    except Exception as e:
        error_msg = f"Error processing bill: {str(e)}"
        print(f"Error details: {traceback.format_exc()}")
        raise ValueError(error_msg) from e

def generate_bill_notes(bill_data):
    """
    Generate bill notes based on bill data
    
    Args:
        bill_data: Dictionary containing bill metadata
        
    Returns:
        str: Formatted bill notes
    """
    notes = []
    notes.append(f"Contractor Name: {bill_data['contractor_name']}")
    notes.append(f"Work Name: {bill_data['work_name']}")
    notes.append(f"Bill Serial: {bill_data['bill_serial']}")
    notes.append(f"Agreement No: {bill_data['agreement_no']}")
    notes.append(f"Work Order Ref: {bill_data['work_order_ref']}")
    notes.append(f"Work Order Amount: {bill_data['work_order_amount']}")
    notes.append(f"Premium: {bill_data['premium_percent']}%")
    
    if bill_data['is_first_bill']:
        notes.append("This is the first bill")
    else:
        notes.append(f"Amount Paid in Last Bill: {bill_data['amount_paid_last_bill']}")
    
    notes.append(f"Bill Type: {bill_data['bill_type']}")
    notes.append(f"Bill Number: {bill_data['bill_number']}")
    
    if bill_data['last_bill']:
        notes.append("This is the last bill")
    
    return "\n".join(notes)

def generate_pdf(output_path, bill_data, bill_result):
    """
    Generate a PDF bill
    
    Args:
        output_path: Path where PDF should be saved
        bill_data: Dictionary containing bill metadata
        bill_result: Processed bill result
    """
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
    
    info_table = Table(info, colWidths=[150, None])
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
    qty_data = [qty_headers] + [[
        str(item['Description']),
        str(item['Quantity']),
        str(item['Rate']),
        str(item['Quantity'] * item['Rate'])
    ] for index, item in bill_result[0]['items'].iterrows()]
    
    qty_table = Table(qty_data, colWidths=[300, 100, 100, 100])
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
    if bill_result[3]:
        elements.append(Paragraph("\nExtra Items", styles['Heading2']))
        
        extra_headers = ["Description", "Quantity", "Rate", "Amount"]
        extra_data = [extra_headers] + [[
            str(item['Description']),
            str(item['Quantity']),
            str(item['Rate']),
            str(item['Quantity'] * item['Rate'])
        ] for item in bill_result[3]['items']]
        
        extra_table = Table(extra_data, colWidths=[300, 100, 100, 100])
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
    total = Paragraph(f"₹ {bill_result[1]['grand_total']:,}", styles['Heading3'])
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
