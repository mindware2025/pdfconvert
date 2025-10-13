import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl.styles import NamedStyle

def process_insurance_excel(file, ageing_filter=True, ageing_threshold=200):
    # Read the Excel file, skipping the first 15 rows
    df = pd.read_excel(file, skiprows=15, engine='openpyxl')
    df.columns = [str(col).strip() for col in df.columns]

    # Convert 'Document Date' to datetime using MM/DD/YYYY format
    df['Document Date'] = pd.to_datetime(df['Document Date'], format='%m/%d/%Y', errors='coerce')

    # Calculate ageing in days
    today = pd.to_datetime(datetime.today())
    df['Ageing'] = (today - df['Document Date']).dt.days

    # Filter rows based on conditions
    filtered_df = df[
        (df['Total Insurance Limit'] > 0) &
        (df['Ar Balance'] > 0)
    ]
    if ageing_filter:
        filtered_df = filtered_df[filtered_df['Ageing'] > ageing_threshold]

    # Add status and reason columns
    filtered_df['Status'] = 'Unpaid'
    filtered_df['reason of edd'] = 'Undergoing reconciliation'

    # Select output columns
    output_columns = [
        'Cust Code', 'Cust Name', 'Document Number', 'Document Date',
        'Document Due Date', 'Ageing', 'Over Due Days',
        'Total Insurance Limit', 'Ar Balance', 'Status', 'reason of edd'
    ]
    final_df = filtered_df[output_columns]

    # Write to Excel with date formatting
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Insurance Filtered')

        # Access workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Insurance Filtered']

        # Create a date style
        date_style = NamedStyle(name="date_style", number_format='MM/DD/YYYY')
        if "date_style" not in workbook.named_styles:
            workbook.add_named_style(date_style)

        # Apply date style to 'Document Date' column (column D = index 4)
        for row in worksheet.iter_rows(min_row=2, min_col=4, max_col=4):
            for cell in row:
                cell.style = date_style

    output.seek(0)
    return output