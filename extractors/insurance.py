import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl.styles import NamedStyle

def process_insurance_excel(file, ageing_filter=True, ageing_threshold=200):
    # Read the Excel file and clean column names
    df = pd.read_excel(file, skiprows=15, engine='openpyxl')
    df.columns = [str(col).strip() for col in df.columns]

    # Convert date columns to datetime format
    for col in df.columns:
        if 'date' in col.lower():
            df[col] = pd.to_datetime(df[col], format='%m/%d/%Y', errors='coerce')

    # Calculate ageing based on today's date
    today = pd.to_datetime(datetime.today())
    if 'Document Date' in df.columns:
        df['Ageing'] = (today - df['Document Date']).dt.days

    # Filter rows based on insurance limit and AR balance
    filtered_df = df[
        (df['Total Insurance Limit'] > 0) &
        (df['Ar Balance'] > 0)
    ]
    if ageing_filter and 'Ageing' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['Ageing'] > ageing_threshold]

    # Add new columns
    filtered_df['Status'] = 'UNPAID'
    filtered_df['reason of edd'] = 'Undergoing reconciliation'
    filtered_df['Paid Amount'] = 0
    filtered_df['Payment Date'] = pd.NaT  # Blank by default

    # Define output columns including new fields
    output_columns = [
        'Cust Code', 'Cust Name', 'Document Number', 'Document Date',
        'Document Due Date', 'Ageing', 'Over Due Days',
        'Total Insurance Limit', 'Ar Balance', 'Paid Amount', 'Payment Date',
        'Status', 'reason of edd'
    ]
    final_df = filtered_df[[col for col in output_columns if col in filtered_df.columns]]

    # Write to Excel with date formatting
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Insurance Filtered')

        workbook = writer.book
        worksheet = writer.sheets['Insurance Filtered']

        # Apply date formatting
        date_style = NamedStyle(name="date_style", number_format='MM/DD/YYYY')
        if "date_style" not in workbook.named_styles:
            workbook.add_named_style(date_style)

        for col_idx, col_name in enumerate(final_df.columns, start=1):
            if 'date' in col_name.lower():
                for row in worksheet.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.style = date_style

    output.seek(0)
    return output