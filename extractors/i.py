
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl.styles import NamedStyle

def process_insurance_excel(
    file,
    ageing_filter=True,
    ageing_min_threshold=200,
    ageing_max_threshold=270
):
    """
    Reads the insurance Excel, calculates Ageing, filters by:
      - Total Insurance Limit > 0
      - Ar Balance >= 1
      - Ageing in [ageing_min_threshold, ageing_max_threshold] if ageing_filter=True
    Returns an in-memory Excel (BytesIO) with formatting.
    """
    # Read the Excel file and clean column names
    df = pd.read_excel(file, skiprows=15, engine='openpyxl')
    df.columns = [str(col).strip() for col in df.columns]

    # Convert date-like columns to datetime
    for col in df.columns:
        if 'date' in col.lower():
            df[col] = pd.to_datetime(df[col], errors='coerce')

    # Calculate ageing based on today's date
    today = pd.to_datetime(datetime.today())
    if 'Document Date' in df.columns:
        df['Ageing'] = (today - df['Document Date']).dt.days

    # Base filters: Total Insurance Limit > 0 and Ar Balance >= 1
    filtered_df = df[
        (df['Total Insurance Limit'] > 0) &
        (df['Ar Balance'] >= 1)
    ].copy()

    # Round Ar Balance to nearest integer
    filtered_df['Ar Balance'] = filtered_df['Ar Balance'].round().astype(int)

    # Apply ageing range filter (inclusive) if enabled and Ageing exists
    if ageing_filter and 'Ageing' in filtered_df.columns:
        # Guardrails: ensure sensible bounds
        if ageing_min_threshold is not None and ageing_max_threshold is not None:
            # Swap if user accidentally sets min > max
            if ageing_min_threshold > ageing_max_threshold:
                ageing_min_threshold, ageing_max_threshold = ageing_max_threshold, ageing_min_threshold
            filtered_df = filtered_df[
                (filtered_df['Ageing'] >= ageing_min_threshold) &
                (filtered_df['Ageing'] <= ageing_max_threshold)
            ]
        elif ageing_min_threshold is not None:
            filtered_df = filtered_df[filtered_df['Ageing'] >= ageing_min_threshold]
        elif ageing_max_threshold is not None:
            filtered_df = filtered_df[filtered_df['Ageing'] <= ageing_max_threshold]

    # Add new columns
    filtered_df['Status'] = 'UNPAID'
    filtered_df['reason of edd'] = 'Undergoing reconciliation'
    filtered_df['Paid Amount'] = 0
    filtered_df['Payment Date'] = pd.NaT

    # Calculate Over Due Days if possible
    if 'Document Due Date' in filtered_df.columns:
        filtered_df['Over Due Days'] = (today - filtered_df['Document Due Date']).dt.days

    # Define output columns
    output_columns = [
        'Cust Code', 'Cust Name', 'Document Number', 'Document Date',
        'Document Due Date', 'Ageing', 'Over Due Days',
        'Total Insurance Limit', 'Ar Balance', 'Paid Amount', 'Payment Date',
        'Status', 'reason of edd'
    ]
    final_df = filtered_df[[col for col in output_columns if col in filtered_df.columns]]

    # Write to Excel with formatting
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Insurance Filtered')
        workbook = writer.book
        worksheet = writer.sheets['Insurance Filtered']

        # Date formatting
        date_style = NamedStyle(name="date_style", number_format='MM/DD/YYYY')
        if "date_style" not in getattr(workbook, "named_styles", []):
            try:
                workbook.add_named_style(date_style)
            except Exception:
                # Some openpyxl versions may have duplicate name issues; ignore safely
                pass

        for col_idx, col_name in enumerate(final_df.columns, start=1):
            if 'date' in col_name.lower():
                for row in worksheet.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.style = "date_style"

        # Numeric formatting
        numeric_style = NamedStyle(name="numeric_style", number_format='0')
        if "numeric_style" not in getattr(workbook, "named_styles", []):
            try:
                workbook.add_named_style(numeric_style)
            except Exception:
                pass

        for col_name in ['Ar Balance', 'Paid Amount']:
            if col_name in final_df.columns:
                col_idx = final_df.columns.get_loc(col_name) + 1
                for row in worksheet.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.style = "numeric_style"

    output.seek(0)
    return output
