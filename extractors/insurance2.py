import pandas as pd
from io import BytesIO
from zipfile import ZipFile

def process_grouped_customer_files(file):
    # Read the Excel file
    df = pd.read_excel(file, engine="openpyxl")

    # Validate 'Status' column
    valid_statuses = {'UNPAID', 'CREDIT', 'PARTLY'}
    if not df['Status'].isin(valid_statuses).all():
        invalid_values = df.loc[~df['Status'].isin(valid_statuses), 'Status'].unique()
        raise ValueError(
            f"Invalid Status values found: {invalid_values}. Allowed values are 'UNPAID', 'CREDIT', 'PARTLY'."
        )

    # Validation for UNPAID status
    unpaid_issues = df[
        (df['Status'] == 'UNPAID') &
        ((df['Paid Amount'] != 0) | (df['Payment Date'].notna()))
    ]
    if not unpaid_issues.empty:
        raise ValueError(
            "Validation error: For status 'UNPAID', 'Paid Amount' must be 0 and 'Payment Date' must be blank."
        )

    # Ensure date columns are parsed correctly
    df['Document Date'] = pd.to_datetime(df['Document Date'], errors='coerce')
    df['Document Due Date'] = pd.to_datetime(df['Document Due Date'], errors='coerce')
    df['Payment Date'] = pd.to_datetime(df['Payment Date'], errors='coerce')

    # Generate formatted output per row (single line)
    df['Formatted Output'] = df.apply(
        lambda row: f"{row['Document Number']};"
                    f"{row['Document Date'].strftime('%d/%m/%Y') if pd.notnull(row['Document Date']) else ''};"
                    f"{row['Document Due Date'].strftime('%d/%m/%Y') if pd.notnull(row['Document Due Date']) else ''};"
                    f"{int(row['Ar Balance']) if pd.notnull(row['Ar Balance']) else ''};"
                    f"{row['Status']};"
                    f"{row['Paid Amount'] if pd.notnull(row['Paid Amount']) else ''};"
                    f"{int(round(row['Paid Amount'])) if pd.notnull(row['Paid Amount']) else ''};"
                    f"{row['Payment Date'].strftime('%d/%m/%Y') if pd.notnull(row['Payment Date']) else ''};"
                    f"{row['reason of edd'] if pd.notnull(row['reason of edd']) else ''}",
        axis=1
    )

    # Group by Cust Code
    grouped = df.groupby('Cust Code')

    # Create a zip file in memory
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, 'w') as zip_file:
        for cust_code, group in grouped:
            output_df = group[['Formatted Output']]

            # Save to CSV
            csv_buffer = BytesIO()
            output_df.to_csv(csv_buffer, index=False, header=False)
            csv_buffer.seek(0)

            # Write to ZIP with custom filename
            filename = f"{cust_code}.csv"
            zip_file.writestr(filename, csv_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer