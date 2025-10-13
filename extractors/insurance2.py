import pandas as pd
import os
from io import BytesIO
from zipfile import ZipFile

def process_grouped_customer_files(file):
    # Read the Excel file
    df = pd.read_excel(file, engine="openpyxl")

    # Ensure date columns are parsed correctly
    df['Document Date'] = pd.to_datetime(df['Document Date'], errors='coerce')
    df['Document Due Date'] = pd.to_datetime(df['Document Due Date'], errors='coerce')

    # Generate formatted output per row
    df['Formatted Output'] = df.apply(lambda row: f"{row['Document Number']};"
                                                  f"{row['Document Date'].strftime('%d/%m/%Y') if pd.notnull(row['Document Date']) else ''};"
                                                  f"{row['Document Due Date'].strftime('%d/%m/%Y') if pd.notnull(row['Document Due Date']) else ''};"
                                                  f"{int(row['Ageing']) if pd.notnull(row['Ageing']) else ''};"
                                                  f"UNPAID;0;;"
                                                  f"{row['reason of edd'] if pd.notnull(row['reason of edd']) else ''}",
                                     axis=1)

    # Group by Cust Code
    grouped = df.groupby('Cust Code')

    # Create a zip file in memory
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, 'w') as zip_file:
        for cust_code, group in grouped:
            output_df = group[['Formatted Output']]

            # Save to Excel
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name='Formatted Output')
            excel_buffer.seek(0)
            zip_file.writestr(f"{cust_code}.xlsx", excel_buffer.read())

            # Save to CSV
            csv_buffer = BytesIO()
            output_df.to_csv(csv_buffer, index=False)
            csv_buffer.seek(0)
            zip_file.writestr(f"{cust_code}.csv", csv_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer