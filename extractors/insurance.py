import pandas as pd
from datetime import datetime
from io import BytesIO
from zipfile import ZipFile
import re

def sanitize_filename(name):
    # Replace invalid characters with underscores
    return re.sub(r'[\\/:*?"<>|]', '_', str(name))

def process_insurance_excel(file, ageing_threshold=200):
    # Read the Excel file
    df = pd.read_excel(file, engine="openpyxl")
    
    # Ensure date columns are parsed correctly
    df['Document Date'] = pd.to_datetime(df['Document Date'], errors='coerce')
    df['Document Due Date'] = pd.to_datetime(df['Document Due Date'], errors='coerce')

    # Calculate ageing
    today = pd.to_datetime(datetime.today())
    if 'Document Date' in df.columns:
        df['Ageing'] = (today - df['Document Date']).dt.days

    # Filter by ageing threshold
    if 'Ageing' in df.columns:
        df = df[df['Ageing'] > ageing_threshold]

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
            cust_name = group['Cust Name'].iloc[0] if 'Cust Name' in group.columns else 'Unknown'

            # Sanitize filename
            filename = sanitize_filename(f"{cust_code}_{cust_name}.csv")

            # Save to CSV
            csv_buffer = BytesIO()
            output_df.to_csv(csv_buffer, index=False, header=False)
            csv_buffer.seek(0)

            # Write to ZIP
            zip_file.writestr(filename, csv_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer
