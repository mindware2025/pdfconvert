# PDF to Excel Table Extractor

This app extracts the 'Summary of costs by domain' table from a PDF and saves it as an Excel file.

## Requirements
- Python 3.7+
- pip

## Setup
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Place your PDF file in the project directory. By default, the script looks for:
   - `5269277427_31-May-2025 - Copy.pdf`

3. Run the script:
   ```bash
   python pdf_to_excel.py
   ```

4. The extracted table will be saved as `summary_of_costs_by_domain.xlsx`.

## Notes // to be updated
- The script only extracts tables with the header:
  - Domain name | Customer ID | Amount(US$)
- If the table is not found, the script will notify you.
- The script works for multiple pages as long as the table structure is the same. 