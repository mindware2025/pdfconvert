# Automation Tracker – Management View

Use this for the **Excel tracker**: copy the **Management description** into your Description column and use the **Tracker columns** section as your sheet layout.

---

## Tracker columns (recommended for management)

| Column | Purpose | Example / formula |
|--------|--------|-------------------|
| **Tool** | Short name for reporting | Google DNTS, Claims Automation, IBM Quotation |
| **Owner team** | Who uses and benefits | Finance, Operations, Credit, Sales |
| **One-line summary** | Single sentence for dashboards | "Turns Google invoice PDFs into ERP-ready Excel in one click." |
| **Description** | Full management-friendly explanation (see below) | 2–4 sentences: problem, solution, outcome |
| **Manual process (min)** | Average time per run before automation | 45 |
| **Automated process (min)** | Average time per run with the tool | 5 |
| **Time saved per run (min)** | Formula: Manual − Automated | `=F2-G2` |
| **Runs per month** | How often the process is run | 20 |
| **Hours saved per month** | Formula: (Time saved per run × Runs) ÷ 60 | `=(H2*I2)/60` |
| **Development (hours)** | One-off build effort | 40 |
| **Payback (months)** | Formula: Development ÷ Hours saved per month | `=J2/K2` |
| **Risk without automation** | What goes wrong if done manually | High error rate, delayed closing |
| **Status** | Live / Planned / Pilot | Live |
| **Priority** | For roadmap discussions | High / Medium / Low |
| **Notes** | Optional (e.g. countries, templates) | UAE, Qatar, KSA; Template 1 & 2 |

---

## 1. Google DNTS Extractor  
**Owner team:** Finance  

**One-line summary:**  
Turns Google Cloud invoice PDFs into the exact Excel format required for our finance system (DNTS/Orion) in one step.

**Description (for management):**  
Google sends invoices as PDFs with a “Summary of costs by domain” table. Previously, staff retyped domain names, customer IDs, and amounts into Excel and then adjusted layout and codes for the finance system. This tool reads the PDF, pulls invoice number and date, extracts every cost line, and produces a ready-to-upload Excel in the correct format. This cuts manual data entry, speeds up posting, and reduces errors that could affect reconciliation and reporting.

---

## 2. Google Invoice Extractor  
**Owner team:** Finance  

**One-line summary:**  
Quickly pulls domain-level costs from Google invoice PDFs into a simple Excel for analysis or checks.

**Description (for management):**  
When the full DNTS format is not needed—for example for a quick check or a report—staff still need the cost breakdown (domain, customer ID, amount) from the PDF. This tool does that in one go: upload the Google invoice PDF and get an Excel with Domain name, Customer ID, and Amount. It saves time on ad-hoc extractions and avoids copy-paste mistakes.

---

## 3. Claims Automation  
**Owner team:** Finance  

**One-line summary:**  
Combines four different claims-related files (SAP export, user list, benefits, account mapping) into one ready-to-use claims file.

**Description (for management):**  
Claims processing used to depend on four separate files: SAP journal data, employee master, benefit details, and account mapping. Staff had to merge them, match employees to accounts, and build narrations by hand. This tool takes all four as input and produces a single, consistent claims file with the right account codes and narrations. It reduces reconciliation errors, speeds up month-end, and keeps audit trails clear.

---

## 4. AWS Invoice Tool  
**Owner team:** Finance  

**One-line summary:**  
Reads AWS invoice and credit-note PDFs and produces both a summary Excel and the correct files for our finance system (DNTS/CNTS) by entity.

**Description (for management):**  
AWS sends different PDF layouts (tax invoice, credit note, direct vs marketplace). Staff used to re-enter amounts, VAT, and billing details and then build DNTS/CNTS files per entity (e.g. Mindware UAE vs FZ). The tool detects the PDF type, extracts amounts and dates, calculates VAT where needed, and generates both a summary Excel and the right DNTS or CNTS files in one ZIP. This reduces re-keying, ensures correct entity and VAT treatment, and speeds up AWS invoice posting.

---

## 5. Dell Invoice Extractor (Pre-Alert Upload)  
**Owner team:** Operations  

**One-line summary:**  
Turns Dell invoice PDFs into the pre-alert upload sheet and highlights lines that need review or master-data updates.

**Description (for management):**  
Pre-alert requires data from Dell PDFs (PO, invoice number, dates, line items, shipping, consolidation) to be entered into our template and matched to internal item codes and prices. Doing this by hand is slow and error-prone. The tool reads one or more Dell PDFs (and an optional master file), fills the pre-alert sheet, matches items where possible, and adds a review sheet that flags “no match,” “price match,” or “ok” so staff can focus on exceptions. This shortens pre-alert turnaround and improves accuracy of item and price matching.

---

## 6. Cloud Invoice Tool  
**Owner team:** Operations  

**One-line summary:**  
Converts the cloud billing export (after CB Convert) into standard cloud invoice Excel, negative-invoice file, and SRCL file with correct codes and versioning.

**Description (for management):**  
Cloud billing data comes from a CB export that must be converted and then mapped to our locations, tax codes, item codes (e.g. M365, Azure), LPO, and end user. Staff used to do this in Excel with a lot of manual lookup and versioning. The tool takes the converted file, applies location and tax rules, derives item codes from descriptions, extracts LPO and end user, and produces the main cloud invoice workbook (with versioning), the negative-invoices file, and the SRCL file. It keeps coding consistent and reduces manual parsing and versioning errors.

---

## 7. Barcode PDF Generator (grouped)  
**Owner team:** Operations  

**One-line summary:**  
Generates pallet barcode labels from a simple list of Pallet ID and IMEI so warehouse can print and apply labels without manual barcode creation.

**Description (for management):**  
Shipping and warehouse need barcode labels per pallet, with each label encoding a set of IMEIs. Creating these manually is tedious and error-prone. Staff upload a CSV with Pallet ID and IMEI; the tool checks that IMEIs are valid and grouped correctly, then produces a PDF of barcode pages (one page per group) with the pallet ID shown. This speeds up labeling and reduces wrong or missing barcodes.

---

## 8. AR to EDD File  
**Owner team:** Credit  

**One-line summary:**  
Filters the insurance/AR export to the right ageing window and adds the columns needed for EDD and reconciliation in one go.

**Description (for management):**  
Credit needs to work from the insurance file with only relevant rows (e.g. positive limit, balance, and ageing between 200–270 days) and standard columns for EDD (status, reason, paid amount, payment date, over-due days). Doing this manually with filters and new columns is repetitive and easy to get wrong. The tool reads the export (from row 16), applies the business rules and ageing range, and outputs a clean Excel ready for EDD and reconciliation. This keeps the process consistent and saves time each cycle.

---

## 9. Coface CSV Uploader  
**Owner team:** Credit  

**One-line summary:**  
Splits one customer-invoice export into separate, correctly formatted CSV files per customer for Coface (or similar) submission.

**Description (for management):**  
Coface (or similar) often requires one file per customer in a specific format: document number, dates, balance, status, paid amount, payment date, reason (semicolon-separated). Manually splitting and formatting from one big export is slow and risks wrong formatting or wrong customer split. The tool checks status rules (e.g. UNPAID must have zero paid amount and blank payment date), then groups by customer code and outputs one CSV per customer in the required format, zipped. This ensures consistent format and reduces manual splitting and copy-paste errors.

---

## 10. IBM Quotation  
**Owner team:** Sales  

**One-line summary:**  
Takes an IBM quotation PDF (and optional Excel), detects the quotation type, and produces a single styled quotation Excel in the right currency for UAE, Qatar, or KSA.

**Description (for management):**  
IBM sends quotation PDFs in different formats (parts with coverage dates vs subscription/SaaS). Sales used to copy data between PDF and Excel, set currency (AED/SAR/USD) by country, and keep terms and layout consistent. The tool lets the user choose country (UAE, Qatar, KSA), uploads the PDF and optionally the Excel, and automatically detects whether it is a “parts” or “subscription” quotation. It then builds one styled quotation Excel (with logo and terms), checks total price vs PDF where relevant, and aligns dates between PDF and Excel. This reduces manual rework, keeps pricing and compliance text consistent across countries, and speeds up quotation turnaround.

---

## 11. MIBB Quotations  
**Owner team:** Sales  

**One-line summary:**  
(Planned) Same idea as IBM Quotation but for MIBB quotation PDFs—one-click extraction and styled Excel with correct terms.

**Description (for management):**  
This tool is planned for MIBB quotation PDFs: extract header and line table (part number, description, dates, qty, price), optionally apply a description master, and produce a styled quotation Excel with MIBB terms. It is not yet connected in the app; list it in the tracker as **Planned** and use it for roadmap and capacity discussions.

---

## How to use in Excel

1. **Header row:** Use the column names from the table (Tool, Owner team, One-line summary, Description, Manual process (min), Automated process (min), etc.).
2. **Description column:** Paste the **Description (for management)** text for each tool (you can trim to fit if needed).
3. **One-line summary:** Use for dashboards, slides, or a summary sheet.
4. **Formulas:**  
   - Time saved per run = Manual − Automated  
   - Hours saved per month = (Time saved per run × Runs per month) ÷ 60  
   - Payback (months) = Development (hours) ÷ Hours saved per month  
5. **Risk without automation:** Fill with 1–2 words or a short phrase (e.g. "High error rate," "Delayed month-end," "Wrong currency in quote") so management can see why the tool matters.
