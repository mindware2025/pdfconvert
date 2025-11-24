import io
import re
from openpyxl import Workbook
import pandas as pd
from datetime import datetime
from dateutil import parser as _parser

# === Header Definition ===
CLOUD_INVOICE_HEADER = [
    "Invoice No.", "Customer Code", "Customer Name", "Invoice Date", "Document Location",
    "Sale Location", "Delivery Location Code", "Delivery Date", "Annotation", "Currency Code",
    "Exchange Rate", "Shipment Mode", "Payment Term", "Mode Of Payment", "Status",
    "Credit Card Transaction No.", "HEADER Discount Code", "HEADER Discount %", "HEADER Currency", "HEADER Basis",
    "HEADER Disc Value", "HEADER Expense Code", "HEADER Expense %", "HEADER Expense Currency", "HEADER Expense Basis",
    "HEADER Expense Value", "Subscription Id", "Billing Cycle Start Date", "Billing Cycle End Date",
    "ITEM Code", "ITEM Name", "UOM", "Grade code-1", "Grade code-2", "Quantity", "Qty Loose",
    "Rate Per Qty", "Gross Value", "ITEM Discount Code", "ITEM Discount %", "ITEM Discount Currency", "ITEM Discount Basis",
    "ITEM Disc Value", "ITEM Expense Code", "ITEM Expense %", "ITEM Expense Currency", "ITEM Expense Basis",
    "ITEM Expense Value", "ITEM Tax Code", "ITEM Tax %", "ITEM Tax Currency", "ITEM Tax Basis", "ITEM Tax Value",
    "LPO Number", "End User", "Cost"
]
GUID_WS = (
    r'[0-9a-fA-F]{8}\s*-\s*'
    r'[0-9a-fA-F]{4}\s*-\s*'
    r'[0-9a-fA-F]{4}\s*-\s*'
    r'[0-9a-fA-F]{4}\s*-\s*'
    r'[0-9a-fA-F]{12}'
)
manual_token_re = re.compile(
    r'(?i)manual.*?((?:[A-Za-z0-9]{1,10}\s*-\s*)?' + GUID_WS + r')',
    flags=re.DOTALL
)
# === Mappings ===
exchange_rate_map = {
    "UJ000": 1,
    "TC000": 0.272294078,
    "QA000": 0.274725274725,
    "OM000": 2.60078023407,
    "KA000": 0.2666666666
}
tax_code_map = {
    "WT000": "", "QA000": "", "TC000": "SLVAT5",
    "OM000": "SLVAT5", "UJ000": "SEVAT0", "KA000": "SLVAT15"
}
tax_percent_map = {
    "WT000": "", "QA000": "", "TC000": 5,
    "OM000": 5, "UJ000": 0, "KA000": 15
}
currency_map = {
    "WT000": "", "QA000": "", "TC000": "AED",
    "OM000": "OMR", "UJ000": "USD", "KA000": "SAR"
}
keyword_map = {
    ("windows server", "window server"): "MSPER-CNS",
    ("ms-azr", "azure subscription"): "MSAZ-CNS",
    ("google workspace",): "GL-WSP-CNS",
    ("m365", "microsoft 365", "office 365", "exchange online","Microsoft Defender for Endpoint P1"): "MS-CNS",
    ("acronis",): "AS-CNS",
    ("windows 11 pro",): "MSPER-CNS",
    ("power bi",): "MS-CNS",
    ("planner", "project plan"): "MS-CNS",
    ("power automate premium",): "MS-CNS",
    ("visio",): "MS-CNS",
    ("dynamics 365",): "MS-CNS"
}

def fmt_date(value):
    try:
        dt = _parser.parse(str(value), dayfirst=False, fuzzy=True)
        return f"{dt.day:02d}/{dt.month:02d}/{dt.year}"
    except Exception:
        return str(value) if value is not None else ""

def extract_digits(s: str) -> str:
    return "".join(ch for ch in str(s) if ch.isdigit())

def build_cloud_invoice_df(df: pd.DataFrame) -> pd.DataFrame:
    today = datetime.today()
    today_str = f"{today.month:02d}/{today.day:02d}/{today.year}"
    out_rows = []
    for _, row in df.iterrows():
        cost = row.get("Cost", "")
        try: cost_val = float(cost)
        except: cost_val = 0
        out_row = {}
        doc_loc = row.get("DocumentLocation", "")
        gross_value = row.get("GrossValue", 0)
        out_row["Invoice No."] = row.get("InvoiceNo", "")
        out_row["Customer Code"] = row.get("CustomerCode", "")
        out_row["Customer Name"] = row.get("CustomerName", "")
        out_row["Invoice Date"] = today_str
        out_row["Document Location"] = doc_loc
        out_row["Sale Location"] = row.get("SaleLocation", "")
        out_row["Delivery Location Code"] = row.get("DeliveryLocationCode", "")
        out_row["Delivery Date"] = today_str
        out_row["Annotation"] = ""
        out_row["Currency Code"] = row.get("CurrencyCode", "")
        out_row["Exchange Rate"] = exchange_rate_map.get(doc_loc, "")
        out_row["Shipment Mode"] = row.get("ShipmentMode", "")
        out_row["Payment Term"] = row.get("PaymentTerm", "")
        out_row["Mode Of Payment"] = row.get("ModeOfPayment", "")
        out_row["Status"] = "Unpaid"
        out_row["Credit Card Transaction No."] = ""
        for field in CLOUD_INVOICE_HEADER[16:26]:
            out_row[field] = ""
        item_code = str(row.get("ITEMCode", "")).strip().lower()
        item_desc_raw = str(row.get("ITEMDescription", ""))
        item_desc_lower = item_desc_raw.lower()
        invoice_desc = str(row.get("InvoiceDescription", "")).strip()
        sub_id_raw = row.get("SubscriptionId", "")
        sub_id = str(sub_id_raw).strip() if pd.notna(sub_id_raw) else ""
        
        
        invoice_desc_clean = re.sub(r"^[#\s]+", "", invoice_desc)
        
        
        sub_id_clean = sub_id[:36] if sub_id else "Sub"
        
        item_name_raw = str(row.get("ITEMName", "")).strip()
        item_name_lower = item_name_raw.lower()
        print("[DBG] ITEMName(raw):", item_name_raw)
        print("[DBG] Contains 'manual'?:", ("manual" in item_name_lower))
        print("[DBG] item_code:", item_code)
        print("[DBG] sub_id_clean (fallback):", sub_id_clean)
        
        # ONE decision tree: Manual first, then az-cns / msri-cns / reserved / default
        if "manual" in item_name_lower:
            m = manual_token_re.search(item_name_raw)
            print("[DBG] Manual GUID/token match?:", bool(m), "| span:", (m.span(1) if m else None))
            if m:
                token = m.group(1)
                token_norm = re.sub(r'\s*-\s*', '-', token)  # normalize whitespace around hyphens
                out_row["Subscription Id"] = token_norm
                print("[DBG] Subscription Id from Manual:", token_norm)
            else:
                m2 = re.search(r'(?i)manual\S*\s+([^\s#:\-]+)', item_name_raw, flags=re.DOTALL)
                print("[DBG] Manual fallback match?:", bool(m2), "| group:", (m2.group(1) if m2 else None))
                if m2:
                    candidate = m2.group(1).strip().strip('#').strip(':').strip('-')
                    out_row["Subscription Id"] = candidate
                    print("[DBG] Subscription Id from Manual fallback:", candidate)
                else:
                    out_row["Subscription Id"] = sub_id_clean
                    print("[DBG] Manual path failed; using sub_id_clean:", out_row["Subscription Id"])
        
        elif item_code == "az-cns":
            digits = extract_digits(invoice_desc_clean)
            out_row["Subscription Id"] = digits[-36:] if digits else sub_id_clean
            print("[DBG] az-cns digits:", digits, "| chosen:", out_row["Subscription Id"])
        
        elif item_code == "msri-cns":
            out_row["Subscription Id"] = invoice_desc_clean[:36] if invoice_desc_clean else sub_id_clean
            print("[DBG] msri-cns chosen:", out_row["Subscription Id"])
        
        elif "reserved vm instance" in item_desc_lower:
            out_row["Subscription Id"] = item_desc_raw[:38] if item_desc_raw else sub_id_clean
            print("[DBG] reserved-vm chosen:", out_row["Subscription Id"])
        
        else:
            out_row["Subscription Id"] = sub_id_clean
            print("[DBG] default chosen:", out_row["Subscription Id"])
        out_row["Billing Cycle Start Date"] = fmt_date(row.get("BillingCycleStartDate", ""))
        out_row["Billing Cycle End Date"] = fmt_date(row.get("BillingCycleEndDate", ""))
       
        item_code = row.get("ITEMCode", "")
        if pd.notna(item_code) and str(item_code).strip():
            out_row["ITEM Code"] = item_code
        else:
            matched = False
            for keywords, code in keyword_map.items():
                for k in keywords:
                    if k in item_desc_lower:
                        out_row["ITEM Code"] = code
                        matched = True
                        break
                if matched:
                    break
            if not matched:
                out_row["ITEM Code"] = ""
        
        # ITEM Name merged description
        # Clean and prepare values
        item_code_upper = out_row.get("ITEM Code", "").strip().upper()
        item_desc_raw = str(row.get("ITEMDescription", "")).strip()
        item_name_raw = str(row.get("ITEMName", "")).strip()
        invoice_desc_clean = re.sub(r"^[#\s]+", "", str(row.get("InvoiceDescription", "")).strip())
        sub_id_raw = row.get("SubscriptionId", "")
        sub_id_full = str(sub_id_raw).strip() if pd.notna(sub_id_raw) else ""
        
        # Choose item_name_detail based on ITEM Code
        if item_code_upper == "MSRI-CNS":
            item_name_detail = invoice_desc_clean
        elif item_code_upper in ["MSAZ-CNS", "AS-CNS", "AWS-UTILITIES-CNS"]:
            item_name_detail = sub_id_full  # full SubscriptionId
        else:
            item_name_detail = out_row.get("Subscription Id", "").strip()
        
        # Billing cycle dates
        billing_start = out_row.get("Billing Cycle Start Date", "").strip()
        billing_end = out_row.get("Billing Cycle End Date", "").strip()
        
        # Format billing_end as MM/YYYY if applicable
        billing_info = ""
        if billing_end and billing_end.lower() != "nan":
            try:
                billing_end_dt = pd.to_datetime(billing_end, dayfirst=True)
                billing_info = billing_end_dt.strftime("%m/%Y")
            except Exception:
                billing_info = billing_end  # fallback
        
        ## Build parts list and skip empty or NaN values
        ##parts = [
         #   item_desc_raw,
         #   item_name_raw,
         #   item_name_detail,
        #]
        #    parts.append(f"{billing_start}-{billing_end}")
       # 
       # 
       # # Join non-empty parts with hyphen
       # out_row["ITEM Name"] = "-".join([p for p in parts if p and p.lower() != "nan"])
        # Special rule for KA000: use only column AE (ITEMName)
        if doc_loc == "KA000":
            out_row["ITEM Name"] = item_desc_raw.strip()
        else:
            # Build parts list and skip empty or NaN values
            parts = [
                item_desc_raw,
                item_name_raw,
                item_name_detail,
            ]
        
            
            if item_code_upper == "MSRI-CNS" and billing_info:
                parts.append(billing_info)
            elif item_code_upper in ["MSAZ-CNS", "AS-CNS", "AWS-UTILITIES-CNS"] and billing_info:
                parts.append(billing_info)
            elif billing_start and billing_end and billing_start.lower() != "nan" and billing_end.lower() != "nan":
                parts.append(f"{billing_start}-{billing_end}")
            
                # Join non-empty parts with hyphen
            out_row["ITEM Name"] = "-".join([p for p in parts if p and p.lower() != "nan"])

            # Join non-empty parts with hyphen
        
        #out_row["ITEM Name"] = f"{item_desc_raw}-{row.get('ITEMName','')}-{out_row['Subscription Id']}#{out_row['Billing Cycle Start Date']}-{out_row['Billing Cycle End Date']}"
        out_row["UOM"] = row.get("UOM", "")
        out_row["Grade code-1"] = "NA"
        out_row["Grade code-2"] = "NA"
        out_row["Quantity"] = row.get("Quantity", "")
        out_row["Qty Loose"] = row.get("QtyLoose", "")
        try:
            out_row["Rate Per Qty"] = float(gross_value) / float(row.get("Quantity", 1))
        except: out_row["Rate Per Qty"] = 0
        try:
            out_row["Gross Value"] = round(float(gross_value), 2)
        except: out_row["Gross Value"] = 0.00
        for field in CLOUD_INVOICE_HEADER[38:48]:
            out_row[field] = ""
        out_row["ITEM Tax Code"] = tax_code_map.get(doc_loc, "")
        out_row["ITEM Tax %"] = tax_percent_map.get(doc_loc, "")
        out_row["ITEM Tax Currency"] = currency_map.get(doc_loc, "")
        out_row["ITEM Tax Basis"] = "R" if doc_loc not in ["WT000", "QA000"] else ""
        try: gross_val = float(gross_value)
        except: gross_val = 0
        tax_value = {
            "TC000": round(gross_val * 0.05, 2),
            "OM000": round(gross_val * 0.05, 2),
            "KA000": round(gross_val * 0.15, 2),
            "UJ000": 0
        }.get(doc_loc, "")
        out_row["ITEM Tax Value"] = tax_value
        lpo = row.get("LPONumber", "")
        out_row["LPO Number"] = "" if pd.isna(lpo) or str(lpo).strip().lower() in ["nan", "none"] else str(lpo)[:30]
       
        end_user = str(row.get("EndUser", "")).strip()
        end_user_country = str(row.get("EndUserCountryCode", "")).strip()
        
        # Define what values are considered invalid
        invalid_values = ["", "nan", "none"]
        
        # Check for invalid End User or Country Code
        if end_user.lower() in invalid_values or end_user_country.lower() in invalid_values:
            out_row["End User"] = ""  # Show empty string if missing
            out_row["_highlight_end_user"] = True  # Flag for red highlight
        else:
            out_row["End User"] = f"{end_user} ; {end_user_country}"
            out_row["_highlight_end_user"] = Falseend_user = str(row.get("EndUser", "")).strip()
        end_user_country = str(row.get("EndUserCountryCode", "")).strip()
        
        if (not end_user or end_user.lower() in ["nan", "none"]) or (not end_user_country or end_user_country.lower() in ["nan", "none"]):

            out_row["End User"] = end_user  # or just empty if you prefer ""
            out_row["_highlight_end_user"] = True
        else:
    # Both present, no highlight
            out_row["End User"] = f"{end_user} ; {end_user_country}"
            out_row["_highlight_end_user"] = False
        

        try: out_row["Cost"] = round(float(cost_val), 2)
        except: out_row["Cost"] = cost
        out_rows.append(out_row)

    result_df = pd.DataFrame(out_rows, columns=CLOUD_INVOICE_HEADER)

    # AS-CNS aggregation
    try:
        if not result_df.empty:
            is_as_cns = result_df["ITEM Code"].astype(str).str.strip().str.upper() == "AS-CNS"
            df_as = result_df[is_as_cns].copy()
            df_other = result_df[~is_as_cns].copy()
            if not df_as.empty:
                df_as["Gross Value"] = pd.to_numeric(df_as["Gross Value"], errors="coerce").fillna(0.0)
                df_as["Cost"] = pd.to_numeric(df_as["Cost"], errors="coerce").fillna(0.0)
                group_cols = ["Invoice No.", "End User", "LPO Number"]
                agg = df_as.groupby(group_cols, as_index=False).agg({"Gross Value": "sum"})
                merged = agg.merge(
                    df_as.drop(columns=["Gross Value"]).drop_duplicates(subset=group_cols),
                    on=group_cols,
                    how="left"
                )
                merged["Gross Value"] = merged["Gross Value"].round(2)
                merged["Quantity"] = 1
                merged["Rate Per Qty"] = merged["Gross Value"]
                merged["Cost"] = merged["Gross Value"]
    
                # <-- FIX: recalculate ITEM Tax Value based on summed Gross Value -->
                tax_rate_map = {"TC000": 0.05, "OM000": 0.05, "KA000": 0.15, "UJ000": 0}
                merged["ITEM Tax Value"] = merged.apply(
                    lambda x: round(x["Gross Value"] * tax_rate_map.get(x["Document Location"], 0), 2),
                    axis=1
                )
    
                result_df = pd.concat([df_other, merged], ignore_index=True)[CLOUD_INVOICE_HEADER]
    except Exception:
        pass

    return result_df

# === Versioning / Invoice Number Mapping ===

def create_summary_sheet(processed_df: pd.DataFrame) -> pd.DataFrame:
    sorted_df = processed_df.sort_values(by=processed_df.columns.tolist()).reset_index(drop=True)
    summary_df = sorted_df[["Invoice No.", "LPO Number", "End User"]].copy()
    summary_df["Combined (D)"] = (
        summary_df["Invoice No."].astype(str) +
        summary_df["LPO Number"].astype(str) +
        summary_df["End User"].astype(str)
    )
    summary_df = summary_df.drop_duplicates(subset=["Combined (D)"]).reset_index(drop=True)
    summary_df["V1"] = summary_df["Invoice No."].ne(summary_df["Invoice No."].shift()).astype(int).replace(0, "")
    v2 = []
    for i, v1 in enumerate(summary_df["V1"]):
        if i == 0:
            v2.append(1 if v1 == 1 else "")
        else:
            if v1 == 1:
                v2.append(1)
            else:
                prev_v2 = v2[-1]
                v2.append(prev_v2 + 1 if isinstance(prev_v2, int) else "")
    summary_df["V2"] = v2
    summary_df["V3"] = "-" + summary_df["V1"].astype(str) + summary_df["V2"].astype(str)
    summary_df["V4"] = summary_df["Invoice No."].astype(str) + summary_df["V3"]
    return summary_df

def map_invoice_numbers(processed_df: pd.DataFrame) -> pd.DataFrame:
    """
    Returns processed_df with a NEW column 'Updated Invoice No.' instead of replacing the original.
    """
    summary_df = create_summary_sheet(processed_df)
    mapping = dict(zip(summary_df["Combined (D)"], summary_df["V4"]))
    processed_df["Combined (D)"] = (
        processed_df["Invoice No."].astype(str) +
        processed_df["LPO Number"].astype(str) +
        processed_df["End User"].astype(str)
    )
    processed_df.drop(columns=["Combined (D)"], inplace=True)
    return processed_df

def create_srcl_file(df):
   

    # --- Sheet 1 headers ---
    headers_head = [
        "S.No",
        "Date - (dd/MM/yyyy)",
        "Cust_Code",
        "Curr_Code",
        "FORM_CODE",
        "Doc_Src_Locn",
        "Location_Code",
        "Delivery_Location",
        "SalesmanID"
    ]

    # --- Sheet 2 headers ---
    headers_item = [
        "S.No",
        "Ref. Key",
        "Item_Code",
        "Item_Name",
        "Grade1",
        "Grade2",
        "UOM",
        "Qty",
        "Qty_Ls",
        "Rate",              # Unit rate
        "CI Number CL",
        "End User CL",
        "Subs ID CL",
        "MPC Billdate CL",   # Dynamic based on Document Location
        "Unit Cost CL",
        "Total"              # Moved to end
    ]

    wb = Workbook()

    # --- Sheet 1 ---
    ws_head = wb.active
    ws_head.title = "SALES_RET_HEAD"
    ws_head.append(headers_head)

    today_str = pd.Timestamp.today().strftime("%d/%m/%Y")

    # Sequential numeric S.No for header
    s_no = 1
    header_sno_map = {}
    for _, row in df.iterrows():
        versioned_inv = row.get("Versioned Invoice No.", row.get("Invoice No.", ""))
        if versioned_inv not in header_sno_map:
            header_sno_map[versioned_inv] = s_no
            ws_head.append([
                s_no,
                today_str,
                row.get("Customer Code", ""),
                row.get("Currency Code", ""),
                "0",  # FORM_CODE
                row.get("Document Location", ""),
                row.get("Document Location", ""),
                row.get("Delivery Location Code", ""),
                "ED068"
            ])
            s_no += 1

    # --- Sheet 2 ---
    ws_item = wb.create_sheet(title="SALES_RET_ITEM")
    ws_item.append(headers_item)

    item_counter = 1
    for _, row in df.iterrows():
        versioned_inv = row.get("Versioned Invoice No.", row.get("Invoice No.", ""))
        ref_key = header_sno_map.get(versioned_inv, "")
        doc_loc = str(row.get("Document Location", "")).strip().upper()

        # --- Clean and sanitize item name ---
        raw_item_name = str(row.get("ITEM Name", "")).strip()
        clean_item_name = re.sub(r"[\r\n]+", " ", raw_item_name)  # remove newlines (Ctrl+J)
        clean_item_name = clean_item_name.replace("'", "").replace('"', "")
        clean_item_name = re.sub(r"\s+", " ", clean_item_name).strip()[:240]

        # --- Numeric cleanup ---
        qty = abs(float(row.get("Quantity", 0) or 0))
        qty_ls = abs(float(row.get("Qty Loose", 0) or 0))
        rate = abs(float(row.get("Rate Per Qty", 0) or 0))
        total = abs(round(qty * rate, 2))
        unit_cost = abs(float(row.get("Cost", 0) or 0))

        # --- MPC Billdate CL mapping ---
        if doc_loc in ["TC000", "UJ000"]:
            mpc_billdate = "UAE - 28"
        elif doc_loc == "QA000":
            mpc_billdate = "QAR - 28"
        elif doc_loc == "WT000":
            mpc_billdate = "KWT - 28"
        else:
            mpc_billdate = ""

        # --- Append cleaned data ---
        ws_item.append([
            item_counter,
            ref_key,
            row.get("ITEM Code", ""),
            clean_item_name,
            row.get("Grade code-1", ""),
            row.get("Grade code-2", ""),
            row.get("UOM", ""),
            qty,
            qty_ls,
            rate,  # ✅ Unit rate only
            versioned_inv,
            row.get("End User", ""),
            row.get("Subscription Id", ""),
            mpc_billdate,  # ✅ new logic
            unit_cost,
            total  # moved to end
        ])

        item_counter += 1

    # --- Save to memory ---
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
