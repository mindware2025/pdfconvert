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
    ("m365", "microsoft 365", "office 365", "exchange online"): "MS-CNS",
    ("acronis",): "AS-CNS",
    ("power bi",): "MS-CNS",
    ("planner", "project plan"): "MS-CNS",
    ("power automate premium",): "MS-CNS",
    ("visio",): "MS-CNS",
    ("dynamics 365",): "MS-CNS"
}

def fmt_date(value):
    try:
        dt = _parser.parse(str(value), dayfirst=False, fuzzy=True)
        return f"{dt.month:02d}/{dt.day:02d}/{dt.year}"
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
        item_desc_raw = str(row.get("ITEMDescription", ""))
        item_desc_lower = item_desc_raw.lower()
        invoice_desc = str(row.get("InvoiceDescription", ""))
        sub_id = row.get("SubscriptionId", "")
        if pd.notna(sub_id) and str(sub_id).strip():
            out_row["Subscription Id"] = str(sub_id).strip()
        elif "msaz-cns" in item_desc_lower:
            digits = extract_digits(invoice_desc)
            out_row["Subscription Id"] = digits[-36:] if len(digits) >= 1 else ""
        elif "msri-cns" in item_desc_lower or "ms-ri-cns" in item_desc_lower:
            out_row["Subscription Id"] = invoice_desc[:38]
        elif "reserved vm instance" in item_desc_lower:
            out_row["Subscription Id"] = item_desc_raw[:38]
        else:
            out_row["Subscription Id"] = "Sub"
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
                if matched: break
            if not matched:
                out_row["ITEM Code"] = ""
        # ITEM Name merged description
        out_row["ITEM Name"] = f"{item_desc_raw}-{row.get('ITEMName','')}-{out_row['Subscription Id']}#{out_row['Billing Cycle Start Date']}-{out_row['Billing Cycle End Date']}"
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
        end_user = str(row.get("EndUser", ""))
        end_user_country = str(row.get("EndUserCountryCode", ""))
        out_row["End User"] = f"{end_user} ; {end_user_country}"
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
    processed_df["Updated Invoice No."] = processed_df["Combined (D)"].map(mapping)
    processed_df.drop(columns=["Combined (D)"], inplace=True)
    return processed_df
