import pandas as pd
import re
from datetime import datetime

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

# === Highlighting Function ===
def highlight_red(val):
    val_str = str(val).strip().lower()
    return 'background-color: red' if val_str in ['', 'nan', 'none'] else ''

# === Main Processing Function ===
def process_cloud_invoice(df):
    today = datetime.today()
    today_str = f"{today.month:02d}-{today.day:02d}-{today.year}"
    out_rows = []

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
    

    for _, row in df.iterrows():
        cost = row.get("Cost", "")
        try:
            cost_val = float(cost)
        except Exception:
            cost_val = 0

        out_row = {}
        doc_loc = row.get("DocumentLocation", "")
        gross_value = row.get("GrossValue", 0)

       
        out_row["Invoice No."] = row.get("InvoiceNo", "") #to beeee updated
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

        for field in [
            "HEADER Discount Code", "HEADER Discount %", "HEADER Currency", "HEADER Basis", "HEADER Disc Value",
            "HEADER Expense Code", "HEADER Expense %", "HEADER Expense Currency", "HEADER Expense Basis", "HEADER Expense Value"
        ]:
            out_row[field] = ""

        item_desc_raw = str(row.get("ITEMDescription", ""))
        item_desc_lower = item_desc_raw.lower()
        invoice_desc = str(row.get("InvoiceDescription", ""))

        def extract_digits(s: str) -> str:
            return "".join(ch for ch in s if ch.isdigit())

        sub_id = row.get("SubscriptionId", "")
        if pd.notna(sub_id) and str(sub_id).strip() != "":
            out_row["Subscription Id"] = str(sub_id).strip()
        elif "msaz-cns" in item_desc_lower:
            # Last 36 digits from InvoiceDescription
            digits = extract_digits(invoice_desc)
            out_row["Subscription Id"] = digits[-36:] if len(digits) >= 1 else ""
        elif "msri-cns" in item_desc_lower or "ms-ri-cns" in item_desc_lower:
            # First 38 characters from InvoiceDescription
            out_row["Subscription Id"] = invoice_desc[:38]
        elif "reserved vm instance" in item_desc_lower:
            out_row["Subscription Id"] = item_desc_raw[:38]
        else:
            out_row["Subscription Id"] = "Sub"

        # Format billing cycle dates as MM-DD-YYYY
        def fmt_date(value):
            try:
                from dateutil import parser as _parser
                dt = _parser.parse(str(value), dayfirst=False, fuzzy=True)
                return f"{dt.month:02d}-{dt.day:02d}-{dt.year}"
            except Exception:
                return str(value) if value is not None else ""

        b_start = row.get("BillingCycleStartDate", "")
        b_end = row.get("BillingCycleEndDate", "")
        out_row["Billing Cycle Start Date"] = fmt_date(b_start)
        out_row["Billing Cycle End Date"] = fmt_date(b_end)

        # === ITEM Code Matcching ===
        item_code = row.get("ITEMCode", "")
        item_desc = str(row.get("ITEMDescription", "")).lower()

        if pd.notna(item_code) and str(item_code).strip() != "":
            out_row["ITEM Code"] = item_code
        else:
            matched = False
            for keywords, code in keyword_map.items():
                for k in keywords:
                    if k in item_desc:
                        out_row["ITEM Code"] = code
                        matched = True
                        break
                if matched:
                    break
            if not matched:
                out_row["ITEM Code"] = ""

        # === ITEM Name composition ===
        item_code_for_name = str(out_row.get("ITEM Code", "")).strip().upper()
        special_codes = {"MSAZ-CNS", "AS-CNS", "AWS-UTILITIES-CNS", "MS-RI-CNS", "MSRI-CNS"}
        if item_code_for_name in special_codes:
            # End date month-year (MM-YYYY)
            end_date_str = str(out_row.get("Billing Cycle End Date", ""))
            try:
                from dateutil import parser as _parser
                dt_end = _parser.parse(str(b_end), dayfirst=False, fuzzy=True)
                mm_yyyy = f"{dt_end.month:02d}-{dt_end.year}"
            except Exception:
                mm_yyyy = end_date_str
            merged_desc = (
                f"{str(row.get('ITEMDescription', ''))}-"
                f"{str(row.get('ITEMName', ''))}-"
                f"{item_code_for_name}-"
                f"{str(out_row.get('Subscription Id', ''))}-"
                f"{mm_yyyy}"
            )
            out_row["ITEM Name"] = merged_desc
        else:
            merged_desc = (
                str(row.get("ITEMDescription", "")) + "-" +
                str(row.get("ITEMName", "")) + "-" +
                str(out_row["Subscription Id"]) + "-" +
                str(out_row["Billing Cycle Start Date"]) + "-" +
                str(out_row["Billing Cycle End Date"])
            )
            out_row["ITEM Name"] = merged_desc

        out_row["UOM"] = row.get("UOM", "")
        out_row["Grade code-1"] = row.get("Gradecode1", "")
        out_row["Grade code-2"] = row.get("GradeCode2", "")
        out_row["Quantity"] = row.get("Quantity", "")
        out_row["Qty Loose"] = row.get("QtyLoose", "")
        quantity = row.get("Quantity", 1)
        try:
            out_row["Rate Per Qty"] = float(gross_value) / float(quantity)
        except Exception:
            out_row["Rate Per Qty"] = 0
        try:
            out_row["Gross Value"] = round(float(gross_value), 2)
        except Exception:
            out_row["Gross Value"] = 0.00

        for field in [
            "ITEM Discount Code", "ITEM Discount %", "ITEM Discount Currency", "ITEM Discount Basis", "ITEM Disc Value",
            "ITEM Expense Code", "ITEM Expense %", "ITEM Expense Currency", "ITEM Expense Basis", "ITEM Expense Value"
        ]:
            out_row[field] = ""

        # === Tax Fields ===
        out_row["ITEM Tax Code"] = tax_code_map.get(doc_loc, "")
        out_row["ITEM Tax %"] = tax_percent_map.get(doc_loc, "")
        out_row["ITEM Tax Currency"] = currency_map.get(doc_loc, "")
        out_row["ITEM Tax Basis"] = "R" if doc_loc not in ["WT000", "QA000"] else ""

        try:
            gross_val = float(gross_value)
        except Exception:
            gross_val = 0

        if doc_loc == "TC000":
            tax_value = round(gross_val * 0.05, 2)
        elif doc_loc == "OM000":
            tax_value = round(gross_val * 0.05, 2)
        elif doc_loc == "KA000":
            tax_value = round(gross_val * 0.15, 2)
        elif doc_loc == "UJ000":
            tax_value = 0
        else:
            tax_value = ""

        out_row["ITEM Tax Value"] = tax_value
        lpo = row.get("LPONumber", "")
        out_row["LPO Number"] = "" if pd.isna(lpo) or str(lpo).strip().lower() in ["nan", "none"] else str(lpo)[:30]


        end_user = str(row.get("EndUser", ""))
        end_user_country = str(row.get("EndUserCountryCode", ""))
        out_row["End User"] = f"{end_user} ; {end_user_country}"


        try:
            cost_val = float(cost)
            out_row["Cost"] = f"{cost_val:.2f}"
        except Exception:
            out_row["Cost"] = cost

        out_rows.append(out_row)

   
    result_df = pd.DataFrame(out_rows, columns=CLOUD_INVOICE_HEADER)

    # === Aggregation for AS-CNS ===
    try:
        if not result_df.empty:
            # Split AS-CNS and non AS-CNS
            is_as_cns = result_df["ITEM Code"].astype(str).str.strip().str.upper() == "AS-CNS"
            df_as = result_df[is_as_cns].copy()
            df_other = result_df[~is_as_cns].copy()

            if not df_as.empty:
                # Ensure numeric for Gross Value and Cost
                df_as["Gross Value"] = pd.to_numeric(df_as["Gross Value"], errors="coerce").fillna(0.0)
                df_as["Cost"] = pd.to_numeric(df_as["Cost"], errors="coerce").fillna(0.0)

                # Group by Invoice No., End User, and LPO Number
                group_cols = ["Invoice No.", "End User", "LPO Number"]
                agg = df_as.groupby(group_cols, as_index=False).agg({
                    "Gross Value": "sum"
                })

                # Take the first row's other fields for each group
                merged = (
                    agg.merge(
                        df_as.drop(columns=["Gross Value"]).drop_duplicates(subset=group_cols),
                        on=group_cols,
                        how="left"
                    )
                )

                # Set Quantity = 1, Rate Per Qty = sum, Cost = sum (numeric)
                merged["Gross Value"] = merged["Gross Value"].round(2)
                merged["Quantity"] = 1
                merged["Rate Per Qty"] = merged["Gross Value"]
                merged["Cost"] = merged["Gross Value"]

                # Recombine
                result_df = pd.concat([df_other, merged], ignore_index=True)[CLOUD_INVOICE_HEADER]
    except Exception:
        # On any unexpected issue, fall back to original result_df
        pass

    styled_df = result_df.style.applymap(highlight_red, subset=["Customer Code", "ITEM Code"])
    return styled_df


def build_cloud_invoice_df(df: pd.DataFrame) -> pd.DataFrame:
    """Build the Cloud Invoice output DataFrame (no styling), including AS-CNS aggregation.

    Returns a plain DataFrame with columns CLOUD_INVOICE_HEADER.
    """
    today = datetime.today()
    today_str = f"{today.month:02d}/{today.day:02d}/{today.year}"

    out_rows: list[dict] = []

    exchange_rate_map = {
        "UJ000": 1,
        "TC000": 0.272294078,
        "QA000": 0.274725274725,
        "OM000": 2.60078023407,
        "KA000": 0.2666666666,
    }
    tax_code_map = {"WT000": "", "QA000": "", "TC000": "SLVAT5", "OM000": "SLVAT5", "UJ000": "SEVAT0", "KA000": "SLVAT15"}
    tax_percent_map = {"WT000": "", "QA000": "", "TC000": 5, "OM000": 5, "UJ000": 0, "KA000": 15}
    currency_map = {"WT000": "", "QA000": "", "TC000": "AED", "OM000": "OMR", "UJ000": "USD", "KA000": "SAR"}
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
        ("dynamics 365",): "MS-CNS",
    }

    for _, row in df.iterrows():
        cost = row.get("Cost", "")
        try:
            cost_val = float(cost)
        except Exception:
            cost_val = 0

        out_row: dict = {}
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
        for field in [
            "HEADER Discount Code",
            "HEADER Discount %",
            "HEADER Currency",
            "HEADER Basis",
            "HEADER Disc Value",
            "HEADER Expense Code",
            "HEADER Expense %",
            "HEADER Expense Currency",
            "HEADER Expense Basis",
            "HEADER Expense Value",
        ]:
            out_row[field] = ""

        item_desc_raw = str(row.get("ITEMDescription", ""))
        item_desc_lower = item_desc_raw.lower()
        invoice_desc = str(row.get("InvoiceDescription", ""))
        def extract_digits(s: str) -> str:
            return "".join(ch for ch in s if ch.isdigit())
        sub_id = row.get("SubscriptionId", "")
        if pd.notna(sub_id) and str(sub_id).strip() != "":
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

        from dateutil import parser as _parser
        def fmt_date(value):
           try:
               from dateutil import parser as _parser
               dt = _parser.parse(str(value), dayfirst=False, fuzzy=True)
               return f"{dt.month:02d}/{dt.day:02d}/{dt.year}"
           except Exception:
               return str(value) if value is not None else ""

        b_start = row.get("BillingCycleStartDate", "")
        b_end = row.get("BillingCycleEndDate", "")
        out_row["Billing Cycle Start Date"] = fmt_date(b_start)
        out_row["Billing Cycle End Date"] = fmt_date(b_end)
        

        item_code = row.get("ITEMCode", "")
        item_desc = str(row.get("ITEMDescription", "")).lower()
        if pd.notna(item_code) and str(item_code).strip() != "":
            out_row["ITEM Code"] = item_code
        else:
            matched = False
            for keywords, code in keyword_map.items():
                for k in keywords:
                    if k in item_desc:
                        out_row["ITEM Code"] = code
                        matched = True
                        break
                if matched:
                    break
            if not matched:
                out_row["ITEM Code"] = ""

        item_code_for_name = str(out_row.get("ITEM Code", "")).strip().upper()
        special_codes = {"MSAZ-CNS", "AS-CNS", "AWS-UTILITIES-CNS", "MS-RI-CNS", "MSRI-CNS"}
        if item_code_for_name in special_codes:
            end_date_str = str(out_row.get("Billing Cycle End Date", ""))
            try:
                dt_end = _parser.parse(str(b_end), dayfirst=False, fuzzy=True)
                mm_yyyy = f"{dt_end.month:02d}-{dt_end.year}"
            except Exception:
                mm_yyyy = end_date_str
            merged_desc = (
                f"{str(row.get('ITEMDescription', ''))}-"
                f"{str(row.get('ITEMName', ''))}-"
                f"{item_code_for_name}-"
                f"{str(out_row.get('Subscription Id', ''))}-"
                f"{mm_yyyy}"
            )
        else:
            merged_desc = (
                str(row.get("ITEMDescription", ""))
                + "-"
                + str(row.get("ITEMName", ""))
                + "-"
                + str(out_row["Subscription Id"])
                + "-"
                + str(out_row["Billing Cycle Start Date"])
                + "-"
                + str(out_row["Billing Cycle End Date"])
            )
        out_row["ITEM Name"] = merged_desc

        out_row["UOM"] = row.get("UOM", "")
        out_row["Grade code-1"] = "N/A"
        out_row["Grade code-2"] = "N/A"
        out_row["Quantity"] = row.get("Quantity", "")
        out_row["Qty Loose"] = row.get("QtyLoose", "")
        quantity = row.get("Quantity", 1)
        try:
            out_row["Rate Per Qty"] = float(gross_value) / float(quantity)
        except Exception:
            out_row["Rate Per Qty"] = 0
        try:
            out_row["Gross Value"] = round(float(gross_value), 2)
        except Exception:
            out_row["Gross Value"] = 0.00

        for field in [
            "ITEM Discount Code",
            "ITEM Discount %",
            "ITEM Discount Currency",
            "ITEM Discount Basis",
            "ITEM Disc Value",
            "ITEM Expense Code",
            "ITEM Expense %",
            "ITEM Expense Currency",
            "ITEM Expense Basis",
            "ITEM Expense Value",
        ]:
            out_row[field] = ""

        out_row["ITEM Tax Code"] = tax_code_map.get(doc_loc, "")
        out_row["ITEM Tax %"] = tax_percent_map.get(doc_loc, "")
        out_row["ITEM Tax Currency"] = currency_map.get(doc_loc, "")
        out_row["ITEM Tax Basis"] = "R" if doc_loc not in ["WT000", "QA000"] else ""

        try:
            gross_val = float(gross_value)
        except Exception:
            gross_val = 0
        if doc_loc == "TC000":
            tax_value = round(gross_val * 0.05, 2)
        elif doc_loc == "OM000":
            tax_value = round(gross_val * 0.05, 2)
        elif doc_loc == "KA000":
            tax_value = round(gross_val * 0.15, 2)
        elif doc_loc == "UJ000":
            tax_value = 0
        else:
            tax_value = ""
        out_row["ITEM Tax Value"] = tax_value

        lpo = row.get("LPONumber", "")
        out_row["LPO Number"] = "" if pd.isna(lpo) or str(lpo).strip().lower() in ["nan", "none"] else str(lpo)[:30]

        end_user = str(row.get("EndUser", ""))
        end_user_country = str(row.get("EndUserCountryCode", ""))
        out_row["End User"] = f"{end_user} ; {end_user_country}"

        try:
            out_row["Cost"] = round(float(cost_val), 2)
        except Exception:
            out_row["Cost"] = cost

        out_rows.append(out_row)

    result_df = pd.DataFrame(out_rows, columns=CLOUD_INVOICE_HEADER)

    # AS-CNS aggregation (same as in process_cloud_invoice)
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
                    how="left",
                )
                merged["Gross Value"] = merged["Gross Value"].round(2)
                merged["Quantity"] = 1
                merged["Rate Per Qty"] = merged["Gross Value"]
                merged["Cost"] = merged["Gross Value"]
                result_df = pd.concat([df_other, merged], ignore_index=True)[CLOUD_INVOICE_HEADER]
    except Exception:
        pass

    return result_df


if __name__ == "__main__":
    input_df = pd.read_excel("your_input_file.xlsx")  
    styled_output = process_cloud_invoice(input_df)
    styled_output.to_excel("cloud_invoice_output.xlsx", index=False, engine="openpyxl")