import streamlit as st
import pandas as pd
import pdf417
from PIL import Image
import fitz  # PyMuPDF
import io

def barcode_tooll():

    # Step 1: Group size input
    group_size = st.number_input(
        "🔢 Enter number of IMEIs to group per barcode (1 to 50):",
        min_value=1,
        max_value=50,
        value=5,
        step=1
    )

    # Step 2: Upload CSV file
    uploaded_file = st.file_uploader("📁 Upload CSV file with PalletID and IMEIs", type=["csv"])
    if not uploaded_file:
        st.info("ℹ️ Please upload a CSV file to begin.")
        return None, False

    try:
        df = pd.read_csv(uploaded_file, dtype=str)
    except Exception as e:
        st.error(f"❌ Failed to read CSV file: {e}")
        return None, False

    # Clean column names
    df.columns = [col.strip() for col in df.columns]

    # Validate required columns
    required_columns = {"PalletID", "IMEI"}
    if not required_columns.issubset(df.columns):
        st.error(f"❌ CSV must contain the following columns: {required_columns}")
        return None, False

    # Convert to string and clean
    df["PalletID"] = df["PalletID"].astype(str).str.strip()
    df["IMEI"] = df["IMEI"].astype(str).str.strip()

    # Remove empty rows
    df = df.dropna(subset=["PalletID", "IMEI"])
    df = df[df["IMEI"] != ""]
    df = df.reset_index(drop=True)

    # Check for duplicate IMEIs
    if df["IMEI"].duplicated().any():
        duplicates = df[df["IMEI"].duplicated(keep=False)]
        st.error("❌ Duplicate IMEIs found. Please remove them before proceeding.")
        st.dataframe(duplicates)
        return None, False

    # Check if total IMEIs is divisible by group size
    total_imeis = len(df)
    if total_imeis % group_size != 0:
        st.error(f"❌ Total IMEIs ({total_imeis}) is not divisible by group size ({group_size}). Please adjust and try again.")
        return None, False

    # Sort by PalletID to ensure correct grouping
    #df = df.sort_values(by="PalletID").reset_index(drop=True)

    st.success("✅ File successfully loaded and validated.")
    st.write("📋 Detected columns:", df.columns.tolist())
    st.write(f"📦 Total IMEIs: {total_imeis}")

    # Step 3: Generate barcodes
    pdf_doc = fitz.open()

    with st.spinner("🔄 Generating barcodes..."):
        for i in range(0, total_imeis, group_size):
            group = df.iloc[i:i+group_size]
            pallet_id = group["PalletID"].iloc[0]
            imeis = group["IMEI"].tolist()
            barcode_data = "\r".join(imeis)

            # Try different column sizes for barcode
            codes = None
            for cols in range(3, 31):
                try:
                    codes = pdf417.encode(barcode_data, columns=cols, security_level=5)
                    break
                except ValueError as e:
                    if "Data too long" in str(e):
                        continue
                    else:
                        st.error(f"❌ Critical error encoding barcode for pallet {pallet_id}: {e}")
                        return None, False

            if not codes:
                st.error(f"❌ Unable to encode barcode for pallet {pallet_id}. Data may be too long or malformed.")
                return None, False

            # Render barcode image
            image = pdf417.render_image(codes, scale=6, ratio=6, padding=10)
            img_byte_arr = io.BytesIO()
            image.save(img_byte_arr, format='PNG')
            img_byte_arr = img_byte_arr.getvalue()

            # Create PDF page
            page = pdf_doc.new_page(width=842, height=595)
            barcode_rect = fitz.Rect(0, 0, 842, 495)
            page.insert_image(barcode_rect, stream=img_byte_arr, keep_proportion=False)

            # Add PalletID label
            page.draw_rect(fitz.Rect(0, 495, 842, 595), color=(1, 1, 1), fill=(1, 1, 1))
            text_rect = fitz.Rect(150, 505, 842, 585)
            page.insert_textbox(
                text_rect,
                f"Pallet ID: {pallet_id}",
                fontsize=36,
                fontname="helv",
                align=1,
                color=(0, 0, 0)
            )

    pdf_bytes = pdf_doc.write()
    pdf_doc.close()

    st.success("✅ Barcode PDF generated successfully.")
    return pdf_bytes, True