import streamlit as st
import pandas as pd
import pdf417
from PIL import Image
import fitz  # PyMuPDF
import io
def barcode_tooll():
    
    group_size = st.number_input(
    "🔢 Enter number of IMEIs to group per barcode (1 to 50):",
    min_value=1,
    max_value=50,
    value=5,
    step=1
)

# Step 2: Upload CSV file
    uploaded_file = st.file_uploader("Upload CSV file with PalletID and IMEIs", type=["csv"])
    if not uploaded_file:
        st.info("Please upload a CSV file to begin.")
        return None, False
    df = pd.read_csv(uploaded_file, dtype=str)
    df.columns = [col.strip() for col in df.columns]
    df["PalletID"] = df["PalletID"].astype(str)
    df["IMEI"] = df["IMEI"].astype(str)
    st.write("✅ Detected columns:", df.columns.tolist())
    df = df.iloc[1:].reset_index(drop=True)
    pdf_doc = fitz.open()
    
    with st.spinner("🔄 Generating barcodes..."):
        for i in range(0, len(df), group_size):
            group = df.iloc[i:i+group_size]
            pallet_id = group["PalletID"].iloc[0]
            imeis = group["IMEI"].tolist()
            barcode_data = "\r".join(imeis)

            for cols in range(3, 31):
                try:
                    codes = pdf417.encode(barcode_data, columns=cols, security_level=5)
                    break
                except ValueError as e:
                    if "Data too long" in str(e):
                        continue
                    else:
                        st.error(f"❌ Error encoding barcode for pallet {pallet_id}: {e}")
                        codes = None
                        break

            if codes:
                image = pdf417.render_image(codes, scale=6, ratio=6, padding=10)
                img_byte_arr = io.BytesIO()
                image.save(img_byte_arr, format='PNG')
                img_byte_arr = img_byte_arr.getvalue()

                page = pdf_doc.new_page(width=842, height=595)
                barcode_rect = fitz.Rect(0, 0, 842, 495)
                page.insert_image(barcode_rect, stream=img_byte_arr, keep_proportion=False)

                page.draw_rect(fitz.Rect(0, 495, 842, 595), color=(1, 1, 1), fill=(1, 1, 1))
                text_rect = fitz.Rect(150, 505, 842, 585)
                page.insert_textbox(
                    text_rect,
                    pallet_id,
                    fontsize=36,
                    fontname="helv",
                    align=1,
                    color=(0, 0, 0)
                )
    pdf_bytes = pdf_doc.write()
    pdf_doc.close()

    return pdf_bytes, True            
# Step 1: Let user select group size before upload
