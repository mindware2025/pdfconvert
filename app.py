
import base64
import csv
import logging
import os
from pathlib import Path
import zipfile
import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl.styles import PatternFill
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import gspread
from google.oauth2.service_account import Credentials
from extractors.barcodeper50 import barcode_tooll
from extractors.aws import AWS_OUTPUT_COLUMNS, build_dnts_cnts_rows, process_multiple_aws_pdfs
from extractors.google_dnts import extract_invoice_info, extract_table_from_text, make_dnts_header_row, DNTS_HEADER_COLS, DNTS_ITEM_COLS
from extractors.insurance import process_insurance_excel
from extractors.insurance2  import process_grouped_customer_files
#from sales.mibb_quotation import create_mibb_excel, extract_mibb_header_from_pdf, extract_mibb_table_from_pdf
from utils.helpers import format_amount, format_invoice_date, format_month_year
from dotenv import load_dotenv
from ibm import extract_ibm_data_from_pdf, create_styled_excel, create_styled_excel_template2, correct_descriptions, extract_last_page_text
from ibm_template2 import extract_ibm_template2_from_pdf, get_extraction_debug
from sales.ibm_v2 import compare_mep_and_cost
from template_detector import detect_ibm_template
import logging
logging.basicConfig(level=logging.INFO)
load_dotenv()
from extractors.google_invoice import extract_table_from_text as extract_invoice_table, extract_invoice_info as extract_invoice_info_invoice, GOOGLE_INVOICE_COLS
from extractors.dell_invoice import (
    extract_invoice_info as extract_dell_invoice_info,
    extract_table_from_text as extract_dell_table,
    DELL_INVOICE_COLS,
    PRE_ALERT_HEADERS,
    build_pre_alert_rows,
    read_master_mapping,
)
from oauth2client.service_account import ServiceAccountCredentials
from extractors.cloud_invoice import create_srcl_file, create_summary_sheet, build_cloud_invoice_df, map_invoice_numbers
from claims_automation import (
    build_output_rows_from_source1,
    write_output_excel,
    read_source1_rows,
    read_master1_map,
    read_source2_rows,
    build_debit_rows_from_source2,
    read_master2_entries,
    derive_defaults_from_source1,
)
import plotly.express as px

SHEET_JSON = "tool-mindware-0d87ca5562ad.json"  # Path to your downloaded JSON
SHEET_NAME = "mindware tool"

scope = ["https://www.googleapis.com/auth/spreadsheets", "https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]

# Read credentials from Streamlit secrets
service_account_info = st.secrets["gcp_service_account"]
creds = Credentials.from_service_account_info(service_account_info, scopes=scope)

# ✅ Authorize and open sheet
gc = gspread.authorize(creds)

tool_sheet = gc.open(SHEET_NAME).worksheet("Sheet1")     # Main usage sheet
env = st.secrets.get("env", "live")  # Default to live if not set

def update_usage(tool_name, team):
    month = datetime.today().strftime("%b-%Y")
    all_rows = tool_sheet.get_all_records()
    found = False

    for i, row in enumerate(all_rows, start=2):
        if row.get("Tool") == tool_name and row.get("Month") == month and row.get("Team") == team:
            current_count = row.get("Usage Count", 0) or 0
            tool_sheet.update_cell(i, 3, current_count + 1)
            found = True
            break

    if not found:
        tool_sheet.append_row([tool_name, month, 1, team])




st.set_page_config(
    page_title="Mindware Tool",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# CSS: hide top-right icons but keep sidebar visible
st.markdown("""
    <style>
    /* Hide Share, GitHub, Settings icons on top-right */
    [data-testid="stToolbar"] {
        display: none !important;
    }
    /* Ensure sidebar is always visible */
    [data-testid="stSidebar"] {
        visibility: visible !important;
        min-width: 280px !important;
    }
    /* Optional: adjust sidebar content font */
    [data-testid="stSidebar"] * {
        font-family: 'Google Sans', sans-serif !important;
    }
    /* General styling */
    html, body, [class*="css"] {
        font-family: 'Google Sans', sans-serif !important;
        background: #f6f8fa;
        color: #202124;
    }
    h1, h2, h3 {
        color: #1a73e8;
        font-weight: 700;
        letter-spacing: -1px;
        margin-bottom: 0.5em;
    }
    .stButton > button, .stDownloadButton > button {
        background: linear-gradient(90deg, #1a73e8, #188038);
        color: white;
        border-radius: 8px;
        font-weight: 600;
        font-size: 16px;
        border: none;
        padding: 12px 28px;
        margin-top: 10px;
        margin-bottom: 10px;
        box-shadow: 0 2px 8px rgba(26, 115, 232, 0.08);
        transition: background 0.2s;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: linear-gradient(90deg, #188038, #1a73e8);
        color: #fff;
    }
    .stDataFrame {
        background: #fff;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(26, 115, 232, 0.06);
        margin-bottom: 2em;
    }
    </style>
""", unsafe_allow_html=True)
# ----------- Constants -----------
DEFAULTS = {
    "supp_code": "SDIG005",
    "curr_code": "USD",
    "form_code": 0,
    "doc_src_locn": "UJ000",
    "location_code": "UJ200"
}
#CORRECT_USERNAME = "admin"
#CORRECT_PASSWORD = "admin"
CORRECT_USERNAME = os.getenv("NAME")
CORRECT_PASSWORD = os.getenv("PASSWORD")

if "login_state" not in st.session_state:
    st.session_state.login_state = "login" 
from pathlib import Path
import streamlit as st

def _img_to_base64(path: str) -> str:
    p = Path(path)
    if not p.exists():
        return ""
    return base64.b64encode(p.read_bytes()).decode("utf-8")
def show_login():  # <-- right-side login
    backg = "im.png"  # your AI-face image file name (put it in same folder as app.py)
    bg_b64 = _img_to_base64(backg)

    if not bg_b64:
        st.warning(f"Background image not found: {backg}. Put it next to your app.py")
        # continue anyway without background

    st.markdown(f"""
    <style>
    /* Make Streamlit background transparent (login only) */
    [data-testid="stAppViewContainer"],
    [data-testid="stApp"],
    [data-testid="stHeader"] {{
        background: transparent !important;
    }}

    /* Fullscreen background image */
    .mw-login-bg {{
        position: fixed;
        inset: 0;
        z-index: -10;
        background-image: url("data:image/png;base64,{bg_b64}");
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
    }}

    /* Dark overlay to improve contrast */
    .mw-login-bg::after {{
        content: "";
        position: absolute;
        inset: 0;
        background: linear-gradient(90deg,
            rgba(0,0,0,0.05) 0%,
            rgba(0,0,0,0.22) 45%,
            rgba(0,0,0,0.62) 100%
        );
    }}

    /* Right-side wrapper */
    .mw-right-wrap {{
        margin-top: 10vh;
        display: flex;
        justify-content: flex-end;
    }}

    /* Glass login card - subtle Ramadan accent */
    .mw-card {{
        width: min(430px, 92%);
        background: rgba(255,255,255,0.10);
        border: 1px solid rgba(245,215,142,0.35);
        border-radius: 18px;
        padding: 28px 28px 18px 28px;
        backdrop-filter: blur(18px) saturate(170%);
        -webkit-backdrop-filter: blur(18px) saturate(170%);
        box-shadow: 0 22px 55px rgba(0,0,0,0.35);
    }}

    .mw-title {{
        font-size: 40px;
        font-weight: 800;
        letter-spacing: -0.7px;
        text-align: center;
        margin: 0 0 4px 0;
        color: #EAF2FF;
        text-shadow: 0 6px 22px rgba(0,0,0,0.45);
    }}

    .mw-subtitle {{
        text-align:center;
        color: rgba(255,255,255,0.78);
        margin: 0 0 18px 0;
        font-weight: 500;
        font-size: 13.5px;
    }}

    /* Ramadan greeting */
    .mw-ramadan {{
        text-align: center;
        color: #f5d78e;
        font-size: 15px;
        font-weight: 600;
        margin: 0 0 8px 0;
        letter-spacing: 0.5px;
        text-shadow: 0 2px 8px rgba(0,0,0,0.3);
    }}

    /* Labels */
    label, .stTextInput label {{
        color: rgba(255,255,255,0.88) !important;
        font-weight: 600 !important;
    }}

    /* Inputs */
    div[data-testid="stTextInput"] input {{
        background: rgba(255,255,255,0.14) !important;
        border-radius: 10px !important;
        border: 1px solid rgba(255,255,255,0.28) !important;
        color: #ffffff !important;
        padding: 12px 14px !important;
    }}
    div[data-testid="stTextInput"] input::placeholder {{
        color: rgba(255,255,255,0.65) !important;
    }}

    /* Fix browser autofill */
    input:-webkit-autofill,
    input:-webkit-autofill:hover,
    input:-webkit-autofill:focus {{
        -webkit-text-fill-color: #ffffff !important;
        transition: background-color 9999s ease-in-out 0s;
        box-shadow: 0 0 0px 1000px rgba(255,255,255,0.12) inset !important;
        border: 1px solid rgba(255,255,255,0.28) !important;
    }}

    /* Button */
    .stButton > button {{
        border-radius: 12px !important;
        padding: 12px 18px !important;
        font-weight: 800 !important;
        background: linear-gradient(90deg, #1a73e8, #2b7de9) !important;
        border: none !important;
        color: white !important;
        box-shadow: 0 10px 24px rgba(26,115,232,0.25);
    }}
    .stButton > button:hover {{
        background: linear-gradient(90deg, #2b7de9, #1a73e8) !important;
    }}

    .mw-footer {{
        text-align: center;
        margin-top: 12px;
        color: rgba(255,255,255,0.60);
        font-size: 12px;
    }}

    /* Hide top toolbar */
    [data-testid="stToolbar"] {{
        display: none !important;
    }}
    </style>

    <div class="mw-login-bg"></div>
    """, unsafe_allow_html=True)

    # Layout: left space for image, right for card
    left, right = st.columns([2.2, 1])
    with right:
        
        st.markdown('<div class="mw-title">Mindbot</div>', unsafe_allow_html=True)
        st.markdown('<div class="mw-ramadan">🌙 Ramadan Kareem • Blessed Ramadan</div>', unsafe_allow_html=True)
        st.markdown('<div class="mw-subtitle">Powered Productivity Tools</div>', unsafe_allow_html=True)

        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("👤 Username", key="login_user", placeholder="Enter your username…")
            password = st.text_input("🔐 Password", type="password", key="login_pass", placeholder="Enter your password…")
            submitted = st.form_submit_button("Login", type="primary")

        if submitted:
            if username == CORRECT_USERNAME and password == CORRECT_PASSWORD:
                st.session_state.login_state = "success"
            else:
                st.session_state.login_state = "fail"
            st.rerun()

        st.markdown('<div class="mw-footer">Made with ❤️ by Mindware • © 2025</div>', unsafe_allow_html=True)
        st.markdown("</div></div>", unsafe_allow_html=True)

# def show_login():
#     # Add Christmas styling and animations
#     st.markdown("""
#     <style>
#     @keyframes snowfall {
#         0% { transform: translateY(-10px) rotate(0deg); opacity: 1; }
#         100% { transform: translateY(100vh) rotate(360deg); opacity: 0; }
#     }
#     @keyframes glow {
#         0%, 100% { text-shadow: 0 0 10px #1a73e8, 0 0 20px #1a73e8; }
#         50% { text-shadow: 0 0 20px #1a73e8, 0 0 30px #1a73e8, 0 0 40px #1a73e8; }
#     }
#     @keyframes float {
#         0%, 100% { transform: translateY(0px); }
#         50% { transform: translateY(-8px); }
#     }
#     .login-snow {
#         position: fixed;
#         top: 0;
#         left: 0;
#         width: 100%;
#         height: 100%;
#         pointer-events: none;
#         z-index: -1;
#     }
#     .snowflake {
#         position: absolute;
#         color: rgba(26, 115, 232, 0.6);
#         user-select: none;
#         animation: snowfall linear infinite;
#     }
#     .login-container {
#         background: linear-gradient(135deg, #f8f9fa 0%, #e3f2fd 100%);
#         border-radius: 20px;
#         padding: 2rem;
#         box-shadow: 0 15px 35px rgba(26, 115, 232, 0.1);
#         border: 2px solid #e3f2fd;
#         margin: 2rem 0;
#         position: relative;
#         overflow: hidden;
#     }
#     .login-title {
#         animation: glow 3s ease-in-out infinite;
#         color: #1a73e8;
#         text-align: center;
#         margin-bottom: 1rem;
#         position: relative;
#     }
#     .christmas-icon {
#         animation: float 2s ease-in-out infinite;
#         display: inline-block;
#         font-size: 1.5rem;
#         margin: 0 0.5rem;
#     }
#     </style>
    
#     <!-- Animated snowflakes -->
#     <div class="login-snow">
#         <div class="snowflake" style="left: 10%; animation-duration: 3s; animation-delay: 0s;">❄️</div>
#         <div class="snowflake" style="left: 20%; animation-duration: 4s; animation-delay: 1s;">🎄</div>
#         <div class="snowflake" style="left: 30%; animation-duration: 3.5s; animation-delay: 0.5s;">❄️</div>
#         <div class="snowflake" style="left: 40%; animation-duration: 5s; animation-delay: 2s;">⭐</div>
#         <div class="snowflake" style="left: 50%; animation-duration: 3.2s; animation-delay: 1.5s;">❄️</div>
#         <div class="snowflake" style="left: 60%; animation-duration: 4.5s; animation-delay: 0.8s;">🎁</div>
#         <div class="snowflake" style="left: 70%; animation-duration: 3.8s; animation-delay: 2.2s;">❄️</div>
#         <div class="snowflake" style="left: 80%; animation-duration: 4.2s; animation-delay: 1.2s;">🌟</div>
#         <div class="snowflake" style="left: 90%; animation-duration: 3.6s; animation-delay: 0.3s;">❄️</div>
#     </div>
#     """, unsafe_allow_html=True)
    
#     for _ in range(10):
#         st.write("")
   
#     col1, col2, col3 = st.columns([1,2,1])
#     with col2:
        
#         # Animated title with Christmas emojis
#         st.markdown("""
#         <h1 class="login-title">
#             <span class="christmas-icon" style="animation-delay: 0s;"></span>
#          Mindbot 
#             <span class="christmas-icon" style="animation-delay: 1s;"></span>
#         </h1>
#         """, unsafe_allow_html=True)
        
#         # Input fields with Christmas emojis
#         username = st.text_input("👤 Username", key="login_user", placeholder="Enter your username...")
#         password = st.text_input("🔐 Password", type="password", key="login_pass", placeholder="Enter your password...")
        
#         # Enhanced login button
#         if st.button(" **Login** ", key="login_btn", use_container_width=True, type="primary"):
#             if username == CORRECT_USERNAME and password == CORRECT_PASSWORD:
#                 st.session_state.login_state = "success"
              
#                 st.snow()
#             else:
#                 st.session_state.login_state = "fail"
        
#         # Christmas footer
#         st.markdown("""
#         <div style="
#             text-align: center;
#             margin-top: 2rem;
#             padding: 1rem;
#             background: linear-gradient(90deg, #e3f2fd, #f3e5f5);
#             border-radius: 10px;
#             border: 1px solid #e1f5fe;
#         ">
#             <p style="margin: 0; color: #1a73e8; font-weight: 500;">
#                 Made with ❤️ by Mindware✨<br>
#             </p>
#         </div>
#         """, unsafe_allow_html=True)
        
#         st.markdown('</div>', unsafe_allow_html=True)  # Close login container

def show_fail():
    st.error("Oops! Wrong credentials... Nice try, but no entry! 😜")
    # Display image from local file
    img_path = Path("img.png")  # or just Path("image.png") if in the root
    if img_path.exists():
        st.image(str(img_path), caption="Access Denied", use_column_width=True)
    else:
        st.warning("image.png not found. Please check the path.")

    if st.button("Back to Login", key="back_login"):
        st.session_state.login_state = "login"

if st.session_state.login_state == "login":
    show_login()
    st.stop()
elif st.session_state.login_state == "fail":
    show_fail()
    st.stop()

# Initialize session state for welcome flow
if "show_team_selection" not in st.session_state:
    st.session_state.show_team_selection = False

# 🎉 Welcome Page (only show if team selection not started)
if not st.session_state.show_team_selection:
   # st.balloons()  # Immediate celebration!

    # Cool welcome message with animation-like styling
    st.markdown("""
    <div style="
        text-align: center;
        padding: 2rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 20px;
        margin: 2rem 0;
        box-shadow: 0 10px 30px rgba(0,0,0,0.2);
    ">
        <h1 style="
            color: white;
            font-size: 3rem;
            margin: 0;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
            animation: pulse 2s infinite;
        ">
             Welcome to Mindbot! 
        </h1>
        <p style="
            color: #f0f0f0;
            font-size: 1.3rem;
            margin-top: 1rem;
        ">
            Ready to supercharge your productivity? Let's get started! ⚡
        </p>
    </div>

    <style>
    @keyframes pulse {
        0% { transform: scale(1); }
        50% { transform: scale(1.05); }
        100% { transform: scale(1); }
    }
    </style>
    """, unsafe_allow_html=True)

    # Fun statistics dashboard
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown("""
        <div style="text-align: center; padding: 1rem; background: linear-gradient(45deg, #FF6B6B, #FF8E53); border-radius: 15px; color: white; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
            <h2 style="margin: 0; font-size: 2.5rem;">🛠️</h2>
            <h3 style="margin: 0;">12</h3>
            <p style="margin: 0;">Tools Available</p>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div style="text-align: center; padding: 1rem; background: linear-gradient(45deg, #4ECDC4, #44A08D); border-radius: 15px; color: white; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
            <h2 style="margin: 0; font-size: 2.5rem;">⚡</h2>
            <h3 style="margin: 0;">99.9%</h3>
            <p style="margin: 0;">Accuracy Rate</p>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown("""
        <div style="text-align: center; padding: 1rem; background: linear-gradient(45deg, #A8EDEA, #00C9FF); border-radius: 15px; color: white; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
            <h2 style="margin: 0; font-size: 2.5rem;">🚀</h2>
            <h3 style="margin: 0;">10x</h3>
            <p style="margin: 0;">Faster Processing</p>
        </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown("""
        <div style="text-align: center; padding: 1rem; background: linear-gradient(45deg, #F093FB, #F5576C); border-radius: 15px; color: white; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
            <h2 style="margin: 0; font-size: 2.5rem;">⏰</h2>
            <h3 style="margin: 0;">24/7</h3>
            <p style="margin: 0;">Always Ready</p>
        </div>
        """, unsafe_allow_html=True)

        # REPLACE the entire motivational messages section with this:
    
        # FIRST - Close the col4 block properly
        # (Remove any 'with col4:' or column indentation before this section)
    
    # Close the statistics section completely and start fresh
    st.markdown("---")  # Add a separator line
    
    # Motivational messages section - FULL WIDTH, NOT in any column
    st.markdown("""
    <div style="
        text-align: center;
        margin: 2rem 0;
        width: 100%;
    ">
        <h2 style="
            color: #667eea;
            font-size: 1.8rem;
            margin-bottom: 1.5rem;
            font-weight: 700;
        ">
          💡 Why Our Tools Will Make You the Office Hero
          🎭 Warning: These Tools May Cause Extreme Productivity
        </h2>
    </div>
    """, unsafe_allow_html=True)
    
    # Create fresh columns for messages - FULL WIDTH
    msg_row1_col1, msg_row1_col2 = st.columns(2)
    
    with msg_row1_col1:
        st.markdown("""
        <div style="
            padding: 1.2rem;
            background: linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%);
            border-radius: 12px;
            margin: 0.5rem;
            color: white;
            text-align: center;
            box-shadow: 0 4px 15px rgba(255,154,158,0.3);
            height: 120px;
            display: flex;
            align-items: center;
            justify-content: center;
        ">
            <div>
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">🧠</div>
                <p style="margin: 0; font-size: 0.95rem; font-weight: 600;">
                    Smart tools for smart people like you
                </p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with msg_row1_col2:
        st.markdown("""
        <div style="
            padding: 1.2rem;
            background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%);
            border-radius: 12px;
            margin: 0.5rem;
            color: white;
            text-align: center;
            box-shadow: 0 4px 15px rgba(168,237,234,0.3);
            height: 120px;
            display: flex;
            align-items: center;
            justify-content: center;
        ">
            <div>
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">💪</div>
                <p style="margin: 0; font-size: 0.95rem; font-weight: 600;">
                    Save hours of manual work!
                </p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    msg_row2_col1, msg_row2_col2 = st.columns(2)
    
    with msg_row2_col1:
        st.markdown("""
        <div style="
            padding: 1.2rem;
            background: linear-gradient(135deg, #d299c2 0%, #fef9d7 100%);
            border-radius: 12px;
            margin: 0.5rem;
            color: #5d4e75;
            text-align: center;
            box-shadow: 0 4px 15px rgba(210,153,194,0.3);
            height: 120px;
            display: flex;
            align-items: center;
            justify-content: center;
        ">
            <div>
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">🎯</div>
                <p style="margin: 0; font-size: 0.95rem; font-weight: 600;">
                    Smart tools for smart people like you
                </p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with msg_row2_col2:
        st.markdown("""
        <div style="
            padding: 1.2rem;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 12px;
            margin: 0.5rem;
            color: white;
            text-align: center;
            box-shadow: 0 4px 15px rgba(102,126,234,0.3);
            height: 120px;
            display: flex;
            align-items: center;
            justify-content: center;
        ">
            <div>
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">🌟</div>
                <p style="margin: 0; font-size: 0.95rem; font-weight: 600;">
                    Every file = efficiency step!
                </p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    msg_row3_col1, msg_row3_col2 = st.columns(2)
    
    with msg_row3_col1:
        st.markdown("""
        <div style="
            padding: 1.2rem;
            background: linear-gradient(135deg, #ffecd2 0%, #fcb69f 100%);
            border-radius: 12px;
            margin: 0.5rem;
            color: #8b4513;
            text-align: center;
            box-shadow: 0 4px 15px rgba(255,236,210,0.3);
            height: 120px;
            display: flex;
            align-items: center;
            justify-content: center;
        ">
            <div>
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">🔥</div>
                <p style="margin: 0; font-size: 0.95rem; font-weight: 600;">
                    Ready to automate the boring stuff?
                </p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with msg_row3_col2:
        st.markdown("""
        <div style="
            padding: 1.2rem;
            background: linear-gradient(135deg, #89f7fe 0%, #66a6ff 100%);
            border-radius: 12px;
            margin: 0.5rem;
            color: white;
            text-align: center;
            box-shadow: 0 4px 15px rgba(137,247,254,0.3);
            height: 120px;
            display: flex;
            align-items: center;
            justify-content: center;
        ">
            <div>
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">✨</div>
                <p style="margin: 0; font-size: 0.95rem; font-weight: 600;">
                    Making the impossible possible, one click at a time!
                </p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    msg_row4_col1, msg_row4_col2 = st.columns(2)
    
    with msg_row4_col1:
        st.markdown("""
        <div style="
            padding: 1.2rem;
            background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
            border-radius: 12px;
            margin: 0.5rem;
            color: white;
            text-align: center;
            box-shadow: 0 4px 15px rgba(250,112,154,0.3);
            height: 120px;
            display: flex;
            align-items: center;
            justify-content: center;
        ">
            <div>
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">🎉</div>
                <p style="margin: 0; font-size: 0.95rem; font-weight: 600;">
                    Time to show those spreadsheets who's boss!
                </p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with msg_row4_col2:
        st.markdown("""
        <div style="
            padding: 1.2rem;
            background: linear-gradient(135deg, #ff9a9e 0%, #fad0c4 100%);
            border-radius: 12px;
            margin: 0.5rem;
            color: white;
            text-align: center;
            box-shadow: 0 4px 15px rgba(255,154,158,0.3);
            height: 120px;
            display: flex;
            align-items: center;
            justify-content: center;
        ">
            <div>
                <div style="font-size: 2rem; margin-bottom: 0.5rem;">🚀</div>
                <p style="margin: 0; font-size: 0.95rem; font-weight: 600;">
                    Blast off to productivity!
                </p>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # Time-based greeting
    import datetime
    current_hour = datetime.datetime.now().hour

    if current_hour < 12:
        greeting = "🌅 Good Morning! Ready to conquer the day?"
        emoji = "☕"
    elif current_hour < 17:
        greeting = "☀️ Good Afternoon! Let's get productive!"
        emoji = "💼"
    else:
        greeting = "🌙 Good Evening! Working late? We've got you covered!"
        emoji = "🌟"

    st.markdown(f"""
    <div style="
        text-align: center;
        padding: 1rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
    ">
        <h3 style="margin: 0;">{emoji} {greeting}</h3>
    </div>
    """, unsafe_allow_html=True)

    # Add a small delay for dramatic effect
    import time
    time.sleep(0.5)
   # st.snow()  # Another fun effect after the delay!

    # Big "Get Started" Button
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🚀 **GET STARTED!** 🚀", 
                     use_container_width=True, 
                     type="primary",
                     help="Click to proceed to team selection!"):
            st.session_state.show_team_selection = True
            #st.balloons()
          #  st.snow()
            st.rerun()

    # Stop here - don't show team selection yet
    st.stop()

# Replace the team selection section with this cleaner, smaller version:

# 🎯 Team Selection Section (only show after "Get Started" is clicked)
if st.session_state.show_team_selection:
    # Smaller, cleaner team selection header
    st.markdown("""
    <div style="
        text-align: center;
        padding: 1.5rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 15px;
        margin: 1.5rem 0;
        box-shadow: 0 8px 25px rgba(102,126,234,0.25);
        color: white;
    ">
        <h2 style="
            font-size: 1.8rem;
            margin: 0;
            text-shadow: 1px 1px 5px rgba(0,0,0,0.3);
            font-weight: 600;
        ">
            🎯 Choose Your Team
        </h2>
        <p style="
            font-size: 1rem;
            margin: 0.5rem 0 0 0;
            opacity: 0.9;
        ">
            Select your department to access specialized tools
        </p>
    </div>
    """, unsafe_allow_html=True)

    # Clean team selection radio buttons - NO extra divs
    team = st.radio(
        "👥 **Select your team:**", 
        ["Finance", "Operations", "Credit", "Sales"], 
        horizontal=True,
        help="Choose your department to see relevant tools!"
    )

    # Show confirmation with team-specific styling
    if team:
        team_colors = {
            "Finance": "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
            "Operations": "linear-gradient(135deg, #ffecd2 0%, #fcb69f 100%)", 
            "Credit": "linear-gradient(135deg, #667eea 0%, #fed6e3 100%)",
            "Sales": "linear-gradient(135deg, #d299c2 0%, #fef9d7 100%)"
        }

        team_emojis = {
            "Finance": "💰",
            "Operations": "⚙️", 
            "Credit": "📊",
            "Sales": "📈"
        }

        st.markdown(f"""
        <div style="
            text-align: center;
            padding: 1.2rem;
            background: {team_colors.get(team, team_colors['Finance'])};
            border-radius: 12px;
            margin: 1rem 0;
            color: white;
            box-shadow: 0 6px 20px rgba(0,0,0,0.1);
        ">
            <h3 style="
                margin: 0; 
                text-shadow: 1px 1px 3px rgba(0,0,0,0.3);
                font-size: 1.3rem;
            ">
                🎉 Perfect! Welcome to team {team}! {team_emojis.get(team, '🎯')}
            </h3>
            <p style="margin: 0.5rem 0 0 0; opacity: 0.9; font-size: 0.9rem;">
                Your specialized tools are ready below ✨
            </p>
        </div>
        """, unsafe_allow_html=True)

# Continue with your existing code...

# Continue with the rest of your existing code for team-specific tools...

def validate_customer_code(df, file_name="File"):
    """
    Validates that Customer Code column has no empty or missing values.
    Shows Streamlit error and stops processing if invalid.
    """
    if "CustomerCode" not in df.columns:
        st.error(f"❌ {file_name}: Missing 'Customer Code' column.")
        st.stop()

    # Check for missing or empty values
    if df["CustomerCode"].isna().any() or (df["CustomerCode"].astype(str).str.strip() == "").any():
        st.error(f"❌ {file_name}: Kindly check the 'Customer Code' column — it cannot be empty.")
        st.stop()
     


def extractor_workflow(
    extractor_name,
    extractor_info,
    file_uploader_label,
    extract_invoice_info_func,
    extract_table_func,
    table_columns,
    file_name_template,
    show_header_df_func=None,
    header_columns=None,
    item_row_builder=None,
    item_columns=None
):
    st.title(f"PDF TO EXCEL ({extractor_name})")
    st.write(extractor_info)
    uploaded_file = st.file_uploader(file_uploader_label, type=["pdf"], accept_multiple_files=False, key=f"uploader_{extractor_name}")
    if uploaded_file:
        invoice_num, invoice_date = extract_invoice_info_func(uploaded_file)
        if invoice_num and invoice_date:
            file_date = format_month_year(invoice_date)
            file_name = file_name_template.format(invoice_num=invoice_num, file_date=file_date)
        else:
            file_name = file_name_template.format(invoice_num='unknown', file_date='unknown')
        rows = extract_table_func(uploaded_file)
        if rows:
            df = pd.DataFrame(rows, columns=table_columns)
            if show_header_df_func and header_columns and item_row_builder and item_columns:
                today_str = datetime.today().strftime("%d/%m/%Y")
                remarks = f"GOOGLE INV-{invoice_num}" if invoice_num else "GOOGLE INV-UNKNOWN"
                header_df = pd.DataFrame([
                    show_header_df_func(invoice_num, invoice_date, today_str, remarks)
                ], columns=header_columns)
                dnts_item_data = [item_row_builder(idx, *row, invoice_num) for idx, row in enumerate(rows, 1)]
                dnts_item_df = pd.DataFrame(dnts_item_data, columns=item_columns)
               
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    header_df.to_excel(writer, sheet_name='DNTS_HEADER', index=False)
                    dnts_item_df.to_excel(writer, sheet_name='DNTS_ITEM', index=False)
                output.seek(0)
                
                st.download_button(
                    label=f"⬇️ Download DNTS Excel",
                    data=output.getvalue(),
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    on_click=lambda: (
                            update_usage("Google Automation", team)
                        ),
                    key=f"download_{extractor_name}"
                )
            else:
                
                towrite = io.BytesIO()
                df.to_excel(towrite, index=False, engine='openpyxl')
                towrite.seek(0)
                st.download_button(
                    label="Download as Excel",
                    data=towrite,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    on_click=lambda: (
                            update_usage("google Automation", team)
                        ),
                )
        else:
            st.warning("No table data found in the uploaded PDF.")
    else:
        st.info(f"Please upload a {extractor_name} PDF file to get started.")
# ----------- Tool Selector UI -----------
st.markdown("""
    <div style='text-align:center; margin-top:2rem; margin-bottom:1.5rem;'>
        <h2 style='color:#1a73e8; font-family:Google Sans, sans-serif; font-weight:700; letter-spacing:-1px;'>🛠️ Tool Selection</h2>
        <p style='font-size:1.2rem; color:#444;'>Choose the tool you want to use for your PDF extraction.</p>
    </div>
""", unsafe_allow_html=True)
if team == "Finance":
    TOOL_OPTIONS = [
        "-- Select a tool --",
        "🟦 Google DNTS Extractor",
        "🟩 Google Invoice Extractor",
        "📄 Claims Automation",
        "🟨 AWS Invoice Tool"
    ]
elif team == "Operations":
    TOOL_OPTIONS = [
        "-- Select a tool --",
        "💻 Dell Invoice Extractor",
        "🧾 Cloud Invoice Tool",
        "📦 Barcode PDF Generator grouped"
    ]
elif team == "Credit":
    TOOL_OPTIONS = [
        "AR to EDD file",
        "Coface CSV Uploader"
    ]
elif team == "Sales":
    TOOL_OPTIONS = [
        "IBM Quotation",
        "MIBB Quotations",
    ]
else:
    TOOL_OPTIONS = ["-- Select a tool --"]
tool = st.selectbox(
    "Select a tool:",
    TOOL_OPTIONS,
    key="tool_selector"
)
if tool == "-- Select a tool --":
    st.info("Please select a tool above to get started.")
elif tool == "🟦 Google DNTS Extractor":
    def dnts_item_row(idx, domain, customer_id, amount, invoice_num):
        formatted_amount = format_amount(amount)
        item_name = (
            f"GOOGLE INV-{invoice_num} / DOMAIN NAME : {domain} / CUSTOMER ID : {customer_id} / AMOUNT - USD - {formatted_amount}"
        ).upper()
        return [
            idx, 1, "NS", item_name, "NA", "NA", "NOS", 1, 0, formatted_amount, 14401, "SDIG005", "PUHO", "GEN", "ZZ-COMM"
        ]
    extractor_workflow(
        extractor_name="Google DNTS Extractor",
        extractor_info="Upload one PDF containing a **'Summary of costs by domain'** table. The app will extract the table and let you download it as Excel.",
        file_uploader_label="Choose your Google DNTS Invoice PDF",
        extract_invoice_info_func=extract_invoice_info,
        extract_table_func=extract_table_from_text,
        table_columns=DNTS_ITEM_COLS[:3],
        file_name_template="{invoice_num}-{file_date}.xlsx",
        show_header_df_func=make_dnts_header_row,
        header_columns=DNTS_HEADER_COLS,
        item_row_builder=dnts_item_row,
        item_columns=DNTS_ITEM_COLS
    )
elif tool == "🟩 Google Invoice Extractor":
    extractor_workflow(
        extractor_name="Google Invoice Extractor",
        extractor_info="Upload a Google Invoice PDF. The app will extract the relevant data and let you download it as Excel.",
        file_uploader_label="Choose your Google Invoice PDF",
        extract_invoice_info_func=extract_invoice_info_invoice,
        extract_table_func=extract_invoice_table,
        table_columns=GOOGLE_INVOICE_COLS,
        file_name_template="{invoice_num}-{file_date}.xlsx"
    )
elif tool == "📄 Claims Automation":
    st.title("Claims Automation")
    
    st.header("📁 Upload Files")
    source1_file = st.file_uploader("JV Orion from SAP (.xlsx)", type=["xlsx"], accept_multiple_files=False, key="claims_source1")
    master1_file = st.file_uploader("User information (.xlsx)", type=["xlsx"], accept_multiple_files=False, key="claims_master1")
    source2_file = st.file_uploader("Employee benefits (.xlsx)", type=["xlsx"], accept_multiple_files=False, key="claims_source2")
    master2_file = st.file_uploader("Main acc file (.xlsx)", type=["xlsx"], accept_multiple_files=False, key="claims_master2")
    
    st.markdown("---")
    run_clicked = st.button("🚀 Generate Output", key="claims_run", use_container_width=True)
    if run_clicked:
        if not source1_file:
            st.error("Please upload Source File 1.")
            st.stop()
        try:
            src_rows = read_source1_rows(source1_file)
            master_map = read_master1_map(master1_file) if master1_file else None
            src2_rows = read_source2_rows(source2_file) if source2_file else None
            master2_entries = read_master2_entries(master2_file) if master2_file else None
            credit_rows = build_output_rows_from_source1(
                src_rows,
                master1_map=master_map,
                source2_rows=src2_rows,
                user_id_col="Sub Acct",
            )
            src1_defaults = derive_defaults_from_source1(src_rows)
            doc_ref = ""
            for r in src_rows:
                val = r.get("Doc Ref.", "")
                if val and str(val).strip():
                    doc_ref = str(val).strip()
                    break
            debit_rows = build_debit_rows_from_source2(
                src2_rows,
                master2_entries=master2_entries,
                master1_map=master_map,
                default_div=src1_defaults.get("Div", ""),
                default_dept=src1_defaults.get("Dept", ""),
                default_anly1=src1_defaults.get("Anly1", ""),
                default_anly2=src1_defaults.get("Anly2", ""),
                default_currency=src_rows[0].get("Currency", "") if src_rows else "",
                doc_ref=doc_ref,
            )
            out_rows = credit_rows + debit_rows
            output_buffer = io.BytesIO()
            write_output_excel(output_buffer, out_rows)
            output_buffer.seek(0)
            st.download_button(
                label="Download claims_output.xlsx",
                data=output_buffer,
                file_name="claims_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                on_click=lambda: (
                            update_usage("Claims Automation", team)
                        ),
                key="claims_download"
            )
        except Exception as e:
            st.error(f"Error: {e}")
elif tool == "🧾 Cloud Invoice Tool":
    st.title("Cloud Invoice Tool")
    st.markdown(
        """
        <div style="
            padding: 18px 20px;
            background: linear-gradient(90deg, #fff3cd, #ffeeba);
            border: 2px solid #ffcc00;
            border-radius: 10px;
            font-weight: 700;
            color: #7a5a00;
            font-size: 16px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            margin-bottom: 12px;
        ">
            <span style="font-size: 18px;">🚨 IMPORTANT:</span>
            <span style="margin-left: 8px;">
            Please make sure to <b>open the CB file</b>, click <b>Convert</b>, then <b>upload the converted file here</b> and use this tool; otherwise <b>you will have missing invoices</b>.
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div style="
            padding: 14px 16px;
            background: #fff;
            border: 1px dashed #ffcc00;
            border-radius: 10px;
            color: #4a4a4a;
            font-size: 15px;
            margin-bottom: 8px;
        ">
        <b>Follow these steps before uploading:</b>
        <ol style="margin-top: 6px;">
            <li>Open the <b>CB file</b>.</li>
            <li>Click <b>Convert</b> to generate the latest output.</li>
            <li>Upload the <b>converted file</b> here.</li>
        </ol>
        </div>
        """,
        unsafe_allow_html=True,
    )
    confirmed = st.checkbox("I confirm I opened the CB file and clicked Convert ✅", key="cloud_cb_confirm")
    if not confirmed:
        st.warning("Please confirm the IMPORTANT notice steps above to proceed.")
        st.stop()
    uploaded_file = st.file_uploader("Upload your CSV file", type=["csv", "xlsx"], key="cloud_invoice_upload")
    if uploaded_file:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, engine="openpyxl")
        else:
            st.error("Unsupported file format. Please upload a CSV or Excel file.")
            st.stop()
            # --- Validate Customer Code before proceeding ---
        validate_customer_code(df, "Cloud Invoice File")

        # Process invoice data
        final_df = build_cloud_invoice_df(df)
        final_df = map_invoice_numbers(final_df)
        sorted_df = final_df.sort_values(by=final_df.columns.tolist()).reset_index(drop=True)
        def highlight_row(row):
            end_user = str(row.get("End User", "")).strip()
            return not (" ; " in end_user)
        sorted_df["_highlight_end_user"] = sorted_df.apply(highlight_row, axis=1)
        # Create unique version rows based on Combined (D)
        unique_rows = sorted_df[["Invoice No.","LPO Number", "End User"]].copy()
        unique_rows["Combined (D)"] = (
            unique_rows["Invoice No."].astype(str) +
            unique_rows["LPO Number"].astype(str) +
            unique_rows["End User"].astype(str)
        )
        unique_rows = unique_rows.drop_duplicates(subset=["Combined (D)"]).reset_index(drop=True)
        # Versioning logic
        unique_rows["Version1 (E)"] = (unique_rows["Invoice No."].ne(unique_rows["Invoice No."].shift()).astype(int))
        v2 = []
        for i, v1 in enumerate(unique_rows["Version1 (E)"]):
            if v1 == 1:
                v2.append(1)
            else:
                prev_v2 = v2[-1]
                v2.append(prev_v2 + 1)
        unique_rows["Version2 (F)"] = v2
        unique_rows["Version3 (G)"] = unique_rows.apply(lambda row: f'-{row["Version2 (F)"]}', axis=1)
        unique_rows["Version4 (H)"] = unique_rows.apply(lambda row: f'{row["Invoice No."]}-{row["Version2 (F)"]}', axis=1)
        # --- MAP Version 4 back to main DataFrame ---
        version_map = dict(zip(unique_rows["Combined (D)"], unique_rows["Version4 (H)"]))
        sorted_df["Combined (D)"] = (
            sorted_df["Invoice No."].astype(str) +
            sorted_df["LPO Number"].astype(str) +
            sorted_df["End User"].astype(str)
        )
        sorted_df["Versioned Invoice No."] = sorted_df["Combined (D)"].map(version_map)
        cols = list(sorted_df.columns)
        cols.append(cols.pop(cols.index("Versioned Invoice No.")))
        sorted_df = sorted_df[cols]
        sorted_df = sorted_df.drop(columns=["Combined (D)"])
        # === ADD HIGHLIGHT FLAG HERE ===
        #sorted_df["_highlight_end_user"] = sorted_df["End User"].astype(str).str.strip() == ""
        # Display metrics
        pos_df = sorted_df[sorted_df["Gross Value"].astype(float) >= 0]
        neg_df = sorted_df[sorted_df["Gross Value"].astype(float) < 0]
        st.success(f"{len(pos_df)} positive, {len(neg_df)} negative, total: {len(sorted_df)}")
        c1, c2, c3 = st.columns(3)
        c1.metric("✅ Positive invoices", len(pos_df))
        c2.metric("❌ Negative invoices", len(neg_df))
        c3.metric("🧮 Total invoices", len(sorted_df))
       
        
        # Create Excel workbook with formulas
        red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
        df_to_write = pos_df.copy()
        
        # Remove highlight flag column before writing
        if "_highlight_end_user" in df_to_write.columns:
            df_to_write = df_to_write.drop(columns=["_highlight_end_user"])
        
        # Get index of 'End User' column (1-based for Excel)
        try:
            end_user_col_idx = df_to_write.columns.get_loc("End User") + 1
        except:
            end_user_col_idx = None
            
        try:
            item_code_col_idx = df_to_write.columns.get_loc("ITEM Code") + 1
        except:
            item_code_col_idx = None    
        
        # Create workbook and write rows
        wb = Workbook()
        ws_invoice = wb.active
        ws_invoice.title = "CLOUD INVOICE"
        
        for r_idx, row in enumerate(dataframe_to_rows(df_to_write, index=False, header=True), start=1):
            ws_invoice.append(row)
            
            # Skip header row
            if r_idx == 1:
              continue
        # Highlight End User
            if end_user_col_idx is not None:
                highlight = sorted_df.iloc[r_idx - 2].get("_highlight_end_user", False)
                if highlight:
                    col_letter = get_column_letter(end_user_col_idx)
                    ws_invoice[f"{col_letter}{r_idx}"].fill = red_fill
 
        # Highlight ITEM Code if empty
            if item_code_col_idx is not None:
                item_code_val = sorted_df.iloc[r_idx - 2].get("ITEM Code", "")
                if not item_code_val or str(item_code_val).strip().lower() in ["", "nan", "none"]:
                   col_letter = get_column_letter(item_code_col_idx)
                   ws_invoice[f"{col_letter}{r_idx}"].fill = red_fill
        
        # Create VERSIONS sheet with formulas
        ws_versions = wb.create_sheet(title="VERSIONS")
        headers = ["Invoice",  "LPO", "End User", "Combined (D)", "Version1 (E)", "Version2 (F)", "Version3 (G)", "Version4 (H)"]
        ws_versions.append(headers)
        for i, row in enumerate(unique_rows.itertuples(index=False, name=None), start=2):
            invoice, lpo, end_user, combined_d = row[:4]
            ws_versions.cell(row=i, column=1, value=invoice)
            ws_versions.cell(row=i, column=2, value=lpo)
            ws_versions.cell(row=i, column=3, value=end_user)
            ws_versions.cell(row=i, column=4, value=combined_d)
            ws_versions.cell(row=i, column=5, value=f'=IF(A{i}=A{i-1},"",1)')
            ws_versions.cell(row=i, column=6, value=f'=IFERROR(IF(E{i}="",E{i-1}+1,""),F{i-1}+1)')
            ws_versions.cell(row=i, column=7, value=f'="-"&E{i}&F{i}')
            ws_versions.cell(row=i, column=8, value=f'=A{i}&G{i}')
        
        # Save to buffer
        output_buffer = io.BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)
        st.download_button(
            label="⬇️ Download Cloud Invoice",
            data=output_buffer.getvalue(),
            file_name="cloud_invoice.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=lambda: (
                            update_usage("Cloud Automation", team)
                        ),

        )
        neg_buffer = io.BytesIO()
        wb_neg = Workbook()
        ws_neg = wb_neg.active
        ws_neg.title = "NEGATIVE INVOICES"
        for row in dataframe_to_rows(neg_df, index=False, header=True):
               ws_neg.append(row)
        wb_neg.save(neg_buffer)
        neg_buffer.seek(0)
        st.download_button(
    label="⬇️ Download Negative Invoices",
    data=neg_buffer.getvalue(),
    file_name="negative_invoices.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
        srcl_buffer = create_srcl_file(neg_df)  # only negative invoices



        st.download_button(
           label="⬇️ Download SRCL File",
           data=srcl_buffer.getvalue(),
           file_name="srcl_file.xlsx",
           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
         
)
     


elif tool == "📦 Barcode PDF Generator grouped":
    st.write("Upload a CSV file with PalletID and IMEIs to generate barcode PDF.")

    pdf_bytes, success = barcode_tooll()

    if success and pdf_bytes:
        st.success("✅ Barcode PDF is ready!")

        # Create ZIP buffer and write PDF into it
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.writestr("pallet_barcodes_fullpage.pdf", pdf_bytes)
        zip_buffer.seek(0)

        st.download_button(
            label="📥 Download Full-Page Barcode PDF (Zipped)",
            data=pdf_bytes,
            file_name="pallet_barcodes_fullpage.zip",
            mime="application/zip",
            on_click=lambda: (
                            update_usage("barcode Automation", team)
                        ),

        )

elif tool == "💻 Dell Invoice Extractor":
    st.title("Dell Invoice Extractor (Pre-Alert Upload)")
    st.write("Upload one or more Dell invoice PDFs. We'll generate a single Excel with sheet 'PRE ALERT UPLOAD'.")
    uploaded_files = st.file_uploader("Choose Dell invoice PDF(s)", type=["pdf"], accept_multiple_files=True, key="dell_upload")
    master_file = st.file_uploader("Master Excel (starts header at row 9)", type=["xlsx"], key="dell_master")
    if uploaded_files:
            from datetime import datetime, timedelta
            tomorrow_date = (datetime.today() + timedelta(days=1)).strftime("%d/%m/%Y")
            all_rows = []
            master_lookup = None
            supplier_counts = None
            orion_counts = None
            supplier_index = None
            orion_index = None
            po_price_index = None
            if master_file is not None:
                try:
                    master_lookup, supplier_counts, orion_counts, supplier_index, orion_index, po_price_index = read_master_mapping(master_file)
                except Exception as e:
                    st.warning(f"Could not read master file: {e}")
            diag: list[dict] = []
            import os
            log_path = os.path.abspath('pdf_extract_debug.log')
            import tempfile
            for f in uploaded_files:
                st.info(f"DEBUG: Processing file: {getattr(f, 'name', str(f))} (type: {type(f)})")
                # Save UploadedFile to a temp file for processing
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                    tmp.write(f.read())
                    tmp_path = tmp.name
                try:
                    rows = build_pre_alert_rows(
                        tmp_path,
                        tomorrow_date,
                        master_lookup=master_lookup,
                        supplier_counts=supplier_counts,
                        orion_counts=orion_counts,
                        supplier_index=supplier_index,
                        orion_index=orion_index,
                        po_price_index=po_price_index,
                        diagnostics=diag,
                    )
                    all_rows.extend(rows)
                except Exception as e:
                    st.warning(f"Failed to parse {getattr(f, 'name', 'file')}: {e}")
                finally:
                    try:
                        os.remove(tmp_path)
                    except Exception:
                        pass
            st.info(f"PDF extraction debug log saved at: {log_path}")
            if all_rows:
                df = pd.DataFrame(all_rows, columns=PRE_ALERT_HEADERS)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='PRE ALERT UPLOAD', index=False)
                    ws = writer.sheets['PRE ALERT UPLOAD']
                    from datetime import datetime as _dt
                    for r in range(2, len(df) + 2):
                        for c in (4, 11, 12):  # D, K, L
                            cell = ws.cell(row=r, column=c)
                            val = cell.value
                            if isinstance(val, str):
                                try:
                                    d = _dt.strptime(val, "%d/%m/%Y")
                                    cell.value = d
                                    cell.number_format = 'dd/mm/yyyy'
                                except Exception:
                                    pass
                    # ====== NEW: Highlight Unit Rate (Q) vs Orion Unit Price (U) in PRE ALERT UPLOAD ======
                    from openpyxl.styles import PatternFill
                    red_fill = PatternFill(fill_type='solid', start_color='FFFFC7CE', end_color='FFFFC7CE')  # light red
                    green_fill = PatternFill(fill_type='solid', start_color='FFC6EFCE', end_color='FFC6EFCE')  # light green
                    
                    def safe_float(val):
                        try:
                            return float(str(val).replace(',', '').strip())
                        except Exception:
                            return None
                    # Columns: Q = 17, U = 21
                    for row_idx in range(2, len(df) + 2):
                        unit_rate_cell = ws.cell(row=row_idx, column=17)
                        orion_price_cell = ws.cell(row=row_idx, column=21)
                        unit_rate_val = safe_float(unit_rate_cell.value)
                        orion_price_val = safe_float(orion_price_cell.value)
                        if unit_rate_val is not None and orion_price_val is not None:
                            if abs(unit_rate_val - orion_price_val) < 1e-6:
                                # Match: highlight green
                                unit_rate_cell.fill = green_fill
                                orion_price_cell.fill = green_fill
                            else:
                                # Mismatch: highlight red
                                unit_rate_cell.fill = red_fill
                                orion_price_cell.fill = red_fill
                    # ====== EXISTING color highlights on PRE ALERT UPLOAD based on diagnostics ======
                    if diag:
                        red = PatternFill(fill_type='solid', start_color='FFFF0000', end_color='FFFF0000')
                        yellow = PatternFill(fill_type='solid', start_color='FFFFFF00', end_color='FFFFFF00')
                        green = PatternFill(fill_type='solid', start_color='FF00FF00', end_color='FF00FF00')
                        for i, d in enumerate(diag, start=2):
                            h = str(d.get('highlight', 'none')).lower()
                            fill = None
                            if h == 'red':
                                fill = red
                            elif h == 'yellow':
                                fill = yellow
                            elif h == 'green':
                                fill = green
                            if fill is not None:
                                ws.cell(row=i, column=13).fill = fill
                                ws.cell(row=i, column=14).fill = fill
                    # Add an extremely user-friendly REVIEW sheet (emoji + clear actions)
                    if diag:
                        def _status_texts(h: str, status: str) -> tuple[str, str, str, int]:
                            h = (h or 'none').lower()
                            status = status or ''
                            if h == 'red':
                                if status in ('C_orion_price_single', 'C_po_price_single'):
                                    return (
                                        '❌ No match (price suggested)',
                                        'Use the suggested Orion code. Confirm Qty & Unit Price.',
                                        'We did not find a supplier match, but price uniquely matched one item.',
                                        0,
                                    )
                                return (
                                    '❌ No match',
                                    'Ask to add mapping in MASTER or correct Supplier Item code in PDF.',
                                    'We could not find this PO + Supplier Item in MASTER.',
                                    0,
                                )
                            if h == 'yellow':
                                if status == 'B_price_single':
                                    return (
                                        '✅ Price matched (needs quick check)',
                                        'Quick check the suggested Orion Code, Qty, and Unit Price.',
                                        'Multiple entries in MASTER; price selected exactly one.',
                                        1,
                                    )
                                return (
                                    '⚠️ Many matches',
                                    'Pick the correct Orion code or ask to refine MASTER.',
                                    'Multiple MASTER entries matched; price did not decide one.',
                                    1,
                                )
                            return (
                                '✅ All good',
                                'No action. Just review and proceed.',
                                'Single clear match found in MASTER.',
                                2,
                            )
                        simple_rows = []
                        for idx0, d in enumerate(diag):
                            h = str(d.get('highlight', 'none')).lower()
                            status = str(d.get('status', '')).strip()
                            label, action, why, sort_key = _status_texts(h, status)
                            try:
                                row_df = df.iloc[idx0]
                            except Exception:
                                row_df = None
                            qty_val = row_df['Qty'] if row_df is not None and 'Qty' in row_df else ''
                            pdf_price = row_df['Unit Rate'] if row_df is not None and 'Unit Rate' in row_df else d.get('pdf_unit_price', '')
                            use_price = d.get('out_orion_unit_price', '') or pdf_price
                            orion_price_u = d.get('out_orion_unit_price', '')
                            use_code = d.get('out_orion_item_code', '') or d.get('mapped_item_code', '')
                            pre_alert_row_num = idx0 + 2  # header row is 1 in PRE ALERT
                            simple_rows.append({
                                'Row in PRE ALERT': pre_alert_row_num,
                                'Status': label,
                                'What to do now': action,
                                'Why this happened': why,
                                'PO': d.get('po', ''),
                                'Supplier Item': d.get('supplier_item_code', ''),
                                'Copy: Orion Code': use_code,
                                'Copy: Qty': qty_val,
                                'Copy: Unit Price': use_price,
                                'Orion Unit Price (col U)': orion_price_u,
                                '_sort': sort_key,
                                '_highlight': h,
                            })
                        simple_rows.sort(key=lambda x: x.get('_sort', 1))
                        for r in simple_rows:
                            r.pop('_sort', None)
                        review_columns = [
                            'Row in PRE ALERT',
                            'Status',
                            'What to do now',
                            'Why this happened',
                            'Copy: Orion Code',
                            'Copy: Qty',
                            'Copy: Unit Price',
                            'Orion Unit Price (col U)',
                            'PO',
                            'Supplier Item',
                        ]
                        review_df = pd.DataFrame(simple_rows, columns=review_columns + ['_highlight'])
                        review_df.to_excel(writer, sheet_name='REVIEW', index=False)
                        ws_review = writer.sheets['REVIEW']
                        ws_review.freeze_panes = 'A2'
                        widths = {
                            'A': 16, 'B': 30, 'C': 38, 'D': 42,
                            'E': 20, 'F': 12, 'G': 18, 'H': 12, 'I': 20,
                        }
                        for col, w in widths.items():
                            try:
                                ws_review.column_dimensions[col].width = w
                            except Exception:
                                pass
                        try:
                            from openpyxl.styles import Alignment
                            wrap_align = Alignment(wrap_text=True, vertical='top')
                            for r in ws_review.iter_rows(min_row=1, max_row=ws_review.max_row, min_col=1, max_col=ws_review.max_column):
                                for cell in r:
                                    if cell.column in (2, 3, 4):  # B:Status, C:What to do, D:Why
                                        cell.alignment = wrap_align
                        except Exception:
                            pass
                        red_fill_row = PatternFill(fill_type='solid', start_color='FFFFE5E5', end_color='FFFFE5E5')
                        yellow_fill_row = PatternFill(fill_type='solid', start_color='FFFFFBE6', end_color='FFFFFBE6')
                        green_fill_row = PatternFill(fill_type='solid', start_color='FFE9FBE9', end_color='FFE9FBE9')
                        for idx1, row_data in enumerate(simple_rows, start=2):
                            hval = str(row_data.get('_highlight', '')).lower()
                            fill = None
                            if hval == 'red':
                                fill = red_fill_row
                            elif hval == 'yellow':
                                fill = yellow_fill_row
                            elif hval == 'green':
                                fill = green_fill_row
                            if fill is not None:
                                for c in range(1, ws_review.max_column + 1):
                                    ws_review.cell(row=idx1, column=c).fill = fill
                        header_map = {str(ws_review.cell(row=1, column=c).value): c for c in range(1, ws_review.max_column + 1)}
                        orion_u_col = header_map.get('Orion Unit Price (col U)')
                        red_cell = PatternFill(fill_type='solid', start_color='FFFFC7CE', end_color='FFFFC7CE')
                        green_cell = PatternFill(fill_type='solid', start_color='FFC6EFCE', end_color='FFC6EFCE')
                        def _as_float(x):
                            try:
                                return float(str(x).replace(',', '').strip())
                            except Exception:
                                return None
                        if orion_u_col:
                            for idx0 in range(len(simple_rows)):
                                r_excel = idx0 + 2
                                try:
                                    row_df = df.iloc[idx0]
                                except Exception:
                                    row_df = None
                                pdf_price_val = _as_float(row_df['Unit Rate']) if row_df is not None and 'Unit Rate' in row_df else _as_float(simple_rows[idx0].get('Copy: Unit Price', ''))
                                orion_price_val = _as_float(simple_rows[idx0].get('Orion Unit Price (col U)', ''))
                                if orion_price_val is not None and pdf_price_val is not None and abs(orion_price_val - pdf_price_val) < 1e-6:
                                    ws_review.cell(row=r_excel, column=orion_u_col).fill = green_cell
                                else:
                                    ws_review.cell(row=r_excel, column=orion_u_col).fill = red_cell
                    # Create COMPONENT UPLOAD sheet with only the header row
                    component_headers = [
                        'PO Txn Code',
                        'PO Number',
                        'Parent Item Code',
                        'Component Item Code',
                        'UOM',
                        'Qty',
                        'Rate',
                    ]
                    pd.DataFrame(columns=component_headers).to_excel(
                        writer, sheet_name='COMPONENT UPLOAD', index=False
                    )
                    # Write master file content as sheet 2 if provided
                    if master_file is not None:
                        try:
                            master_file.seek(0)
                        except Exception:
                            pass
                        try:
                            df_master = pd.read_excel(master_file, header=8)
                            df_master.to_excel(writer, sheet_name='MASTER', index=False)
                        except Exception:
                            pass
                output.seek(0)
                st.download_button(
                    label="⬇️ Download PRE ALERT UPLOAD",
                    data=output.getvalue(),
                    file_name="pre_alert_upload.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_dell_pre_alert",
                   
                )
            else:
                st.warning("No items found in the uploaded PDF(s).")

    
elif tool == "🟨 AWS Invoice Tool":
        st.title("AWS Invoice Tool")
        st.write("Upload AWS invoice PDF(s) and download the extracted data as Excel.")
    
        uploaded_files = st.file_uploader(
            "Choose AWS invoice PDF(s)", type=["pdf"], key="aws_upload", accept_multiple_files=True
        )
    
        if uploaded_files:
            rows, template_map, text_map = process_multiple_aws_pdfs(uploaded_files)
            if rows:
                df = pd.DataFrame(rows, columns=AWS_OUTPUT_COLUMNS)
                
    
                output_original = io.BytesIO()
                with pd.ExcelWriter(output_original, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='AWS_INVOICE', index=False)
                output_original.seek(0)
                st.download_button(
                    label="⬇️ Download Extracted AWS Invoice Data",
                    data=output_original.getvalue(),
                    file_name="aws_invoice_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    
                )
    
                output_files = build_dnts_cnts_rows(rows, template_map, text_map)
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for file_key, data in output_files.items():
                        bill_to, file_type = file_key.split("__")
                        safe_bill_to = bill_to.replace(" ", "_").replace(".", "").replace(",", "")
                        file_name = f"{file_type}_{safe_bill_to}.xlsx"
                
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            pd.DataFrame(data["header"], columns=[
                                "S.No", "Date - (dd/MM/yyyy)", "Supp_Code", "Curr_Code", "Form_Code",
                                "Doc_Src_Locn", "Location_Code", "Remarks", "Supplier_Ref", "Supplier_Ref_Date - (dd/MM/yyyy)"
                            ]).to_excel(writer, sheet_name=f"{file_type}_HEADER", index=False)
                
                            pd.DataFrame(data["item"], columns=[
                                "S.No", "Ref. Key", "Item_Code", "Item_Name", "Grade1", "Grade2", "UOM",
                                "Qty", "Qty_Ls", "Rate", "Main_Account", "Sub_Account", "Division", "Department", "Analysis2"
                            ]).to_excel(writer, sheet_name=f"{file_type}_ITEM", index=False)
                
                        output.seek(0)
                        zip_file.writestr(file_name, output.read())
                
                zip_buffer.seek(0)
                
                st.download_button(
                    label="⬇️ Download All DNTS/CNTS Files as ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="aws_dnts_cnts_files.zip",
                    mime="application/zip",
             
                    on_click=lambda: (
                            update_usage("AWS Automation", team)
                        ),
                )
            else:
                st.warning("No data extracted from the uploaded AWS PDFs.")
        else:
            st.info("Please upload one or more AWS invoice PDFs to begin.")
            
elif tool == "Coface CSV Uploader":
    
    st.write("Upload an Excel file with customer invoice data to generate grouped outputs by customer code.")
    
    uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])
    
    if uploaded_file:
        st.success("✅ File uploaded successfully.")
        zip_output = process_grouped_customer_files(uploaded_file)
    
        st.download_button(
            label="⬇️ Download All Customer Files (ZIP)",
            data=zip_output.getvalue(),
            file_name="customer_outputs.zip",
            mime="application/zip",
            on_click=lambda: (
                            update_usage("credit format by customer", team)
                        ),
        )
elif tool == "AR to EDD file":
    st.title("AR to EDD file")
    st.write("Upload the insurance Excel file (starting from row 16) to filter and extract relevant data.")

    ageing_min_threshold = st.number_input(
        label="📅 Minimum Ageing Threshold (days)",
        min_value=0,
        value=200,
        step=1,
        help="Only include records with ageing greater than this number"
    )
    
    ageing_max_threshold = st.number_input(
        label="⏱️ Maximum Ageing Threshold (days)",
        min_value=0,
        value=270,  # default as requested
        step=1,
        help="Only include records with ageing less than or equal to this number"
    )


    uploaded_file = st.file_uploader(
        "Choose Insurance Excel File", type=["xlsx"], key="insurance_upload"
    )

    if uploaded_file:
        
        if ageing_min_threshold > ageing_max_threshold:
            st.error("Minimum ageing threshold cannot be greater than the maximum threshold.")
        else:

           output_excel = process_insurance_excel(
            uploaded_file,
            ageing_filter=True,
            
            ageing_min_threshold=ageing_min_threshold,
            ageing_max_threshold=ageing_max_threshold

        )

        st.download_button(
            label="⬇️ Download AR to EDD file",
            data=output_excel.getvalue(),
            file_name="EDD.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            on_click=lambda: (
                            update_usage("credit Automation ", team)
                            
                        ),
        )
elif tool == "IBM Quotation":

    st.header("🆕 IBM Excel to Excel + PDF to Excel (Combo)")
    st.info("Upload an IBM quotation PDF and (optionally) an Excel file. The tool will auto-detect the template and use the best logic for each.")

    # Country selection
    country = st.selectbox("Choose a country:", ["UAE", "Qatar", "KSA"])

    logo_path = "image.png"
    compliance_text = ""  # Add compliance text if needed

    st.subheader("📤 Upload IBM Quotation Files")

    uploaded_pdf = st.file_uploader(
        "Upload IBM Quotation PDF (.pdf)",
        type=["pdf"],
        help="Supports .pdf files. The tool will extract header information from the PDF."
    )

    uploaded_excel = st.file_uploader(
        "Upload IBM Quotation Excel (.xlsx, .xlsm, .xls)",
        type=["xlsx", "xlsm", "xls"],
        help="Supports .xlsx, .xlsm, and .xls files. The tool will extract line items from the second sheet."
    )

    if uploaded_pdf:
        from sales.ibm_v2_combo import process_ibm_combo
        import io
        pdf_bytes = io.BytesIO(uploaded_pdf.getbuffer())
        excel_bytes = io.BytesIO(uploaded_excel.getbuffer()) if uploaded_excel else None
        result = process_ibm_combo(pdf_bytes, excel_bytes, country=country)

        if result['error']:
            st.error(f"❌ {result['error']}")
        else:
            st.success(f"✅ Detected Template: {result['template']}")
            if result['mep_cost_msg']:
                st.info(result['mep_cost_msg'])
            if result['bid_number_error']:
                st.error(result['bid_number_error'])
            if result.get('date_validation_msg'):
                st.info(f"📅 Date Validation:\n{result['date_validation_msg']}")
            if result['data']:
                if result.get('columns'):
                    st.dataframe(pd.DataFrame(result['data'], columns=result['columns']))
                else:
                    st.dataframe(pd.DataFrame(result['data']))
            if result.get('excel_bytes'):
                st.download_button(
                    label="📥 Download Styled Excel File",
                    data=result['excel_bytes'],
                    file_name="Styled_Quotation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    on_click=lambda: (
                            update_usage("IBM Automation", team)
                            
                        ),
                    
                )



st.markdown("""
<footer style='text-align:center; margin-top:3rem; color:#1a73e8; font-size:20px; font-weight:bold; font-family: Google Sans, sans-serif;'>
    Made with ❤️ by Mindware | © 2025
</footer>
""", unsafe_allow_html=True)
