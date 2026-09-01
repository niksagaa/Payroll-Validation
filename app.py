import io
import os
import msoffcrypto
import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Page Config
st.set_page_config(
    page_title="Payroll Hours Validator",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Minimalist Ultra-Clean CSS
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    /* Hide Streamlit Native Overlays */
    #MainMenu, header, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] {
        visibility: hidden !important;
        display: none !important;
    }

    /* Global Canvas */
    html, body, [class*="css"], .stApp {
        font-family: 'Inter', -apple-system, sans-serif !important;
        background-color: #F8FAFC !important;
        color: #0F172A !important;
    }

    .main .block-container {
        padding-top: 4rem !important;
        padding-bottom: 4rem !important;
        max-width: 540px;
    }

    /* Floating Card Container Styling */
    .app-header {
        text-align: center;
        margin-bottom: 2rem;
    }

    .app-title {
        font-size: 1.75rem;
        font-weight: 700;
        color: #0F172A;
        letter-spacing: -0.025em;
        margin-bottom: 0.35rem;
    }

    .app-subtitle {
        font-size: 0.875rem;
        color: #64748B;
        line-height: 1.4;
    }

    /* Form Labels */
    .field-label {
        font-size: 0.825rem;
        font-weight: 600;
        color: #334155;
        margin-bottom: 0.4rem;
        display: block;
    }

    /* File Uploader Clean Styling */
    [data-testid="stFileUploader"] {
        margin-bottom: 1.25rem;
    }

    [data-testid="stFileUploaderDropzone"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 12px !important;
        padding: 1rem !important;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05) !important;
        transition: border-color 0.15s ease, box-shadow 0.15s ease !important;
    }

    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: #94A3B8 !important;
    }

    [data-testid="stFileUploaderDropzone"] span, 
    [data-testid="stFileUploaderDropzone"] small,
    [data-testid="stFileUploaderDropzone"] p {
        color: #64748B !important;
    }

    [data-testid="stFileUploaderDropzone"] button {
        background-color: #F1F5F9 !important;
        color: #0F172A !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
        font-size: 0.85rem !important;
    }

    /* Password Field Light Theme Fix */
    div[data-baseweb="input"],
    div[data-baseweb="input"] > div,
    input[type="password"],
    input[type="text"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border-radius: 12px !important;
    }

    div[data-baseweb="input"] {
        border: 1px solid #E2E8F0 !important;
        height: 2.85rem !important;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05) !important;
    }

    div[data-baseweb="input"]:focus-within {
        border-color: #0284C7 !important;
        box-shadow: 0 0 0 3px rgba(2, 132, 199, 0.12) !important;
    }

    /* Primary Action Button Clean Full-Width */
    [data-testid="stButton"],
    div.stButton {
        width: 100% !important;
        margin-top: 1.25rem;
    }

    [data-testid="stButton"] > button,
    div.stButton > button {
        width: 100% !important;
        border-radius: 12px !important;
        height: 3.1rem !important;
        font-weight: 600 !important;
        font-size: 0.925rem !important;
        background-color: #0F172A !important;
        border: none !important;
        color: #FFFFFF !important;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1) !important;
        transition: background-color 0.15s ease !important;
    }

    [data-testid="stButton"] > button:hover,
    div.stButton > button:hover {
        background-color: #1E293B !important;
    }

    /* Download Button */
    [data-testid="stDownloadButton"],
    div[data-testid="stDownloadButton"] {
        width: 100% !important;
    }

    [data-testid="stDownloadButton"] > button {
        background-color: #0284C7 !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        height: 3.1rem !important;
        border-radius: 12px !important;
        border: none !important;
        width: 100% !important;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1) !important;
    }

    [data-testid="stDownloadButton"] > button:hover {
        background-color: #0369A1 !important;
    }

    /* Success Card Banner */
    .success-card {
        background-color: #F0FDF4;
        border: 1px solid #BBF7D0;
        border-radius: 12px;
        padding: 1rem;
        text-align: center;
        margin-bottom: 1.25rem;
    }

    .success-title {
        font-weight: 600;
        font-size: 0.95rem;
        color: #166534;
        margin-bottom: 0.15rem;
    }

    .success-desc {
        font-size: 0.825rem;
        color: #15803D;
    }
    </style>
""", unsafe_allow_html=True)

ACCOUNTING_FORMAT = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

COLUMN_MAPPINGS = [
    ("Sum of Regular Hours", "Salaries and Wages (Units)"),
    ("Vacation Hours", "Vacation Leave (Hours)"),
    ("Regular ND Hours", "Night Differential (Hours)"),
    ("Regular OT Hours", "Restday OT > 8 (Hours)"),
    ("Sum of Rest Day Hours", "Restday OT (Hours)"),
]

def load_encrypted_excel(file_bytes, password, sheet_identifier):
    decrypted_stream = io.BytesIO()
    office_file = msoffcrypto.OfficeFile(file_bytes)
    office_file.load_key(password=password)
    office_file.decrypt(decrypted_stream)
    
    decrypted_stream.seek(0)
    engines = ["openpyxl", "pyxlsb", "xlrd"]
    excel_file = None
    
    for engine in engines:
        try:
            decrypted_stream.seek(0)
            excel_file = pd.ExcelFile(decrypted_stream, engine=engine)
            break
        except Exception:
            continue

    if not excel_file:
        raise ValueError("Cannot open file. Please check password or file format.")

    if isinstance(sheet_identifier, int):
        target_sheet = excel_file.sheet_names[sheet_identifier]
    else:
        target_clean = str(sheet_identifier).strip().lower()
        matched = [s for s in excel_file.sheet_names if s.strip().lower() == target_clean]
        target_sheet = matched[0] if matched else excel_file.sheet_names[0]

    df = pd.read_excel(excel_file, sheet_name=target_sheet)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def find_matching_column(df_columns, target_name):
    cols_clean = {str(c).strip().lower(): c for c in df_columns}
    target_clean = target_name.strip().lower()

    if target_clean in cols_clean:
        return cols_clean[target_clean]

    with_sum = f"sum of {target_clean}"
    if with_sum in cols_clean:
        return cols_clean[with_sum]

    if target_clean.startswith("sum of "):
        without_sum = target_clean.replace("sum of ", "").strip()
        if without_sum in cols_clean:
            return cols_clean[without_sum]

    return None

def clean_id(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str.replace(" ", "")

def process_validation(main_bytes, dr2_bytes, pwd):
    df_main = load_encrypted_excel(main_bytes, pwd, "Hourly Checker")
    df_dr2 = load_encrypted_excel(dr2_bytes, pwd, 0)

    actual_key_main = find_matching_column(df_main.columns, "Row Labels") or "Row Labels"
    actual_key_dr2 = find_matching_column(df_dr2.columns, "ID") or "ID"

    df_main[actual_key_main] = df_main[actual_key_main].apply(clean_id)
    df_dr2[actual_key_dr2] = df_dr2[actual_key_dr2].apply(clean_id)

    final_cols = []
    display_header_map = {}

    for col in df_main.columns:
        final_cols.append(col)
        display_header_map[col] = col

        mapping = next((m for m in COLUMN_MAPPINGS if find_matching_column([col], m[0])), None)

        if mapping:
            main_target, dr2_target = mapping
            dr2_col_match = find_matching_column(df_dr2.columns, dr2_target)

            dr2_val_col = f"__DR2_{col}"
            chk_col = f"__CHK_{col}"
            rem_col = f"__REM_{col}"

            final_cols.extend([dr2_val_col, chk_col, rem_col])
            display_header_map[dr2_val_col] = "DR2"
            display_header_map[chk_col] = "CHECKER"
            display_header_map[rem_col] = "REMARKS"

            dr2_grouped = df_dr2.groupby(actual_key_dr2)[dr2_col_match].sum().to_dict() if dr2_col_match else {}

            dr2_vals, chk_vals, rem_vals = [], [], []

            for _, row in df_main.iterrows():
                emp_id = row[actual_key_main]
                main_val = pd.to_numeric(row[col], errors='coerce') or 0.0

                if dr2_col_match and emp_id in dr2_grouped:
                    dr2_val = float(dr2_grouped[emp_id])
                    diff = main_val - dr2_val

                    dr2_vals.append(dr2_val)
                    if abs(diff) < 0.001:
                        chk_vals.append("-")
                        rem_vals.append("OK; NO VARIANCE")
                    else:
                        chk_vals.append(diff)
                        rem_vals.append("OK; NOT ELIGIBLE")
                else:
                    dr2_vals.append(None)
                    chk_vals.append("#N/A")
                    rem_vals.append("NOT IN DR2")

            df_main[dr2_val_col] = dr2_vals
            df_main[chk_col] = chk_vals
            df_main[rem_col] = rem_vals

    df_final = df_main[final_cols]

    output_stream = io.BytesIO()
    writer = pd.ExcelWriter(output_stream, engine="openpyxl")
    df_final.to_excel(writer, sheet_name="Validated", index=False, startrow=0)
    
    wb = writer.book
    ws = writer.sheets["Validated"]

    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )

    header_display_names = [display_header_map[col] for col in final_cols]
    for col_idx, text in enumerate(header_display_names, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = text
        cell.border = thin_border

        if text in ["DR2", "CHECKER", "REMARKS"]:
            cell.fill = PatternFill(start_color="D4AC0D", end_color="D4AC0D", fill_type="solid")
            cell.font = Font(name="Calibri", size=10, bold=True, color="000000")
        else:
            cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

        cell.alignment = Alignment(horizontal="center", vertical="center")

    key_col_index = final_cols.index(actual_key_main) + 1

    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = Font(name="Calibri", size=10)
            cell.border = thin_border

            if col_idx == key_col_index:
                cell.number_format = '@'
                cell.alignment = Alignment(horizontal="left", vertical="center")
                continue

            if isinstance(cell.value, (int, float)):
                cell.number_format = ACCOUNTING_FORMAT
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif cell.value in ["-", "#N/A"]:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or '')
            if cell.number_format == ACCOUNTING_FORMAT and isinstance(cell.value, (int, float)):
                val = f"{cell.value:,.2f}"
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 12)

    writer.close()
    output_stream.seek(0)
    return output_stream

# --- HEADER SECTION ---
st.markdown("""
    <div class="app-header">
        <div class="app-title">Payroll Hours Validator</div>
        <div class="app-subtitle">Cross-reference payroll ledgers against masterfile records.</div>
    </div>
""", unsafe_allow_html=True)

# --- INPUT FIELDS ---
st.markdown('<span class="field-label">1. Input File</span>', unsafe_allow_html=True)
main_file = st.file_uploader(
    "1. Input File",
    type=["xlsx", "xlsb", "xls"],
    key="main_file_key",
    label_visibility="collapsed"
)

st.markdown('<span class="field-label">2. Masterfile</span>', unsafe_allow_html=True)
dr2_file = st.file_uploader(
    "2. Masterfile",
    type=["xlsx", "xlsb", "xls"],
    key="dr2_file_key",
    label_visibility="collapsed"
)

st.markdown('<span class="field-label">3. Decryption Password</span>', unsafe_allow_html=True)
password = st.text_input(
    "3. Decryption Password",
    value="tp_paseo",
    type="password",
    label_visibility="collapsed"
)

# --- ACTION & OUTPUT ---
if st.button("Run Validation", type="primary"):
    if not main_file or not dr2_file:
        st.warning("Please upload both the Input File and Masterfile.")
    else:
        with st.spinner("Processing files..."):
            try:
                result_excel = process_validation(main_file, dr2_file, password)
                
                st.markdown("""
                    <div class="success-card">
                        <div class="success-title">Validation Complete</div>
                        <div class="success-desc">The report has been generated successfully.</div>
                    </div>
                """, unsafe_allow_html=True)
                
                st.download_button(
                    label="Download Validated File (.xlsx)",
                    data=result_excel,
                    file_name="Hourly_Regular_Hours_Validated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error processing files: {str(e)}")