import io
import os
import msoffcrypto
import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Payroll Hours Validator",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- MODERN UI STYLING (CSS) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    #MainMenu, header, footer, [data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"] {
        visibility: hidden !important;
        display: none !important;
    }

    html, body, [class*="css"], .stApp {
        font-family: 'Plus Jakarta Sans', -apple-system, sans-serif !important;
        background: #F4F6F3 !important;
        color: #1A2E26 !important;
    }

    .main .block-container {
        padding-top: 2rem !important;
        padding-bottom: 3rem !important;
        max-width: 1000px;
    }

    /* Header Styling */
    .zen-header {
        text-align: center;
        margin-bottom: 2rem;
    }

    .zen-badge {
        display: inline-block;
        background-color: #E2ECE6;
        color: #2E5A44;
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.8px;
        padding: 5px 16px;
        border-radius: 20px;
        margin-bottom: 0.8rem;
    }

    .zen-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #12231C;
        letter-spacing: -0.5px;
        margin-bottom: 0.4rem;
    }

    .zen-subtitle {
        font-size: 0.95rem;
        color: #5A6E65;
    }

    .field-label {
        font-size: 0.88rem;
        font-weight: 700;
        color: #1E3E2F;
        margin-bottom: 0.4rem;
        display: block;
    }

    /* File Uploader styling */
    [data-testid="stFileUploader"] {
        background-color: transparent !important;
    }

    [data-testid="stFileUploaderDropzone"] {
        background-color: #FFFFFF !important;
        border: 1.5px dashed #C2D3C9 !important;
        border-radius: 14px !important;
        padding: 1rem !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.02) !important;
    }

    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: #2E5A44 !important;
        background-color: #F8FAF9 !important;
    }

    /* Input Password styling */
    div[data-baseweb="input"] {
        background-color: #FFFFFF !important;
        border: 1.5px solid #C2D3C9 !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.02) !important;
    }

    div[data-baseweb="input"]:focus-within {
        border-color: #2E5A44 !important;
    }

    /* Primary Buttons */
    div.stButton > button, [data-testid="stDownloadButton"] > button {
        width: 100% !important;
        border-radius: 12px !important;
        height: 3.4rem !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        background: #2E5A44 !important;
        border: none !important;
        color: #FFFFFF !important;
        transition: all 0.25s ease !important;
        box-shadow: 0 6px 16px rgba(46, 90, 68, 0.2) !important;
    }

    div.stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
        background: #234735 !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 10px 20px rgba(46, 90, 68, 0.3) !important;
    }

    /* Success Card */
    .success-card {
        background-color: #FFFFFF;
        border: 1.5px solid #2E5A44;
        border-radius: 14px;
        padding: 1.2rem;
        text-align: center;
        margin-top: 1rem;
        margin-bottom: 1rem;
        box-shadow: 0 6px 16px rgba(46, 90, 68, 0.06);
    }

    .success-title {
        font-weight: 700;
        font-size: 1.05rem;
        color: #2E5A44;
    }

    .success-desc {
        font-size: 0.85rem;
        color: #5A6E65;
    }
    </style>
""", unsafe_allow_html=True)

# --- CONSTANTS ---
ACCOUNTING_FORMAT = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'

COLUMN_MAPPINGS = [
    ("Sum of Regular Hours", "Salaries and Wages (Units)"),
    ("Vacation Hours", "Vacation Leave (Hours)"),
    ("Regular ND Hours", "Night Differential (Hours)"),
    ("Regular OT Hours", "Restday OT > 8 (Hours)"),
    ("Sum of Rest Day Hours", "Restday OT (Hours)"),
]

# --- HELPER FUNCTIONS ---
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

def load_encrypted_excel(file_bytes, password, sheet_identifier):
    decrypted_stream = io.BytesIO()
    try:
        office_file = msoffcrypto.OfficeFile(file_bytes)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted_stream)
    except Exception as e:
        raise ValueError("Maling password o sira ang encrypted file structure.")

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
        raise ValueError("Hindi mabuksan ang Excel stream. Suriin ang format o password.")

    if isinstance(sheet_identifier, int):
        target_sheet = excel_file.sheet_names[sheet_identifier]
    else:
        target_clean = str(sheet_identifier).strip().lower()
        matched = [s for s in excel_file.sheet_names if s.strip().lower() == target_clean]
        target_sheet = matched[0] if matched else excel_file.sheet_names[0]

    df = pd.read_excel(excel_file, sheet_name=target_sheet)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def process_validation(main_file_obj, dr2_file_obj, pwd):
    main_bytes = io.BytesIO(main_file_obj.getvalue())
    dr2_bytes = io.BytesIO(dr2_file_obj.getvalue())

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
            _, dr2_target = mapping
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
    with pd.ExcelWriter(output_stream, engine="openpyxl") as writer:
        df_final.to_excel(writer, sheet_name="Validated", index=False)
        
        ws = writer.sheets["Validated"]
        thin_border = Border(
            left=Side(style='thin', color='D3D3D3'),
            right=Side(style='thin', color='D3D3D3'),
            top=Side(style='thin', color='D3D3D3'),
            bottom=Side(style='thin', color='D3D3D3')
        )

        header_display_names = [display_header_map[c] for c in final_cols]
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

    output_stream.seek(0)
    return output_stream

# --- UI HEADER ---
st.markdown("""
    <div class="zen-header">
        <div class="zen-badge">🌿 AUTOMATED VALIDATION WORKSPACE</div>
        <div class="zen-title">Payroll Hours Validator</div>
        <div class="zen-subtitle">Seamlessly reconcile input ledgers with your masterfile in seconds.</div>
    </div>
""", unsafe_allow_html=True)

# --- FORM LAYOUT (2 COLUMNS) ---
col1, col2 = st.columns(2)

with col1:
    st.markdown('<span class="field-label">1. Input File</span>', unsafe_allow_html=True)
    main_file = st.file_uploader(
        "1. Input File",
        type=["xlsx", "xlsb", "xls"],
        key="main_file_key",
        label_visibility="collapsed"
    )

with col2:
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

st.write("") # Spacing

# --- ACTION & EXECUTION ---
if st.button("Start Validation Engine 🌿", type="primary"):
    if not main_file or not dr2_file:
        st.warning("Pakisuri ang files: kailangan parehong naka-upload ang Input File at Masterfile.")
    elif not password:
        st.warning("Pakilagay ang decryption password.")
    else:
        with st.spinner("Reconciling payroll records..."):
            try:
                result_excel = process_validation(main_file, dr2_file, password)
                
                st.markdown("""
                    <div class="success-card">
                        <div class="success-title">Validation Successful</div>
                        <div class="success-desc">All calculations processed seamlessly. Ready for download.</div>
                    </div>
                """, unsafe_allow_html=True)
                
                st.download_button(
                    label="Download Validated Report (.xlsx)",
                    data=result_excel,
                    file_name="Hourly_Regular_Hours_Validated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error sa Pagproseso: {str(e)}")