import io
import os
import msoffcrypto
import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Page Config
st.set_page_config(
    page_title="Payroll Hourly Validator",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom Light Widget Card CSS (Exact Match to Reference UI)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    /* Global Light Canvas Overrides */
    html, body, [class*="css"], .stApp {
        font-family: 'Inter', -apple-system, sans-serif !important;
        background-color: #F8FAFC !important;
        color: #1E293B !important;
    }

    .main .block-container {
        padding-top: 2.5rem;
        padding-bottom: 3rem;
        max-width: 680px;
    }

    /* Main Container Card (Floating White Box with Soft Shadow) */
    div.element-container {
        font-family: 'Inter', sans-serif !important;
    }

    /* Form Container Setup */
    .widget-card {
        background: #FFFFFF;
        border-radius: 20px;
        padding: 2rem;
        box-shadow: 0 10px 30px -5px rgba(0, 0, 0, 0.05), 0 4px 12px -2px rgba(0, 0, 0, 0.03);
        border: 1px solid #F1F5F9;
        margin-bottom: 1.5rem;
    }

    /* Header Typography */
    .card-title {
        font-size: 1.5rem;
        font-weight: 700;
        color: #0F172A;
        letter-spacing: -0.5px;
        margin-bottom: 0.2rem;
    }

    .card-subtitle {
        font-size: 0.9rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }

    /* Section Labels */
    .input-label {
        font-size: 0.82rem;
        font-weight: 600;
        color: #475569;
        margin-bottom: 0.4rem;
        display: block;
    }

    /* File Upload Area Customization */
    [data-testid="stFileUploader"] {
        background-color: #FFFFFF !important;
        border: 1.5px solid #E2E8F0 !important;
        border-radius: 12px !important;
        padding: 8px !important;
        transition: all 0.2s ease;
    }

    [data-testid="stFileUploader"]:hover {
        border-color: #16A34A !important;
        background-color: #F8FAF9 !important;
    }

    /* Password Text Input Override */
    div[data-baseweb="input"] {
        background-color: #FFFFFF !important;
        border: 1.5px solid #E2E8F0 !important;
        border-radius: 12px !important;
        color: #0F172A !important;
        height: 2.8rem;
    }

    div[data-baseweb="input"]:focus-within {
        border-color: #16A34A !important;
        box-shadow: 0 0 0 3px rgba(22, 163, 74, 0.15) !important;
    }

    /* Green Primary Action Button (Matching Reference Image) */
    div.stButton > button:first-child {
        width: 100%;
        border-radius: 12px;
        height: 3.4rem;
        font-weight: 600;
        font-size: 1rem;
        background-color: #16A34A !important;
        border: none !important;
        color: #FFFFFF !important;
        transition: all 0.2s ease;
        margin-top: 0.5rem;
        box-shadow: 0 4px 12px rgba(22, 163, 74, 0.25);
    }

    div.stButton > button:first-child:hover {
        background-color: #15803D !important;
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(22, 163, 74, 0.35);
    }

    /* Download Button Specific Green Pill */
    div[data-testid="stDownloadButton"] > button {
        background-color: #16A34A !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        height: 3.4rem !important;
        border-radius: 12px !important;
        border: none !important;
        width: 100% !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 4px 12px rgba(22, 163, 74, 0.25) !important;
    }

    div[data-testid="stDownloadButton"] > button:hover {
        background-color: #15803D !important;
        transform: translateY(-1px) !important;
    }

    /* Result Metric Boxes (Soft Colored Badges) */
    .metric-card-green {
        background-color: #F0FDF4;
        border: 1px solid #DCFCE7;
        border-radius: 12px;
        padding: 1rem;
        text-align: center;
        margin-bottom: 0.8rem;
    }

    .metric-value-green {
        font-size: 1.4rem;
        font-weight: 700;
        color: #16A34A;
    }

    .metric-label-green {
        font-size: 0.78rem;
        font-weight: 500;
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
    """Decrypts and loads Excel files safely in memory."""
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
    """Smart column identification."""
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
    """Cleans up Employee IDs."""
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str.replace(" ", "")

def process_validation(main_bytes, dr2_bytes, pwd):
    """Core validation processing."""
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
st.markdown('<div class="card-title">Validate Payroll Hours</div>', unsafe_allow_html=True)
st.markdown('<div class="card-subtitle">Cross-check input files against masterfile records instantly.</div>', unsafe_allow_html=True)

# --- UPLOAD FORM CONTAINER ---
st.markdown('<span class="input-label">Select Input File</span>', unsafe_allow_html=True)
main_file = st.file_uploader(
    "Upload Input File",
    type=["xlsx", "xlsb", "xls"],
    key="main_file_key",
    label_visibility="collapsed"
)

st.markdown("<br>", unsafe_allow_html=True)

st.markdown('<span class="input-label">Select Masterfile</span>', unsafe_allow_html=True)
dr2_file = st.file_uploader(
    "Upload Masterfile",
    type=["xlsx", "xlsb", "xls"],
    key="dr2_file_key",
    label_visibility="collapsed"
)

st.markdown("<br>", unsafe_allow_html=True)

st.markdown('<span class="input-label">Excel Security Password</span>', unsafe_allow_html=True)
password = st.text_input(
    "File Password",
    value="tp_paseo",
    type="password",
    label_visibility="collapsed"
)

st.markdown("<br>", unsafe_allow_html=True)

# --- ACTION & RESULTS ---
if st.button("Run Validation", type="primary"):
    if not main_file or not dr2_file:
        st.error("Please upload both the Input File and Masterfile to proceed.")
    else:
        with st.spinner("Analyzing payroll records..."):
            try:
                result_excel = process_validation(main_file, dr2_file, password)
                
                # Soft Colored Metric Cards (Matching Reference Style)
                m1, m2 = st.columns(2)
                with m1:
                    st.markdown("""
                        <div class="metric-card-green">
                            <div class="metric-value-green">Ready</div>
                            <div class="metric-label-green">Validation Status</div>
                        </div>
                    """, unsafe_allow_html=True)
                with m2:
                    st.markdown("""
                        <div class="metric-card-green">
                            <div class="metric-value-green">.XLSX</div>
                            <div class="metric-label-green">Output Format</div>
                        </div>
                    """, unsafe_allow_html=True)
                
                st.download_button(
                    label="Download Validated Report",
                    data=result_excel,
                    file_name="Hourly_Regular_Hours_Validated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error processing files: {str(e)}")