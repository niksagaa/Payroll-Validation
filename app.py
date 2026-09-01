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
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom High-End Technical / Portfolio Dashboard Aesthetics CSS
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Space+Grotesk:wght@400;500;600;700&display=swap');

    /* Global Typography Reset */
    html, body, [class*="css"], .stMarkdown, button, input {
        font-family: 'Space Grotesk', -apple-system, sans-serif !important;
        background-color: #0A0A0C !important;
        color: #E2E8F0 !important;
    }

    .main .block-container {
        padding-top: 3rem;
        padding-bottom: 4rem;
        max-width: 820px;
    }

    /* Top Header Section */
    .portfolio-tag {
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.72rem;
        letter-spacing: 1.5px;
        color: #94A3B8;
        text-transform: uppercase;
        margin-bottom: 0.5rem;
    }

    .app-title {
        font-size: 2.2rem;
        font-weight: 700;
        letter-spacing: -0.8px;
        color: #FFFFFF;
        margin-bottom: 0.4rem;
    }

    .app-subtitle {
        font-size: 0.95rem;
        color: #8E8E93;
        margin-bottom: 2rem;
    }

    /* Technical Card Containers */
    div[data-testid="stColumn"] {
        background: #141417;
        border: 1px solid #242429;
        border-radius: 14px;
        padding: 1.2rem;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
    }

    div[data-testid="stColumn"]:hover {
        border-color: #3B82F6;
        transform: translateY(-2px);
        box-shadow: 0 10px 30px -10px rgba(0, 0, 0, 0.8);
    }

    /* Mono Section Badges */
    .mono-badge {
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.75rem;
        font-weight: 500;
        letter-spacing: 1px;
        color: #3B82F6;
        background: rgba(59, 130, 246, 0.1);
        padding: 4px 10px;
        border-radius: 20px;
        border: 1px solid rgba(59, 130, 246, 0.2);
        display: inline-block;
        margin-bottom: 0.8rem;
    }

    /* Streamlit Uploaders Overrides */
    [data-testid="stFileUploader"] {
        background: transparent !important;
        border: 1px dashed #2E2E35 !important;
        border-radius: 10px !important;
        padding: 8px !important;
        transition: border 0.2s ease;
    }

    [data-testid="stFileUploader"]:hover {
        border-color: #3B82F6 !important;
    }

    /* Custom Inputs */
    div[data-baseweb="input"] {
        background-color: #141417 !important;
        border: 1px solid #242429 !important;
        border-radius: 8px !important;
        color: #FFFFFF !important;
    }

    /* Action Buttons (High Contrast Minimalist Accent) */
    div.stButton > button:first-child {
        width: 100%;
        border-radius: 10px;
        height: 3.4rem;
        font-family: 'JetBrains Mono', monospace !important;
        font-weight: 700;
        font-size: 0.95rem;
        letter-spacing: 1px;
        background: #FFFFFF;
        border: none;
        color: #0A0A0C;
        transition: all 0.2s ease;
        margin-top: 1rem;
    }

    div.stButton > button:first-child:hover {
        background: #E2E8F0;
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(255, 255, 255, 0.15);
    }

    /* Download Button Styling */
    div[data-testid="stDownloadButton"] > button {
        background: #10B981 !important;
        color: #000000 !important;
        font-family: 'JetBrains Mono', monospace !important;
        font-weight: 700 !important;
        height: 3.4rem !important;
        border-radius: 10px !important;
        border: none !important;
        transition: all 0.2s ease !important;
    }

    div[data-testid="stDownloadButton"] > button:hover {
        background: #059669 !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(16, 185, 129, 0.25) !important;
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
        raise ValueError("Cannot open file. Check password or file format.")

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
st.markdown('<div class="portfolio-tag">[ PAYROLL UTILITY // ENGINE v2.0 ]</div>', unsafe_allow_html=True)
st.markdown('<div class="app-title">Hourly Payroll Validator</div>', unsafe_allow_html=True)
st.markdown('<div class="app-subtitle">Automated ledger cross-validation system for non-technical teams.</div>', unsafe_allow_html=True)

# --- STEP 1: FILE ATTACHMENTS ---
col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="mono-badge">01. INPUT FILE</div>', unsafe_allow_html=True)
    main_file = st.file_uploader(
        "Upload Input File",
        type=["xlsx", "xlsb", "xls"],
        key="main_file_key",
        label_visibility="collapsed"
    )

with col2:
    st.markdown('<div class="mono-badge">02. MASTERFILE</div>', unsafe_allow_html=True)
    dr2_file = st.file_uploader(
        "Upload Masterfile",
        type=["xlsx", "xlsb", "xls"],
        key="dr2_file_key",
        label_visibility="collapsed"
    )

st.markdown("<br>", unsafe_allow_html=True)

# --- STEP 2: SECURITY CONFIG ---
st.markdown('<div class="mono-badge">03. SECURITY KEY</div>', unsafe_allow_html=True)

password = st.text_input(
    "File Password",
    value="tp_paseo",
    type="password",
    help="Decryption password for protected workbooks."
)

st.markdown("<br>", unsafe_allow_html=True)

# --- STEP 3: EXECUTION ---
if st.button("RUN VALIDATION ENGINE ⚡", type="primary"):
    if not main_file or not dr2_file:
        st.error("Please upload both required Excel files to begin.")
    else:
        with st.spinner("Executing cross-validation logic..."):
            try:
                result_excel = process_validation(main_file, dr2_file, password)
                st.success("Validation completed successfully.")
                
                st.download_button(
                    label="DOWNLOAD VALIDATED REPORT (.XLSX)",
                    data=result_excel,
                    file_name="Hourly_Regular_Hours_Validated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Execution Error: {str(e)}")