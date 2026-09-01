import io
import os
import msoffcrypto
import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Set layout with a clean title and icon
st.set_page_config(page_title="Payroll Validator", page_icon="⚡", layout="centered")

# Modern, clean CSS that adapts to both Dark Mode and Light Mode
st.markdown("""
    <style>
    /* Clean typography and padding adjustment */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 800px;
    }
    
    /* Sleek card-like styling for drag-and-drop file upload containers */
    [data-testid="stFileUploader"] {
        border-radius: 8px;
        padding: 8px;
    }
    
    /* Sleek, full-width primary button styling */
    div.stButton > button:first-child {
        width: 100%;
        border-radius: 6px;
        height: 3em;
        font-weight: 600;
        letter-spacing: 0.5px;
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
    """Decrypts and loads password-protected Excel files into a DataFrame."""
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
        raise ValueError("Failed to open encrypted file. Check format or password.")

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
    """Finds matching column names accounting for minor variations."""
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
    """Normalizes employee ID format."""
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str.replace(" ", "")

def process_validation(main_bytes, dr2_bytes, pwd):
    """Performs validation between Input File and Masterfile."""
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

# --- Minimalist Header ---
st.title("Payroll Validator")
st.caption("Upload source files to generate validation report.")

st.divider()

# --- Inputs Section ---
col1, col2 = st.columns(2)
with col1:
    main_file = st.file_uploader("UPLOAD INPUT FILE", type=["xlsx", "xlsb", "xls"])
with col2:
    dr2_file = st.file_uploader("UPLOAD MASTERFILE", type=["xlsx", "xlsb", "xls"])

password = st.text_input("File Password", value="tp_paseo", type="password")

st.markdown("<br>", unsafe_allow_html=True)

# --- Action Button ---
if st.button("RUN VALIDATION", type="primary"):
    if not main_file or not dr2_file:
        st.warning("Please upload both the Input File and Masterfile.")
    else:
        with st.spinner("Processing..."):
            try:
                result_excel = process_validation(main_file, dr2_file, password)
                st.success("Validation complete.")
                st.download_button(
                    label="Download Report",
                    data=result_excel,
                    file_name="Hourly_Regular_Hours_Validated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error: {str(e)}")