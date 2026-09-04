import io
import os
import json
import sqlite3
import uvicorn
import msoffcrypto
import pandas as pd
from difflib import SequenceMatcher
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="Payroll Hours Validator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/", response_class=HTMLResponse)
async def serve_frontend():
    if os.path.exists("index.html"):
        with open("index.html", "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>index.html not found! Ilagay ang index.html sa parehong folder.</h1>"

DB_FILE = "payroll_memory.db"

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS learned_mappings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            input_col_name TEXT UNIQUE,
            master_col_name TEXT,
            frequency INTEGER DEFAULT 1
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def get_smart_suggestion(input_col: str) -> str:
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('SELECT master_col_name FROM learned_mappings WHERE input_col_name = ?', (str(input_col).strip().lower(),))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else None

def normalize_text(text):
    return str(text).lower().replace("sum of ", "").replace("  ", " ").strip()

def get_similarity(a, b):
    return SequenceMatcher(None, normalize_text(a), normalize_text(b)).ratio()

def smart_match_columns(main_cols, dr2_cols, threshold=0.55):
    matched_map = {}
    ambiguous_cols = []
    for m_col in main_cols:
        smart_match = get_smart_suggestion(m_col)
        if smart_match and smart_match in dr2_cols:
            matched_map[m_col] = smart_match
            continue

        best_match, highest_score = None, 0.0
        for d_col in dr2_cols:
            score = get_similarity(m_col, d_col)
            norm_m, norm_d = normalize_text(m_col), normalize_text(d_col)
            if norm_m in norm_d or norm_d in norm_m:
                score = max(score, 0.85)
            if score > highest_score:
                highest_score, best_match = score, d_col
        if highest_score >= threshold:
            matched_map[m_col] = best_match
        else:
            ambiguous_cols.append(m_col)
            matched_map[m_col] = "-- Skip / Do Not Compare --"
    return matched_map, ambiguous_cols

def read_excel_file(file_bytes, sheet_name=0, password=""):
    buffer = io.BytesIO(file_bytes)
    try:
        if password:
            decrypted = io.BytesIO()
            office_file = msoffcrypto.OfficeFile(buffer)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            return pd.read_excel(decrypted, sheet_name=sheet_name)
        else:
            return pd.read_excel(buffer, sheet_name=sheet_name)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error reading file/password: {str(e)}")

def get_excel_sheets(file_bytes, password=""):
    buffer = io.BytesIO(file_bytes)
    try:
        if password:
            decrypted = io.BytesIO()
            office_file = msoffcrypto.OfficeFile(buffer)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
            xls = pd.ExcelFile(decrypted)
        else:
            xls = pd.ExcelFile(buffer)
        return xls.sheet_names
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid password or corrupt file: {str(e)}")

@app.post("/get-sheets")
async def api_get_sheets(file: UploadFile = File(...), password: str = Form("")):
    contents = await file.read()
    sheets = get_excel_sheets(contents, password)
    return {"sheet_names": sheets}

@app.post("/analyze-mapping")
async def api_analyze_mapping(
    payroll_inputs: UploadFile = File(...),
    payroll_masterfile: UploadFile = File(...),
    password: str = Form(""),
    sheet_name: str = Form("Sheet1"),
    employee_id_col: str = Form("")
):
    input_bytes = await payroll_inputs.read()
    master_bytes = await payroll_masterfile.read()

    df_input = read_excel_file(input_bytes, sheet_name=sheet_name, password=password)
    df_master = read_excel_file(master_bytes, sheet_name=0, password=password)

    main_cols = [str(c) for c in df_input.columns]
    dr2_cols = [str(c) for c in df_master.columns]

    detected_key = employee_id_col
    if not detected_key:
        for col in main_cols:
            if any(k in col.lower() for k in ["id", "employee", "emp"]):
                detected_key = col
                break
        if not detected_key and main_cols:
            detected_key = main_cols[0]

    initial_mapping, _ = smart_match_columns(main_cols, dr2_cols)

    return {
        "main_columns": main_cols,
        "data_cols": main_cols,
        "dr2_columns": dr2_cols,
        "detected_key_main": detected_key,
        "initial_mapping": initial_mapping
    }

@app.post("/validate-mapped")
async def api_validate_mapped(
    payroll_inputs: UploadFile = File(...),
    payroll_masterfile: UploadFile = File(...),
    password: str = Form(""),
    all_mappings_json: str = Form("{}"),
    all_id_keys_json: str = Form("{}")
):
    input_bytes = await payroll_inputs.read()
    master_bytes = await payroll_masterfile.read()
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_master = read_excel_file(master_bytes, sheet_name=0, password=password)
        df_master.to_excel(writer, sheet_name="Validation_Summary", index=False)
    output.seek(0)
    
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=Payroll_Validated_Report.xlsx"}
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("app:app", host="0.0.0.0", port=port)