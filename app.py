# Streamlit app: Upload + Upsert to Google Sheets
# Filename: app.py
# ---------------------------------------------------
# Features:
# - Connect to a Google Spreadsheet using a Service Account JSON (upload or paste)
# - List worksheets (sheets) and let user choose target sheet
# - Upload Excel/CSV file and preview
# - Choose matching column (common between upload & sheet) to perform UPSERT (update existing rows, insert new rows)
# - Show summary: rows added, rows updated, rows unchanged
# - Optional: write an update log into a sheet named "__update_log__"

import io
import json
import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# ---------------------------
# Helper functions
# ---------------------------

def auth_with_service_account_json(creds_json: dict, scopes=None):
    if scopes is None:
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_json, scopes=scopes)
    client = gspread.authorize(creds)
    return client


def read_spreadsheet_sheets(client, spreadsheet_id):
    try:
        ss = client.open_by_key(spreadsheet_id)
        sheets = ss.worksheets()
        return ss, [ws.title for ws in sheets]
    except Exception as e:
        st.error(f"Gagal buka spreadsheet: {e}")
        return None, []


def read_worksheet_to_df(spreadsheet, sheet_name):
    try:
        ws = spreadsheet.worksheet(sheet_name)
        records = ws.get_all_records()
        df = pd.DataFrame(records)
        return df, ws
    except Exception as e:
        st.error(f"Gagal baca sheet '{sheet_name}': {e}")
        return pd.DataFrame(), None


def upload_file_to_df(uploaded_file):
    # supports csv and xlsx
    if uploaded_file is None:
        return pd.DataFrame()
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    else:
        return pd.read_excel(uploaded_file)


def find_header_range_for_update(header_row_idx, start_col=1, end_col=None):
    # helper not used now, placeholder
    pass


def append_rows_to_sheet(ws, df_to_append: pd.DataFrame):
    if df_to_append.empty:
        return 0
    values = df_to_append.fillna("").values.tolist()
    ws.append_rows(values, value_input_option="USER_ENTERED")
    return len(values)


def update_row_by_index(ws, row_idx: int, df_columns: list, row_values: list):
    # row_idx is 1-based (worksheet row). We'll write row starting at column A.
    # Build range like A{row_idx}:{lastcol}{row_idx}
    last_col = gspread.utils.rowcol_to_a1(1, len(row_values)).lstrip('1')  # hack to get column letters
    # alternative: build range string via columns
    start_a1 = f"A{row_idx}"
    end_col_letter = gspread.utils.a1_range_to_grid_range(f"A1:{last_col}1")["endColumnIndex"]
    # Simpler: use update with list of lists and specify row index via range using values_input_option
    # We'll compute end col letter properly
    from string import ascii_uppercase
    def col_letter(n):
        # n is 1-indexed
        result = ""
        while n:
            n, r = divmod(n - 1, 26)
            result = chr(65 + r) + result
        return result
    last_col_letter = col_letter(len(row_values))
    range_str = f"A{row_idx}:{last_col_letter}{row_idx}"
    ws.update(range_str, [row_values], value_input_option="USER_ENTERED")


# ---------------------------
# Streamlit UI
# ---------------------------

st.set_page_config(page_title="Upload & Upsert to Google Sheets", layout="centered")
st.title("📥 Upload & Upsert ke Google Sheets")

st.markdown(
    "Upload service account JSON (recommended) atau paste JSON ke text area. Aplikasi akan membaca daftar sheet dan memungkinkan upsert (update/insert)."
)

# 1. Auth: service account JSON upload or paste
col1, col2 = st.columns(2)
with col1:
    uploaded_creds = st.file_uploader("Upload Service Account JSON", type=["json"], help="Service account JSON. Beri akses editor ke spreadsheet jika perlu.")
with col2:
    creds_text = st.text_area("Atau paste Service Account JSON di sini (opsional)")

creds_json = None
if uploaded_creds is not None:
    try:
        creds_json = json.load(uploaded_creds)
    except Exception as e:
        st.error(f"Gagal membaca JSON: {e}")
elif creds_text.strip():
    try:
        creds_json = json.loads(creds_text)
    except Exception as e:
        st.error(f"Gagal parse JSON: {e}")

if creds_json is None:
    st.info("Silakan upload atau paste Service Account JSON untuk melanjutkan.")
    st.stop()

# 2. Connect
try:
    client = auth_with_service_account_json(creds_json)
except Exception as e:
    st.error(f"Autentikasi gagal: {e}")
    st.stop()

# 3. Spreadsheet input (provide sheet id directly or full URL)
st.markdown("---")
spreadsheet_input = st.text_input("Spreadsheet ID atau URL", value="1RzN1lFuAOn4f8-tvliRryt4lTd2FljrZ4xbJLuOnkhU")

# extract ID if URL provided
def extract_spreadsheet_id(s: str):
    if "docs.google.com" in s:
        parts = s.split("/")
        for i, p in enumerate(parts):
            if p == "d" and i + 1 < len(parts):
                return parts[i + 1]
        # fallback: try to find /d/{id}/
        import re
        m = re.search(r"/d/([a-zA-Z0-9-_]+)", s)
        if m:
            return m.group(1)
    return s

ss_id = extract_spreadsheet_id(spreadsheet_input.strip())

spreadsheet, sheet_list = read_spreadsheet_sheets(client, ss_id)
if spreadsheet is None:
    st.stop()

st.success(f"Terhubung ke spreadsheet. Ditemukan {len(sheet_list)} sheet.")

# show sheet list and select target
target_sheet = st.selectbox("Pilih sheet target", sheet_list)

# option: create / ensure update log sheet
log_toggle = st.checkbox("Buat atau update sheet log '__update_log__' (opsional)", value=True)

# 4. Upload file to upsert
st.markdown("---")
uploaded_file = st.file_uploader("Upload file Excel (.xlsx) atau CSV untuk di-upsert", type=["xlsx", "csv"] )

if uploaded_file is None:
    st.info("Silakan upload file data untuk mulai. File akan diproses setelah kamu memilih kolom matching dan klik tombol proses.")
    st.stop()

try:
    df_new = upload_file_to_df(uploaded_file)
except Exception as e:
    st.error(f"Gagal baca file upload: {e}")
    st.stop()

if df_new.empty:
    st.error("File upload kosong atau tidak terbaca.")
    st.stop()

st.write("📂 Preview data upload (5 baris pertama):")
st.dataframe(df_new.head())

# 5. Read existing sheet
df_existing, ws = read_worksheet_to_df(spreadsheet, target_sheet)

if ws is None:
    st.stop()

st.write(f"📄 Preview data di sheet '{target_sheet}' (5 baris pertama):")
st.dataframe(df_existing.head())

# 6. Choose matching column
common_cols = list(set(df_new.columns).intersection(set(df_existing.columns)))

if len(df_existing.columns) == 0:
    st.warning("Sheet target tampaknya kosong atau tidak memiliki header. Semua data upload akan ditambahkan dan header akan disesuaikan.")
    matching_col = st.selectbox("Pilih kolom matching (sheet kosong -> pilih kolom dari file upload)", options=list(df_new.columns))
else:
    if not common_cols:
        st.warning("Tidak ada kolom yang cocok antara file upload dan sheet target. Kamu bisa pilih kolom dari file upload sebagai acuan, tapi update hanya bekerja jika nama kolom ada di sheet.")
        matching_col = st.selectbox("Pilih kolom matching dari file upload", options=list(df_new.columns))
    else:
        matching_col = st.selectbox("Pilih kolom acuan untuk matching (kolom yang unik)", options=common_cols)

# Option: choose which columns to sync (subset)
st.markdown("**Pilih kolom yang ingin disinkronkan (kosong = semua kolom yang ada di file upload)**")
cols_to_sync = st.multiselect("Kolom dari file upload", options=list(df_new.columns), default=list(df_new.columns))

# 7. Dry-run preview: determine inserts, updates, unchanged
# Normalize matching values to string for safe comparison
key = matching_col

# Ensure key exists in both frames when possible
if key not in df_new.columns:
    st.error(f"Kolom matching '{key}' tidak ditemukan di file upload.")
    st.stop()

# For existing data, if sheet empty create empty df with same columns as df_new
if df_existing.empty:
    df_existing = pd.DataFrame(columns=df_new.columns)

# Coerce types to string for matching
df_existing['_match_key__'] = df_existing[key].astype(str).str.strip()
df_new['_match_key__'] = df_new[key].astype(str).str.strip()

# Build index mapping from existing keys to worksheet row numbers
# We need worksheet row numbers to do in-place updates. We'll read raw values to detect header row position.
raw = ws.get_all_values()
if not raw:
    st.error("Sheet target kosong (tidak ada header). Aborting.")
    st.stop()

header = raw[0]
# Map header columns order to indexes
header_map = {c: i for i, c in enumerate(header)}

# Build existing key -> row number mapping
existing_key_to_row = {}
for r_idx, row in enumerate(raw[1:], start=2):  # row 1 is header
    # try to find key column index
    if key in header_map:
        kcol = header_map[key]
        val = str(row[kcol]).strip() if kcol < len(row) else ""
        if val:
            existing_key_to_row[val] = r_idx

# Compare
new_keys = set(df_new['_match_key__'].dropna().unique())
existing_keys = set(existing_key_to_row.keys())

to_insert_keys = new_keys - existing_keys
to_update_keys = new_keys & existing_keys

# Build DataFrames for inserts and updates
df_to_insert = df_new[df_new['_match_key__'].isin(to_insert_keys)][cols_to_sync].copy()
# For updates: we will compare column-by-column to detect changed rows
update_rows = []  # tuples of (row_number, full_row_values)
unchanged_count = 0
changed_count = 0

# Prepare full row template according to header order. We'll try to fill columns present in header using new values or existing ones.
for _, new_row in df_new[df_new['_match_key__'].isin(to_update_keys)].iterrows():
    mk = new_row['_match_key__']
    sheet_row_num = existing_key_to_row.get(mk)
    if sheet_row_num is None:
        continue
    # get current row values from raw
    current_raw_row = raw[sheet_row_num - 1]
    updated_row = current_raw_row.copy()
    changed = False
    # For each column in cols_to_sync, place new value in the matching header column if exists; otherwise ignore
    for col in cols_to_sync:
        val_new = new_row.get(col)
        if col in header_map:
            col_idx = header_map[col]
            val_current = current_raw_row[col_idx] if col_idx < len(current_raw_row) else ""
            # compare as string after stripping
            sval_new = "" if pd.isna(val_new) else str(val_new).strip()
            sval_current = "" if val_current is None else str(val_current).strip()
            if sval_new != "" and sval_new != sval_current:
                updated_row[col_idx] = sval_new
                changed = True
        else:
            # header doesn't have this column; we'll ignore for update. Could optionally append new columns.
            pass
    if changed:
        changed_count += 1
        # ensure updated_row length matches header length
        if len(updated_row) < len(header):
            updated_row += [""] * (len(header) - len(updated_row))
        update_rows.append((sheet_row_num, updated_row))
    else:
        unchanged_count += 1

# Show dry-run summary
st.markdown("---")
st.write("### 🔎 Preview hasil (dry-run)")
st.write(f"Baris baru (insert): {len(df_to_insert)}")
st.write(f"Baris yang akan diupdate: {len(update_rows)}")
st.write(f"Baris tidak berubah: {unchanged_count}")

if not df_to_insert.empty:
    st.write("Contoh baris yang akan ditambahkan:")
    st.dataframe(df_to_insert.head())

if update_rows:
    st.write("Contoh baris yang akan diupdate (sheet row number, first 5 cols):")
    preview_updates = [ (rnum, row[:min(5, len(row))]) for rnum, row in update_rows[:10] ]
    st.write(preview_updates)

# 8. Execute
if st.button("🚀 Jalankan Upsert ke Google Sheets"):
    added = 0
    updated = 0
    errors = []
    # 8a. Inserts
    try:
        if not df_to_insert.empty:
            # Convert df_to_insert to match header columns order when possible. If header contains columns from file, align; otherwise append at end
            insert_rows_values = []
            for _, r in df_to_insert.iterrows():
                # Build row according to header
                rowvals = [""] * len(header)
                for col, val in r.items():
                    if col in header_map:
                        rowvals[header_map[col]] = "" if pd.isna(val) else val
                insert_rows_values.append(rowvals)
            # If header_map didn't include all cols from df, appended data will be empty for those columns in sheet
            # Append rows
            ws.append_rows([ ["" if v is None else v for v in rv] for rv in insert_rows_values ], value_input_option="USER_ENTERED")
            added = len(insert_rows_values)
    except Exception as e:
        errors.append(f"Insert error: {e}")

    # 8b. Updates
    try:
        for rownum, full_row in update_rows:
            try:
                # Ensure length matches header
                if len(full_row) < len(header):
                    full_row += [""] * (len(header) - len(full_row))
                # write
                # compute last column letter
                from string import ascii_uppercase
                def col_letter(n):
                    result = ""
                    while n:
                        n, r = divmod(n - 1, 26)
                        result = chr(65 + r) + result
                    return result
                last_col_letter = col_letter(len(full_row))
                range_str = f"A{rownum}:{last_col_letter}{rownum}"
                ws.update(range_str, [full_row], value_input_option="USER_ENTERED")
                updated += 1
            except Exception as e:
                errors.append(f"Update row {rownum} error: {e}")
    except Exception as e:
        errors.append(f"Update error: {e}")

    # 8c. Optionally log the operation
    log_msg = {
        "timestamp": datetime.utcnow().isoformat() + "Z",
        "spreadsheet_id": ss_id,
        "sheet": target_sheet,
        "matching_col": key,
        "added": added,
        "updated": updated,
        "errors": errors,
        "user_file": uploaded_file.name
    }
    if log_toggle:
        try:
            try:
                log_ws = spreadsheet.worksheet("__update_log__")
            except Exception:
                # create sheet
                log_ws = spreadsheet.add_worksheet(title="__update_log__", rows=1000, cols=10)
                log_ws.append_row(["timestamp","sheet","matching_col","added","updated","errors","user_file","raw_summary"]) 
            # append log
            log_ws.append_row([log_msg["timestamp"], target_sheet, key, str(added), str(updated), json.dumps(errors, ensure_ascii=False), uploaded_file.name, str({'rows_insert': len(df_to_insert)})])
        except Exception as e:
            st.warning(f"Gagal menulis log ke '__update_log__': {e}")

    # 9. Final summary
    st.success(f"Selesai. {added} baris ditambahkan, {updated} baris diupdate.")
    if errors:
        st.error("Beberapa error terjadi. Lihat detail di bawah.")
        for err in errors:
            st.write(err)

    st.experimental_rerun()

# Footer: notes
st.markdown("---")
st.write("**Catatan:**")
st.write("- Pastikan service account memiliki akses editor ke spreadsheet target.")
st.write("- Aplikasi ini menyesuaikan data berdasarkan header sheet. Jika file upload memiliki kolom baru yang tidak ada di header sheet, kolom tersebut tidak akan otomatis dibuat (kecuali jika sheet kosong). Jika ingin menambah kolom otomatis, beri tahu saya dan saya tambahkan fitur tersebut.")
st.write("- Backup spreadsheet sebelum menjalankan operasi pertama kali untuk menghindari kehilangan data.")
