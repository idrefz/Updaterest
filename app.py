import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ==============================================
# AUTH GOOGLE SHEETS
# ==============================================
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
client = gspread.authorize(creds)

st.title("📊 Upload & Update Google Sheets (Upsert Mode)")

# Input spreadsheet
spreadsheet_name = st.text_input("Nama Spreadsheet Google Sheets", "DATA SIIS")

try:
    spreadsheet = client.open(spreadsheet_name)
    sheet_list = [ws.title for ws in spreadsheet.worksheets()]
except Exception as e:
    st.error(f"❌ Gagal membuka spreadsheet: {e}")
    st.stop()

# Pilih sheet target
target_sheet = st.selectbox("Pilih Sheet Target", sheet_list)

# Upload file
uploaded_file = st.file_uploader("Upload file Excel/CSV", type=["xlsx", "csv"])

if uploaded_file and target_sheet:
    # ==============================================
    # Baca file upload
    # ==============================================
    if uploaded_file.name.endswith(".csv"):
        data_new = pd.read_csv(uploaded_file)
    else:
        data_new = pd.read_excel(uploaded_file)

    st.write("📂 **Preview Data Baru:**")
    st.dataframe(data_new.head())

    # ==============================================
    # Baca data existing dari Google Sheet
    # ==============================================
    worksheet = spreadsheet.worksheet(target_sheet)
    data_existing = pd.DataFrame(worksheet.get_all_records())

    st.write(f"📊 **Preview Data Existing ({target_sheet}):**")
    st.dataframe(data_existing.head())

    if data_existing.empty:
        st.warning("⚠ Sheet kosong. Semua data akan ditambahkan.")
        matching_column = None
    else:
        # Pilih kolom acuan untuk matching
        common_columns = list(set(data_new.columns).intersection(set(data_existing.columns)))
        matching_column = st.selectbox(
            "Pilih kolom acuan untuk matching (ID unik)",
            options=common_columns
        )

    if matching_column or data_existing.empty:
        # ==============================================
        # Proses UPSERT
        # ==============================================
        if data_existing.empty:
            # Jika sheet kosong → langsung tambah semua data
            rows_to_add = data_new.values.tolist()
            worksheet.append_rows(rows_to_add)
            st.success(f"✅ Sheet kosong. {len(data_new)} baris ditambahkan.")
        else:
            df_existing = data_existing.copy()
            df_new = data_new.copy()

            # Gabungkan data lama & baru berdasarkan kolom matching
            df_combined = pd.merge(
                df_existing,
                df_new,
                on=matching_column,
                how="outer",
                suffixes=("_old", "_new"),
                indicator=True
            )

            # Baris untuk UPDATE → data yang ada di kedua tabel, tapi ada perbedaan di kolom lain
            update_rows = []
            update_indices = []

            for idx, row in df_combined.iterrows():
                if row["_merge"] == "both":
                    # Cek apakah ada perubahan selain kolom matching
                    changed = False
                    updated_values = []
                    for col in df_existing.columns:
                        if col == matching_column:
                            updated_values.append(row[matching_column])
                        else:
                            old_val = row[f"{col}_old"]
                            new_val = row[f"{col}_new"]
                            # Update jika nilai baru berbeda dan tidak NaN
                            if pd.notna(new_val) and old_val != new_val:
                                changed = True
                                updated_values.append(new_val)
                            else:
                                updated_values.append(old_val)
                    if changed:
                        update_rows.append(updated_values)
                        update_indices.append(idx)

            # Baris untuk INSERT → hanya ada di file baru
            insert_rows = df_combined[df_combined["_merge"] == "right_only"]
            insert_rows = insert_rows[df_existing.columns].values.tolist()

            # ==============================================
            # Eksekusi Update
            # ==============================================
            if st.button("🚀 Update ke Google Sheets"):
                # 1. Update baris yang berubah
                if update_rows:
                    for row in update_rows:
                        match_value = row[0]  # nilai dari kolom matching
                        # Cari index baris di sheet
                        cell = worksheet.find(str(match_value))
                        if cell:
                            worksheet.update(f"A{cell.row}:Z{cell.row}", [row])
                    st.success(f"🔄 {len(update_rows)} baris diupdate.")

                # 2. Insert baris baru
                if insert_rows:
                    worksheet.append_rows(insert_rows)
                    st.success(f"➕ {len(insert_rows)} baris baru ditambahkan.")

                if not update_rows and not insert_rows:
                    st.info("Tidak ada perubahan data yang perlu diupdate.")
