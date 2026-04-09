import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import math

st.set_page_config(page_title="UPDL Jakarta - Dynamic System", layout="wide")

def init_connection():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_dict = st.secrets["gcp_service_account"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    return client

st.title("🚀 UPDL Jakarta - Dynamic Column System")

# --- SIDEBAR CONFIG ---
st.sidebar.header("Konfigurasi")
target_sheet = st.sidebar.selectbox("Pilih Tujuan Sheet:", ["Detail L1", "Detail L2"])
uploaded_file = st.sidebar.file_uploader("Upload File Excel PLN", type=["xlsx"])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file)
        
        # --- 1. FITUR PEMILIHAN KOLOM (DIKEMBALIKAN) ---
        st.subheader(f"🛠️ Pilih & Urutkan Kolom untuk {target_sheet}")
        st.info("Urutan pilihan Anda akan menentukan urutan kolom di Google Sheets.")
        
        all_cols = df_raw.columns.tolist()
        
        # Tambahkan opsi untuk membuat Kode Unik secara otomatis jika belum ada di Excel
        if st.checkbox("Tambahkan Kode Unik Otomatis (Kode Judul + Angkatan)"):
            if 'Kode Judul' in df_raw.columns and 'Angkatan' in df_raw.columns:
                df_raw['Kode Unik'] = df_raw['Kode Judul'].astype(str).str.strip() + "." + df_raw['Angkatan'].astype(str).str.strip()
                all_cols = df_raw.columns.tolist()

        selected_cols = st.multiselect("Pilih kolom yang akan dikirim (Urutkan sesuai kolom di Sheets):", 
                                       options=all_cols)

        if selected_cols:
            df_selected = df_raw[selected_cols].copy()

            # --- 2. FILTER TANGGAL ---
            st.divider()
            date_ref = st.selectbox("Pilih kolom referensi tanggal untuk filter:", selected_cols)
            df_selected[date_ref] = pd.to_datetime(df_selected[date_ref], errors='coerce')
            
            c1, c2 = st.columns(2)
            with c1:
                start_d = st.date_input("Mulai:", df_selected[date_ref].min())
            with c2:
                end_d = st.date_input("Selesai:", df_selected[date_ref].max())

            mask = (df_selected[date_ref].dt.date >= start_d) & (df_selected[date_ref].dt.date <= end_d)
            df_final = df_selected.loc[mask].copy()

            st.write(f"Preview Data ({len(df_final)} baris):")
            st.dataframe(df_final)

            # --- 3. PROSES PENGIRIMAN ---
            if st.button(f"Kirim {len(selected_cols)} Kolom ke {target_sheet}"):
                with st.spinner('Mengirim data...'):
                    client = init_connection()
                    sheet = client.open("Copy of Monitoring Evaluasi Pembelajaran").worksheet(target_sheet)
                    
                    # Konversi ke List & Pembersihan Nilai Non-JSON (NaN)
                    raw_lists = df_final.values.tolist()
                    clean_lists = []
                    for row in raw_lists:
                        # Membersihkan nilai NaN agar tidak error dan angka tetap angka
                        clean_row = [ "" if (isinstance(val, float) and math.isnan(val)) else val for val in row ]
                        clean_lists.append(clean_row)

                    # Kirim dengan USER_ENTERED agar angka tidak ada tanda petik (')
                    sheet.append_rows(clean_lists, value_input_option='USER_ENTERED')
                    
                    st.success("✅ Data berhasil masuk sesuai urutan kolom yang Anda pilih!")
                    st.balloons()
        else:
            st.warning("Silakan pilih kolom terlebih dahulu melalui menu di atas.")

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
