import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import math

st.set_page_config(page_title="UPDL Jakarta - Advanced System", layout="wide")

def init_connection():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
    client = gspread.authorize(creds)
    return client

st.title("🚀 UPDL Jakarta Multi-Function System")

# --- 1. PILIH TUJUAN SHEET ---
target_sheet = st.sidebar.radio("Pilih Tujuan Sheet:", ["Detail L1", "Detail L2"])

uploaded_file = st.sidebar.file_uploader("Upload Excel PLN", type=["xlsx"])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file)
        
        # --- 2. FITUR PILIH KOLOM (Skenario 1) ---
        st.subheader("⚙️ Pengaturan Kolom & Filter")
        all_columns = df_raw.columns.tolist()
        selected_cols = st.multiselect("Pilih kolom yang ingin dimasukkan ke Google Sheets:", 
                                       all_columns, 
                                       default=all_columns[:5] if len(all_columns) > 5 else all_columns)

        if selected_cols:
            df_filtered = df_raw[selected_cols].copy()

            # --- 3. FITUR FILTER TANGGAL (Skenario 3) ---
            # Pastikan ada kolom tanggal untuk difilter
            date_col = st.selectbox("Pilih kolom referensi tanggal untuk filter:", selected_cols)
            
            # Ubah ke datetime agar bisa difilter
            df_filtered[date_col] = pd.to_datetime(df_filtered[date_col], errors='coerce')
            
            col1, col2 = st.columns(2)
            with col1:
                start_date = st.date_input("Rentang Mulai:", df_filtered[date_col].min())
            with col2:
                end_date = st.date_input("Rentang Selesai:", df_filtered[date_col].max())

            # Eksekusi Filter Tanggal
            mask = (df_filtered[date_col].dt.date >= start_date) & (df_filtered[date_col].dt.date <= end_date)
            df_final = df_filtered.loc[mask].copy()

            # Preview
            st.write(f"📊 Menampilkan {len(df_final)} baris hasil filter.")
            st.dataframe(df_final)

            # --- 4. TOMBOL KIRIM (Skenario 2) ---
            if st.button(f"Kirim Data ke sheet '{target_sheet}'"):
                with st.spinner('Proses pengiriman...'):
                    client = init_connection()
                    # Membuka sheet berdasarkan pilihan di sidebar
                    sheet = client.open("Copy of Monitoring Evaluasi Pembelajaran").worksheet(target_sheet)
                    
                    # Konversi ke format yang aman bagi JSON (menghapus NaN)
                    raw_lists = df_final.astype(str).values.tolist() # Gunakan str untuk keamanan multikolom
                    
                    sheet.append_rows(raw_lists, value_input_option='USER_ENTERED')
                    
                    st.success(f"✅ Berhasil! Data masuk ke '{target_sheet}'.")
                    st.balloons()
        else:
            st.warning("Silakan pilih minimal satu kolom.")

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
