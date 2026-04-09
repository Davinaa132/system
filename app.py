import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import math

st.set_page_config(page_title="UPDL Jakarta Data System", layout="wide")

# --- KONEKSI GOOGLE SHEETS ---
def init_connection():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    # Pastikan sudah setting st.secrets di Streamlit Cloud/lokal
    creds_dict = st.secrets["gcp_service_account"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    return client

# --- FUNGSI PEMROSESAN DATA ---
def process_data(df, target):
    # Bersihkan spasi dan standarisasi tipe data kunci
    df['Kode Judul'] = df['Kode Judul'].astype(str).str.strip()
    df['Angkatan'] = df['Angkatan'].astype(str).str.strip()
    
    # Membuat Kode Unik (Standard: Kode.Batch)
    df['Kode Unik'] = df['Kode Judul'] + "." + df['Angkatan']

    # Standarisasi format Tanggal agar seragam YYYY-MM-DD
    for col in ['Tgl Mulai', 'Tgl Selesai']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')

    if target == "Detail L1":
        cols_order = [
            'Kode Judul', 'Judul Pembelajaran', 'Bidang', 'Tgl Mulai', 
            'Tgl Selesai', 'Angkatan', 'Kode Unik', 'UPDL Penyelenggara',
            'Jenis Diklat', 'Strategi Pelaksana', 'P.Isi', 'P.Hadir',
            'Ins-Eng-1 of 2', 'Ins-Eng-2 of 2', 'Ins-Rel-1 of 2', 'Ins-Rel-2 of 2',
            'Ins-Sat-1 of 4', 'Ins-Sat-2 of 4', 'Ins-Sat-3 of 4', 'Ins-Sat-4 of 4',
            'Ins-Rat', 'Ins-Val', 'Mat-Eng-1 of 2', 'Mat-Eng-2 of 2', 
            'Mat-Rel-1 of 2', 'Mat-Rel-2 of 2', 'Mat-Sat-1 of 2', 'Mat-Sat-2 of 2', 
            'Mat-Rat', 'Mat-Val', 'Sarpras-Sas-1 of 5', 'Sarpras-Sas-2 of 5', 
            'Sarpras-Sas-3 of 5', 'Sarpras-Sas-4 of 5', 'Sarpras-Sas-5 of 5', 
            'Sarpras-Rat', 'Dig-Sas-1 of 5', 'Dig-Sas-2 of 5', 'Dig-Sas-3 of 5', 
            'Dig-Sas-4 of 5', 'Dig-Sas-5 of 5', 'Dig Rat'
        ]
    else:  
        cols_order = [
            'Kode Judul', 'Jenis Permintaan Diklat', 'UPDL Code', 'UPDL', 
            'Angkatan', 'Tgl Mulai', 'Tgl Selesai', 'Kode Unik', 
            'Jumlah Peserta Hadir', 'Jumlah Peserta Lulus', 'Jumlah Peserta Isi', 'Confidence Level', 'Commitment Level'
        ]

    # Reindex: Mengambil kolom sesuai urutan, jika tidak ada di Excel akan jadi NaN
    df_reindexed = df.reindex(columns=cols_order)

    # Konversi kolom angka agar tidak ada tanda petik (') di Google Sheets
    # Mencari kolom angka: dimulai dari 'P.Isi' (L1) atau 'Jumlah Peserta Hadir' (L2)
    start_numeric = 'P.Isi' if target == "Detail L1" else 'Jumlah Peserta Hadir'
    if start_numeric in cols_order:
        idx_start = cols_order.index(start_numeric)
        for col in cols_order[idx_start:]:
            df_reindexed[col] = pd.to_numeric(df_reindexed[col], errors='coerce')

    return df_reindexed

# --- INTERFACE STREAMLIT ---
st.title("🚀 UPDL Jakarta - Unified Data Integration")
st.markdown("Sistem integrasi data otomatis untuk monitoring evaluasi L1 & L2.")

# Sidebar untuk navigasi
st.sidebar.header("Konfigurasi")
target_sheet = st.sidebar.selectbox("Pilih Tujuan Sheet:", ["Detail L1", "Detail L2"])
uploaded_file = st.sidebar.file_uploader("Upload File Excel PLN", type=["xlsx"])

if uploaded_file:
    try:
        # Load Data
        df_raw = pd.read_excel(uploaded_file)
        
        # --- FITUR FILTER TANGGAL ---
        st.subheader(f"📅 Filter Data untuk {target_sheet}")
        col_date = st.selectbox("Pilih kolom referensi tanggal:", df_raw.columns.tolist())
        
        # Pastikan kolom referensi tanggal valid
        df_raw[col_date] = pd.to_datetime(df_raw[col_date], errors='coerce')
        
        c1, c2 = st.columns(2)
        with c1:
            start_d = st.date_input("Dari Tanggal:", df_raw[col_date].min())
        with c2:
            end_d = st.date_input("Sampai Tanggal:", df_raw[col_date].max())

        # Filter Berdasarkan Tanggal
        mask = (df_raw[col_date].dt.date >= start_d) & (df_raw[col_date].dt.date <= end_d)
        df_filtered = df_raw.loc[mask].copy()

        # Proses Data Sesuai Target
        df_final = process_data(df_filtered, target_sheet)

        st.write(f"Skenario terdeteksi: **{target_sheet}** | Jumlah data: **{len(df_final)} baris**")
        st.dataframe(df_final)

        # --- TOMBOL KIRIM ---
        if st.button(f"Kirim Data ke {target_sheet}"):
            with st.spinner('Menghubungkan ke Google Sheets...'):
                client = init_connection()
                # Sesuaikan nama Spreadsheet Utama Anda
                sheet = client.open("Copy of Monitoring Evaluasi Pembelajaran").worksheet(target_sheet)
                
                # Pembersihan manual NaN agar JSON Compliant (Anti Error NaN)
                raw_lists = df_final.values.tolist()
                clean_lists = []
                for row in raw_lists:
                    clean_row = [ "" if (isinstance(val, float) and math.isnan(val)) else val for val in row ]
                    clean_lists.append(clean_row)

                # Kirim data dengan USER_ENTERED agar angka murni & tanggal terbaca otomatis
                sheet.append_rows(clean_lists, value_input_option='USER_ENTERED')
                
                st.success(f"✅ Data Berhasil Masuk ke Tab {target_sheet}!")
                st.balloons()

    except Exception as e:
        st.error(f"Terjadi Kesalahan: {e}")
