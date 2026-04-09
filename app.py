import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Data System", layout="wide")

def init_connection():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
    client = gspread.authorize(creds)
    return client

def clean_and_map(df):
    # 1. Pastikan kolom kunci ada dan bersih
    df['Kode Judul'] = df['Kode Judul'].astype(str).str.strip()
    df['Angkatan'] = df['Angkatan'].astype(str).str.strip()

    # 2. Penanganan Tanggal (Supaya tidak error JSON Timestamp)
    df['Tgl Mulai'] = pd.to_datetime(df['Tgl Mulai'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['Tgl Selesai'] = pd.to_datetime(df['Tgl Selesai'], errors='coerce').dt.strftime('%Y-%m-%d')

    # 3. Buat Kode Unik
    df['Kode Unik'] = df['Kode Judul'] + "." + df['Angkatan']

    # 4. Definisikan urutan kolom (INI ADALAH cols_order YANG TADI ERROR)
    cols_order = [
        'Kode Judul', 'Judul Pembelajaran', 'Bidang', 'Tgl Mulai', 
        'Tgl Selesai', 'Angkatan', 'Kode Unik', 'UPDL Penyelenggara',
        'Jenis Diklat', 'Strategi Pelaksana', 'P.Isi', 'P.Hadir',
        'Ins-Eng-1 of 2', 'Ins-Eng-2 of 2', 'Ins-Rel-1 of 2', 'Ins-Rel-2 of 2',
        'Ins-Sat-1 of 4', 'Ins-Sat-2 of 4', 'Ins-Sat-3 of 4', 'Ins-Sat-4 of 4',
        'Ins-Rat', 'Ins-Val',
        'Mat-Eng-1 of 2', 'Mat-Eng-2 of 2', 'Mat-Rel-1 of 2', 'Mat-Rel-2 of 2',
        'Mat-Sat-1 of 2', 'Mat-Sat-2 of 2', 'Mat-Rat', 'Mat-Val',
        'Sarpras-Sas-1 of 5', 'Sarpras-Sas-2 of 5', 'Sarpras-Sas-3 of 5',
        'Sarpras-Sas-4 of 5', 'Sarpras-Sas-5 of 5', 'Sarpras-Rat',
        'Dig-Sas-1 of 5', 'Dig-Sas-2 of 5', 'Dig-Sas-3 of 5', 'Dig-Sas-4 of 5',
        'Dig-Sas-5 of 5', 'Dig Rat'
    ]

    # 5. Reindex dan bersihkan NaN (Agar aman bagi API Google Sheets) dan LOGIKA ANTI-TANDA PETIK: Ubah kolom skor kembali ke angka (Float)
    # Gunakan cols_order yang sudah didefinisikan di atas
    df_final = df.reindex(columns=cols_order).fillna("")

    # Kita mulai dari kolom 'P.Isi' sampai kolom terakhir
    cols_numeric = cols_order[cols_order.index('P.Isi'):]
    for col in cols_numeric:
        # Ubah ke angka, jika gagal (kosong/nan) biarkan tetap NaN (bukan "")
        df_final[col] = pd.to_numeric(df_final[col], errors='coerce')
    
    # Pastikan semua data dikirim sebagai string
    return df_final.astype(str)


st.title("🚀 UPDL Jakarta Data Integration")
st.markdown("Upload file Excel dari PLN Pusat untuk memperbarui Database Google Sheets.")

uploaded_file = st.file_uploader("Pilih file Excel (xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df_processed = clean_and_map(df_raw)
        
        st.subheader("Preview Data yang akan di-upload:")
        st.dataframe(df_processed)

        if st.button("Kirim ke Google Sheets"):
            with st.spinner('Sedang mengirim data...'):
                client = init_connection()
                sheet = client.open("Copy of Monitoring Evaluasi Pembelajaran").worksheet("Detail L1")
                
                data_to_push = df_processed.values.tolist()
                sheet.append_rows(data_to_push)
                
                st.success(f"✅ Berhasil! {len(data_to_push)} baris ditambahkan ke Database.")
                st.balloons()
                
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
