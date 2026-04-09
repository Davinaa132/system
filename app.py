import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import math

st.set_page_config(page_title="Data System", layout="wide")

def init_connection():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
    client = gspread.authorize(creds)
    return client

def clean_and_map(df):
    # 1. Standarisasi Kolom Identitas
    df['Kode Judul'] = df['Kode Judul'].astype(str).str.strip()
    df['Angkatan'] = df['Angkatan'].astype(str).str.strip()

    # 2. Penanganan Tanggal
    df['Tgl Mulai'] = pd.to_datetime(df['Tgl Mulai'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['Tgl Selesai'] = pd.to_datetime(df['Tgl Selesai'], errors='coerce').dt.strftime('%Y-%m-%d')

    # 3. Buat Kode Unik
    df['Kode Unik'] = df['Kode Judul'] + "." + df['Angkatan']

    # 4. Urutan kolom sesuai Google Sheet 'Detail L1'
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

    # Reindex kolom agar urutan konsisten
    df_final = df.reindex(columns=cols_order)

    # 5. Konversi kolom skor menjadi angka (Numeric)
    # Dimulai dari kolom 'P.Isi' sampai terakhir
    idx_pisi = cols_order.index('P.Isi')
    cols_numeric = cols_order[idx_pisi:]
    for col in cols_numeric:
        df_final[col] = pd.to_numeric(df_final[col], errors='coerce')

    return df_final

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
                # Pastikan nama file dan worksheet tepat
                sheet = client.open("Copy of Monitoring Evaluasi Pembelajaran").worksheet("Detail L1")
                
                # 6. PEMBERSIHAN AKHIR (List Conversion & NaN Handling)
                # Mengubah dataframe menjadi list of lists
                raw_lists = df_processed.values.tolist()
                clean_lists = []
                
                for row in raw_lists:
                    clean_row = []
                    for val in row:
                        # Jika nilainya NaN (bukan angka), ubah jadi string kosong
                        if isinstance(val, float) and math.isnan(val):
                            clean_row.append("")
                        else:
                            clean_row.append(val)
                    clean_lists.append(clean_row)

                # 7. KIRIM DATA dengan USER_ENTERED (Anti-Tanda Petik)
                sheet.append_rows(clean_lists, value_input_option='USER_ENTERED')
                
                st.success(f"✅ Berhasil! {len(clean_lists)} baris ditambahkan ke Database.")
                st.balloons()
                
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
