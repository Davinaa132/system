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
    # 1. Pastikan kolom kunci ada dan bersih dari spasi gaib
    # Di gambar kamu kolomnya adalah 'Kode Judul' dan 'Angkatan'
    df['Kode Judul'] = df['Kode Judul'].astype(str).str.strip()
    df['Angkatan'] = df['Angkatan'].astype(str).str.strip()

    # Convert date columns to datetime first, then to string format
    # Use .apply with pd.notna to handle NaT values correctly
    df['Tgl Mulai'] = pd.to_datetime(df['Tgl Mulai'], errors='coerce').apply(
        lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else None
    )
    df['Tgl Selesai'] = pd.to_datetime(df['Tgl Selesai'], errors='coerce').apply(
        lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else None
    )

    # 2. Buat Kode Unik (Gunakan variabel yang sudah dibersihkan)
    df['Kode Unik'] = df['Kode Judul'] + "." + df['Angkatan']

    # 3. Ambil kolom yang dibutuhkan (Pastikan 'Kode Unik' ada di urutan pertama)
    # Urutan ini harus sama persis dengan urutan kolom di Google Sheets 'detail L1'
    df_final = df[[
        'Kode Judul',
        'Judul Pembelajaran',
        'Bidang',
        'Tgl Mulai',
        'Tgl Selesai',
        'Angkatan',
        'Kode Unik',
        'UPDL Penyelenggara',
        'Jenis Diklat',
        'Strategi Pelaksana',
        'P.Isi',
        'P.Hadir',
        'Ins-Eng-1 of 2', 'Ins-Eng-2 of 2', 'Ins-Rel-1 of 2', 'Ins-Rel-2 of 2',
        'Ins-Sat-1 of 4', 'Ins-Sat-2 of 4', 'Ins-Sat-3 of 4', 'Ins-Sat-4 of 4',
        'Ins-Rat', 'Ins-Val',
        'Mat-Eng-1 of 2', 'Mat-Eng-2 of 2', 'Mat-Rel-1 of 2', 'Mat-Rel-2 of 2',
        'Mat-Sat-1 of 2', 'Mat-Sat-2 of 2', 'Mat-Rat', 'Mat-Val',
        'Sarpras-Sas-1 of 5', 'Sarpras-Sas-2 of 5', 'Sarpras-Sas-3 of 5',
        'Sarpras-Sas-4 of 5', 'Sarpras-Sas-5 of 5', 'Sarpras-Rat',
        'Dig-Sas-1 of 5', 'Dig-Sas-2 of 5', 'Dig-Sas-3 of 5', 'Dig-Sas-4 of 5',
        'Dig-Sas-5 of 5', 'Dig Rat'
    ]]
    df_final = df.reindex(columns=cols_order).fillna("")
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
                sheet = client.open("Monitoring Evaluasi Pembelajaran").worksheet("Detail L1")
                
                data_to_push = df_processed.values.tolist()
                sheet.append_rows(data_to_push)
                
                st.success(f"✅ Berhasil! {len(data_to_push)} baris ditambahkan ke Database.")
                st.balloons()
                
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
