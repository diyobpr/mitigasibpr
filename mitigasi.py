import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from io import BytesIO

# Function to format numerical values
def format_angka(x):
    if isinstance(x, (int, float)) and not pd.isnull(x):
        return '{:.0f}'.format(x)
    return x

def process_bpr_data(df_bpr):
    # Daftar kolom yang diharapkan
    required_columns = ['_KOLEK', 'ACCNODR', 'NOREKENING','NAMA','NAMA_INST','PETUGAS','ALAMAT','TGL_BUKA','TGL_JT','PLAFOND','BAKIDEBET',
                        'ANGSURPK', 'ANGSURBNG','TGKPOKOK','TGKBUNGA','HR_TGKP','HR_TGKB','PPAP','CADBUNGA','DENDA']
    
    # Memeriksa apakah kolom yang diperlukan ada di DataFrame
    missing_columns = [col for col in required_columns if col not in df_bpr.columns]
    if missing_columns:
        raise ValueError(f"Missing columns in the input DataFrame: {missing_columns}")

    # Mengambil hanya kolom-kolom yang spesifik
    bpr_data = df_bpr[required_columns]

    # Menghapus baris dengan nilai kosong di kolom 'NAMA'
    bpr_data = bpr_data.dropna(subset=['NAMA'])

    # Menambahkan kolom baru 'jumlah_angsuran' yang merupakan hasil penjumlahan dari kolom 'ANGSURPK' dan 'ANGSURBNG'
    bpr_data['jumlah_angsuran'] = bpr_data['ANGSURPK'] + bpr_data['ANGSURBNG']

    # Konversi kolom-kolom yang relevan ke tipe numerik
    kolom_numerik = ['ANGSURPK', 'ANGSURBNG', 'jumlah_angsuran', 'TGKPOKOK', 'TGKBUNGA']
    bpr_data[kolom_numerik] = bpr_data[kolom_numerik].apply(pd.to_numeric, errors='coerce')

    # Menghapus baris di mana semua nilai adalah NaN
    bpr_data = bpr_data.dropna(how='all')

    # Mengisi nilai NaN di kolom 'TGKPOKOK' dan 'TGKBUNGA' dengan 0
    bpr_data['TGKPOKOK'] = bpr_data['TGKPOKOK'].fillna(0)
    bpr_data['TGKBUNGA'] = bpr_data['TGKBUNGA'].fillna(0)

    # Format angka dalam kolom-kolom numerik
    kolom_angka = [kolom for kolom in bpr_data.columns if kolom != 'NOREKENING']
    bpr_data[kolom_angka] = bpr_data[kolom_angka].applymap(format_angka)

    # Mengurutkan data berdasarkan kolom '_KOLEK'
    bpr_data = bpr_data.sort_values(by='_KOLEK')

    # Definisikan kriteria perubahan warna
    kriteria = [
        {'kolek': 1, 'tgkp_threshold': 25, 'tgkb_threshold': 25},
        {'kolek': 2, 'tgkp_threshold': 85, 'tgkb_threshold': 85},
        {'kolek': 3, 'tgkp_threshold': 175, 'tgkb_threshold': 175},
        {'kolek': 4, 'tgkp_threshold': 355, 'tgkb_threshold': 355}
    ]

    # Tambahkan kolom 'CEK DATA' dengan nilai default
    bpr_data['CEK DATA'] = ""

    # Tandai baris berdasarkan kriteria
    for kriteria_item in kriteria:
        kolek = kriteria_item['kolek']
        tgkp_threshold = kriteria_item['tgkp_threshold']
        tgkb_threshold = kriteria_item['tgkb_threshold']

        mask = (bpr_data['_KOLEK'] == kolek) | \
               (pd.to_numeric(bpr_data['HR_TGKP'], errors='coerce') >= tgkp_threshold) | \
               (pd.to_numeric(bpr_data['HR_TGKB'], errors='coerce') >= tgkb_threshold)

        bpr_data.loc[mask, 'NAMA_INST'] = bpr_data.loc[mask, 'NAMA_INST'].str.replace('BPR', '')
        bpr_data.loc[mask, 'CEK DATA'] = 'HARUS DI CEK, AKAN MELEBIHI JATUH TEMPO'

    return bpr_data

# Function to plot percentage of each '_KOLEK'
def plot_kolektibilitas(df):
    sns.set(style="darkgrid")
    plt.figure(figsize=(8, 6))
    df['_KOLEK'].value_counts().plot(kind='bar')
    plt.xlabel('Kolektibilitas')
    plt.xticks(rotation=0)
    plt.ylabel('Jumlah')
    plt.title('Persentase Kolektibilitas')
    st.pyplot()

# Function to download processed data as Excel
def download_excel(df):
    # Create a unique file name based on the current date
    tanggal_sekarang = pd.to_datetime('today').strftime("%Y%m%d")
    nama_file = f"{tanggal_sekarang}_MitigasiCibadak.xlsx"

    # Save DataFrame to Excel using BytesIO as Excel writer
    excel_data = BytesIO()
    with pd.ExcelWriter(excel_data, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)

    # Display the download button
    st.download_button(label="Download Excel", key="download_excel", data=excel_data, file_name=nama_file,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Streamlit app
def main():
    # Set Streamlit app title and description
    st.title("Aplikasi Proses Data Mitigasi BPR Cibadak")
    st.write("Unggah file Excel dan aplikasi akan memproses data serta menampilkan persentase kolektibilitas.")

    # Upload Excel file using Streamlit file uploader
    uploaded_file = st.file_uploader("Unggah file Excel", type=["xlsx"])

    # Process and display the data if a file is uploaded
    if uploaded_file is not None:
        # Read the Excel file into a DataFrame
        df_bpr = pd.read_excel(uploaded_file)

        # Process BPR data using the defined function
        processed_data = process_bpr_data(df_bpr)

        # Display the processed DataFrame
        st.write("DataFrame Hasil Pemrosesan Data:")
        st.write(processed_data)

        # Add a button to download the processed data as an Excel file
        download_excel(processed_data)

# Run the Streamlit app
if __name__ == "__main__":
    main()
