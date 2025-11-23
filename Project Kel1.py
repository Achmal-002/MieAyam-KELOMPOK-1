import pandas as pd
import os

# ===============================
# DATA CLEANING SEKOLAH DASAR
# ===============================

# Path file Excel
file_path = "gambaran-umum-keadaan-sekolah-dasar-tiap-propinsi-indonesia-sd-2024.xlsx"

# Baca file Excel
df = pd.read_excel(file_path, sheet_name="Sheet1")

# ===============================
# DATA CLEANING
# ===============================

# Hapus baris kosong dan metadata
df = df.dropna(how='all')

# Cari header yang benar (baris yang mengandung 'Provinsi')
for idx, row in df.iterrows():
    if 'Provinsi' in str(row[0]):
        df.columns = df.iloc[idx]
        df = df.iloc[idx+1:].reset_index(drop=True)
        break

# Rename kolom
df.columns = [
    'Provinsi', 'Sekolah', 'Siswa', 'Mengulang', 'Putus Sekolah', 
    'Kepala Sekolah & Guru', 'Tenaga Kependidikan', 'Rombel', 
    'Ruang Kelas', 'Status'
]

# Hapus kolom yang tidak diperlukan
kolom_dihapus = ['Kepala Sekolah & Guru', 'Tenaga Kependidikan', 'Rombel', 'Ruang Kelas']
df = df.drop(columns=kolom_dihapus, errors='ignore')

# Hapus baris metadata
df = df[~df['Provinsi'].str.contains('Tanggal cutoff|Sumber|Gambaran Umum', na=False)]
df = df.dropna(subset=['Provinsi'])

# Cleaning data
df['Provinsi'] = df['Provinsi'].str.replace(r'Prov\.', '', regex=True).str.strip()
numeric_cols = ['Sekolah', 'Siswa', 'Mengulang', 'Putus Sekolah']
df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

# ===============================
# HASIL
# ===============================
print("DATA CLEANING SELESAI")
print(f"Dimensi: {df.shape}")
print(f"Kolom: {df.columns.tolist()}")
print(f"\nPreview:\n{df.head()}")

# Simpan hasil
output_file = "data_sekolah_dasar_clean_2024.xlsx"
df.to_excel(output_file, index=False)
print(f"\nData disimpan: {output_file}")