import pandas as pd

# Baca file Excel
file_path = r"gambaran-umum-keadaan-sekolah-dasar-tiap-propinsi-indonesia-sd-2024.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1")

# ===============================
# DATA CLEANING AWAL (sebelumnya)
# ===============================
# Hapus baris kosong
df = df.dropna(how='all')

# Cari dan set header yang benar
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

# Hapus baris metadata
df = df[~df['Provinsi'].str.contains('Tanggal cutoff', na=False)]
df = df[~df['Provinsi'].str.contains('Sumber', na=False)]
df = df[~df['Provinsi'].str.contains('Gambaran Umum', na=False)]
df = df.dropna(subset=['Provinsi'])

# ===============================
# HAPUS KOLOM YANG DITENTUKAN
# ===============================
# Daftar kolom yang akan dihapus
kolom_yang_dihapus = [
    'Kepala Sekolah & Guru', 
    'Tenaga Kependidikan', 
    'Rombel', 
    'Ruang Kelas'
]

# Cek apakah kolom tersebut ada dalam dataframe
print("Kolom sebelum dihapus:")
print(df.columns.tolist())
print(f"\nJumlah kolom sebelum dihapus: {len(df.columns)}")

# Hapus kolom yang ditentukan
df = df.drop(columns=kolom_yang_dihapus, errors='ignore')

# ===============================
# KONVERSI TIPE DATA & CLEANING LANJUTAN
# ===============================
# Konversi kolom numerik yang tersisa
numeric_columns = ['Sekolah', 'Siswa', 'Mengulang', 'Putus Sekolah']
for col in numeric_columns:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# Handling missing values
df[numeric_columns] = df[numeric_columns].fillna(0)
df['Status'] = df['Status'].fillna('Tidak Diketahui')

# Standardisasi nama provinsi
df['Provinsi'] = df['Provinsi'].str.replace('Prov\.', '', regex=True)
df['Provinsi'] = df['Provinsi'].str.strip()

# Reset index final
df = df.reset_index(drop=True)

# ===============================
# TAMPILKAN HASIL
# ===============================
print("\n" + "="*50)
print("HASIL SETELAH MENGHAPUS KOLOM")
print("="*50)

print(f"\nKolom setelah dihapus:")
print(df.columns.tolist())
print(f"Jumlah kolom setelah dihapus: {len(df.columns)}")

print(f"\nKolom yang dihapus: {kolom_yang_dihapus}")

print("\nStruktur data sekarang:")
print(df.info())

print("\nPreview data (5 baris pertama):")
print(df.head())

print(f"\nDimensi data: {df.shape}")  # (baris, kolom)

# ===============================
# SIMPAN DATA YANG SUDAH DIPERBAIKI
# ===============================
output_file = "data_sekolah_dasar_simplified_2024.xlsx"
df.to_excel(output_file, index=False)

print(f"\nData telah disimpan ke: {output_file}")

# ===============================
# STATISTIK DATA TERKINI
# ===============================
print("\n" + "="*50)
print("STATISTIK DATA TERKINI")
print("="*50)

print(f"\nTotal Provinsi: {df['Provinsi'].nunique()}")
print(f"Status Sekolah: {df['Status'].unique().tolist()}")

print("\nStatistik deskriptif:")
print(df[numeric_columns].describe())

# Hitung beberapa metrik penting
total_sekolah = df['Sekolah'].sum()
total_siswa = df['Siswa'].sum()
total_mengulang = df['Mengulang'].sum()
total_putus_sekolah = df['Putus Sekolah'].sum()

print(f"\nTOTAL NASIONAL:")
print(f"Jumlah Sekolah: {total_sekolah:,}")
print(f"Jumlah Siswa: {total_siswa:,}")
print(f"Jumlah Mengulang: {total_mengulang:,}")
print(f"Jumlah Putus Sekolah: {total_putus_sekolah:,}")