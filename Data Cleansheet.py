import pandas as pd

# Baca file Excel
file_path = "gambaran-umum-keadaan-sekolah-dasar-tiap-propinsi-indonesia-sd-2024.xlsx"
df = pd.read_excel(file_path, sheet_name="Sheet1")

# ===============================
# 1. INSPEKSI DATA AWAL
# ===============================
print("Info Data Awal:")
print(df.info())
print("\n5 Data Teratas:")
print(df.head())
print("\nCek Nilai Null:")
print(df.isnull().sum())

# ===============================
# 2. HAPUS BARIS KOSONG DAN HEADER GANDA
# ===============================
# Hapus baris yang seluruhnya kosong
df = df.dropna(how='all')

# Reset index setelah menghapus baris
df = df.reset_index(drop=True)

# Cari baris yang berisi nama kolom sebenarnya
for idx, row in df.iterrows():
    if 'Provinsi' in str(row[0]):
        df.columns = df.iloc[idx]
        df = df.iloc[idx+1:].reset_index(drop=True)
        break

# ===============================
# 3. RENAME KOLOM UNTUK KONSISTENSI
# ===============================
df.columns = [
    'Provinsi', 'Sekolah', 'Siswa', 'Mengulang', 'Putus Sekolah', 
    'Kepala Sekolah & Guru', 'Tenaga Kependidikan', 'Rombel', 
    'Ruang Kelas', 'Status'
]

# ===============================
# 4. HAPUS BARIS YANG TIDAK DIPERLUKAN
# ===============================
# Hapus baris yang berisi metadata (seperti tanggal cutoff dan sumber)
df = df[~df['Provinsi'].str.contains('Tanggal cutoff', na=False)]
df = df[~df['Provinsi'].str.contains('Sumber', na=False)]
df = df[~df['Provinsi'].str.contains('Gambaran Umum', na=False)]

# Hapus baris kosong lagi setelah pembersihan
df = df.dropna(subset=['Provinsi'])

# ===============================
# 5. KONVERSI TIPE DATA
# ===============================
# Konversi kolom numerik (handle non-numeric values)
numeric_columns = ['Sekolah', 'Siswa', 'Mengulang', 'Putus Sekolah', 
                   'Kepala Sekolah & Guru', 'Tenaga Kependidikan', 
                   'Rombel', 'Ruang Kelas']

for col in numeric_columns:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# ===============================
# 6. HANDLING MISSING VALUES
# ===============================
# Isi missing values dengan 0 untuk kolom numerik
df[numeric_columns] = df[numeric_columns].fillna(0)

# Untuk kolom Status, isi dengan 'Tidak Diketahui' jika kosong
df['Status'] = df['Status'].fillna('Tidak Diketahui')

# ===============================
# 7. STANDARDISASI NAMA PROVINSI
# ===============================
# Hapus prefix "Prov." dan "Prov " untuk konsistensi
df['Provinsi'] = df['Provinsi'].str.replace('Prov\.', '', regex=True)
df['Provinsi'] = df['Provinsi'].str.replace('Prov ', '', regex=True)
df['Provinsi'] = df['Provinsi'].str.strip()

# ===============================
# 8. VALIDASI DATA
# ===============================
# Cek duplikat berdasarkan Provinsi dan Status
duplikat = df.duplicated(subset=['Provinsi', 'Status'], keep=False)
if duplikat.any():
    print(f"\nDitemukan {duplikat.sum()} baris duplikat")
    print(df[duplikat][['Provinsi', 'Status']])
else:
    print("\nTidak ada data duplikat")

# Cek konsistensi data numerik (tidak boleh negatif)
for col in numeric_columns:
    if (df[col] < 0).any():
        print(f"Peringatan: Nilai negatif ditemukan di kolom {col}")

# ===============================
# 9. SIMPAN DATA BERSIH
# ===============================
# Reset index final
df = df.reset_index(drop=True)

# Simpan ke file Excel baru
output_file = "data_sekolah_dasar_bersih_2024.xlsx"
df.to_excel(output_file, index=False)

# ===============================
# 10. TAMPILKAN INFORMASI DATA BERSIH
# ===============================
print("\n" + "="*50)
print("INFORMASI DATA BERSIH")
print("="*50)
print(f"Jumlah Baris: {len(df)}")
print(f"Jumlah Kolom: {len(df.columns)}")
print(f"Provinsi Unik: {df['Provinsi'].nunique()}")
print(f"Status Sekolah: {df['Status'].unique().tolist()}")

print("\nStatistik Deskriptif:")
print(df[numeric_columns].describe())

print(f"\nData bersih telah disimpan di: {output_file}")

# Tampilkan preview data bersih
print("\nPreview Data Bersih:")
print(df.head(10))                                                                                                      