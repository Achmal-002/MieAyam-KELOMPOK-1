import pandas as pd
import numpy as np

# ===============================
# DATA CLEANING UNTUK FILE SEKOLAH DASAR
# ===============================

# Data dari file Excel yang diberikan
data = {
    'Provinsi': [
        'Prov. D.K.I. Jakarta', 'Prov. Jawa Barat', 'Prov. Jawa Tengah', 'Prov. D.I. Yogyakarta', 
        'Prov. Jawa Timur', 'Prov. Aceh', 'Prov. Sumatera Utara', 'Prov. Sumatera Barat', 
        'Prov. Riau', 'Prov. Jambi', 'Prov. Sumatera Selatan', 'Prov. Lampung', 
        'Prov. Kalimantan Barat', 'Prov. Kalimantan Tengah', 'Prov. Kalimantan Selatan', 
        'Prov. Kalimantan Timur', 'Prov. Sulawesi Utara', 'Prov. Sulawesi Tengah', 
        'Prov. Sulawesi Selatan', 'Prov. Sulawesi Tenggara', 'Prov. Maluku', 'Prov. Bali', 
        'Prov. Nusa Tenggara Barat', 'Prov. Nusa Tenggara Timur', 'Prov. Papua', 
        'Prov. Bengkulu', 'Prov. Maluku Utara', 'Prov. Banten', 'Prov. Kepulauan Bangka Belitung', 
        'Prov. Gorontalo', 'Prov. Kepulauan Riau', 'Prov. Papua Barat', 'Prov. Sulawesi Barat', 
        'Prov. Kalimantan Utara', 'Luar Negeri', 'Prov. Papua Tengah', 'Prov. Papua Selatan', 
        'Prov. Papua Pegunungan', 'Prov. Papua Barat Daya', 'Prov. D.K.I. Jakarta', 
        'Prov. Jawa Barat', 'Prov. Jawa Tengah', 'Prov. D.I. Yogyakarta', 'Prov. Jawa Timur', 
        'Prov. Aceh', 'Prov. Sumatera Utara', 'Prov. Sumatera Barat', 'Prov. Riau', 
        'Prov. Jambi', 'Prov. Sumatera Selatan', 'Prov. Lampung', 'Prov. Kalimantan Barat', 
        'Prov. Kalimantan Tengah', 'Prov. Kalimantan Selatan', 'Prov. Kalimantan Timur', 
        'Prov. Sulawesi Utara', 'Prov. Sulawesi Tengah', 'Prov. Sulawesi Selatan', 
        'Prov. Sulawesi Tenggara', 'Prov. Maluku', 'Prov. Bali', 'Prov. Nusa Tenggara Barat', 
        'Prov. Nusa Tenggara Timur', 'Prov. Papua', 'Prov. Bengkulu', 'Prov. Maluku Utara', 
        'Prov. Banten', 'Prov. Kepulauan Bangka Belitung', 'Prov. Gorontalo', 
        'Prov. Kepulauan Riau', 'Prov. Papua Barat', 'Prov. Sulawesi Barat', 
        'Prov. Kalimantan Utara', 'Luar Negeri', 'Prov. Papua Tengah', 'Prov. Papua Selatan', 
        'Prov. Papua Pegunungan', 'Prov. Papua Barat Daya'
    ],
    'Sekolah': [
        1305, 16980, 17254, 1418, 16864, 3336, 8116, 3892, 3232, 2302, 4250, 4295, 4139, 2412, 
        2737, 1664, 1349, 2676, 6071, 2265, 1299, 2251, 3008, 3415, 515, 1289, 1116, 3853, 761, 
        897, 683, 412, 1296, 437, 83, 295, 377, 518, 305, 918, 2558, 1343, 423, 2111, 207, 1631, 
        327, 605, 172, 489, 447, 316, 241, 186, 280, 861, 273, 381, 87, 537, 156, 344, 1832, 335, 
        119, 202, 799, 71, 36, 305, 162, 32, 49, 41, 265, 241, 160, 249
    ],
    'Siswa': [
        534084, 3892560, 2259455, 182927, 2097016, 453928, 1149616, 512533, 626689, 331425, 
        790920, 711692, 493992, 231066, 315349, 352693, 137150, 287380, 802533, 282647, 148643, 
        338502, 474323, 411297, 70293, 171467, 118093, 1020399, 142879, 103294, 153988, 48552, 
        138333, 70242, 12536, 71294, 57447, 112406, 36044, 212135, 620397, 319228, 83347, 404776, 
        40154, 346857, 73064, 146184, 36401, 92261, 86276, 64197, 45359, 35716, 70396, 82629, 
        28740, 76723, 14089, 55552, 43992, 43262, 242072, 47619, 23246, 22598, 202372, 16775, 
        5077, 73836, 22759, 5114, 10486, 6253, 73273, 39392, 34363, 24564
    ],
    'Mengulang': [
        911, 3221, 4511, 166, 2976, 881, 2722, 2949, 2733, 1303, 4199, 1411, 6421, 1859, 1690, 
        1113, 437, 1691, 1989, 1131, 822, 94, 875, 3635, 870, 875, 715, 1550, 597, 962, 369, 
        1010, 650, 211, 51, 1292, 2501, 5358, 898, 330, 215, 110, 26, 399, 37, 562, 117, 263, 
        59, 128, 69, 249, 132, 55, 78, 221, 44, 65, 6, 327, 9, 49, 2222, 425, 6, 70, 99, 3, 1, 
        52, 702, 3, 26, 19, 1733, 2550, 114, 354
    ],
    'Putus Sekolah': [
        277, 4681, 1185, 28, 1618, 695, 3030, 668, 824, 739, 1767, 696, 982, 370, 346, 489, 283, 
        812, 1411, 637, 905, 107, 944, 1737, 372, 336, 610, 1552, 209, 392, 154, 164, 293, 114, 
        35, 420, 486, 988, 205, 165, 508, 168, 26, 523, 87, 955, 38, 131, 42, 155, 66, 61, 63, 
        30, 61, 162, 53, 97, 28, 364, 29, 124, 1028, 297, 25, 85, 222, 6, 11, 70, 92, 1, 34, 7, 
        592, 290, 216, 67
    ],
    'Status': [
        'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 
        'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 
        'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 
        'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 
        'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Negeri', 'Swasta', 
        'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 
        'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 
        'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 
        'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 
        'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta', 'Swasta'
    ]
}

# Buat DataFrame
df = pd.DataFrame(data)

# ===============================
# DATA CLEANING
# ===============================

# 1. Hapus kolom yang tidak diperlukan (Kepala Sekolah & Guru, Tenaga Kependidikan, Rombel, Ruang Kelas)
# Kolom-kolom ini sudah tidak ada dalam data yang diberikan

# 2. Cleaning nama provinsi
df['Provinsi'] = df['Provinsi'].str.replace('Prov\.', '', regex=True)
df['Provinsi'] = df['Provinsi'].str.strip()

# 3. Konversi tipe data numerik
numeric_columns = ['Sekolah', 'Siswa', 'Mengulang', 'Putus Sekolah']
for col in numeric_columns:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# 4. Handling missing values (jika ada)
df[numeric_columns] = df[numeric_columns].fillna(0)

# 5. Reset index
df = df.reset_index(drop=True)

# ===============================
# TAMPILKAN HASIL
# ===============================
print("=" * 70)
print("DATA CLEANING SEKOLAH DASAR INDONESIA 2024")
print("=" * 70)

print(f"\nDimensi data: {df.shape}")
print(f"\nKolom yang tersisa: {df.columns.tolist()}")

print(f"\nTipe data:")
print(df.dtypes)

print(f"\nPreview data (10 baris pertama):")
print(df.head(10))

print(f"\nStatistik deskriptif:")
print(df[numeric_columns].describe())

# ===============================
# ANALISIS DATA
# ===============================
print("\n" + "=" * 70)
print("ANALISIS DATA SEKOLAH DASAR")
print("=" * 70)

# Total per status
status_summary = df.groupby('Status').agg({
    'Sekolah': 'sum',
    'Siswa': 'sum',
    'Mengulang': 'sum',
    'Putus Sekolah': 'sum'
}).reset_index()

print("\nRINGKASAN PER STATUS SEKOLAH:")
print(status_summary)

# Provinsi dengan siswa terbanyak
top_provinsi_siswa = df.groupby('Provinsi')['Siswa'].sum().sort_values(ascending=False).head(10)
print("\n10 PROVINSI DENGAN SISWA TERBANYAK:")
print(top_provinsi_siswa)

# Rasio putus sekolah
df_total = df.groupby('Provinsi').agg({
    'Siswa': 'sum',
    'Putus Sekolah': 'sum'
}).reset_index()
df_total['Rasio_Putus_Sekolah'] = (df_total['Putus Sekolah'] / df_total['Siswa'] * 100).round(4)

print("\n10 PROVINSI DENGAN RASIO PUTUS SEKOLAH TERTINGGI:")
print(df_total.nlargest(10, 'Rasio_Putus_Sekolah')[['Provinsi', 'Siswa', 'Putus Sekolah', 'Rasio_Putus_Sekolah']])

# ===============================
# SIMPAN DATA CLEAN
# ===============================
output_file = "data_sekolah_dasar_clean_2024.xlsx"
df.to_excel(output_file, index=False)

print(f"\n" + "=" * 70)
print(f"Data berhasil disimpan ke: {output_file}")
print("=" * 70)

# ===============================
# INFORMASI TAMBAHAN
# ===============================
print(f"\nINFORMASI DATA:")
print(f"Total Provinsi: {df['Provinsi'].nunique()}")
print(f"Total Sekolah: {df['Sekolah'].sum():,}")
print(f"Total Siswa: {df['Siswa'].sum():,}")
print(f"Total Mengulang: {df['Mengulang'].sum():,}")
print(f"Total Putus Sekolah: {df['Putus Sekolah'].sum():,}")

print(f"\nDistribusi Status:")
status_counts = df['Status'].value_counts()
print(status_counts)