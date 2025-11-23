import pandas as pd

# Baca file Excel
file_path = 'gambaran-umum-keadaan-sekolah-dasar-tiap-propinsi-indonesia-sd-2024.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1', header=2)

# Filter hanya kolom yang diperlukan
df = df[['Provinsi', 'Putus Sekolah', 'Status']]

# Hapus baris yang tidak memiliki nama provinsi
df = df.dropna(subset=['Provinsi'])

# Kelompokkan data berdasarkan provinsi dan status, jumlahkan putus sekolah
result = df.groupby(['Provinsi', 'Status'])['Putus Sekolah'].sum().reset_index()

# Pisahkan data negeri dan swasta
negeri = result[result['Status'] == 'Negeri'][['Provinsi', 'Putus Sekolah']]
swasta = result[result['Status'] == 'Swasta'][['Provinsi', 'Putus Sekolah']]

# Gabungkan data negeri dan swasta
merged = pd.merge(negeri, swasta, on='Provinsi', how='outer', suffixes=('_Negeri', '_Swasta'))

# Isi nilai NaN dengan 0
merged = merged.fillna(0)

# Hitung total putus sekolah per provinsi
merged['Total Putus Sekolah'] = merged['Putus Sekolah_Negeri'] + merged['Putus Sekolah_Swasta']

# Urutkan berdasarkan total putus sekolah tertinggi
merged = merged.sort_values('Total Putus Sekolah', ascending=False)

# Buat DataFrame untuk output
output_df = merged[['Provinsi', 'Putus Sekolah_Negeri', 'Putus Sekolah_Swasta', 'Total Putus Sekolah']]

# Simpan ke file Excel
output_file = 'total_putus_sekolah_per_provinsi.xlsx'
output_df.to_excel(output_file, index=False)

print("Data berhasil diproses dan disimpan ke:", output_file)
print("\n10 Provinsi dengan Putus Sekolah Tertinggi:")
print(output_df.head(10).to_string(index=False))

print(f"\nTotal keseluruhan anak putus sekolah: {output_df['Total Putus Sekolah'].sum():,}")