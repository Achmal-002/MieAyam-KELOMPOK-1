import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# ==============================================================================
# KONFIGURASI FILE
# ==============================================================================
FILE_INPUT = "gambaran-umum-keadaan-sekolah-dasar-tiap-propinsi-indonesia-sd-2024.xlsx"
FILE_OUTPUT_CLEAN = "data_sekolah_dasar_clean_2024.xlsx"

# ==============================================================================
# DATA CLEANING SEKOLAH DASAR (VERSI SIMPLE)
# ==============================================================================

def clean_and_prepare_data(file_path):
    """Melakukan pembersihan dan perhitungan rasio data."""
    print("⏳ Memulai pembersihan data...")
    
    try:
        # Baca file Excel
        df = pd.read_excel(file_path, sheet_name="Sheet1")
    except FileNotFoundError:
        print(f"\n❌ ERROR: File input '{file_path}' tidak ditemukan.")
        print("Pastikan file Excel ada di folder yang sama dengan skrip ini.")
        return None

    # Hapus baris kosong
    df = df.dropna(how='all')

    # Cari header yang benar (baris yang mengandung 'Provinsi')
    header_found = False
    for idx, row in df.iterrows():
        if 'Provinsi' in str(row[0]):
            df.columns = df.iloc[idx]
            df = df.iloc[idx+1:].reset_index(drop=True)
            header_found = True
            break
    
    if not header_found:
        print("❌ ERROR: Header 'Provinsi' tidak ditemukan.")
        return None

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
    
    # Konversi kolom numerik
    numeric_cols = ['Sekolah', 'Siswa', 'Mengulang', 'Putus Sekolah']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Hitung rasio
    df['Rasio Mengulang (%)'] = (df['Mengulang'] / df['Siswa'] * 100).round(3)
    df['Rasio Putus Sekolah (%)'] = (df['Putus Sekolah'] / df['Siswa'] * 100).round(3)
    
    # Kategorikan berdasarkan rasio putus sekolah
    conditions = [
        df['Rasio Putus Sekolah (%)'] <= 0.1,
        (df['Rasio Putus Sekolah (%)'] > 0.1) & (df['Rasio Putus Sekolah (%)'] <= 0.5),
        df['Rasio Putus Sekolah (%)'] > 0.5
    ]
    choices = ['Rendah', 'Sedang', 'Tinggi']
    df['Tingkat Putus Sekolah'] = np.select(conditions, choices, default='Tidak Terdefinisi')
    
    print("✅ Data berhasil dibersihkan!")
    return df

def create_pie_chart(df):
    """Membuat pie chart untuk distribusi tingkat putus sekolah"""
    print("📊 Membuat pie chart...")
    
    # Hitung jumlah provinsi per kategori
    kategori_counts = df['Tingkat Putus Sekolah'].value_counts()
    
    # Warna untuk setiap kategori
    colors = ['#2ecc71', '#f39c12', '#e74c3c']  # Hijau, Orange, Merah
    
    # Buat pie chart
    plt.figure(figsize=(10, 8))
    
    # Plot pie chart dengan pengaturan yang lebih jelas
    wedges, texts, autotexts = plt.pie(
        kategori_counts.values, 
        labels=kategori_counts.index,
        colors=colors[:len(kategori_counts)],
        autopct='%1.1f%%',
        startangle=90,
        textprops={'fontsize': 12, 'weight': 'bold'}
    )
    
    # Perbaiki tampilan persentase
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
        autotext.set_fontsize(11)
    
    # Tambahkan judul
    plt.title('DISTRIBUSI TINGKAT PUTUS SEKOLAH SD PER PROVINSI (2024)', 
              fontsize=14, fontweight='bold', pad=20)
    
    # Tambahkan legenda
    plt.legend(wedges, [f'{label}: {value} provinsi' 
                       for label, value in kategori_counts.items()],
              title="Kategori:",
              loc="center left",
              bbox_to_anchor=(1, 0, 0.5, 1))
    
    # Simpan gambar
    plt.tight_layout()
    plt.savefig('pie_chart_tingkat_putus_sekolah.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    print("✅ Pie chart berhasil disimpan sebagai 'pie_chart_tingkat_putus_sekolah.png'")

def main():
    """Fungsi utama"""
    # Bersihkan data
    df_clean = clean_and_prepare_data(FILE_INPUT)
    
    if df_clean is not None:
        # Tampilkan preview data
        print("\n📋 Preview Data Hasil Cleaning:")
        print(df_clean.head())
        
        # Tampilkan statistik kategori
        print("\n📊 Distribusi Tingkat Putus Sekolah:")
        print(df_clean['Tingkat Putus Sekolah'].value_counts())
        
        # Buat pie chart
        create_pie_chart(df_clean)
        
        # Simpan ke Excel
        try:
            # Buat worksheet dengan data lengkap
            with pd.ExcelWriter(FILE_OUTPUT_CLEAN, engine='openpyxl') as writer:
                # Sheet 1: Data lengkap
                df_clean.to_excel(writer, sheet_name='Data Lengkap', index=False)
                
                # Sheet 2: Ringkasan per kategori
                ringkasan = df_clean.groupby('Tingkat Putus Sekolah').agg({
                    'Provinsi': 'count',
                    'Rasio Putus Sekolah (%)': ['min', 'max', 'mean'],
                    'Siswa': 'sum'
                }).round(2)
                ringkasan.columns = ['Jumlah Provinsi', 'Rasio Min (%)', 'Rasio Max (%)', 'Rasio Rata-rata (%)', 'Total Siswa']
                ringkasan.to_excel(writer, sheet_name='Ringkasan Kategori')
                
                # Sheet 3: Data per kategori (Rendah, Sedang, Tinggi)
                for kategori in ['Rendah', 'Sedang', 'Tinggi']:
                    df_kategori = df_clean[df_clean['Tingkat Putus Sekolah'] == kategori]
                    if not df_kategori.empty:
                        df_kategori.to_excel(writer, sheet_name=f'Data {kategori}', index=False)
            
            print(f"\n💾 Data berhasil disimpan ke: {FILE_OUTPUT_CLEAN}")
            print("📁 File Excel berisi:")
            print("   - Sheet 'Data Lengkap': Semua data provinsi")
            print("   - Sheet 'Ringkasan Kategori': Statistik per kategori")
            print("   - Sheet 'Data Rendah/Sedang/Tinggi': Data terpisah per kategori")
            
        except Exception as e:
            print(f"❌ Error saat menyimpan file Excel: {e}")
    
    else:
        print("❌ Gagal memproses data.")

# Jalankan program
if _name_ == "_main_":
    main()
    # Selesai