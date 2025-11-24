import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import FancyBboxPatch # Diperlukan untuk Slide 4

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
        # Asumsi header data sebenarnya berada di baris ke-X, 
        # kita akan mencari 'Provinsi' untuk menentukan header yang benar.
        df = pd.read_excel(file_path, sheet_name="Sheet1", header=None)
    except FileNotFoundError:
        print(f"\n❌ ERROR: File input '{file_path}' tidak ditemukan.")
        print("Pastikan file Excel ada di folder yang sama dengan skrip ini.")
        return None

    # Hapus baris kosong
    df = df.dropna(how='all')

    # Cari header yang benar (baris yang mengandung 'Provinsi')
    header_found = False
    for idx, row in df.iterrows():
        # Asumsi kolom pertama (indeks 0) yang berisi 'Provinsi'
        if isinstance(row[0], str) and 'Provinsi' in row[0]:
            df.columns = df.iloc[idx]
            df = df.iloc[idx+1:].reset_index(drop=True)
            header_found = True
            break
    
    if not header_found:
        print("❌ ERROR: Header 'Provinsi' tidak ditemukan. Cek struktur file Excel Anda.")
        return None

    # Rename kolom
    # Catatan: Jumlah kolom harus sesuai dengan data yang Anda ambil setelah header
    try:
        df.columns = [
            'Provinsi', 'Sekolah', 'Siswa', 'Mengulang', 'Putus Sekolah', 
            'Kepala Sekolah & Guru', 'Tenaga Kependidikan', 'Rombel', 
            'Ruang Kelas', 'Status'
        ]
    except ValueError:
        print("❌ ERROR: Jumlah kolom yang terdeteksi tidak sesuai (harap periksa file Excel).")
        print(f"Jumlah kolom yang terdeteksi: {len(df.columns)}")
        return None

    # Hapus kolom yang tidak diperlukan
    kolom_dihapus = ['Kepala Sekolah & Guru', 'Tenaga Kependidikan', 'Rombel', 'Ruang Kelas']
    df = df.drop(columns=kolom_dihapus, errors='ignore')

    # Hapus baris metadata atau total (jika ada)
    df = df[~df['Provinsi'].astype(str).str.contains('Tanggal cutoff|Sumber|Gambaran Umum|TOTAL', na=False)]
    df = df.dropna(subset=['Provinsi'])

    # Cleaning data
    df['Provinsi'] = df['Provinsi'].astype(str).str.replace(r'Prov\.', '', regex=True).str.strip()
    
    # Konversi kolom numerik
    numeric_cols = ['Sekolah', 'Siswa', 'Mengulang', 'Putus Sekolah']
    for col in numeric_cols:
        # Konversi ke numerik, ganti nilai non-numerik dengan NaN, lalu isi NaN dengan 0
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
    # Hapus baris yang memiliki total siswa 0 untuk menghindari ZeroDivisionError pada ratio
    df = df[df['Siswa'] > 0].copy()

# ===============================
    # TAMBAH KOLOM RATIO MENGULANG DAN PUTUS SEKOLAH
    # ===============================
    
    # Hitung ratio mengulang dalam persen (Mengulang / Siswa * 100)
    df['Ratio_Mengulang_Num'] = (df['Mengulang'] / df['Siswa'] * 100).round(2)
    
    # Hitung ratio putus sekolah dalam persen (Putus Sekolah / Siswa * 100)
    df['Ratio_Putus_Num'] = (df['Putus Sekolah'] / df['Siswa'] * 100).round(2)
    
    # Format sebagai string dengan simbol %
    df['Ratio Mengulang (%)'] = df['Ratio_Mengulang_Num'].astype(str) + '%'
    df['Ratio Putus Sekolah (%)'] = df['Ratio_Putus_Num'].astype(str) + '%'
    
    # Reorder kolom untuk menempatkan ratio mengulang di samping kolom mengulang
    kolom_akhir = df.columns.tolist()
    
    # Pindahkan 'Ratio Mengulang (%)' ke setelah 'Mengulang'
    if 'Mengulang' in kolom_akhir and 'Ratio Mengulang (%)' in kolom_akhir:
        idx_mengulang = kolom_akhir.index('Mengulang')
        kolom_akhir.insert(idx_mengulang + 1, kolom_akhir.pop(kolom_akhir.index('Ratio Mengulang (%)')))
    
    # Pindahkan 'Ratio Putus Sekolah (%)' ke setelah 'Putus Sekolah'
    if 'Putus Sekolah' in kolom_akhir and 'Ratio Putus Sekolah (%)' in kolom_akhir:
        idx_putus = kolom_akhir.index('Putus Sekolah')
        kolom_akhir.insert(idx_putus + 1, kolom_akhir.pop(kolom_akhir.index('Ratio Putus Sekolah (%)')))
    
    df = df[kolom_akhir]
    
    print("Pembersihan data selesai & Penambahan Ratio Mengulang & Putus Sekolah.")
    print(f"Dimensi data: {df.shape}")
    print(f"Kolom: {df.columns.tolist()}")
    
    # Pastikan return df berada di dalam fungsi
    return df

# ==============================================================================
# FUNGSI VISUALISASI PER SLIDE
# ==============================================================================

def slide_1_distribusi_mengulang(df):
    """Membuat slide 1: Distribusi Rasio Mengulang"""
    fig, ax = plt.subplots(figsize=(10, 8))
    
    # Kategorisasi data mengulang
    mengulang_categories = {
        'Rendah (<1%)': (df['Ratio_Mengulang_Num'] < 1).sum(),
        'Sedang (1-3%)': ((df['Ratio_Mengulang_Num'] >= 1) & (df['Ratio_Mengulang_Num'] < 3)).sum(),
        'Tinggi (≥3%)': (df['Ratio_Mengulang_Num'] >= 3).sum()
    }
    
    # Hapus kategori yang nilainya 0 untuk Pie Chart yang lebih bersih
    mengulang_categories = {k: v for k, v in mengulang_categories.items() if v > 0}
    
    if not mengulang_categories:
        print("Tidak ada data untuk Pie Chart Mengulang. Melanjutkan.")
        return None
        
    colors = ['#90EE90', '#FFD700', '#FF6B6B']  # Hijau, Kuning, Merah
    
    # Buat pie chart
    ax.pie(mengulang_categories.values(), 
           labels=mengulang_categories.keys(),
           autopct='%1.1f%%', 
           colors=colors[:len(mengulang_categories)], # Sesuaikan warna dengan jumlah kategori
           startangle=90)
    
    ax.set_title('SLIDE 1: 🍎 DISTRIBUSI RASIO MENGULANG\n(Jumlah Provinsi per Kategori)', 
                 fontweight='bold', fontsize=16, pad=20)
    
    output_file = 'slide_1_distribusi_mengulang.png'
    fig.savefig(output_file, dpi=300, bbox_inches='tight')
    plt.close(fig)
    print(f"✅ Slide 1 disimpan: {output_file}")
    return output_file

def slide_2_distribusi_putus_sekolah(df):
    """Membuat slide 2: Distribusi Rasio Putus Sekolah"""
    fig, ax = plt.subplots(figsize=(10, 8))
    
    # Kategorisasi data putus sekolah
    putus_categories = {
        'Rendah (<0.5%)': (df['Ratio_Putus_Num'] < 0.5).sum(),
        'Sedang (0.5-1%)': ((df['Ratio_Putus_Num'] >= 0.5) & (df['Ratio_Putus_Num'] < 1)).sum(),
        'Tinggi (≥1%)': (df['Ratio_Putus_Num'] >= 1).sum()
    }
    
    # Hapus kategori yang nilainya 0 untuk Pie Chart yang lebih bersih
    putus_categories = {k: v for k, v in putus_categories.items() if v > 0}
    
    if not putus_categories:
        print("Tidak ada data untuk Pie Chart Putus Sekolah. Melanjutkan.")
        return None

    colors = ['#87CEEB', '#FFA500', '#DC143C']  # Biru, Oranye, Merah
    
    # Buat pie chart
    ax.pie(putus_categories.values(), 
           labels=putus_categories.keys(),
           autopct='%1.1f%%', 
           colors=colors[:len(putus_categories)], # Sesuaikan warna dengan jumlah kategori
           startangle=90)
    
    ax.set_title('SLIDE 2: 🎓 DISTRIBUSI RASIO PUTUS SEKOLAH\n(Jumlah Provinsi per Kategori)', 
                 fontweight='bold', fontsize=16, pad=20)
    
    output_file = 'slide_2_distribusi_putus_sekolah.png'
    fig.savefig(output_file, dpi=300, bbox_inches='tight')
    plt.close(fig)
    print(f"✅ Slide 2 disimpan: {output_file}")
    return output_file

def slide_3_perbandingan_status(df):
    """Membuat slide 3: Perbandingan Negeri vs Swasta"""
    
    # Pastikan kolom 'Status' hanya berisi 'Negeri' dan 'Swasta'
    df_filtered = df[df['Status'].isin(['Negeri', 'Swasta'])].copy()
    
    if df_filtered.empty:
        print("Tidak ada data yang valid untuk perbandingan Negeri vs Swasta. Melanjutkan.")
        return None
        
    fig, ax = plt.subplots(figsize=(10, 8))
    
    # Hitung rata-rata per status sekolah
    status_avg = df_filtered.groupby('Status').agg({
        'Ratio_Mengulang_Num': 'mean',
        'Ratio_Putus_Num': 'mean'
    }).round(2)
    
    # Setup untuk bar chart
    status_list = status_avg.index.tolist()
    x = np.arange(len(status_list))  # Posisi untuk Negeri dan Swasta
    width = 0.35  # Lebar setiap bar
    
    # Buat bar untuk mengulang dan putus sekolah
    bars_mengulang = ax.bar(x - width/2, status_avg['Ratio_Mengulang_Num'], 
                            width, label='Mengulang (Rata-rata)', 
                            color='#FF9E6D', alpha=0.9)
    
    bars_putus = ax.bar(x + width/2, status_avg['Ratio_Putus_Num'], 
                        width, label='Putus Sekolah (Rata-rata)', 
                        color='#6DCFF6', alpha=0.9)
    
    # Atur judul dan label
    ax.set_title('SLIDE 3: 🏫 PERBANDINGAN RATA-RATA RASIO PROVINSI\nNEGERI vs SWASTA', 
                 fontweight='bold', pad=20, fontsize=16)
    ax.set_xticks(x)
    ax.set_xticklabels(status_list, fontsize=12, fontweight='bold')
    ax.set_ylabel('Rata-rata Rasio (%)', fontsize=12)
    ax.legend()
    
    # Tambahkan nilai di atas setiap bar
    for bars in [bars_mengulang, bars_putus]:
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height + 0.05, 
                    f'{height:.2f}%', ha='center', va='bottom', 
                    fontsize=10, fontweight='bold')
    
    output_file = 'slide_3_perbandingan_status.png'
    fig.savefig(output_file, dpi=300, bbox_inches='tight')
    plt.close(fig)
    print(f"✅ Slide 3 disimpan: {output_file}")
    return output_file

def slide_4_summary_statistics(df):
    """Membuat slide 4: Ringkasan Statistik"""
    fig, ax = plt.subplots(figsize=(10, 8))
    ax.axis('off')  # Sembunyikan axes
    
    # Hitung statistik utama
    total_sekolah = df['Sekolah'].sum()
    total_siswa = df['Siswa'].sum()
    avg_mengulang = df['Ratio_Mengulang_Num'].mean()
    avg_putus = df['Ratio_Putus_Num'].mean()
    
    # Cari performer terbaik/terburuk (yang rasio angkanya terendah dan tertinggi)
    # Gunakan .idxmin() dan .idxmax()
    prov_mengulang_terendah = df.loc[df['Ratio_Mengulang_Num'].idxmin(), 'Provinsi']
    prov_mengulang_tertinggi = df.loc[df['Ratio_Mengulang_Num'].idxmax(), 'Provinsi']
    
    prov_putus_terendah = df.loc[df['Ratio_Putus_Num'].idxmin(), 'Provinsi']
    prov_putus_tertinggi = df.loc[df['Ratio_Putus_Num'].idxmax(), 'Provinsi']
    
    
    # Buat teks ringkasan
    summary_text = f"""
SLIDE 4: 📋 RINGKASAN STATISTIK UTAMA SD INDONESIA
--------------------------------------------------------
 
[Image of Elementary School Building]

SEKOLAH & SISWA NASIONAL
🏫 TOTAL SEKOLAH: {total_sekolah:,.0f}
👨‍🎓 TOTAL SISWA: {total_siswa:,.0f}
 
RATA-RATA RASIO PROVINSI
📊 Mengulang: {avg_mengulang:.2f}%
📊 Putus Sekolah: {avg_putus:.2f}%
 
KINERJA PROVINSI (Mengulang / Putus Sekolah)
 
🏆 RASIO TERENDAH (TERBAIK)
 • Mengulang: {prov_mengulang_terendah}
 • Putus Sekolah: {prov_putus_terendah}
 
📉 RASIO TERTINGGI (TERBURUK)
 • Mengulang: {prov_mengulang_tertinggi}
 • Putus Sekolah: {prov_putus_tertinggi}

--------------------------------------------------------
"""
    
    # Buat kotak dekorasi
    from matplotlib.patches import FancyBboxPatch # Import lokal untuk memastikan ketersediaan
    bbox = FancyBboxPatch((0.05, 0.05), 0.9, 0.9, 
                          boxstyle="round,pad=0.02", 
                          facecolor='#E0F7FA', alpha=0.8, 
                          edgecolor='#00BCD4', linewidth=1.5)
    ax.add_patch(bbox)
    
    # Tambahkan teks ke dalam kotak
    ax.text(0.1, 0.95, summary_text, fontsize=12, fontfamily='monospace', 
            verticalalignment='top', linespacing=2, fontweight='normal', color='#263238')
    
    output_file = 'slide_4_summary_statistics.png'
    fig.savefig(output_file, dpi=300, bbox_inches='tight')
    plt.close(fig)
    print(f"✅ Slide 4 disimpan: {output_file}")
    return output_file

# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

def main():
    """Fungsi utama untuk menjalankan program dan membuat 4 slide."""
    
    # 1. Bersihkan dan siapkan data
    df = clean_and_prepare_data(FILE_INPUT)
    
    # 2. Jika data berhasil dibersihkan, buat visualisasi
    if df is not None:
        print("\n" + "="*50)
        print("MEMBUAT 4 SLIDE VISUALISASI")
        print("="*50)
        
        # Buat 4 slide visualisasi
        slide_1_distribusi_mengulang(df)
        slide_2_distribusi_putus_sekolah(df)
        slide_3_perbandingan_status(df)
        slide_4_summary_statistics(df)
        
        # 3. Simpan data bersih (tanpa kolom perhitungan internal)
        df_to_save = df.drop(columns=['Ratio_Mengulang_Num', 'Ratio_Putus_Num'], errors='ignore')
        df_to_save.to_excel(FILE_OUTPUT_CLEAN, index=False)
        print(f"\n💾 Data clean disimpan: {FILE_OUTPUT_CLEAN}")
        
        # 4. Tampilkan pesan sukses
        print("\n" + "="*50)
        print("PROGRAM SELESAI! 4 SLIDE TELAH DIBUAT. 🎉")
        print("="*50)
        
        # Tampilkan preview data
        print(f"\n📊 PREVIEW DATA BERSIH:")
        print(df_to_save.head())
        
    else:
        print("\n❌ Program dihentikan karena ada error dalam pembersihan data.")

# Jalankan program (Sintaks diperbaiki: __name__ == "__main__")
if __name__ == "__main__":
    main()