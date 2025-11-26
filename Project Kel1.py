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
    
    # Identifikasi status sekolah (Negeri vs Swasta)
    df['Status Sekolah'] = df['Status'].apply(lambda x: 'Negeri' if 'Negeri' in str(x) else 'Swasta' if 'Swasta' in str(x) else 'Tidak Diketahui')
    
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

def create_top10_visualizations(df):
    """Membuat visualisasi Top 10 untuk Mengulang dan Putus Sekolah"""
    print("📈 Membuat visualisasi Top 10...")
    
    # Filter data yang valid
    df_valid = df[(df['Siswa'] > 0) & (df['Mengulang'] >= 0) & (df['Putus Sekolah'] >= 0)]
    
    # 1. Top 10 Siswa Mengulang Tertinggi
    plt.figure(figsize=(14, 10))
    
    # Subplot 1: Top 10 Mengulang Tertinggi
    plt.subplot(2, 2, 1)
    top10_mengulang_tinggi = df_valid.nlargest(10, 'Mengulang')[['Provinsi', 'Mengulang', 'Rasio Mengulang (%)']]
    bars = plt.barh(top10_mengulang_tinggi['Provinsi'], top10_mengulang_tinggi['Mengulang'], 
                    color='#e74c3c', alpha=0.7)
    plt.xlabel('Jumlah Siswa Mengulang')
    plt.title('TOP 10 PROVINSI: SISWA MENGULANG TERTINGGI', fontweight='bold')
    plt.gca().invert_yaxis()
    
    # Tambahkan nilai di bar
    for bar, nilai in zip(bars, top10_mengulang_tinggi['Mengulang']):
        plt.text(bar.get_width() + bar.get_width()*0.01, bar.get_y() + bar.get_height()/2, 
                f'{int(nilai):,}', ha='left', va='center', fontweight='bold')
    
    # Subplot 2: Top 10 Mengulang Terendah
    plt.subplot(2, 2, 2)
    top10_mengulang_rendah = df_valid.nsmallest(10, 'Mengulang')[['Provinsi', 'Mengulang', 'Rasio Mengulang (%)']]
    bars = plt.barh(top10_mengulang_rendah['Provinsi'], top10_mengulang_rendah['Mengulang'], 
                    color='#2ecc71', alpha=0.7)
    plt.xlabel('Jumlah Siswa Mengulang')
    plt.title('TOP 10 PROVINSI: SISWA MENGULANG TERENDAH', fontweight='bold')
    plt.gca().invert_yaxis()
    
    # Tambahkan nilai di bar
    for bar, nilai in zip(bars, top10_mengulang_rendah['Mengulang']):
        plt.text(bar.get_width() + bar.get_width()*0.01, bar.get_y() + bar.get_height()/2, 
                f'{int(nilai):,}', ha='left', va='center', fontweight='bold')
    
    # Subplot 3: Top 10 Putus Sekolah Tertinggi
    plt.subplot(2, 2, 3)
    top10_putus_tinggi = df_valid.nlargest(10, 'Putus Sekolah')[['Provinsi', 'Putus Sekolah', 'Rasio Putus Sekolah (%)']]
    bars = plt.barh(top10_putus_tinggi['Provinsi'], top10_putus_tinggi['Putus Sekolah'], 
                    color='#c0392b', alpha=0.7)
    plt.xlabel('Jumlah Siswa Putus Sekolah')
    plt.title('TOP 10 PROVINSI: SISWA PUTUS SEKOLAH TERTINGGI', fontweight='bold')
    plt.gca().invert_yaxis()
    
    # Tambahkan nilai di bar
    for bar, nilai in zip(bars, top10_putus_tinggi['Putus Sekolah']):
        plt.text(bar.get_width() + bar.get_width()*0.01, bar.get_y() + bar.get_height()/2, 
                f'{int(nilai):,}', ha='left', va='center', fontweight='bold')
    
    # Subplot 4: Top 10 Putus Sekolah Terendah
    plt.subplot(2, 2, 4)
    top10_putus_rendah = df_valid.nsmallest(10, 'Putus Sekolah')[['Provinsi', 'Putus Sekolah', 'Rasio Putus Sekolah (%)']]
    bars = plt.barh(top10_putus_rendah['Provinsi'], top10_putus_rendah['Putus Sekolah'], 
                    color='#27ae60', alpha=0.7)
    plt.xlabel('Jumlah Siswa Putus Sekolah')
    plt.title('TOP 10 PROVINSI: SISWA PUTUS SEKOLAH TERENDAH', fontweight='bold')
    plt.gca().invert_yaxis()
    
    # Tambahkan nilai di bar
    for bar, nilai in zip(bars, top10_putus_rendah['Putus Sekolah']):
        plt.text(bar.get_width() + bar.get_width()*0.01, bar.get_y() + bar.get_height()/2, 
                f'{int(nilai):,}', ha='left', va='center', fontweight='bold')
    
    plt.tight_layout()
    plt.savefig('top10_analysis.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    print("✅ Visualisasi Top 10 berhasil disimpan sebagai 'top10_analysis.png'")
    
    return {
        'top10_mengulang_tinggi': top10_mengulang_tinggi,
        'top10_mengulang_rendah': top10_mengulang_rendah,
        'top10_putus_tinggi': top10_putus_tinggi,
        'top10_putus_rendah': top10_putus_rendah
    }

def create_swasta_vs_negeri_comparison(df):
    """Membuat visualisasi perbandingan Swasta vs Negeri"""
    print("🏫 Membuat perbandingan Swasta vs Negeri...")
    
    # Aggregasi data berdasarkan status sekolah
    status_comparison = df.groupby('Status Sekolah').agg({
        'Sekolah': 'sum',
        'Siswa': 'sum',
        'Mengulang': 'sum',
        'Putus Sekolah': 'sum'
    }).reset_index()
    
    # Hitung rasio
    status_comparison['Rasio Mengulang (%)'] = (status_comparison['Mengulang'] / status_comparison['Siswa'] * 100).round(3)
    status_comparison['Rasio Putus Sekolah (%)'] = (status_comparison['Putus Sekolah'] / status_comparison['Siswa'] * 100).round(3)
    
    # Visualisasi
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(15, 12))
    
    # Plot 1: Perbandingan Jumlah Sekolah
    colors1 = ['#3498db', '#9b59b6', '#95a5a6']
    ax1.bar(status_comparison['Status Sekolah'], status_comparison['Sekolah'], 
            color=colors1[:len(status_comparison)], alpha=0.7)
    ax1.set_title('JUMLAH SEKOLAH: SWASTA vs NEGERI', fontweight='bold')
    ax1.set_ylabel('Jumlah Sekolah')
    ax1.tick_params(axis='x', rotation=45)
    
    # Tambahkan nilai di bar
    for i, v in enumerate(status_comparison['Sekolah']):
        ax1.text(i, v + v*0.01, f'{int(v):,}', ha='center', va='bottom', fontweight='bold')
    
    # Plot 2: Perbandingan Jumlah Siswa
    colors2 = ['#2980b9', '#8e44ad', '#7f8c8d']
    ax2.bar(status_comparison['Status Sekolah'], status_comparison['Siswa'], 
            color=colors2[:len(status_comparison)], alpha=0.7)
    ax2.set_title('JUMLAH SISWA: SWASTA vs NEGERI', fontweight='bold')
    ax2.set_ylabel('Jumlah Siswa')
    ax2.tick_params(axis='x', rotation=45)
    
    # Tambahkan nilai di bar
    for i, v in enumerate(status_comparison['Siswa']):
        ax2.text(i, v + v*0.01, f'{int(v):,}', ha='center', va='bottom', fontweight='bold')
    
    # Plot 3: Perbandingan Rasio Mengulang
    colors3 = ['#e67e22', '#d35400', '#bdc3c7']
    ax3.bar(status_comparison['Status Sekolah'], status_comparison['Rasio Mengulang (%)'], 
            color=colors3[:len(status_comparison)], alpha=0.7)
    ax3.set_title('RASIO MENGULANG: SWASTA vs NEGERI', fontweight='bold')
    ax3.set_ylabel('Rasio Mengulang (%)')
    ax3.tick_params(axis='x', rotation=45)
    
    # Tambahkan nilai di bar
    for i, v in enumerate(status_comparison['Rasio Mengulang (%)']):
        ax3.text(i, v + 0.01, f'{v}%', ha='center', va='bottom', fontweight='bold')
    
    # Plot 4: Perbandingan Rasio Putus Sekolah
    colors4 = ['#c0392b', '#a93226', '#a6acaf']
    ax4.bar(status_comparison['Status Sekolah'], status_comparison['Rasio Putus Sekolah (%)'], 
            color=colors4[:len(status_comparison)], alpha=0.7)
    ax4.set_title('RASIO PUTUS SEKOLAH: SWASTA vs NEGERI', fontweight='bold')
    ax4.set_ylabel('Rasio Putus Sekolah (%)')
    ax4.tick_params(axis='x', rotation=45)
    
    # Tambahkan nilai di bar
    for i, v in enumerate(status_comparison['Rasio Putus Sekolah (%)']):
        ax4.text(i, v + 0.001, f'{v}%', ha='center', va='bottom', fontweight='bold')
    
    plt.tight_layout()
    plt.savefig('swasta_vs_negeri_comparison.png', dpi=300, bbox_inches='tight')
    plt.show()
    
    print("✅ Perbandingan Swasta vs Negeri disimpan sebagai 'swasta_vs_negeri_comparison.png'")
    
    return status_comparison

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
        
        # Buat visualisasi Top 10
        top10_data = create_top10_visualizations(df_clean)
        
        # Buat perbandingan Swasta vs Negeri
        status_data = create_swasta_vs_negeri_comparison(df_clean)
        
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
                
                # Sheet 4: Data Top 10
                with pd.ExcelWriter('data_top10_analysis.xlsx', engine='openpyxl') as writer_top10:
                    top10_data['top10_mengulang_tinggi'].to_excel(writer_top10, sheet_name='Top10 Mengulang Tinggi', index=False)
                    top10_data['top10_mengulang_rendah'].to_excel(writer_top10, sheet_name='Top10 Mengulang Rendah', index=False)
                    top10_data['top10_putus_tinggi'].to_excel(writer_top10, sheet_name='Top10 Putus Tinggi', index=False)
                    top10_data['top10_putus_rendah'].to_excel(writer_top10, sheet_name='Top10 Putus Rendah', index=False)
                    status_data.to_excel(writer_top10, sheet_name='Swasta vs Negeri', index=False)
            
            print(f"\n💾 Data berhasil disimpan ke: {FILE_OUTPUT_CLEAN}")
            print("💾 Data Top 10 disimpan ke: data_top10_analysis.xlsx")
            print("📁 File Excel berisi:")
            print("   - Sheet 'Data Lengkap': Semua data provinsi")
            print("   - Sheet 'Ringkasan Kategori': Statistik per kategori")
            print("   - Sheet 'Data Rendah/Sedang/Tinggi': Data terpisah per kategori")
            print("\n📊 Visualisasi yang dihasilkan:")
            print("   - pie_chart_tingkat_putus_sekolah.png")
            print("   - top10_analysis.png")
            print("   - swasta_vs_negeri_comparison.png")
            
        except Exception as e:
            print(f"❌ Error saat menyimpan file Excel: {e}")
    
    else:
        print("❌ Gagal memproses data.")

# Jalankan program
if __name__ == "__main__":
    main()
