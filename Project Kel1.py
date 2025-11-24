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
    # dan ratio putus sekolah di samping kolom putus sekolah
    kolom_akhir = df.columns.tolist()
    
    # Pindahkan 'Ratio Mengulang (%)' ke setelah 'Mengulang'
    idx_mengulang = kolom_akhir.index('Mengulang')
    kolom_akhir.insert(idx_mengulang + 1, kolom_akhir.pop(kolom_akhir.index('Ratio Mengulang (%)')))
    
    # Pindahkan 'Ratio Putus Sekolah (%)' ke setelah 'Putus Sekolah'
    idx_putus = kolom_akhir.index('Putus Sekolah')
    kolom_akhir.insert(idx_putus + 1, kolom_akhir.pop(kolom_akhir.index('Ratio Putus Sekolah (%)')))
    
    df = df[kolom_akhir]
    
    print("Pembersihan data selesai & Penambahan Ratio Mengulang & Putus Sekolah.")
    print(f"Dimensi data: {df.shape}")
    print(f"Kolom: {df.columns.tolist()}")
    
    return df