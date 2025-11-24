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
    
    print("✅ Pembersihan data selesai.")
    print(f"Dimensi data: {df.shape}")
    print(f"Kolom: {df.columns.tolist()}")
    
    return df