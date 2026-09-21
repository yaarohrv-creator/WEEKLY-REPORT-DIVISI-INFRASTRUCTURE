import streamlit as st
import pandas as pd
import sqlite3

# -----------------------------------------------------------------------------
# FUNGSI KONEKSI DATABASE (Pastikan disesuaikan dengan nama database Anda)
# -----------------------------------------------------------------------------
def get_db_connection():
    # Ganti 'database.db' dengan nama file database SQLite Anda
    conn = sqlite3.connect('database.db', check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

# -----------------------------------------------------------------------------
# FUNGSI RENDER DASHBOARD TABLE (LENGKAP)
# -----------------------------------------------------------------------------
def render_dashboard_table(df_data, tab_key_prefix):
    column_order = [
        'No', 'Nomor SPK', 'Nama Kontraktor', 'Jenis Pekerjaan', 'Unit Proyek', 'Jumlah',
        'Nilai Kontrak Pekerjaan Ini (Rp)', 'Progress Minggu Lalu (%)', 'Progress Minggu Ini (%)',
        'Selisih / Varian (%)', 'Catatan Pekerjaan Terbaru', 'Pratinjau Foto 1', 'Pratinjau Foto 2'
    ]

    # Handle dataframe kosong
    if df_data is None or df_data.empty:
        st.info("💡 Belum ada data progress untuk wilayah/kategori ini.")
        df_display = pd.DataFrame(columns=['real_id'] + column_order)
    else:
        df_display = df_data.copy().reset_index(drop=True)
        if 'No' not in df_display.columns:
            df_display.insert(0, 'No', range(1, len(df_display) + 1))

    existing_cols = [c for c in ['real_id'] + column_order if c in df_display.columns]

    editor_key = f"editor_{tab_key_prefix}"
    
    # Render Komponen Data Editor Streamlit
    edited_df = st.data_editor(
        df_display[existing_cols],
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "real_id": None,  # Menyembunyikan kolom ID internal dari UI
            "Jumlah": st.column_config.NumberColumn("Jumlah", format="%d"),
            "Nilai Kontrak Pekerjaan Ini (Rp)": st.column_config.NumberColumn("Nilai Kontrak (Rp)", format="Rp %d"),
            "Progress Minggu Lalu (%)": st.column_config.NumberColumn("Progress Minggu Lalu (%)", format="%.2f %%"),
            "Progress Minggu Ini (%)": st.column_config.NumberColumn("Progress Minggu Ini (%)", format="%.2f %%"),
            "Selisih / Varian (%)": st.column_config.NumberColumn("Selisih / Varian (%)", format="%.2f %%"),
            "Pratinjau Foto 1": st.column_config.ImageColumn("Pratinjau Foto 1"),
            "Pratinjau Foto 2": st.column_config.ImageColumn("Pratinjau Foto 2"),
        },
        key=editor_key
    )

    # -------------------------------------------------------------------------
    # 1. DETEKSI & EKSEKUSI HAPUS PERMANEN DARI DATABASE
    # -------------------------------------------------------------------------
    if df_data is not None and not df_data.empty:
        # Ambil daftar indeks baris yang dihapus/dicentang oleh pengguna
        deleted_indices = []
        if editor_key in st.session_state and "deleted_rows" in st.session_state[editor_key]:
            deleted_indices = st.session_state[editor_key]["deleted_rows"]

        # Eksekusi Otomatis jika pengguna menghapus baris di UI Streamlit
        if deleted_indices:
            with get_db_connection() as conn:
                cursor = conn.cursor()
                for idx in deleted_indices:
                    if idx < len(df_display):
                        row_to_del = df_display.iloc[idx]
                        real_id = row_to_del.get('real_id')
                        no_spk = str(row_to_del.get('Nomor SPK', '')).strip()
                        j_pek = str(row_to_del.get('Jenis Pekerjaan', '')).strip()

                        # A. Hapus berdasarkan ID Utama jika tersedia
                        if pd.notna(real_id) and str(real_id).isdigit():
                            cursor.execute("DELETE FROM laporan_mingguan WHERE id = ?", (int(real_id),))

                        # B. Hapus dari seluruh tabel terkait menggunakan TRIM & LOWER agar akurat 100%
                        if no_spk:
                            # Hapus dari tabel Laporan Mingguan
                            cursor.execute("""
                                DELETE FROM laporan_mingguan 
                                WHERE LOWER(TRIM(no_spk)) = LOWER(?) 
                                   OR (LOWER(TRIM(no_spk)) = LOWER(?) AND LOWER(TRIM(jenis_pekerjaan)) = LOWER(?))
                            """, (no_spk, no_spk, j_pek))

                            # Hapus dari tabel Master SPK (PENTING: Agar data tidak muncul lagi dari JOIN master)
                            cursor.execute("""
                                DELETE FROM master_spk 
                                WHERE LOWER(TRIM(no_spk)) = LOWER(?)
                            """, (no_spk,))

                            # Hapus dari tabel History Progress
                            cursor.execute("""
                                DELETE FROM history_progress 
                                WHERE LOWER(TRIM(no_spk)) = LOWER(?)
                            """, (no_spk,))

                conn.commit()

            # Bersihkan session state editor untuk menghindari loop
            if editor_key in st.session_state:
                del st.session_state[editor_key]

            st.success("✅ Data berhasil dibersihkan dan dihapus permanen dari seluruh tabel database!")
            st.rerun()

        # -------------------------------------------------------------------------
        # 2. TOMBOL SIMPAN EDITAN TEXT / ANGKA
        # -------------------------------------------------------------------------
        if st.button("💾 Simpan Perubahan Data", key=f"btn_save_{tab_key_prefix}"):
            with get_db_connection() as conn:
                cursor = conn.cursor()
                for idx, row in edited_df.iterrows():
                    if idx < len(df_display):
                        real_id = df_display.iloc[idx].get('real_id')
                        if pd.notna(real_id) and str(real_id).isdigit():
                            cursor.execute("""
                                UPDATE laporan_mingguan
                                SET progress_minggu_ini = ?, catatan = ?
                                WHERE id = ?
                            """, (
                                row.get('Progress Minggu Ini (%)'), 
                                row.get('Catatan Pekerjaan Terbaru'), 
                                int(real_id)
                            ))
                conn.commit()

            st.success("✅ Perubahan data berhasil disimpan!")
            st.rerun()

        # -------------------------------------------------------------------------
        # 3. EXPORT & DOWNLOAD LAPORAN EXCEL
        # -------------------------------------------------------------------------
        st.markdown("---")
        st.subheader("📥 Export & Download Laporan")
        try:
            excel_bytes = generate_excel_full_feature(df_display)
            st.download_button(
                label=f"📥 Download Laporan ({tab_key_prefix.capitalize()}) - Excel",
                data=excel_bytes,
                file_name=f"Laporan_Progress_{tab_key_prefix}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{tab_key_prefix}"
            )
        except Exception as e:
            st.error(f"Gagal memproses file Excel: {e}")
