def render_dashboard_table(df_data, tab_key_prefix):
        if df_data.empty:
            st.info("💡 Belum ada data progress untuk wilayah/kategori ini.")
            return

        df_display = df_data.copy()
        if 'No' not in df_display.columns:
            df_display.insert(0, 'No', range(1, len(df_display) + 1))

        column_order = [
            'No', 'Nomor SPK', 'Nama Kontraktor', 'Jenis Pekerjaan', 'Unit Proyek', 'Jumlah',
            'Nilai Kontrak Pekerjaan Ini (Rp)', 'Progress Minggu Lalu (%)', 'Progress Minggu Ini (%)',
            'Selisih / Varian (%)', 'Catatan Pekerjaan Terbaru', 'Pratinjau Foto 1', 'Pratinjau Foto 2'
        ]
        existing_cols = [c for c in column_order if c in df_display.columns]

        # data_editor dengan penanganan baris dihapus (deleted_rows)
        editor_key = f"editor_{tab_key_prefix}"
        edited_df = st.data_editor(
            df_display[existing_cols],
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "real_id": None,
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

        # Cek apakah ada aksi hapus baris langsung via UI centang/tombol hapus
        if editor_key in st.session_state and "deleted_rows" in st.session_state[editor_key]:
            deleted_indices = st.session_state[editor_key]["deleted_rows"]
            if deleted_indices:
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    for idx in deleted_indices:
                        row_to_del = df_data.iloc[idx]
                        real_id = row_to_del['real_id']
                        no_spk = row_to_del['Nomor SPK']
                        j_pek = row_to_del['Jenis Pekerjaan']
                        
                        # Hapus dari database laporan & master
                        cursor.execute("DELETE FROM laporan_mingguan WHERE id = ?", (real_id,))
                        cursor.execute("DELETE FROM master_spk WHERE no_spk = ? AND jenis_pekerjaan = ?", (no_spk, j_pek))
                    conn.commit()
                st.success("✅ Data yang dicentang/dihapus berhasil dibersihkan dari database!")
                st.rerun()

        # Tombol Simpan Perubahan Data
        if st.button("💾 Simpan Perubahan Data", key=f"btn_save_{tab_key_prefix}"):
            with get_db_connection() as conn:
                cursor = conn.cursor()
                for idx, row in edited_df.iterrows():
                    if idx < len(df_data):
                        real_id = df_data.iloc[idx]['real_id']
                        if pd.notna(real_id):
                            cursor.execute("""
                                UPDATE laporan_mingguan
                                SET progress_minggu_ini = ?, catatan = ?
                                WHERE id = ?
                            """, (row.get('Progress Minggu Ini (%)'), row.get('Catatan Pekerjaan Terbaru'), real_id))
                conn.commit()
            st.success("Perubahan data berhasil disimpan!")
            st.rerun()

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
