import os
import io
import sqlite3
import pandas as pd
import streamlit as st

# Modul untuk styling dan menyisipkan gambar ke Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Konfigurasi Halaman Streamlit
st.set_page_config(page_title="Sistem Progress Proyek", layout="wide")

# ---------------------------------------------------------
# FUNGSIONALITAS DATABASE SQLITE (proyek_v2.db)
# ---------------------------------------------------------
def init_db():
    conn = sqlite3.connect('proyek_v2.db')
    cursor = conn.cursor()

    # 1. TABEL MASTER_SPK
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS master_spk (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            no_spk TEXT,
            kontraktor TEXT,
            jenis_pekerjaan TEXT,
            unit TEXT,
            jumlah INTEGER DEFAULT 1,
            nilai_pekerjaan REAL,
            UNIQUE(no_spk, jenis_pekerjaan)
        )
    ''')

    # 2. TABEL LAPORAN_MINGGUAN
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS laporan_mingguan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            waktu_input DATETIME DEFAULT CURRENT_TIMESTAMP,
            no_spk TEXT,
            jenis_pekerjaan TEXT,
            kontraktor TEXT,
            unit TEXT,
            jumlah INTEGER,
            nilai_pekerjaan REAL,
            progress_minggu_lalu REAL,
            progress_minggu_ini REAL,
            catatan TEXT
        )
    ''')

    # 3. TABEL DOKUMENTASI FOTO
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS dokumentasi_foto (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tipe_dok TEXT,
            unit TEXT,
            nama_file TEXT
        )
    ''')

    conn.commit()
    return conn

conn = init_db()

# ---------------------------------------------------------
# NAVIGASI SIDEBAR
# ---------------------------------------------------------
st.sidebar.title("Navigasi")
menu = st.sidebar.selectbox("Pilih Menu", [
    "Dashboard Progress", 
    "Input Progress Mingguan",
    "Kelola Master SPK", 
    "Upload Dokumentasi Foto"
])

# ---------------------------------------------------------
# MENU 1: DASHBOARD PROGRESS
# ---------------------------------------------------------
if menu == "Dashboard Progress":
    st.title("📊 Dashboard Progress Proyek")

    query_view = """
        SELECT 
            id AS real_id,
            waktu_input AS [Waktu Input],
            no_spk AS [Nomor SPK],
            kontraktor AS [Nama Kontraktor],
            jenis_pekerjaan AS [Jenis Pekerjaan],
            unit AS [Unit Proyek],
            jumlah AS [Jumlah],
            nilai_pekerjaan AS [Nilai Kontrak Pekerjaan Ini (Rp)],
            progress_minggu_lalu AS [Progress Minggu Lalu (%)],
            progress_minggu_ini AS [Progress Minggu Ini (%)],
            (COALESCE(progress_minggu_ini, 0) - COALESCE(progress_minggu_lalu, 0)) AS [Selisih / Varian (%)],
            catatan AS [Catatan Pekerjaan]
        FROM laporan_mingguan
        ORDER BY id ASC
    """

    try:
        df_view = pd.read_sql_query(query_view, conn)
    except Exception:
        df_view = pd.DataFrame()

    if df_view.empty:
        df_view = pd.DataFrame(columns=[
            'No', 'real_id', 'Waktu Input', 'Nomor SPK', 'Nama Kontraktor',
            'Jenis Pekerjaan', 'Unit Proyek', 'Jumlah', 'Nilai Kontrak Pekerjaan Ini (Rp)',
            'Progress Minggu Lalu (%)', 'Progress Minggu Ini (%)',
            'Selisih / Varian (%)', 'Catatan Pekerjaan'
        ])
        st.info("💡 Belum ada data progress. Silakan kelola Master SPK atau input progress mingguan.")
    else:
        if 'No' not in df_view.columns:
            df_view.insert(0, 'No', range(1, len(df_view) + 1))

    edited_df = st.data_editor(
        df_view,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="editor_dashboard"
    )

    if st.button("💾 Simpan Perubahan & Hapus Data"):
        cursor = conn.cursor()
        current_ids = [row['real_id'] for idx, row in edited_df.iterrows() if pd.notna(row.get('real_id'))]

        if current_ids:
            format_strings = ','.join(['?'] * len(current_ids))
            cursor.execute(f"DELETE FROM laporan_mingguan WHERE id NOT IN ({format_strings})", current_ids)
        else:
            cursor.execute("DELETE FROM laporan_mingguan")

        for idx, row in edited_df.iterrows():
            real_id = row.get('real_id')
            if pd.notna(real_id) and real_id != "":
                cursor.execute("""
                    UPDATE laporan_mingguan
                    SET no_spk = ?,
                        kontraktor = ?,
                        jenis_pekerjaan = ?,
                        unit = ?,
                        jumlah = ?,
                        nilai_pekerjaan = ?,
                        progress_minggu_lalu = ?,
                        progress_minggu_ini = ?,
                        catatan = ?
                    WHERE id = ?
                """, (
                    row.get('Nomor SPK'),
                    row.get('Nama Kontraktor'),
                    row.get('Jenis Pekerjaan'),
                    row.get('Unit Proyek'),
                    row.get('Jumlah'),
                    row.get('Nilai Kontrak Pekerjaan Ini (Rp)'),
                    row.get('Progress Minggu Lalu (%)'),
                    row.get('Progress Minggu Ini (%)'),
                    row.get('Catatan Pekerjaan'),
                    real_id
                ))

        conn.commit()
        st.success("Perubahan data berhasil disimpan!")
        st.rerun()

    st.markdown("---")
    st.subheader("📥 Export & Download Laporan")

    def generate_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_excel = df_view.drop(columns=['real_id'], errors='ignore')
            df_excel.to_excel(writer, index=False, sheet_name='Laporan Progress')
            
            worksheet = writer.sheets['Laporan Progress']
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
            )

            for col_num in range(1, len(df_excel.columns) + 1):
                cell = worksheet.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for row_idx in range(2, len(df_excel) + 2):
                for col_idx in range(1, len(df_excel.columns) + 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="center")

            for col in worksheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

        return output.getvalue()

    try:
        excel_data = generate_excel()
        st.download_button(
            label="📥 Download Laporan (Excel)",
            data=excel_data,
            file_name="Laporan_Progress_Proyek.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Gagal memproses file Excel: {e}")

# ---------------------------------------------------------
# MENU 2: KELOLA MASTER SPK
# ---------------------------------------------------------
elif menu == "Kelola Master SPK":
    st.title("📑 Kelola Master SPK Proyek")
    st.write("Daftarkan rincian jenis pekerjaan untuk setiap Nomor SPK.")

    with st.form("form_master_spk"):
        c1, c2 = st.columns(2)
        with c1:
            no_spk = st.text_input("Nomor SPK", placeholder="Contoh: 048/PSM2/BPRE/BPSL/JKTO/INF/III/2026")
            kontraktor = st.text_input("Nama Kontraktor", placeholder="Contoh: CV. Selamat Jaya")
            jenis_pekerjaan = st.text_input("Jenis Pekerjaan", placeholder="Contoh: Renovasi atap R. G2 No 18 Tahun 1996 - BPRE")
        with c2:
            unit = st.text_input("Unit Proyek", placeholder="Contoh: BPRE")
            jumlah = st.number_input("Jumlah", min_value=1, step=1, value=1)
            nilai_pekerjaan = st.number_input("Nilai Kontrak Pekerjaan Ini (Rp)", min_value=0.0, step=1000000.0, format="%.2f")

        submit_master = st.form_submit_button("➕ Tambah Ke Master SPK")

        if submit_master:
            if not no_spk or not jenis_pekerjaan or not kontraktor:
                st.warning("Nomor SPK, Nama Kontraktor, dan Jenis Pekerjaan wajib diisi.")
            else:
                try:
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO master_spk (no_spk, kontraktor, jenis_pekerjaan, unit, jumlah, nilai_pekerjaan)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (no_spk, kontraktor, jenis_pekerjaan, unit, jumlah, nilai_pekerjaan))
                    conn.commit()
                    st.success(f"✅ Item '{jenis_pekerjaan}' berhasil ditambahkan ke SPK '{no_spk}'!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("⚠️ Jenis pekerjaan ini sudah ada di dalam SPK tersebut!")

    st.markdown("---")
    st.subheader("📋 Edit & Kelola Master SPK")
    st.caption("💡 **Cara Edit:** Klik langsung pada cell yang ingin diubah. Klik 'Simpan Perubahan Master SPK' untuk memperbarui database.")

    try:
        df_master = pd.read_sql_query("""
            SELECT 
                id AS real_id,
                no_spk AS [Nomor SPK], 
                kontraktor AS [Nama Kontraktor], 
                jenis_pekerjaan AS [Jenis Pekerjaan], 
                unit AS [Unit Proyek], 
                jumlah AS [Jumlah],
                nilai_pekerjaan AS [Nilai Kontrak Pekerjaan Ini (Rp)]
            FROM master_spk 
            ORDER BY id ASC
        """, conn)
    except Exception:
        df_master = pd.DataFrame()

    if df_master.empty:
        df_master = pd.DataFrame(columns=[
            'real_id', 'Nomor SPK', 'Nama Kontraktor', 'Jenis Pekerjaan', 'Unit Proyek', 'Jumlah', 'Nilai Kontrak Pekerjaan Ini (Rp)', 'Total Nilai Kontrak (Rp)'
        ])
    else:
        spk_totals = df_master.groupby('Nomor SPK')['Nilai Kontrak Pekerjaan Ini (Rp)'].transform('sum')
        df_master['Total Nilai Kontrak (Rp)'] = spk_totals

    edited_master = st.data_editor(
        df_master,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "real_id": None,
            "Nomor SPK": st.column_config.TextColumn("Nomor SPK"),
            "Nama Kontraktor": st.column_config.TextColumn("Nama Kontraktor"),
            "Jenis Pekerjaan": st.column_config.TextColumn("Jenis Pekerjaan"),
            "Unit Proyek": st.column_config.TextColumn("Unit Proyek"),
            "Jumlah": st.column_config.NumberColumn("Jumlah", min_value=1, step=1),
            "Nilai Kontrak Pekerjaan Ini (Rp)": st.column_config.NumberColumn(
                "Nilai Kontrak Pekerjaan Ini (Rp)",
                format="Rp %d"
            ),
            "Total Nilai Kontrak (Rp)": st.column_config.NumberColumn(
                "Total Nilai Kontrak (Rp)",
                format="Rp %d",
                disabled=True
            )
        },
        key="editor_master_spk"
    )

    if st.button("💾 Simpan Perubahan Master SPK"):
        cursor = conn.cursor()
        
        current_ids = [row['real_id'] for idx, row in edited_master.iterrows() if pd.notna(row.get('real_id'))]
        if current_ids:
            format_strings = ','.join(['?'] * len(current_ids))
            cursor.execute(f"DELETE FROM master_spk WHERE id NOT IN ({format_strings})", current_ids)
        else:
            cursor.execute("DELETE FROM master_spk")

        for idx, row in edited_master.iterrows():
            real_id = row.get('real_id')
            if pd.notna(real_id) and real_id != "":
                cursor.execute("""
                    UPDATE master_spk
                    SET no_spk = ?,
                        kontraktor = ?,
                        jenis_pekerjaan = ?,
                        unit = ?,
                        jumlah = ?,
                        nilai_pekerjaan = ?
                    WHERE id = ?
                """, (
                    row.get('Nomor SPK'),
                    row.get('Nama Kontraktor'),
                    row.get('Jenis Pekerjaan'),
                    row.get('Unit Proyek'),
                    row.get('Jumlah'),
                    row.get('Nilai Kontrak Pekerjaan Ini (Rp)'),
                    real_id
                ))
            else:
                if row.get('Nomor SPK') and row.get('Jenis Pekerjaan'):
                    cursor.execute("""
                        INSERT INTO master_spk (no_spk, kontraktor, jenis_pekerjaan, unit, jumlah, nilai_pekerjaan)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (
                        row.get('Nomor SPK'),
                        row.get('Nama Kontraktor'),
                        row.get('Jenis Pekerjaan'),
                        row.get('Unit Proyek'),
                        row.get('Jumlah', 1),
                        row.get('Nilai Kontrak Pekerjaan Ini (Rp)', 0.0)
                    ))

        conn.commit()
        st.success("✅ Master SPK berhasil diperbarui!")
        st.rerun()

# ---------------------------------------------------------
# MENU 3: INPUT PROGRESS MINGGUAN (UPDATE OTOMATIS)
# ---------------------------------------------------------
elif menu == "Input Progress Mingguan":
    st.title("📝 Input Progress Mingguan Berdasarkan SPK")

    try:
        spk_list = pd.read_sql_query("SELECT DISTINCT no_spk FROM master_spk", conn)['no_spk'].tolist()
    except Exception:
        spk_list = []

    if not spk_list:
        st.warning("⚠️ Belum ada data Master SPK. Harap daftarkan SPK di menu 'Kelola Master SPK' terlebih dahulu.")
    else:
        col_spk, col_job = st.columns(2)

        with col_spk:
            selected_spk = st.selectbox("Pilih Nomor SPK", spk_list)

        job_list = pd.read_sql_query(
            "SELECT jenis_pekerjaan FROM master_spk WHERE no_spk = ?", 
            conn, 
            params=(selected_spk,)
        )['jenis_pekerjaan'].tolist()

        with col_job:
            selected_job = st.selectbox("Pilih Jenis Pekerjaan", job_list)

        spk_detail = pd.read_sql_query(
            "SELECT * FROM master_spk WHERE no_spk = ? AND jenis_pekerjaan = ?", 
            conn, 
            params=(selected_spk, selected_job)
        ).iloc[0]

        # Cek apakah SPK + Jenis Pekerjaan ini sudah ada di laporan
        default_progress_lalu = 0.0
        default_progress_ini = 0.0
        default_catatan = ""
        already_exists = False

        try:
            existing_df = pd.read_sql_query(
                "SELECT progress_minggu_ini, catatan FROM laporan_mingguan WHERE no_spk = ? AND jenis_pekerjaan = ?", 
                conn, 
                params=(selected_spk, selected_job)
            )
            if not existing_df.empty:
                already_exists = True
                # Progress minggu ini yang lama OTOMATIS menjadi progress minggu lalu
                last_progress = float(existing_df.iloc[0]['progress_minggu_ini'] or 0.0)
                default_progress_lalu = last_progress
                default_progress_ini = last_progress
                default_catatan = str(existing_df.iloc[0]['catatan'] or "")
        except Exception:
            pass

        nilai_peks = spk_detail['nilai_pekerjaan'] if pd.notna(spk_detail['nilai_pekerjaan']) else 0.0
        st.info(f"📌 **Detail:** {spk_detail['kontraktor']} | Unit: **{spk_detail['unit']}** | Jumlah: **{spk_detail['jumlah']}** | Nilai Pekerjaan: **Rp {nilai_peks:,.2f}**")

        if already_exists:
            st.caption("🔄 *Data sudah terdaftar di laporan. Input ini akan meng-UPDATE progress terbaru tanpa menambah baris/duplikat baru.*")

        with st.form("form_update_progress_mingguan"):
            col1, col2 = st.columns(2)

            with col1:
                st.text_input("Nama Kontraktor", value=spk_detail['kontraktor'], disabled=True)
                st.text_input("Unit Proyek", value=spk_detail['unit'], disabled=True)
                st.number_input("Jumlah", value=int(spk_detail['jumlah']), disabled=True)
                st.number_input("Nilai Pekerjaan (Rp)", value=float(nilai_peks), disabled=True)

            with col2:
                # Field ini terkunci (disabled=True) agar Progress Minggu Lalu otomatis mengambil dari Progress Minggu Ini sebelumnya
                prog_lalu = st.number_input(
                    "Progress Minggu Lalu (%) [Otomatis]", 
                    value=default_progress_lalu, 
                    min_value=0.0, 
                    max_value=100.0,
                    disabled=True
                )
                
                # Masukkan nilai Progress Minggu Ini yang baru
                prog_ini = st.number_input(
                    "Progress Minggu Ini (%)", 
                    value=default_progress_ini, 
                    min_value=0.0, 
                    max_value=100.0, 
                    step=0.1
                )
                
                catatan = st.text_area("Catatan Pekerjaan Minggu Ini", value=default_catatan, placeholder="Masukkan kendala / progres pekerjaan...")

            submit_progress = st.form_submit_button("💾 Update Progress Minggu Ini")

            if submit_progress:
                cursor = conn.cursor()
                
                # BILA SUDAH ADA, UPDATE BARIS YANG ADA
                if already_exists:
                    cursor.execute("""
                        UPDATE laporan_mingguan
                        SET kontraktor = ?,
                            unit = ?,
                            jumlah = ?,
                            nilai_pekerjaan = ?,
                            progress_minggu_lalu = ?,
                            progress_minggu_ini = ?,
                            catatan = ?,
                            waktu_input = CURRENT_TIMESTAMP
                        WHERE no_spk = ? AND jenis_pekerjaan = ?
                    """, (
                        str(spk_detail['kontraktor']),
                        str(spk_detail['unit']),
                        int(spk_detail['jumlah']),
                        float(spk_detail['nilai_pekerjaan']),
                        float(prog_lalu),
                        float(prog_ini),
                        str(catatan),
                        selected_spk,
                        str(selected_job)
                    ))
                    st.success(f"✅ Progress untuk '{selected_job}' (SPK: {selected_spk}) BERHASIL DI-UPDATE!")
                
                # BILA BELUM ADA, BUAT BARIS BARU
                else:
                    cursor.execute("""
                        INSERT INTO laporan_mingguan (
                            no_spk, jenis_pekerjaan, kontraktor, unit, jumlah, nilai_pekerjaan,
                            progress_minggu_lalu, progress_minggu_ini, catatan
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        selected_spk,
                        str(selected_job),
                        str(spk_detail['kontraktor']),
                        str(spk_detail['unit']),
                        int(spk_detail['jumlah']),
                        float(spk_detail['nilai_pekerjaan']),
                        float(prog_lalu),
                        float(prog_ini),
                        str(catatan)
                    ))
                    st.success(f"✅ Progress baru untuk '{selected_job}' (SPK: {selected_spk}) berhasil ditambahkan!")

                conn.commit()
                st.rerun()

# ---------------------------------------------------------
# MENU 4: UPLOAD DOKUMENTASI FOTO
# ---------------------------------------------------------
elif menu == "Upload Dokumentasi Foto":
    st.title("📷 Upload Dokumentasi Foto Pekerjaan")

    with st.form("form_upload_foto"):
        tipe_dok = st.text_input("Tipe Dokumentasi (misal: Progress Atap, Pondasi)")
        unit_proyek = st.text_input("Nama Unit Proyek (Harus sama persis dengan di laporan)")
        uploaded_files = st.file_uploader("Pilih Foto Dokumentasi", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)
        
        submit_foto = st.form_submit_button("📷 Simpan Semua Foto")

        if submit_foto:
            if not unit_proyek or not uploaded_files:
                st.warning("Mohon isi Nama Unit Proyek dan unggah minimal satu foto.")
            else:
                cursor = conn.cursor()
                for file in uploaded_files:
                    file_path = os.path.join(".", file.name)
                    with open(file_path, "wb") as f:
                        f.write(file.getbuffer())

                    cursor.execute("""
                        INSERT INTO dokumentasi_foto (tipe_dok, unit, nama_file)
                        VALUES (?, ?, ?)
                    """, (tipe_dok, unit_proyek, file.name))

                conn.commit()
                st.success("📷 Foto berhasil diunggah dan disimpan ke database!")
