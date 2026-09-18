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

# Folder Penyimpanan Foto
UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# ---------------------------------------------------------
# FUNGSIONALITAS DATABASE SQLITE (proyek_v2.db)
# ---------------------------------------------------------
def init_db():
    conn = sqlite3.connect('proyek_v2.db')
    cursor = conn.cursor()

    # 1. TABEL MASTER_SPK (Ditambahkan kolom 'foto')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS master_spk (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            no_spk TEXT,
            kontraktor TEXT,
            jenis_pekerjaan TEXT,
            unit TEXT,
            jumlah INTEGER DEFAULT 1,
            nilai_pekerjaan REAL,
            foto TEXT,
            UNIQUE(no_spk, jenis_pekerjaan)
        )
    ''')

    # Migrasi Otomatis jika kolom 'foto' belum ada pada database lama
    cursor.execute("PRAGMA table_info(master_spk)")
    columns = [column[1] for column in cursor.fetchall()]
    if 'foto' not in columns:
        cursor.execute("ALTER TABLE master_spk ADD COLUMN foto TEXT")

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

    # 3. TABEL HISTORY_PROGRESS
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS history_progress (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            waktu_input DATETIME DEFAULT CURRENT_TIMESTAMP,
            no_spk TEXT,
            jenis_pekerjaan TEXT,
            kontraktor TEXT,
            unit TEXT,
            progress_minggu_lalu REAL,
            progress_minggu_ini REAL,
            progres_penambahan REAL,
            catatan TEXT
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
    "Kelola Master SPK"
])

# ---------------------------------------------------------
# MENU 1: DASHBOARD PROGRESS & HISTORY
# ---------------------------------------------------------
if menu == "Dashboard Progress":
    st.title("📊 Dashboard Progress & History Proyek")

    tab1, tab2 = st.tabs(["📌 Status Progress Terkini", "📜 Riwayat / History Perubahan Progress"])

    with tab1:
        query_view = """
            SELECT 
                l.id AS real_id,
                l.waktu_input AS [Waktu Input Terbaru],
                l.no_spk AS [Nomor SPK],
                l.kontraktor AS [Nama Kontraktor],
                l.jenis_pekerjaan AS [Jenis Pekerjaan],
                l.unit AS [Unit Proyek],
                l.jumlah AS [Jumlah],
                l.nilai_pekerjaan AS [Nilai Kontrak Pekerjaan Ini (Rp)],
                l.progress_minggu_lalu AS [Progress Minggu Lalu (%)],
                l.progress_minggu_ini AS [Progress Minggu Ini (%)],
                (COALESCE(l.progress_minggu_ini, 0) - COALESCE(l.progress_minggu_lalu, 0)) AS [Selisih / Varian (%)],
                m.foto AS [Path Foto],
                l.catatan AS [Catatan Pekerjaan Terbaru]
            FROM laporan_mingguan l
            LEFT JOIN master_spk m ON l.no_spk = m.no_spk AND l.jenis_pekerjaan = m.jenis_pekerjaan
            ORDER BY l.id ASC
        """

        try:
            df_view = pd.read_sql_query(query_view, conn)
        except Exception:
            df_view = pd.DataFrame()

        if df_view.empty:
            st.info("💡 Belum ada data progress terkini. Silakan input progress mingguan.")
        else:
            if 'No' not in df_view.columns:
                df_view.insert(0, 'No', range(1, len(df_view) + 1))

            edited_df = st.data_editor(
                df_view,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                column_config={
                    "real_id": None,
                    "Path Foto": st.column_config.ImageColumn("Dokumentasi", help="Foto dokumentasi pekerjaan")
                },
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
                            row.get('Catatan Pekerjaan Terbaru'),
                            real_id
                        ))

                conn.commit()
                st.success("Perubahan data berhasil disimpan!")
                st.rerun()

    with tab2:
        st.subheader("📜 Log Riwayat Input Progress Pekerjaan")
        query_history = """
            SELECT 
                waktu_input AS [Tanggal / Waktu Update],
                no_spk AS [Nomor SPK],
                kontraktor AS [Kontraktor],
                jenis_pekerjaan AS [Jenis Pekerjaan],
                unit AS [Unit Proyek],
                progress_minggu_lalu AS [Progress Lalu (%)],
                progress_minggu_ini AS [Progress Ini (%)],
                progres_penambahan AS [Penambahan (%)],
                catatan AS [Catatan Pada Tanggal Tersebut]
            FROM history_progress
            ORDER BY id DESC
        """
        try:
            df_history = pd.read_sql_query(query_history, conn)
        except Exception:
            df_history = pd.DataFrame()

        if df_history.empty:
            st.info("💡 Belum ada riwayat update progress.")
        else:
            if 'No' not in df_history.columns:
                df_history.insert(0, 'No', range(1, len(df_history) + 1))

            filter_spk = st.selectbox("Filter Berdasarkan SPK:", ["Semua SPK"] + df_history["Nomor SPK"].unique().tolist())
            if filter_spk != "Semua SPK":
                df_history_filtered = df_history[df_history["Nomor SPK"] == filter_spk]
            else:
                df_history_filtered = df_history

            st.dataframe(df_history_filtered, use_container_width=True, hide_index=True)

# ---------------------------------------------------------
# MENU 2: KELOLA MASTER SPK (DENGAN FOTO DOKUMENTASI)
# ---------------------------------------------------------
elif menu == "Kelola Master SPK":
    st.title("📑 Kelola Master SPK Proyek")
    st.write("Daftarkan rincian jenis pekerjaan beserta foto dokumentasi posisi awal untuk setiap Nomor SPK.")

    with st.form("form_master_spk", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            no_spk = st.text_input("Nomor SPK", placeholder="Contoh: 048/PSM2/BPRE/BPSL/JKTO/INF/III/2026")
            kontraktor = st.text_input("Nama Kontraktor", placeholder="Contoh: CV. Selamat Jaya")
            jenis_pekerjaan = st.text_input("Jenis Pekerjaan", placeholder="Contoh: Renovasi atap R. G2 No 18 Tahun 1996 - BPRE")
            uploaded_file = st.file_uploader("Upload Foto Dokumentasi", type=['jpg', 'jpeg', 'png'])
        with c2:
            unit = st.text_input("Unit Proyek", placeholder="Contoh: BPRE")
            jumlah = st.number_input("Jumlah", min_value=1, step=1, value=1)
            nilai_pekerjaan = st.number_input("Nilai Kontrak Pekerjaan Ini (Rp)", min_value=0.0, step=1000000.0, format="%.2f")

        submit_master = st.form_submit_button("➕ Tambah Ke Master SPK")

        if submit_master:
            if not no_spk or not jenis_pekerjaan or not kontraktor:
                st.warning("Nomor SPK, Nama Kontraktor, dan Jenis Pekerjaan wajib diisi.")
            else:
                file_path = None
                if uploaded_file is not None:
                    file_path = os.path.join(UPLOAD_DIR, uploaded_file.name)
                    with open(file_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                try:
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO master_spk (no_spk, kontraktor, jenis_pekerjaan, unit, jumlah, nilai_pekerjaan, foto)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (no_spk, kontraktor, jenis_pekerjaan, unit, jumlah, nilai_pekerjaan, file_path))
                    conn.commit()
                    st.success(f"✅ Item '{jenis_pekerjaan}' berhasil ditambahkan ke SPK '{no_spk}'!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("⚠️ Jenis pekerjaan ini sudah ada di dalam SPK tersebut!")

    st.markdown("---")
    st.subheader("📋 Edit & Kelola Master SPK")

    try:
        df_master = pd.read_sql_query("""
            SELECT 
                id AS real_id,
                no_spk AS [Nomor SPK], 
                kontraktor AS [Nama Kontraktor], 
                jenis_pekerjaan AS [Jenis Pekerjaan], 
                unit AS [Unit Proyek], 
                jumlah AS [Jumlah],
                nilai_pekerjaan AS [Nilai Kontrak Pekerjaan Ini (Rp)],
                foto AS [Dokumentasi Foto]
            FROM master_spk 
            ORDER BY id ASC
        """, conn)
    except Exception:
        df_master = pd.DataFrame()

    if df_master.empty:
        df_master = pd.DataFrame(columns=[
            'real_id', 'Nomor SPK', 'Nama Kontraktor', 'Jenis Pekerjaan', 'Unit Proyek', 'Jumlah', 'Nilai Kontrak Pekerjaan Ini (Rp)', 'Total Nilai Kontrak (Rp)', 'Dokumentasi Foto'
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
            ),
            "Dokumentasi Foto": st.column_config.ImageColumn(
                "Dokumentasi Foto", 
                help="Pratinjau Foto Pekerjaan"
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

        conn.commit()
        st.success("✅ Master SPK berhasil diperbarui!")
        st.rerun()

# ---------------------------------------------------------
# MENU 3: INPUT PROGRESS MINGGUAN
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
                last_progress = float(existing_df.iloc[0]['progress_minggu_ini'] or 0.0)
                default_progress_lalu = last_progress
                default_progress_ini = last_progress
                default_catatan = str(existing_df.iloc[0]['catatan'] or "")
        except Exception:
            pass

        nilai_peks = spk_detail['nilai_pekerjaan'] if pd.notna(spk_detail['nilai_pekerjaan']) else 0.0
        st.info(f"📌 **Detail:** {spk_detail['kontraktor']} | Unit: **{spk_detail['unit']}** | Jumlah: **{spk_detail['jumlah']}** | Nilai Pekerjaan: **Rp {nilai_peks:,.2f}**")

        # Menampilkan foto posisi pekerjaan jika ada
        if pd.notna(spk_detail.get('foto')) and os.path.exists(str(spk_detail['foto'])):
            st.image(spk_detail['foto'], caption=f"Dokumentasi Pekerjaan: {selected_job}", width=350)

        with st.form("form_update_progress_mingguan"):
            col1, col2 = st.columns(2)

            with col1:
                st.text_input("Nama Kontraktor", value=spk_detail['kontraktor'], disabled=True)
                st.text_input("Unit Proyek", value=spk_detail['unit'], disabled=True)
                st.number_input("Jumlah", value=int(spk_detail['jumlah']), disabled=True)
                st.number_input("Nilai Pekerjaan (Rp)", value=float(nilai_peks), disabled=True)

            with col2:
                prog_lalu = st.number_input(
                    "Progress Minggu Lalu (%) [Otomatis]", 
                    value=default_progress_lalu, 
                    min_value=0.0, 
                    max_value=100.0,
                    disabled=True
                )
                
                prog_ini = st.number_input(
                    "Progress Minggu Ini (%)", 
                    value=default_progress_ini, 
                    min_value=0.0, 
                    max_value=100.0, 
                    step=0.1
                )
                
                catatan = st.text_area("Catatan Pekerjaan Minggu Ini", value=default_catatan, placeholder="Masukkan kendala / progres pekerjaan...")

            submit_progress = st.form_submit_button("💾 Simpan Progress Minggu Ini")

            if submit_progress:
                cursor = conn.cursor()
                penambahan = float(prog_ini) - float(prog_lalu)

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

                cursor.execute("""
                    INSERT INTO history_progress (
                        no_spk, jenis_pekerjaan, kontraktor, unit,
                        progress_minggu_lalu, progress_minggu_ini, progres_penambahan, catatan
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    selected_spk,
                    str(selected_job),
                    str(spk_detail['kontraktor']),
                    str(spk_detail['unit']),
                    float(prog_lalu),
                    float(prog_ini),
                    float(penambahan),
                    str(catatan)
                ))

                conn.commit()
                st.success(f"✅ Progress & History untuk '{selected_job}' (SPK: {selected_spk}) BERHASIL DISIMPAN!")
                st.rerun()
