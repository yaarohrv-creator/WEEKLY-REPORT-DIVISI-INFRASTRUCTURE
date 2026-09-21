import os
import io
import re
import sqlite3
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st

# Modul untuk styling dan export Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------
# KONFIGURASI HALAMAN & DIREKTORI FOTO
# ---------------------------------------------------------
st.set_page_config(page_title="Sistem Progress Proyek", layout="wide")

DB_NAME = "proyek_v2.db"
BASE_UPLOAD_DIR = Path("uploads_dokumentasi")
BASE_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# Logo Aplikasi
logo_path = Path("logo.png")
if logo_path.exists():
    st.sidebar.image(str(logo_path), use_container_width=True)

# ---------------------------------------------------------
# AUTENTIKASI PASSWORD
# ---------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

def check_password():
    password_benar = st.secrets.get("APP_PASSWORD", "123456")
    if st.session_state.get("password_input") == password_benar:
        st.session_state["authenticated"] = True
        if "password_input" in st.session_state:
            del st.session_state["password_input"]
    else:
        st.session_state["authenticated"] = False
        st.error("🔑 Password salah! Silakan coba lagi.")

if not st.session_state["authenticated"]:
    st.title("🔒 Akses Terbatas - Laporan Progress Proyek")
    st.write("Silakan masukkan password tim untuk mengakses aplikasi.")
    st.text_input("Password Akses:", type="password", key="password_input", on_change=check_password)
    st.info("💡 Silakan hubungi admin untuk mendapatkan password akses.")
    st.stop()

# ---------------------------------------------------------
# FUNGSIONALITAS DATABASE & HELPER FOTO
# ---------------------------------------------------------
def get_connection():
    return sqlite3.connect(DB_NAME)

def sanitize_filename(text: str) -> str:
    """Membersihkan teks dari karakter ilegal untuk dijadikan nama folder/file."""
    text = re.sub(r'[^\w\s-]', '', str(text))
    return re.sub(r'[-\s]+', '_', text).strip('_')

def save_uploaded_photo(uploaded_file, no_spk: str, jenis_pekerjaan: str, index_foto: int) -> str:
    """
    Menyimpan foto ke folder khusus nama Jenis Pekerjaan 
    dengan nama file berformat: SPK_JenisPekerjaan_FotoX_Timestamp.ext
    """
    if uploaded_file is None:
        return None

    # 1. Buat folder khusus berdasarkan Jenis Pekerjaan
    folder_name = sanitize_filename(jenis_pekerjaan)
    target_dir = BASE_UPLOAD_DIR / folder_name
    target_dir.mkdir(parents=True, exist_ok=True)

    # 2. Format nama file unik & deskriptif
    spk_clean = sanitize_filename(no_spk)
    ext = Path(uploaded_file.name).suffix
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"{spk_clean}_{folder_name}_Foto{index_foto}_{timestamp}{ext}"

    # 3. Simpan file
    file_path = target_dir / file_name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    return str(file_path)

def init_db():
    with get_connection() as conn:
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
                foto TEXT,
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
                catatan TEXT,
                foto_1 TEXT,
                foto_2 TEXT
            )
        ''')

        # Migrasi Kolom
        cursor.execute("PRAGMA table_info(laporan_mingguan)")
        cols_lap = [col[1] for col in cursor.fetchall()]
        if 'foto_1' not in cols_lap:
            cursor.execute("ALTER TABLE laporan_mingguan ADD COLUMN foto_1 TEXT")
        if 'foto_2' not in cols_lap:
            cursor.execute("ALTER TABLE laporan_mingguan ADD COLUMN foto_2 TEXT")

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
                catatan TEXT,
                foto_1 TEXT,
                foto_2 TEXT
            )
        ''')

        cursor.execute("PRAGMA table_info(history_progress)")
        cols_hist = [col[1] for col in cursor.fetchall()]
        if 'foto_1' not in cols_hist:
            cursor.execute("ALTER TABLE history_progress ADD COLUMN foto_1 TEXT")
        if 'foto_2' not in cols_hist:
            cursor.execute("ALTER TABLE history_progress ADD COLUMN foto_2 TEXT")

        conn.commit()

init_db()

# ---------------------------------------------------------
# NAVIGASI SIDEBAR
# ---------------------------------------------------------
st.sidebar.title("Navigasi")
menu = st.sidebar.selectbox("Pilih Menu", [
    "Dashboard Progress", 
    "Input Progress Mingguan",
    "Kelola Master SPK"
])

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Logout"):
    st.session_state["authenticated"] = False
    st.rerun()

# ---------------------------------------------------------
# MENU 1: DASHBOARD PROGRESS & HISTORY
# ---------------------------------------------------------
if menu == "Dashboard Progress":
    st.title("📊 WEEKLY REPORT DIVISI INFRASTRUCTURE")

    tab1, tab2 = st.tabs(["📌 Status Progress Terkini", "📜 Riwayat / History Perubahan Progress"])

    with tab1:
        query_view = """
            SELECT 
                id AS real_id,
                waktu_input AS [Waktu Input Terbaru],
                no_spk AS [Nomor SPK],
                kontraktor AS [Nama Kontraktor],
                jenis_pekerjaan AS [Jenis Pekerjaan],
                unit AS [Unit Proyek],
                jumlah AS [Jumlah],
                nilai_pekerjaan AS [Nilai Kontrak Pekerjaan Ini (Rp)],
                progress_minggu_lalu AS [Progress Minggu Lalu (%)],
                progress_minggu_ini AS [Progress Minggu Ini (%)],
                (COALESCE(progress_minggu_ini, 0) - COALESCE(progress_minggu_lalu, 0)) AS [Selisih / Varian (%)],
                foto_1 AS [Link Path Foto 1],
                foto_2 AS [Link Path Foto 2],
                catatan AS [Catatan Pekerjaan Terbaru]
            FROM laporan_mingguan
            ORDER BY id ASC
        """

        try:
            with get_connection() as conn:
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
                    "Link Path Foto 1": st.column_config.ImageColumn("Dokumentasi 1", help="Preview foto dokumentasi 1"),
                    "Link Path Foto 2": st.column_config.ImageColumn("Dokumentasi 2", help="Preview foto dokumentasi 2")
                },
                key="editor_dashboard"
            )

            if st.button("💾 Simpan Perubahan & Hapus Data"):
                with get_connection() as conn:
                    cursor = conn.cursor()
                    current_ids = [int(row['real_id']) for idx, row in edited_df.iterrows() if pd.notna(row.get('real_id'))]

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
                                int(real_id)
                            ))

                    conn.commit()
                st.success("Perubahan data berhasil disimpan!")
                st.rerun()

            st.markdown("---")
            st.subheader("📥 Export & Download Laporan")

            def generate_excel(df):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_excel = df.drop(columns=['real_id'], errors='ignore')
                    df_excel.to_excel(writer, index=False, sheet_name='Laporan Progress')
                    
                    worksheet = writer.sheets['Laporan Progress']
                    
                    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                    
                    thin_border = Border(
                        left=Side(style='thin', color='000000'),
                        right=Side(style='thin', color='000000'),
                        top=Side(style='thin', color='000000'),
                        bottom=Side(style='thin', color='000000')
                    )

                    for col_num in range(1, len(df_excel.columns) + 1):
                        cell = worksheet.cell(row=1, column=col_num)
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        cell.border = thin_border

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
                excel_data = generate_excel(df_view)
                st.download_button(
                    label="📥 Download Laporan (Excel)",
                    data=excel_data,
                    file_name="Laporan_Progress_Proyek.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Gagal memproses file Excel: {e}")

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
                foto_1 AS [Link Path Foto 1],
                foto_2 AS [Link Path Foto 2],
                catatan AS [Catatan Pada Tanggal Tersebut]
            FROM history_progress
            ORDER BY id DESC
        """
        try:
            with get_connection() as conn:
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

            st.dataframe(
                df_history_filtered, 
                use_container_width=True, 
                hide_index=True,
                column_config={
                    "Link Path Foto 1": st.column_config.ImageColumn("Dokumentasi 1"),
                    "Link Path Foto 2": st.column_config.ImageColumn("Dokumentasi 2")
                }
            )

# ---------------------------------------------------------
# MENU 2: KELOLA MASTER SPK
# ---------------------------------------------------------
elif menu == "Kelola Master SPK":
    st.title("📑 Kelola Master SPK Proyek")
    st.write("Daftarkan rincian jenis pekerjaan untuk setiap Nomor SPK.")

    with st.form("form_master_spk", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            no_spk = st.text_input("Nomor SPK", placeholder="Contoh: 048/PSM2/BPRE/BPSL/JKTO/INF/III/2026")
            kontraktor = st.text_input("Nama Kontraktor", placeholder="Contoh: CV. Selamat Jaya")
            jenis_pekerjaan = st.text_input("Jenis Pekerjaan", placeholder="Contoh: Renovasi Atap R. G2")
        with c2:
            unit = st.text_input("Unit Proyek", placeholder="Contoh: BPRE")
            jumlah = st.number_input("Jumlah", min_value=1, step=1, value=1)
            nilai_pekerjaan = st.number_input("Nilai Kontrak Pekerjaan Ini (Rp)", min_value=0.0, step=1000000.0, format="%.2f")

        submit_master = st.form_submit_button("➕ Tambah Ke Master SPK")

        if submit_master:
            if not no_spk.strip() or not jenis_pekerjaan.strip() or not kontraktor.strip():
                st.warning("Nomor SPK, Nama Kontraktor, dan Jenis Pekerjaan wajib diisi.")
            else:
                try:
                    with get_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("""
                            INSERT INTO master_spk (no_spk, kontraktor, jenis_pekerjaan, unit, jumlah, nilai_pekerjaan)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (no_spk.strip(), kontraktor.strip(), jenis_pekerjaan.strip(), unit.strip(), int(jumlah), float(nilai_pekerjaan)))
                        conn.commit()
                    st.success(f"✅ Item '{jenis_pekerjaan}' berhasil ditambahkan ke SPK '{no_spk}'!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("⚠️ Jenis pekerjaan ini sudah ada di dalam SPK tersebut!")

    st.markdown("---")
    st.subheader("📋 Edit & Kelola Master SPK")

    try:
        with get_connection() as conn:
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

    if not df_master.empty:
        spk_totals = df_master.groupby('Nomor SPK')['Nilai Kontrak Pekerjaan Ini (Rp)'].transform('sum')
        df_master['Total Nilai Kontrak (Rp)'] = spk_totals

    edited_master = st.data_editor(
        df_master,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "real_id": None,
            "Nilai Kontrak Pekerjaan Ini (Rp)": st.column_config.NumberColumn("Nilai Kontrak Pekerjaan Ini (Rp)", format="Rp %d"),
            "Total Nilai Kontrak (Rp)": st.column_config.NumberColumn("Total Nilai Kontrak (Rp)", format="Rp %d", disabled=True)
        },
        key="editor_master_spk"
    )

    if st.button("💾 Simpan Perubahan Master SPK"):
        with get_connection() as conn:
            cursor = conn.cursor()
            current_ids = [int(row['real_id']) for idx, row in edited_master.iterrows() if pd.notna(row.get('real_id'))]
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
                        SET no_spk = ?, kontraktor = ?, jenis_pekerjaan = ?, unit = ?, jumlah = ?, nilai_pekerjaan = ?
                        WHERE id = ?
                    """, (
                        row.get('Nomor SPK'), row.get('Nama Kontraktor'), row.get('Jenis Pekerjaan'),
                        row.get('Unit Proyek'), row.get('Jumlah'), row.get('Nilai Kontrak Pekerjaan Ini (Rp)'), int(real_id)
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
        with get_connection() as conn:
            spk_list = pd.read_sql_query("SELECT DISTINCT no_spk FROM master_spk", conn)['no_spk'].tolist()
    except Exception:
        spk_list = []

    if not spk_list:
        st.warning("⚠️ Belum ada data Master SPK. Harap daftarkan SPK di menu 'Kelola Master SPK' terlebih dahulu.")
    else:
        col_spk, col_job = st.columns(2)

        with col_spk:
            selected_spk = st.selectbox("Pilih Nomor SPK", spk_list)

        with get_connection() as conn:
            job_list = pd.read_sql_query(
                "SELECT jenis_pekerjaan FROM master_spk WHERE no_spk = ?", 
                conn, 
                params=(selected_spk,)
            )['jenis_pekerjaan'].tolist()

        with col_job:
            selected_job = st.selectbox("Pilih Jenis Pekerjaan", job_list)

        with get_connection() as conn:
            spk_detail = pd.read_sql_query(
                "SELECT * FROM master_spk WHERE no_spk = ? AND jenis_pekerjaan = ?", 
                conn, 
                params=(selected_spk, selected_job)
            ).iloc[0]

        default_progress_lalu = 0.0
        default_progress_ini = 0.0
        default_catatan = ""
        existing_foto_1 = None
        existing_foto_2 = None
        already_exists = False

        try:
            with get_connection() as conn:
                existing_df = pd.read_sql_query(
                    "SELECT progress_minggu_ini, catatan, foto_1, foto_2 FROM laporan_mingguan WHERE no_spk = ? AND jenis_pekerjaan = ?", 
                    conn, 
                    params=(selected_spk, selected_job)
                )
            if not existing_df.empty:
                already_exists = True
                last_progress = float(existing_df.iloc[0]['progress_minggu_ini'] or 0.0)
                default_progress_lalu = last_progress
                default_progress_ini = last_progress
                default_catatan = str(existing_df.iloc[0]['catatan'] or "")
                existing_foto_1 = existing_df.iloc[0]['foto_1']
                existing_foto_2 = existing_df.iloc[0]['foto_2']
        except Exception:
            pass

        nilai_peks = spk_detail['nilai_pekerjaan'] if pd.notna(spk_detail['nilai_pekerjaan']) else 0.0
        st.info(f"📌 **Detail:** {spk_detail['kontraktor']} | Unit: **{spk_detail['unit']}** | Jumlah: **{spk_detail['jumlah']}** | Nilai Pekerjaan: **Rp {nilai_peks:,.2f}**")

        # Tampilkan foto tersimpan sebelumnya
        if existing_foto_1 or existing_foto_2:
            st.markdown("**📸 Foto Dokumentasi Minggu Lalu/Terkini:**")
            img_col1, img_col2 = st.columns(2)
            with img_col1:
                if existing_foto_1 and Path(str(existing_foto_1)).exists():
                    st.image(existing_foto_1, caption=f"Dokumentasi 1 ({Path(existing_foto_1).name})", use_container_width=True)
            with img_col2:
                if existing_foto_2 and Path(str(existing_foto_2)).exists():
                    st.image(existing_foto_2, caption=f"Dokumentasi 2 ({Path(existing_foto_2).name})", use_container_width=True)

        with st.form("form_update_progress_mingguan"):
            col1, col2 = st.columns(2)

            with col1:
                st.text_input("Nama Kontraktor", value=spk_detail['kontraktor'], disabled=True)
                st.text_input("Unit Proyek", value=spk_detail['unit'], disabled=True)
                st.number_input("Jumlah", value=int(spk_detail['jumlah']), disabled=True)
                st.number_input("Nilai Pekerjaan (Rp)", value=float(nilai_peks), disabled=True)

                st.subheader("📷 Upload Foto Dokumentasi Baru")
                file_foto_1 = st.file_uploader("Upload Foto Dokumentasi 1", type=['jpg', 'jpeg', 'png'], key="up_foto_1")
                file_foto_2 = st.file_uploader("Upload Foto Dokumentasi 2", type=['jpg', 'jpeg', 'png'], key="up_foto_2")

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

            submit_progress = st.form_submit_button("💾 Simpan Progress & Foto Minggu Ini")

            if submit_progress:
                penambahan = float(prog_ini) - float(prog_lalu)
                
                # Proses simpan foto ke folder khusus jenis pekerjaan
                path_foto_1 = save_uploaded_photo(file_foto_1, selected_spk, selected_job, 1) or existing_foto_1
                path_foto_2 = save_uploaded_photo(file_foto_2, selected_spk, selected_job, 2) or existing_foto_2

                with get_connection() as conn:
                    cursor = conn.cursor()
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
                                foto_1 = ?,
                                foto_2 = ?,
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
                            path_foto_1,
                            path_foto_2,
                            selected_spk,
                            str(selected_job)
                        ))
                    else:
                        cursor.execute("""
                            INSERT INTO laporan_mingguan (
                                no_spk, jenis_pekerjaan, kontraktor, unit, jumlah, nilai_pekerjaan,
                                progress_minggu_lalu, progress_minggu_ini, catatan, foto_1, foto_2
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            selected_spk,
                            str(selected_job),
                            str(spk_detail['kontraktor']),
                            str(spk_detail['unit']),
                            int(spk_detail['jumlah']),
                            float(spk_detail['nilai_pekerjaan']),
                            float(prog_lalu),
                            float(prog_ini),
                            str(catatan),
                            path_foto_1,
                            path_foto_2
                        ))

                    cursor.execute("""
                        INSERT INTO history_progress (
                            no_spk, jenis_pekerjaan, kontraktor, unit,
                            progress_minggu_lalu, progress_minggu_ini, progres_penambahan, catatan, foto_1, foto_2
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        selected_spk,
                        str(selected_job),
                        str(spk_detail['kontraktor']),
                        str(spk_detail['unit']),
                        float(prog_lalu),
                        float(prog_ini),
                        float(penambahan),
                        str(catatan),
                        path_foto_1,
                        path_foto_2
                    ))

                    conn.commit()

                st.success(f"✅ Progress & Foto Dokumentasi untuk '{selected_job}' BERHASIL DISIMPAN!")
                st.rerun()
