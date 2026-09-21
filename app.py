import os
import io
import sqlite3
import pandas as pd
import streamlit as st
import re # Digunakan untuk sanitasi nama file

# Modul untuk styling dan export Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- LIBRARY UNTUK MEMPROSES GAMBAR (WAJIB INSTAL PILImage) ---
# Perintah instalasi: pip install Pillow
try:
    from PIL import Image as PILImage
    from openpyxl.drawing.image import Image as OpenPyXLImage
    has_pil = True
except ImportError:
    st.error("⚠️ Library 'Pillow' belum terinstal. Gambar fisik tidak akan muncul di Excel. Silakan instal dengan perintah: pip install Pillow")
    has_pil = False
# --------------------------------------------------------------

# Konfigurasi Halaman Streamlit
st.set_page_config(page_title="Sistem Progress Proyek", layout="wide")

# Logo Aplikasi
if os.path.exists("logo.png"):
    st.sidebar.image("logo.png", use_container_width=True)

# Folder Penyimpanan Foto
UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# ---------------------------------------------------------
# AUTENTIKASI PASSWORD
# ---------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

def check_password():
    password_benar = st.secrets.get("APP_PASSWORD", "123456")
    if st.session_state["password_input"] == password_benar:
        st.session_state["authenticated"] = True
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
# FUNGSIONALITAS DATABASE SQLITE (proyek_v2.db)
# ---------------------------------------------------------
# Helper function untuk membersihkan nama file dari karakter ilegal
def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", filename).replace(" ", "_")

def get_db_connection():
    return sqlite3.connect('proyek_v2.db')

def init_db():
    conn = get_db_connection()
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
            catatan TEXT,
            foto_1 TEXT,
            foto_2 TEXT
        )
    ''')

    # Migrasi otomatis jika kolom foto_1 / foto_2 belum ada
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
    return conn

conn = init_db()

# ==========================================
# FUNGSI EXPORT EXCEL (LINK INTERAKTIF + FOTO RAPI)
# ==========================================
def generate_excel_full_feature(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # ---------------------------------------------------------
        # Persiapan Data
        # ---------------------------------------------------------
        df_excel = df.drop(columns=['real_id'], errors='ignore').copy()
        
        # Lembar Utama (Laporan Progress) - Tanpa Kolom Path Foto Panjang
        df_progress = df_excel.drop(columns=['Foto 1', 'Foto 2'], errors='ignore').copy()
        
        # --- PERBAIKAN 1: Sisipkan Kolom 'Dokumentasi' Kosong ---
        # Kita sisipkan kolom 'Dokumentasi' kosong setelah 'Catatan Pekerjaan Terbaru'
        try:
            target_col_idx = df_progress.columns.get_loc('Catatan Pekerjaan Terbaru') + 1
            df_progress.insert(target_col_idx, 'Dokumentasi', '') 
        except Exception:
            # Jika kolom catatan tidak ditemukan, tambahkan di paling akhir
            df_progress['Dokumentasi'] = ''

        df_progress.to_excel(writer, index=False, sheet_name='Laporan Progress')
        
        worksheet_progress = writer.sheets['Laporan Progress']
        
        # Lembar Kedua (Foto Dokumentasi) - Layout Khusus Gambar
        workbook = writer.book
        worksheet_foto = workbook.create_sheet(title='Foto Dokumentasi')
        
        # ---------------------------------------------------------
        # Styling & Hyperlink Internal (Sheet Progress)
        # ---------------------------------------------------------
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        border_standard = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        align_center = Alignment(horizontal="center", vertical="center")
        blue_link_font = Font(color="0000FF", underline="single")

        # Format Header Sheet Progress
        for col_num in range(1, worksheet_progress.max_column + 1):
            cell = worksheet_progress.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_standard

        # Format Sel Data & Auto-adjust lebar kolom (Sheet Progress)
        for row_idx in range(2, worksheet_progress.max_row + 1):
            for col_idx in range(1, worksheet_progress.max_column + 1):
                cell = worksheet_progress.cell(row=row_idx, column=col_idx)
                cell.border = border_standard
                cell.alignment = Alignment(vertical="center")

        # Auto-adjust lebar kolom (Sheet Progress) - Kecuali Dokumentasi
        for col in worksheet_progress.columns:
            header_name = col[0].value
            if header_name != 'Dokumentasi':
                max_len = max(len(str(cell.value or '')) for cell in col)
                worksheet_progress.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 12)
            else:
                # Set lebar fix untuk kolom Dokumentasi
                worksheet_progress.column_dimensions[get_column_letter(col[0].column)].width = 15

        # ---------------------------------------------------------
        # Layout & Penyisipan Gambar Visual (Sheet Foto Dokumentasi)
        # ---------------------------------------------------------
        LEBAR_KOLOM_FOTO = 50
        worksheet_foto.column_dimensions['A'].width = 40 # Kolom Judul
        worksheet_foto.column_dimensions['B'].width = LEBAR_KOLOM_FOTO # Kolom Foto 1
        worksheet_foto.column_dimensions['C'].width = LEBAR_KOLOM_FOTO # Kolom Foto 2

        # Header Sheet Foto
        headers_foto = ["Jenis Pekerjaan / SPK", "Visual Foto Dokumentasi 1", "Visual Foto Dokumentasi 2"]
        for col_num, header_text in enumerate(headers_foto, 1):
            cell_h = worksheet_foto.cell(row=1, column=col_num, value=header_text)
            cell_h.fill = header_fill
            cell_h.font = header_font
            cell_h.alignment = align_center
            cell_h.border = border_standard

        # Map untuk menyimpan lokasi tujuan hyperlink (Key: "No", Value: "No Baris Excel Sheet Foto")
        job_map_targets = {}
        
        # Loop data untuk menyisipkan gambar fisik
        foto_row_idx = 2
        
        for index, row in df_excel.iterrows():
            # Tulis Judul Pekerjaan (Kolom A)
            judul_gabungan = f"SPK: {row['Nomor SPK']}\n\nPekerjaan: {row['Jenis Pekerjaan']}"
            cell_j = worksheet_foto.cell(row=foto_row_idx, column=1, value=judul_gabungan)
            cell_j.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
            cell_j.border = border_standard
            
            # --- PERBAIKAN 2: Simpan Target Lokasi untuk Hyperlink ---
            # Kita simpan Nomor Baris Excel saat ini di Sheet Foto untuk setiap 'No' urut
            if 'No' in row:
                job_map_targets[row['No']] = foto_row_idx

            # Set Tinggi Baris agar foto muat
            worksheet_foto.row_dimensions[foto_row_idx].height = 250

            # Fungsi Helper untuk menyisipkan satu gambar fisik (dengan resize)
            def insert_image_visual_resized(path, ws, current_row, current_col, target_col_width):
                cell_p = ws.cell(row=current_row, column=current_col)
                cell_p.border = border_standard
                
                if path and os.path.exists(str(path)) and has_pil:
                    try:
                        # Resize proporsional: set tinggi 300px, lebar menyesuaikan
                        pil_img = PILImage.open(path)
                        orig_w, orig_h = pil_img.size
                        
                        # Hitung target pixel (konversi kasar karakter ke pixel, kurangi padding)
                        target_width_px = int((target_col_width * 7.5) - 5)
                        
                        # Hitung tinggi target secara proporsional
                        target_height_px = int((orig_h / orig_w) * target_width_px)
                        
                        pil_img_resized = pil_img.resize((target_width_px, target_height_px), PILImage.Resampling.LANCZOS)
                        
                        img_buffer = io.BytesIO()
                        img_format = pil_img.format if pil_img.format else 'JPEG'
                        pil_img_resized.save(img_buffer, format=img_format)
                        img_buffer.seek(0)
                        
                        opx_img = OpenPyXLImage(img_buffer)
                        col_letter = get_column_letter(current_col)
                        ws.add_image(opx_img, f'{col_letter}{current_row}')
                        
                    except Exception as e:
                        cell_p.value = f"Eror load gambar: {e}"
                        cell_p.alignment = align_center
                else:
                    cell_p.value = "Foto tidak tersedia / Pillow belum diinstal"
                    cell_p.alignment = align_center

            # Sisipkan Foto 1 (Kolom B = 2)
            insert_image_visual_resized(row['Foto 1'], worksheet_foto, foto_row_idx, 2, LEBAR_KOLOM_FOTO)
            
            # Sisipkan Foto 2 (Kolom C = 3)
            insert_image_visual_resized(row['Foto 2'], worksheet_foto, foto_row_idx, 3, LEBAR_KOLOM_FOTO)

            foto_row_idx += 1

        # ---------------------------------------------------------
        # --- PERBAIKAN 3: Buat Hyperlink di Sheet 'Laporan Progress' ---
        # ---------------------------------------------------------
        # Temukan indeks kolom 'No' dan 'Dokumentasi' (1-based untuk openpyxl)
        try:
            no_col_idx = df_progress.columns.get_loc('No') + 1
            doc_col_idx = df_progress.columns.get_loc('Dokumentasi') + 1
        except Exception:
            no_col_idx, doc_col_idx = None, None

        if no_col_idx and doc_col_idx:
            # Loop data di Sheet Laporan Progress mulai dari baris 2
            for p_row_idx in range(2, worksheet_progress.max_row + 1):
                # Ambil nilai 'No' di baris ini
                no_value = worksheet_progress.cell(row=p_row_idx, column=no_col_idx).value
                
                # Cek apakah 'No' ini ada di map target kita (apakah punya foto)
                if no_value in job_map_targets:
                    # Ambil baris tujuan di Sheet Foto
                    target_photo_row = job_map_targets[no_value]
                    
                    # Tulis teks link di kolom 'Dokumentasi'
                    cell_link = worksheet_progress.cell(row=p_row_idx, column=doc_col_idx, value="Lihat Foto")
                    
                    # Jadikan hyperlink internal mengarah ke Sheet Foto, Kolom A
                    cell_link.hyperlink = f"#'Foto Dokumentasi'!A{target_photo_row}"
                    
                    # Terapkan styling link biru
                    cell_link.font = blue_link_font
                    cell_link.alignment = align_center

    return output.getvalue()

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
                foto_1 AS [Foto 1],
                foto_2 AS [Foto 2],
                catatan AS [Catatan Pekerjaan Terbaru]
            FROM laporan_mingguan
            ORDER BY id ASC
        """

        with get_db_connection() as conn:
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
                    "Foto 1": st.column_config.ImageColumn("Pratinjau Foto 1", help="Foto dokumentasi 1"),
                    "Foto 2": st.column_config.ImageColumn("Pratinjau Foto 2", help="Foto dokumentasi 2")
                },
                key="editor_dashboard"
            )

            if st.button("💾 Simpan Perubahan & Hapus Data"):
                with get_db_connection() as conn:
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

            # ---------------------------------------------------------
            # EXPORT DATA KE EXCEL (STYLING OPENPYXL)
            # ---------------------------------------------------------
            st.markdown("---")
            st.subheader("📥 Export & Download Laporan")
            st.info("File Excel akan berisi 2 Sheet: Sheet 1 (Data Teks) dan Sheet 2 (Visual Foto).")

            try:
                excel_bytes = generate_excel_full_feature(df_view)
                st.download_button(
                    label="📥 Download Laporan (Excel)",
                    data=excel_bytes,
                    file_name="Laporan_Progress_dan_Foto.xlsx",
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
                foto_1 AS [Foto 1],
                foto_2 AS [Foto 2],
                catatan AS [Catatan Pada Tanggal Tersebut]
            FROM history_progress
            ORDER BY id DESC
        """
        with get_db_connection() as conn:
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

            st.dataframe(
                df_history_filtered, 
                use_container_width=True, 
                hide_index=True,
                column_config={
                    "Foto 1": st.column_config.ImageColumn("Dokumentasi 1"),
                    "Foto 2": st.column_config.ImageColumn("Dokumentasi 2")
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
                with get_db_connection() as conn:
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

    with get_db_connection() as conn:
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
        with get_db_connection() as conn:
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
                        SET no_spk = ?, kontraktor = ?, jenis_pekerjaan = ?, unit = ?, jumlah = ?, nilai_pekerjaan = ?
                        WHERE id = ?
                    """, (
                        row.get('Nomor SPK'), row.get('Nama Kontraktor'), row.get('Jenis Pekerjaan'),
                        row.get('Unit Proyek'), row.get('Jumlah'), row.get('Nilai Kontrak Pekerjaan Ini (Rp)'), real_id
                    ))

            conn.commit()
        st.success("✅ Master SPK berhasil diperbarui!")
        st.rerun()

# ---------------------------------------------------------
# MENU 3: INPUT PROGRESS MINGGUAN (DENGAN 2 FOTO UPLOAD)
# ---------------------------------------------------------
elif menu == "Input Progress Mingguan":
    st.title("📝 Input Progress Mingguan Berdasarkan SPK")

    with get_db_connection() as conn:
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

        with get_db_connection() as conn:
            job_list = pd.read_sql_query(
                "SELECT jenis_pekerjaan FROM master_spk WHERE no_spk = ?", 
                conn, 
                params=(selected_spk,)
            )['jenis_pekerjaan'].tolist()

        with col_job:
            selected_job = st.selectbox("Pilih Jenis Pekerjaan", job_list)

        with get_db_connection() as conn:
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

        with get_db_connection() as conn:
            try:
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

        # Menampilkan foto terkini jika sudah pernah diunggah
        if existing_foto_1 or existing_foto_2:
            st.markdown("**📸 Foto Dokumentasi Minggu Lalu/Terkini:**")
            img_col1, img_col2 = st.columns(2)
            with img_col1:
                if existing_foto_1 and os.path.exists(str(existing_foto_1)):
                    st.image(existing_foto_1, caption="Dokumentasi 1", use_container_width=True)
            with img_col2:
                if existing_foto_2 and os.path.exists(str(existing_foto_2)):
                    st.image(existing_foto_2, caption="Dokumentasi 2", use_container_width=True)

        with st.form("form_update_progress_mingguan"):
            col1, col2 = st.columns(2)

            with col1:
                st.text_input("Nama Kontraktor", value=spk_detail['kontraktor'], disabled=True)
                st.text_input("Unit Proyek", value=spk_detail['unit'], disabled=True)
                st.number_input("Jumlah", value=int(spk_detail['jumlah']), disabled=True)
                st.number_input("Nilai Pekerjaan (Rp)", value=float(nilai_peks), disabled=True)

                st.subheader("📷 Update Foto Dokumentasi Minggu Ini")
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
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    penambahan = float(prog_ini) - float(prog_lalu)

                    path_foto_1 = existing_foto_1
                    path_foto_2 = existing_foto_2

                    # Sanitasi nama untuk keamanan file system
                    spk_clean = sanitize_filename(selected_spk)
                    job_clean = sanitize_filename(selected_job)

                    # Simpan Foto 1 jika di-upload baru
                    if file_foto_1 is not None:
                        # Buat nama file unik dengan timestamp
                        import time
                        timestamp = int(time.time())
                        path_foto_1 = os.path.join(UPLOAD_DIR, f"{spk_clean}_{job_clean}_f1_{timestamp}.jpg")
                        with open(path_foto_1, "wb") as f:
                            f.write(file_foto_1.getbuffer())

                    # Simpan Foto 2 jika di-upload baru
                    if file_foto_2 is not None:
                        import time
                        timestamp = int(time.time())
                        path_foto_2 = os.path.join(UPLOAD_DIR, f"{spk_clean}_{job_clean}_f2_{timestamp}.jpg")
                        with open(path_foto_2, "wb") as f:
                            f.write(file_foto_2.getbuffer())

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

                    # Rekam ke History
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
                st.success(f"✅ Progress & Foto untuk '{selected_job}' BERHASIL DISIMPAN!")
                st.rerun()
