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
    # Password default adalah 123456, bisa diubah di streamlit secrets
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
def get_db_connection():
    return sqlite3.connect('proyek_v2.db')

# Helper function untuk membersihkan nama file dari karakter ilegal
def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", filename).replace(" ", "_")

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
            catatan TEXT,
            UNIQUE(no_spk, jenis_pekerjaan)
        )
    ''')

    # Migrasi otomatis jika kolom catatan belum ada
    cursor.execute("PRAGMA table_info(master_spk)")
    cols_mast = [col[1] for col in cursor.fetchall()]
    if 'catatan' not in cols_mast:
        cursor.execute("ALTER TABLE master_spk ADD COLUMN catatan TEXT")

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
    conn.commit()
    conn.close()

init_db()

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
        
        # Sisipkan Kolom 'Dokumentasi' Kosong setelah 'Catatan Pekerjaan Terbaru'
        try:
            target_col_idx = df_progress.columns.get_loc('Catatan Pekerjaan Terbaru') + 1
            df_progress.insert(target_col_idx, 'Dokumentasi', '') 
        except Exception:
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

        for col in worksheet_progress.columns:
            header_name = col[0].value
            if header_name != 'Dokumentasi':
                max_len = max(len(str(cell.value or '')) for cell in col)
                worksheet_progress.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 12)
            else:
                worksheet_progress.column_dimensions[get_column_letter(col[0].column)].width = 15

        # ---------------------------------------------------------
        # Layout & Penyisipan Gambar Visual yang Rapi (Sheet Foto)
        # ---------------------------------------------------------
        LEBAR_KOLOM_FOTO = 50
        worksheet_foto.column_dimensions['A'].width = 40 
        worksheet_foto.column_dimensions['B'].width = LEBAR_KOLOM_FOTO 
        worksheet_foto.column_dimensions['C'].width = LEBAR_KOLOM_FOTO 

        # Header Sheet Foto
        headers_foto = ["Jenis Pekerjaan / SPK", "Visual Foto Dokumentasi 1", "Visual Foto Dokumentasi 2"]
        for col_num, header_text in enumerate(headers_foto, 1):
            cell_h = worksheet_foto.cell(row=1, column=col_num, value=header_text)
            cell_h.fill = header_fill
            cell_h.font = header_font
            cell_h.alignment = align_center
            cell_h.border = border_standard

        # Map untuk menyimpan lokasi tujuan hyperlink
        job_map_targets = {}
        
        # Loop data untuk menyisipkan gambar fisik
        foto_row_idx = 2
        
        for index, row in df_excel.iterrows():
            # Tulis Judul Pekerjaan (Kolom A)
            judul_gabungan = f"SPK: {row['Nomor SPK']}\n\nPekerjaan: {row['Jenis Pekerjaan']}"
            cell_j = worksheet_foto.cell(row=foto_row_idx, column=1, value=judul_gabungan)
            cell_j.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
            cell_j.border = border_standard
            
            if 'No' in row:
                job_map_targets[row['No']] = foto_row_idx

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
                        
                        target_width_px = int((target_col_width * 7.5) - 5)
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

            insert_image_visual_resized(row['Foto 1'], worksheet_foto, foto_row_idx, 2, LEBAR_KOLOM_FOTO)
            insert_image_visual_resized(row['Foto 2'], worksheet_foto, foto_row_idx, 3, LEBAR_KOLOM_FOTO)

            foto_row_idx += 1

        # Buat Hyperlink di Sheet 'Laporan Progress'
        try:
            no_col_idx = df_progress.columns.get_loc('No') + 1
            doc_col_idx = df_progress.columns.get_loc('Dokumentasi') + 1
        except Exception:
            no_col_idx, doc_col_idx = None, None

        if no_col_idx and doc_col_idx:
            for p_row_idx in range(2, worksheet_progress.max_row + 1):
                no_value = worksheet_progress.cell(row=p_row_idx, column=no_col_idx).value
                if no_value in job_map_targets:
                    target_photo_row = job_map_targets[no_value]
                    cell_link = worksheet_progress.cell(row=p_row_idx, column=doc_col_idx, value="Lihat Foto")
                    cell_link.hyperlink = f"#'Foto Dokumentasi'!A{target_photo_row}"
                    cell_link.font = blue_link_font
                    cell_link.alignment = align_center

    return output.getvalue()

# ---------------------------------------------------------
# NAVIGASI SIDEBAR
# ---------------------------------------------------------
st.sidebar.title("Navigasi")
# DEFINE MENU NAMES AS VARIABLES FOR CONSISTENCY
MENU_DASHBOARD = "Dashboard Progress"
MENU_INPUT = "Input Progress Mingguan"
MENU_MASTER = "Kelola Master SPK" # INI NAMA MENU YANG DIPERBAIKI

menu = st.sidebar.selectbox("Pilih Menu", [
    MENU_DASHBOARD, 
    MENU_INPUT,
    MENU_MASTER
])

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Logout"):
    st.session_state["authenticated"] = False
    st.rerun()

# ---------------------------------------------------------
# MENU 1: DASHBOARD PROGRESS (BERDASARKAN WILAYAH)
# ---------------------------------------------------------
if menu == MENU_DASHBOARD:
    st.title("📊 WEEKLY REPORT DIVISI INFRASTRUCTURE")

    tab_bangka, tab_belitung, tab_semua, tab_history = st.tabs([
        "🏝️ Laporan Progress Bangka", 
        "🏖️ Laporan Progress Belitung", 
        "📋 Semua Progress Proyek", 
        "📜 Riwayat / History Update"
    ])

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
            catatan AS [Catatan Pekerjaan Terbaru],
            foto_1 AS [Pratinjau Foto 1],
            foto_2 AS [Pratinjau Foto 2]
        FROM laporan_mingguan
        ORDER BY id ASC
    """

    with get_db_connection() as conn:
        try:
            df_all = pd.read_sql_query(query_view, conn)
        except Exception:
            df_all = pd.DataFrame()

    # Function Helper untuk Menampilkan Tabel & Tombol Export per Wilayah
    def render_dashboard_table(df_data, tab_key_prefix):
        if df_data.empty:
            st.info("💡 Belum ada data progress untuk wilayah/kategori ini.")
            return

        df_display = df_data.copy()
        if 'No' not in df_display.columns:
            df_display.insert(0, 'No', range(1, len(df_display) + 1))

        # Susunan kolom dengan foto di paling kanan
        column_order = [
            'No', 'Nomor SPK', 'Nama Kontraktor', 'Jenis Pekerjaan', 'Unit Proyek', 'Jumlah',
            'Nilai Kontrak Pekerjaan Ini (Rp)', 'Progress Minggu Lalu (%)', 'Progress Minggu Ini (%)',
            'Selisih / Varian (%)', 'Catatan Pekerjaan Terbaru', 'Pratinjau Foto 1', 'Pratinjau Foto 2'
        ]
        existing_cols = [c for c in column_order if c in df_display.columns]

        edited_df = st.data_editor(
            df_display[existing_cols],
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "real_id": None,
                "Pratinjau Foto 1": st.column_config.ImageColumn("Pratinjau Foto 1"),
                "Pratinjau Foto 2": st.column_config.ImageColumn("Pratinjau Foto 2"),
                "Nilai Kontrak Pekerjaan Ini (Rp)": st.column_config.NumberColumn(format="Rp %d"),
            },
            key=f"editor_{tab_key_prefix}"
        )

        if st.button("💾 Simpan Perubahan Data", key=f"btn_save_{tab_key_prefix}"):
            with get_db_connection() as conn:
                cursor = conn.cursor()
                for idx, row in edited_df.iterrows():
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

    # --- TAB 1: BANGKA ---
    with tab_bangka:
        st.subheader("📍 Laporan Progress Proyek - Wilayah Bangka")
        if not df_all.empty:
            df_bangka = df_all[df_all['Unit Proyek'].str.contains('BANGKA|BKA', case=False, na=False)]
            render_dashboard_table(df_bangka, "bangka")

    # --- TAB 2: BELITUNG ---
    with tab_belitung:
        st.subheader("📍 Laporan Progress Proyek - Wilayah Belitung")
        if not df_all.empty:
            df_belitung = df_all[df_all['Unit Proyek'].str.contains('BELITUNG|BLT', case=False, na=False)]
            render_dashboard_table(df_belitung, "belitung")

    # --- TAB 3: SEMUA ---
    with tab_semua:
        st.subheader("🌐 Semua Laporan Progress Proyek")
        render_dashboard_table(df_all, "semua")

    # --- TAB 4: HISTORY ---
    with tab_history:
        st.subheader("📜 Log Riwayat Update")
        with get_db_connection() as conn:
            df_history = pd.read_sql_query("SELECT * FROM history_progress ORDER BY waktu_input DESC", conn)
        st.dataframe(df_history, use_container_width=True)

# ---------------------------------------------------------
# MENU 2: INPUT PROGRESS (BERDASARKAN TAB WILAYAH)
# ---------------------------------------------------------
elif menu == MENU_INPUT:
    st.title("📝 Input Laporan Progress Mingguan Berdasarkan SPK")
    st.markdown("---")

    # Ambil data Master SPK untuk saringan
    with get_db_connection() as conn:
        df_master_all = pd.read_sql_query("SELECT * FROM master_spk", conn)

    if df_master_all.empty:
        st.warning("⚠️ Belum ada data Master SPK. Harap daftarkan SPK terlebih dahulu di menu Kelola Master SPK.")
    else:
        # --- PERBAIKAN: BUAT TAB WILAYAH UNTUK INPUT ---
        tab_i_bangka, tab_i_belitung = st.tabs([
            "🏝️ Input Progress Bangka", 
            "🏖️ Input Progress Belitung"
        ])

        # Function Helper internal untuk merender Form Input per Wilayah
        def render_input_form(df_master_wilayah, tab_key_prefix):
            if df_master_wilayah.empty:
                st.info("💡 Belum ada data master pekerjaan terdaftar untuk wilayah ini.")
                return

            # Siapkan Dropdown Pilihan SPK (Formatted Display)
            df_master_wilayah['display'] = df_master_wilayah['no_spk'] + ' - ' + df_master_wilayah['jenis_pekerjaan']
           # Function Helper internal untuk merender Form Input per Wilayah
        def render_input_form(df_master_wilayah, tab_key_prefix):
            if df_master_wilayah.empty:
                st.info("💡 Belum ada data master pekerjaan terdaftar untuk wilayah ini.")
                return

            # Siapkan Dropdown Pilihan SPK (Formatted Display)
            df_master_wilayah['display'] = df_master_wilayah['no_spk'] + ' - ' + df_master_wilayah['jenis_pekerjaan']
            
            selected_spk_text = st.selectbox(
                "Pilih SPK/Pekerjaan yang akan dilaporkan:", 
                df_master_wilayah['display'].tolist(),
                key=f"selectbox_spk_{tab_key_prefix}"
            )

            # Ambil detail SPK yang dipilih
            spk_data_selected = df_master_wilayah[df_master_wilayah['display'] == selected_spk_text].iloc[0]

            # --- PERBAIKAN FORMATTING ANGKA DI SINI ---
            try:
                val_num = float(spk_data_selected['nilai_pekerjaan'])
                nilai_formatted = f"Rp {val_num:,.2f}"
            except (ValueError, TypeError):
                nilai_formatted = "-"

            # Tampilkan Ringkasan Detail (Info)
            st.info(f"""📌 **Detail SPK:** 
*   Kontraktor: **{spk_data_selected['kontraktor']}**
*   Jenis Pekerjaan: **{spk_data_selected['jenis_pekerjaan']}**
*   Unit/Wilayah: **{spk_data_selected['unit']}**
*   Nilai Kontrak: **{nilai_formatted}**
""")

            # Ambil progress terakhir untuk SPK ini (dari db laporan_mingguan)
            prog_terakhir = 0.0
            catatan_terakhir = ""
            existing_foto_1 = None
            existing_foto_2 = None
            
            with get_db_connection() as conn:
                query_last = "SELECT progress_minggu_ini, catatan, foto_1, foto_2 FROM laporan_mingguan WHERE no_spk=? AND jenis_pekerjaan=?"
                existing_prog_df = pd.read_sql_query(query_last, conn, params=(spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan']))
                
                if not existing_prog_df.empty:
                    prog_terakhir = existing_prog_df.iloc[0]['progress_minggu_ini']
                    catatan_terakhir = existing_prog_df.iloc[0]['catatan'] or ""
                    existing_foto_1 = existing_prog_df.iloc[0]['foto_1']
                    existing_foto_2 = existing_prog_df.iloc[0]['foto_2']

            # Tampilkan foto terakhir jika ada
            if existing_foto_1 or existing_foto_2:
                st.markdown("**📸 Pratinjau Foto Dokumentasi Terakhir:**")
                c_img1, c_img2 = st.columns(2)
                with c_img1:
                    if existing_foto_1 and os.path.exists(str(existing_foto_1)):
                        st.image(existing_foto_1, caption="Foto Dokumentasi 1 (Minggu Lalu)", use_container_width=True)
                with c_img2:
                    if existing_foto_2 and os.path.exists(str(existing_foto_2)):
                        st.image(existing_foto_2, caption="Foto Dokumentasi 2 (Minggu Lalu)", use_container_width=True)

            # --- FORM INPUT WEEKLY REPORT ---
            with st.form(f"form_input_week_{tab_key_prefix}", clear_on_submit=True):
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("📝 Progress Minggu Ini")
                    prog_ini = st.number_input(
                        f"Progress Akumulatif Minggu Ini (%) - (Hingga Minggu Lalu: {prog_terakhir:.2f}%)", 
                        min_value=prog_terakhir, # Minimal harus sama dengan minggu lalu
                        max_value=100.0, 
                        value=prog_terakhir,
                        step=0.1
                    )
                    catatan_lap = st.text_area("Catatan/Kendala Pekerjaan Minggu Ini", value=catatan_terakhir)
                
                with col2:
                    st.subheader("📷 Update Foto Dokumentasi (Upload Baru)")
                    f_upload_1 = st.file_uploader("Upload Foto 1", type=["jpg", "jpeg", "png"], key=f"f1_{tab_key_prefix}")
                    f_upload_2 = st.file_uploader("Upload Foto 2", type=["jpg", "jpeg", "png"], key=f"f2_{tab_key_prefix}")
                
                # Jaga data foto lama jika tidak diupload baru
                path_f1_final = existing_foto_1
                path_f2_final = existing_foto_2

                submit_btn = st.form_submit_button(f"💾 Simpan Laporan Minggu Ini")
                
                if submit_btn:
                    if prog_ini < prog_lalu:
                        st.error("⚠️ Progress minggu ini tidak boleh lebih kecil dari minggu lalu (progress bersifat akumulatif)!")
                    else:
                        cursor = conn.cursor()
                        import time
                        ts = int(time.time())
                        # Sanitasi Nomor SPK untuk nama file yang aman
                        spk_fniz = sanitize_filename(spk_data_selected['no_spk'])
                        
                        # Simpan Foto 1 jika di-upload baru
                        if f_upload_1:
                            path_f1_final = os.path.join(UPLOAD_DIR, f"{spk_fniz}_f1_{ts}.jpg")
                            with open(path_f1_final, "wb") as f: f.write(f_upload_1.getbuffer())

                        # Simpan Foto 2 jika di-upload baru
                        if f_upload_2:
                            path_f2_final = os.path.join(UPLOAD_DIR, f"{spk_fniz}_f2_{ts}.jpg")
                            with open(path_f2_final, "wb") as f: f.write(f_upload_2.getbuffer())
                        
                        # Hitung penambahan minggu ini (varian)
                        penambahan_week = prog_ini - prog_terakhir

                        # Masukkan/Overwrite ke Laporan Mingguan Utama
                        with get_db_connection() as conn:
                            cursor = conn.cursor()
                            cursor.execute("DELETE FROM laporan_mingguan WHERE no_spk=? AND jenis_pekerjaan=?", (spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan']))
                            cursor.execute("""
                                INSERT INTO laporan_mingguan (
                                    no_spk, jenis_pekerjaan, kontraktor, unit, jumlah, nilai_pekerjaan, 
                                    progress_minggu_lalu, progress_minggu_ini, catatan, foto_1, foto_2, waktu_input
                                ) VALUES (?,?,?,?,?,?,?,?,?,?,?, CURRENT_TIMESTAMP)""", 
                                (spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan'], spk_data_selected['kontraktor'], spk_data_selected['unit'], spk_data_selected['jumlah'], spk_data_selected['nilai_pekerjaan'], prog_terakhir, prog_ini, catatan_lap, path_f1_final, path_f2_final))
                            
                            # Rekam ke History
                            cursor.execute("""
                                INSERT INTO history_progress (
                                    no_spk, jenis_pekerjaan, kontraktor, unit, 
                                    progress_minggu_lalu, progress_minggu_ini, progres_penambahan, catatan, foto_1, foto_2, waktu_input
                                ) VALUES (?,?,?,?,?,?,?,?,?,?, CURRENT_TIMESTAMP)""",
                                (spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan'], spk_data_selected['kontraktor'], spk_data_selected['unit'], prog_terakhir, prog_ini, penambahan_week, catatan_lap, path_f1_final, path_f2_final))
                            
                            conn.commit()
                        st.success(f"✅ Laporan mingguan untuk '{selected_spk_text}' berhasil disimpan!")
                        st.rerun()

# --- TAB 1: INPUT PROGRESS BANGKA ---
        with tab_i_bangka:
            st.subheader("🏝️ Input Progress - Wilayah Bangka")
            # Saring Master SPK untuk Bangka (Unit mengandung BKA/BANGKA)
            df_bangka_master = df_master_all[df_master_all['unit'].str.contains('BANGKA|BKA', case=False, na=False)]
            render_input_form(df_bangka_master, "bangka")

        # --- TAB 2: INPUT PROGRESS BELITUNG ---
        with tab_i_belitung:
            st.subheader("🏖️ Input Progress - Wilayah Belitung")
            # Saring Master SPK untuk Belitung (Unit mengandung BLT/BELITUNG)
            df_belitung_master = df_master_all[df_master_all['unit'].str.contains('BELITUNG|BLT', case=False, na=False)]
            render_input_form(df_belitung_master, "belitung")     
# ---------------------------------------------------------
# MENU 3: KELOLA MASTER (DENGAN TAB WILAYAH)
# ---------------------------------------------------------
elif menu == MENU_MASTER: # NAMA MENU KONSISTEN DENGAN SIDEBAR
    st.title("⚙️ Kelola Master Data Pekerjaan / SPK")

    tab_m_bangka, tab_m_belitung, tab_m_tambah = st.tabs([
        "🏝️ Master Data Bangka", 
        "🏖️ Master Data Belitung", 
        "➕ Tambah SPK / Pekerjaan Baru"
    ])

    query_master = """
        SELECT 
            id AS real_id,
            no_spk AS [Nomor SPK],
            kontraktor AS [Nama Kontraktor],
            jenis_pekerjaan AS [Jenis Pekerjaan],
            unit AS [Unit Proyek],
            jumlah AS [Jumlah],
            nilai_pekerjaan AS [Nilai Kontrak (Rp)],
            catatan AS [Catatan / Keterangan]
        FROM master_spk
        ORDER BY id ASC
    """

    with get_db_connection() as conn:
        try:
            df_master = pd.read_sql_query(query_master, conn)
        except Exception:
            df_master = pd.DataFrame()

    # Helper untuk menampilkan tabel master
    def render_master_table(df_data, tab_key_prefix):
        if df_data.empty:
            st.info("Belum ada data master untuk wilayah ini.")
            return

        edited_df = st.data_editor(
            df_data.drop(columns=['real_id'], errors='ignore'),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Nilai Kontrak (Rp)": st.column_config.NumberColumn(format="Rp %d"),
            },
            key=f"editor_m_{tab_key_prefix}"
        )

        # Tombol Simpan Edit & Hapus
        col_s, col_d = st.columns([2,2])
        
        with col_s:
            if st.button("💾 Simpan Perubahan Master", key=f"btn_save_m_{tab_key_prefix}"):
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    for idx, row in edited_df.iterrows():
                        # Ambil real_id asli
                        real_id = df_data.iloc[idx]['real_id']
                        cursor.execute("""
                            UPDATE master_spk
                            SET no_spk=?, kontraktor=?, jenis_pekerjaan=?, unit=?, jumlah=?, nilai_pekerjaan=?, catatan=?
                            WHERE id=?
                        """, (row['Nomor SPK'], row['Nama Kontraktor'], row['Jenis Pekerjaan'], row['Unit Proyek'], row['Jumlah'], row['Nilai Kontrak (Rp)'], row['Catatan / Keterangan'], real_id))
                    conn.commit()
                st.success("Master data diperbarui.")
                st.rerun()

        with col_d:
            # Fitur Hapus data master
            spk_to_del = st.selectbox("Pilih SPK yang akan dihapus:", ["-- Pilih SPK --"] + df_data['Nomor SPK'].tolist(), key=f"select_del_m_{tab_key_prefix}")
            if st.button("🗑️ Hapus SPK Dipilih", key=f"btn_del_m_{tab_key_prefix}", type="primary"):
                if spk_to_del != "-- Pilih SPK --":
                    with get_db_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM master_spk WHERE no_spk=?", (spk_to_del,))
                        conn.commit()
                    st.success(f"SPK {spk_to_del} berhasil dihapus.")
                    st.rerun()

    # --- TAB 1: MASTER BANGKA ---
    with tab_m_bangka:
        st.subheader("📍 Master Data - Wilayah Bangka")
        if not df_master.empty:
            # Filter unit mengandung BANGKA atau BKA
            df_m_bangka = df_master[df_master['Unit Proyek'].str.contains('BANGKA|BKA', case=False, na=False)]
            render_master_table(df_m_bangka, "bangka")

    # --- TAB 2: MASTER BELITUNG ---
    with tab_m_belitung:
        st.subheader("📍 Master Data - Wilayah Belitung")
        if not df_master.empty:
            # Filter unit mengandung BELITUNG atau BLT
            df_m_belitung = df_master[df_master['Unit Proyek'].str.contains('BELITUNG|BLT', case=False, na=False)]
            render_master_table(df_m_belitung, "belitung")

    # --- TAB 3: TAMBAH SPK ---
    with tab_m_tambah:
        st.subheader("➕ Form Tambah SPK Baru")
        with st.form("form_tambah_master", clear_on_submit=True):
            col_a, col_b = st.columns(2)
            with col_a:
                wilayah = st.selectbox("Pilih Wilayah Proyek:", ["BANGKA", "BELITUNG"])
                no_spk_input = st.text_input("Nomor SPK", placeholder="Contoh: 001/SPK/BANGKA/BPRE/2026")
                kontraktor_input = st.text_input("Nama Kontraktor")
                jenis_input = st.text_area("Jenis Pekerjaan")
            with col_b:
                unit_input = st.text_input("Unit Proyek", value=f"UNIT {wilayah}")
                jumlah_input = st.number_input("Jumlah Unit/Lokasi", min_value=1, value=1)
                nilai_input = st.number_input("Nilai Kontrak (Rp)", min_value=0.0, step=1000000.0)
                catatan_input = st.text_input("Catatan Tambahan")

            if st.form_submit_button("➕ Tambahkan ke Master Data"):
                if not no_spk_input or not jenis_input or not kontraktor_input:
                    st.error("Wajib mengisi Nomor SPK, Kontraktor, dan Jenis Pekerjaan.")
                else:
                    with get_db_connection() as conn:
                        try:
                            cursor = conn.cursor()
                            cursor.execute("""
                                INSERT INTO master_spk (no_spk, kontraktor, jenis_pekerjaan, unit, jumlah, nilai_pekerjaan, catatan)
                                VALUES (?,?,?,?,?,?,?)""", 
                                (no_spk_input, kontraktor_input, jenis_input, unit_input, jumlah_input, nilai_input, catatan_input))
                            conn.commit()
                            st.success(f"Master SPK {no_spk_input} wilayah {wilayah} berhasil ditambahkan.")
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error(f"⚠️ SPK {no_spk_input} sudah terdaftar.")
