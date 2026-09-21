import os
import io
import sqlite3
import pandas as pd
import streamlit as st
import re

# Modul untuk styling dan export Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- LIBRARY UNTUK MEMPROSES GAMBAR ---
try:
    from PIL import Image as PILImage
    from openpyxl.drawing.image import Image as OpenPyXLImage
    has_pil = True
except ImportError:
    st.error("⚠️ Library 'Pillow' belum terinstal. Gambar fisik tidak akan muncul di Excel. Silakan instal dengan perintah: pip install Pillow")
    has_pil = False

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
def get_db_connection():
    return sqlite3.connect('proyek_v2.db')

def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", str(filename)).replace(" ", "_")

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
# FUNGSI EXPORT EXCEL
# ==========================================
def generate_excel_full_feature(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_excel = df.drop(columns=['real_id'], errors='ignore').copy()
        
        df_progress = df_excel.drop(
            columns=['Foto 1', 'Foto 2', 'Pratinjau Foto 1', 'Pratinjau Foto 2'], 
            errors='ignore'
        ).copy()
        
        try:
            target_col_idx = df_progress.columns.get_loc('Catatan Pekerjaan Terbaru') + 1
            df_progress.insert(target_col_idx, 'Dokumentasi', '') 
        except Exception:
            df_progress['Dokumentasi'] = ''

        df_progress.to_excel(writer, index=False, sheet_name='Laporan Progress')
        
        worksheet_progress = writer.sheets['Laporan Progress']
        workbook = writer.book
        worksheet_foto = workbook.create_sheet(title='Foto Dokumentasi')
        
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        border_standard = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        align_center = Alignment(horizontal="center", vertical="center")
        blue_link_font = Font(color="0000FF", underline="single")

        for col_num in range(1, worksheet_progress.max_column + 1):
            cell = worksheet_progress.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_standard

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

        LEBAR_KOLOM_FOTO = 50
        worksheet_foto.column_dimensions['A'].width = 40 
        worksheet_foto.column_dimensions['B'].width = LEBAR_KOLOM_FOTO 
        worksheet_foto.column_dimensions['C'].width = LEBAR_KOLOM_FOTO 

        headers_foto = ["Jenis Pekerjaan / SPK", "Visual Foto Dokumentasi 1", "Visual Foto Dokumentasi 2"]
        for col_num, header_text in enumerate(headers_foto, 1):
            cell_h = worksheet_foto.cell(row=1, column=col_num, value=header_text)
            cell_h.fill = header_fill
            cell_h.font = header_font
            cell_h.alignment = align_center
            cell_h.border = border_standard

        job_map_targets = {}
        foto_row_idx = 2
        
        for index, row in df_excel.iterrows():
            judul_gabungan = f"SPK: {row.get('Nomor SPK', '')}\n\nPekerjaan: {row.get('Jenis Pekerjaan', '')}"
            cell_j = worksheet_foto.cell(row=foto_row_idx, column=1, value=judul_gabungan)
            cell_j.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
            cell_j.border = border_standard
            
            if 'No' in row:
                job_map_targets[row['No']] = foto_row_idx

            worksheet_foto.row_dimensions[foto_row_idx].height = 250

            def insert_image_visual_resized(path, ws, current_row, current_col, target_col_width):
                cell_p = ws.cell(row=current_row, column=current_col)
                cell_p.border = border_standard
                
                if path and os.path.exists(str(path)) and has_pil:
                    try:
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
                    cell_p.value = "Foto tidak tersedia"
                    cell_p.alignment = align_center

            insert_image_visual_resized(row.get('Pratinjau Foto 1'), worksheet_foto, foto_row_idx, 2, LEBAR_KOLOM_FOTO)
            insert_image_visual_resized(row.get('Pratinjau Foto 2'), worksheet_foto, foto_row_idx, 3, LEBAR_KOLOM_FOTO)

            foto_row_idx += 1

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
MENU_DASHBOARD = "Dashboard Progress"
MENU_INPUT = "Input Progress Mingguan"
MENU_MASTER = "Kelola Master SPK"

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
# MENU 1: DASHBOARD PROGRESS
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
            CAST(COALESCE(jumlah, 1) AS INTEGER) AS [Jumlah],
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

    def render_dashboard_table(df_data, tab_key_prefix):
        column_order = [
            'No', 'Nomor SPK', 'Nama Kontraktor', 'Jenis Pekerjaan', 'Unit Proyek', 'Jumlah',
            'Nilai Kontrak Pekerjaan Ini (Rp)', 'Progress Minggu Lalu (%)', 'Progress Minggu Ini (%)',
            'Selisih / Varian (%)', 'Catatan Pekerjaan Terbaru', 'Pratinjau Foto 1', 'Pratinjau Foto 2'
        ]

        if df_data.empty:
            st.info("💡 Belum ada data progress untuk wilayah/kategori ini.")
            df_display = pd.DataFrame(columns=['real_id'] + column_order)
        else:
            df_display = df_data.copy().reset_index(drop=True)
            if 'No' not in df_display.columns:
                df_display.insert(0, 'No', range(1, len(df_display) + 1))

        existing_cols = [c for c in ['real_id'] + column_order if c in df_display.columns]

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

        # -------------------------------------------------------------------------
        # DETEKSI BARIS DIHAPUS & PROSES HAPUS PERMANEN
        # -------------------------------------------------------------------------
        if not df_data.empty:
            deleted_indices = []
            if editor_key in st.session_state and "deleted_rows" in st.session_state[editor_key]:
                deleted_indices = st.session_state[editor_key]["deleted_rows"]

            if deleted_indices:
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    for idx in deleted_indices:
                        if idx < len(df_display):
                            row_to_del = df_display.iloc[idx]
                            real_id = row_to_del.get('real_id')
                            no_spk = str(row_to_del.get('Nomor SPK', '')).strip()
                            j_pek = str(row_to_del.get('Jenis Pekerjaan', '')).strip()

                            if pd.notna(real_id):
                                cursor.execute("DELETE FROM laporan_mingguan WHERE id = ?", (int(real_id),))

                            if no_spk:
                                cursor.execute("""
                                    DELETE FROM laporan_mingguan 
                                    WHERE LOWER(TRIM(no_spk)) = LOWER(?) 
                                       OR (LOWER(TRIM(no_spk)) = LOWER(?) AND LOWER(TRIM(jenis_pekerjaan)) = LOWER(?))
                                """, (no_spk, no_spk, j_pek))

                                cursor.execute("""
                                    DELETE FROM master_spk 
                                    WHERE LOWER(TRIM(no_spk)) = LOWER(?)
                                """, (no_spk,))

                                cursor.execute("""
                                    DELETE FROM history_progress 
                                    WHERE LOWER(TRIM(no_spk)) = LOWER(?)
                                """, (no_spk,))

                    conn.commit()

                if editor_key in st.session_state:
                    del st.session_state[editor_key]

                st.success("✅ Data berhasil dibersihkan permanen dari database!")
                st.rerun()

            # -------------------------------------------------------------------------
            # TOMBOL SIMPAN UNTUK EDIT NILAI / TEXT
            # -------------------------------------------------------------------------
            if st.button("💾 Simpan Perubahan Data", key=f"btn_save_{tab_key_prefix}"):
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    for idx, row in edited_df.iterrows():
                        if idx < len(df_display):
                            real_id = df_display.iloc[idx].get('real_id')
                            if pd.notna(real_id):
                                cursor.execute("""
                                    UPDATE laporan_mingguan
                                    SET progress_minggu_ini = ?, catatan = ?
                                    WHERE id = ?
                                """, (row.get('Progress Minggu Ini (%)'), row.get('Catatan Pekerjaan Terbaru'), int(real_id)))
                    conn.commit()

                st.success("✅ Perubahan data berhasil disimpan!")
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

    # --- TAB BANGKA ---
    with tab_bangka:
        st.subheader("📍 Laporan Progress Proyek - Wilayah Bangka")
        if not df_all.empty:
            df_bangka = df_all[df_all['Unit Proyek'].astype(str).str.contains('BANGKA|BKA', case=False, na=False)]
            render_dashboard_table(df_bangka, "bangka")

    # --- TAB BELITUNG ---
    with tab_belitung:
        st.subheader("📍 Laporan Progress Proyek - Wilayah Belitung")
        if not df_all.empty:
            pola_belitung = 'BELITUNG|BLT|BPSL|BPRE|BPT'
            df_belitung = df_all[
                df_all['Unit Proyek'].astype(str).str.contains(pola_belitung, case=False, na=False) |
                (~df_all['Unit Proyek'].astype(str).str.contains('BANGKA|BKA', case=False, na=False))
            ]
            render_dashboard_table(df_belitung, "belitung")

    # --- TAB SEMUA ---
    with tab_semua:
        st.subheader("🌐 Semua Laporan Progress Proyek")
        render_dashboard_table(df_all, "semua")

    # --- TAB RIWAYAT ---
    with tab_history:
        st.subheader("📜 Log Riwayat Update")
        with get_db_connection() as conn:
            df_history = pd.read_sql_query("SELECT * FROM history_progress ORDER BY waktu_input DESC", conn)
        
        st.dataframe(
            df_history, 
            use_container_width=True,
            column_config={
                "progress_minggu_lalu": st.column_config.NumberColumn(format="%.2f %%"),
                "progress_minggu_ini": st.column_config.NumberColumn(format="%.2f %%"),
                "progres_penambahan": st.column_config.NumberColumn(format="%.2f %%"),
            }
        )

        if not df_history.empty:
            st.markdown("---")
            st.subheader("🗑️ Pengelolaan Data Riwayat Log")
            
            col_del_single, col_del_all = st.columns([2, 1])

            with col_del_single:
                id_pilihan = st.selectbox(
                    "Pilih ID Log Riwayat yang ingin dihapus:",
                    df_history['id'].tolist(),
                    key="select_history_id_del"
                )
                if st.button("🗑️ Hapus ID Dipilih", type="secondary", key="btn_del_single_hist"):
                    with get_db_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM history_progress WHERE id = ?", (id_pilihan,))
                        conn.commit()
                    st.success(f"Data riwayat ID {id_pilihan} berhasil dihapus!")
                    st.rerun()

            with col_del_all:
                st.write("")
                st.write("")
                if st.button("🚨 Hapus Semua Log Riwayat", type="primary", key="btn_del_all_hist"):
                    with get_db_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM history_progress")
                        conn.commit()
                    st.success("Seluruh data riwayat berhasil dibersihkan!")
                    st.rerun()

# ---------------------------------------------------------
# MENU 2: INPUT PROGRESS
# ---------------------------------------------------------
elif menu == MENU_INPUT:
    st.title("📝 Input Laporan Progress Mingguan Berdasarkan SPK")
    st.markdown("---")

    with get_db_connection() as conn:
        df_master_all = pd.read_sql_query("SELECT * FROM master_spk", conn)

    tab_i_bangka, tab_i_belitung = st.tabs([
        "🏝️ Input Progress Bangka", 
        "🏖️ Input Progress Belitung"
    ])

    def render_input_form(df_master_wilayah, tab_key_prefix):
        if df_master_wilayah.empty:
            st.info("💡 Belum ada data master pekerjaan terdaftar untuk wilayah ini. Silakan daftarkan SPK terlebih dahulu di menu Kelola Master SPK.")
            return

        list_spk_unique = sorted(df_master_wilayah['no_spk'].unique().tolist())
        selected_spk_no = st.selectbox(
            "Pilih Nomor SPK:",
            list_spk_unique,
            key=f"select_spk_no_{tab_key_prefix}"
        )

        df_spk_filtered = df_master_wilayah[df_master_wilayah['no_spk'] == selected_spk_no]

        list_pekerjaan = df_spk_filtered['jenis_pekerjaan'].unique().tolist()
        selected_pekerjaan = st.selectbox(
            "Pilih Jenis Pekerjaan (Tergroup berdasarkan SPK):",
            list_pekerjaan,
            key=f"select_pekerjaan_{tab_key_prefix}"
        )

        spk_data_selected = df_spk_filtered[df_spk_filtered['jenis_pekerjaan'] == selected_pekerjaan].iloc[0]

        try:
            val_num = float(spk_data_selected['nilai_pekerjaan'])
            nilai_formatted = f"Rp {val_num:,.2f}"
        except (ValueError, TypeError):
            nilai_formatted = "-"

        st.info(f"""📌 **Detail SPK Dipilih:** 
*   **Nomor SPK:** {spk_data_selected['no_spk']}
*   **Kontraktor:** {spk_data_selected['kontraktor']}
*   **Jenis Pekerjaan:** {spk_data_selected['jenis_pekerjaan']}
*   **Unit/Wilayah:** {spk_data_selected['unit']}
*   **Jumlah:** {int(spk_data_selected['jumlah'] or 1)}
*   **Nilai Kontrak:** {nilai_formatted}
""")

        prog_terakhir = 0.0
        catatan_terakhir = ""
        existing_foto_1 = None
        existing_foto_2 = None
        
        with get_db_connection() as conn:
            query_last = "SELECT progress_minggu_ini, catatan, foto_1, foto_2 FROM laporan_mingguan WHERE no_spk=? AND jenis_pekerjaan=?"
            existing_prog_df = pd.read_sql_query(query_last, conn, params=(spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan']))
            
            if not existing_prog_df.empty:
                prog_terakhir = float(existing_prog_df.iloc[0]['progress_minggu_ini'] or 0.0)
                catatan_terakhir = existing_prog_df.iloc[0]['catatan'] or ""
                existing_foto_1 = existing_prog_df.iloc[0]['foto_1']
                existing_foto_2 = existing_prog_df.iloc[0]['foto_2']

        if existing_foto_1 or existing_foto_2:
            st.markdown("**📸 Pratinjau Foto Dokumentasi Terakhir:**")
            c_img1, c_img2 = st.columns(2)
            with c_img1:
                if existing_foto_1 and os.path.exists(str(existing_foto_1)):
                    st.image(existing_foto_1, caption="Foto Dokumentasi 1 (Minggu Lalu)", use_container_width=True)
            with c_img2:
                if existing_foto_2 and os.path.exists(str(existing_foto_2)):
                    st.image(existing_foto_2, caption="Foto Dokumentasi 2 (Minggu Lalu)", use_container_width=True)

        with st.form(f"form_input_week_{tab_key_prefix}", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📝 Progress Minggu Ini")
                prog_ini = st.number_input(
                    f"Progress Akumulatif Minggu Ini (%) - (Hingga Minggu Lalu: {prog_terakhir:.2f}%)", 
                    min_value=prog_terakhir,
                    max_value=100.0, 
                    value=prog_terakhir,
                    step=0.01,
                    format="%.2f"
                )
                catatan_lap = st.text_area("Catatan/Kendala Pekerjaan Minggu Ini", value=catatan_terakhir)
            
            with col2:
                st.subheader("📷 Update Foto Dokumentasi (Upload Baru)")
                f_upload_1 = st.file_uploader("Upload Foto 1", type=["jpg", "jpeg", "png"], key=f"f1_{tab_key_prefix}")
                f_upload_2 = st.file_uploader("Upload Foto 2", type=["jpg", "jpeg", "png"], key=f"f2_{tab_key_prefix}")
            
            path_f1_final = existing_foto_1
            path_f2_final = existing_foto_2

            submit_btn = st.form_submit_button("💾 Simpan Laporan Minggu Ini")
            
            if submit_btn:
                if prog_ini < prog_terakhir:
                    st.error("⚠️ Progress minggu ini tidak boleh lebih kecil dari minggu lalu (progress bersifat akumulatif)!")
                else:
                    import time
                    ts = int(time.time())
                    spk_fniz = sanitize_filename(spk_data_selected['no_spk'])
                    
                    if f_upload_1:
                        path_f1_final = os.path.join(UPLOAD_DIR, f"{spk_fniz}_f1_{ts}.jpg")
                        with open(path_f1_final, "wb") as f: 
                            f.write(f_upload_1.getbuffer())

                    if f_upload_2:
                        path_f2_final = os.path.join(UPLOAD_DIR, f"{spk_fniz}_f2_{ts}.jpg")
                        with open(path_f2_final, "wb") as f: 
                            f.write(f_upload_2.getbuffer())
                    
                    penambahan_week = prog_ini - prog_terakhir

                    with get_db_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM laporan_mingguan WHERE no_spk=? AND jenis_pekerjaan=?", (spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan']))
                        cursor.execute("""
                            INSERT INTO laporan_mingguan (
                                no_spk, jenis_pekerjaan, kontraktor, unit, jumlah, nilai_pekerjaan, 
                                progress_minggu_lalu, progress_minggu_ini, catatan, foto_1, foto_2, waktu_input
                            ) VALUES (?,?,?,?,?,?,?,?,?,?,?, CURRENT_TIMESTAMP)""", 
                            (spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan'], spk_data_selected['kontraktor'], spk_data_selected['unit'], int(spk_data_selected['jumlah'] or 1), spk_data_selected['nilai_pekerjaan'], prog_terakhir, prog_ini, catatan_lap, path_f1_final, path_f2_final))
                        
                        cursor.execute("""
                            INSERT INTO history_progress (
                                no_spk, jenis_pekerjaan, kontraktor, unit, 
                                progress_minggu_lalu, progress_minggu_ini, progres_penambahan, catatan, foto_1, foto_2, waktu_input
                            ) VALUES (?,?,?,?,?,?,?,?,?,?, CURRENT_TIMESTAMP)""",
                            (spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan'], spk_data_selected['kontraktor'], spk_data_selected['unit'], prog_terakhir, prog_ini, penambahan_week, catatan_lap, path_f1_final, path_f2_final))
                        
                        conn.commit()
                    st.success(f"✅ Laporan mingguan untuk SPK '{selected_spk_no}' - '{selected_pekerjaan}' berhasil disimpan!")
                    st.rerun()

    with tab_i_bangka:
        st.subheader("🏝️ Input Progress - Wilayah Bangka")
        if df_master_all.empty:
            df_bangka_master = pd.DataFrame()
        else:
            df_bangka_master = df_master_all[df_master_all['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False)]
        render_input_form(df_bangka_master, "bangka")

    with tab_i_belitung:
        st.subheader("🏖️ Input Progress - Wilayah Belitung")
        if df_master_all.empty:
            df_belitung_master = pd.DataFrame()
        else:
            pola_belitung = 'BELITUNG|BLT|BPSL|BPRE|BPT'
            df_belitung_master = df_master_all[
                df_master_all['unit'].astype(str).str.contains(pola_belitung, case=False, na=False) |
                (~df_master_all['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False))
            ]
        render_input_form(df_belitung_master, "belitung")  

# ---------------------------------------------------------
# MENU 3: KELOLA MASTER
# ---------------------------------------------------------
elif menu == MENU_MASTER:
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
            CAST(COALESCE(jumlah, 1) AS INTEGER) AS [Jumlah],
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

    def render_master_table(df_data, tab_key_prefix):
        if df_data.empty:
            st.info("Belum ada data master untuk wilayah ini.")
            return

        edited_df = st.data_editor(
            df_data.drop(columns=['real_id'], errors='ignore'),
            use_container_width=True,
            hide_index=True,
            column_config={
                "Jumlah": st.column_config.NumberColumn("Jumlah", format="%d"),
                "Nilai Kontrak (Rp)": st.column_config.NumberColumn("Nilai Kontrak (Rp)", format="Rp %d"),
            },
            key=f"editor_m_{tab_key_prefix}"
        )

        col_s, col_d = st.columns([2, 2])
        
        with col_s:
            if st.button("💾 Simpan Perubahan Master", key=f"btn_save_m_{tab_key_prefix}"):
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    for idx, row in edited_df.iterrows():
                        real_id = df_data.iloc[idx]['real_id']
                        cursor.execute("""
                            UPDATE master_spk
                            SET no_spk=?, kontraktor=?, jenis_pekerjaan=?, unit=?, jumlah=?, nilai_pekerjaan=?, catatan=?
                            WHERE id=?
                        """, (row['Nomor SPK'], row['Nama Kontraktor'], row['Jenis Pekerjaan'], row['Unit Proyek'], int(row['Jumlah'] or 1), row['Nilai Kontrak (Rp)'], row['Catatan / Keterangan'], real_id))
                    conn.commit()
                st.success("Master data diperbarui.")
                st.rerun()

        with col_d:
            spk_to_del = st.selectbox("Pilih SPK yang akan dihapus:", ["-- Pilih SPK --"] + df_data['Nomor SPK'].unique().tolist(), key=f"select_del_m_{tab_key_prefix}")
            if st.button("🗑️ Hapus SPK Dipilih", key=f"btn_del_m_{tab_key_prefix}", type="primary"):
                if spk_to_del != "-- Pilih SPK --":
                    with get_db_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM master_spk WHERE LOWER(TRIM(no_spk)) = LOWER(?)", (spk_to_del.strip(),))
                        cursor.execute("DELETE FROM laporan_mingguan WHERE LOWER(TRIM(no_spk)) = LOWER(?)", (spk_to_del.strip(),))
                        cursor.execute("DELETE FROM history_progress WHERE LOWER(TRIM(no_spk)) = LOWER(?)", (spk_to_del.strip(),))
                        conn.commit()
                    st.success(f"✅ Master SPK '{spk_to_del}' beserta laporan terkait berhasil dihapus!")
                    st.rerun()
                else:
                    st.warning("⚠️ Silakan pilih Nomor SPK yang ingin dihapus terlebih dahulu.")

    # --- TAB MASTER BANGKA ---
    with tab_m_bangka:
        st.subheader("📍 Master Data SPK - Wilayah Bangka")
        if not df_master.empty:
            df_m_bangka = df_master[df_master['Unit Proyek'].astype(str).str.contains('BANGKA|BKA', case=False, na=False)]
            render_master_table(df_m_bangka, "m_bangka")
        else:
            st.info("Belum ada data master untuk wilayah Bangka.")

    # --- TAB MASTER BELITUNG ---
    with tab_m_belitung:
        st.subheader("📍 Master Data SPK - Wilayah Belitung")
        if not df_master.empty:
            pola_belitung = 'BELITUNG|BLT|BPSL|BPRE|BPT'
            df_m_belitung = df_master[
                df_master['Unit Proyek'].astype(str).str.contains(pola_belitung, case=False, na=False) |
                (~df_master['Unit Proyek'].astype(str).str.contains('BANGKA|BKA', case=False, na=False))
            ]
            render_master_table(df_m_belitung, "m_belitung")
        else:
            st.info("Belum ada data master untuk wilayah Belitung.")

    # --- TAB TAMBAH SPK BARU ---
    with tab_m_tambah:
        st.subheader("➕ Form Tambah SPK / Pekerjaan Baru")
        
        with st.form("form_tambah_master_spk", clear_on_submit=True):
            col_a, col_b = st.columns(2)
            
            with col_a:
                new_no_spk = st.text_input("Nomor SPK:", placeholder="Contoh: SPK/INFRA/2026/001")
                new_kontraktor = st.text_input("Nama Kontraktor:", placeholder="Contoh: PT. Sinar Bangka")
                new_jenis_pekerjaan = st.text_input("Jenis Pekerjaan:", placeholder="Contoh: Pekerjaan Cor Jalan")
            
            with col_b:
                new_unit = st.selectbox("Unit / Wilayah Proyek:", ["BANGKA", "BELITUNG", "BPSL", "BPRE", "BPT"])
                new_jumlah = st.number_input("Jumlah Unit/Volume:", min_value=1, value=1, step=1)
                new_nilai = st.number_input("Nilai Kontrak Pekerjaan (Rp):", min_value=0.0, value=0.0, step=1000000.0, format="%.2f")
            
            new_catatan = st.text_area("Catatan / Keterangan Tambahan:", placeholder="Opsional...")
            
            btn_submit_master = st.form_submit_button("➕ Simpan Master SPK Baru")
            
            if btn_submit_master:
                if not new_no_spk.strip() or not new_jenis_pekerjaan.strip():
                    st.error("⚠️ Nomor SPK dan Jenis Pekerjaan wajib diisi!")
                else:
                    try:
                        with get_db_connection() as conn:
                            cursor = conn.cursor()
                            cursor.execute("""
                                INSERT INTO master_spk (no_spk, kontraktor, jenis_pekerjaan, unit, jumlah, nilai_pekerjaan, catatan)
                                VALUES (?, ?, ?, ?, ?, ?, ?)
                            """, (new_no_spk.strip(), new_kontraktor.strip(), new_jenis_pekerjaan.strip(), new_unit, int(new_jumlah), float(new_nilai), new_catatan.strip()))
                            conn.commit()
                        st.success(f"✅ Master SPK '{new_no_spk}' - '{new_jenis_pekerjaan}' berhasil ditambahkan!")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("⚠️ Kombinasi Nomor SPK dan Jenis Pekerjaan tersebut sudah terdaftar di database!")
                    except Exception as e:
                        st.error(f"⚠️ Terjadi kesalahan saat menyimpan data: {e}")
