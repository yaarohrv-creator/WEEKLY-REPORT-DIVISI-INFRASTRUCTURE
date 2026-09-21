import os
import io
import sqlite3
import pandas as pd
import streamlit as st
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# KONFIGURASI HALAMAN & DIREKTORI
# ==========================================
st.set_page_config(page_title="Sistem Progress Proyek", layout="wide")

# Folder penyimpanan foto
UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# Logo Aplikasi
if os.path.exists("logo.png"):
    st.sidebar.image("logo.png", use_container_width=True)

# ==========================================
# 🔐 AUTENTIKASI PASSWORD
# ==========================================
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

# ==========================================
# 🗄️ FUNGSIONALITAS DATABASE
# ==========================================
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

# Helper untuk membersihkan nama file
def sanitize_filename(text):
    return re.sub(r'[^a-zA-Z0-9]', '_', text)

# ==========================================
# 📊 FUNGSI EXPORT EXCEL
# ==========================================
def generate_excel_interaktif(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # ---------------------------------------------------------
        # Persiapan Data
        # ---------------------------------------------------------
        # Buat kopi data untuk ekspor dan drop real_id
        df_excel = df.drop(columns=['real_id'], errors='ignore').copy()
        
        # Simpan file ke sheet "Laporan Progress" sebagai teks biasa
        df_excel.to_excel(writer, index=False, sheet_name='Laporan Progress')
        worksheet_progress = writer.sheets['Laporan Progress']
        
        # Buat sheet baru khusus untuk tautan dokumentasi foto
        workbook = writer.book
        worksheet_foto = workbook.create_sheet(title='Foto Dokumentasi')
        
        # Tambahkan data header ke sheet foto
        header_foto = ["Nomor SPK", "Jenis Pekerjaan", "Tautan Foto 1", "Tautan Foto 2"]
        worksheet_foto.append(header_foto)

        # ---------------------------------------------------------
        # Styling & Hyperlink
        # ---------------------------------------------------------
        # Style Header
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        border_standard = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        align_center = Alignment(horizontal="center", vertical="center")
        
        # Style untuk Hyperlink Biru
        blue_link_font = Font(color="0000FF", underline="single")

        # Indeks Kolom Foto (1-based)
        col_f1_idx = df_excel.columns.get_loc("Foto 1") + 1
        col_f2_idx = df_excel.columns.get_loc("Foto 2") + 1
        col_job_idx = df_excel.columns.get_loc("Jenis Pekerjaan") + 1
        col_spk_idx = df_excel.columns.get_loc("Nomor SPK") + 1

        # Variabel untuk menempelkan data ke Sheet Foto
        foto_sheet_row_current = 2
        
        # Simpan peta lokasi (Map) Jenis Pekerjaan -> Alamat Sel di Sheet Foto
        # Key: "SPK | Pekerjaan", Value: Alamat Sel Tautan Foto 1 di Sheet Foto
        job_map = {}

        # Loop setiap baris data di Sheet Progress
        for row_idx in range(2, worksheet_progress.max_row + 1):
            for col_idx in range(1, worksheet_progress.max_column + 1):
                cell_p = worksheet_progress.cell(row=row_idx, column=col_idx)
                cell_p.border = border_standard
                cell_p.alignment = Alignment(vertical="center")
                
                # Format Header Progress
                if row_idx == 1:
                    cell_p.fill = header_fill
                    cell_p.font = header_font
                    cell_p.alignment = Alignment(horizontal="center", wrap_text=True)

                # Ambil nilai Jenis Pekerjaan & SPK baris ini untuk map
                current_job_name = worksheet_progress.cell(row=row_idx, column=col_job_idx).value
                current_spk_id = worksheet_progress.cell(row=row_idx, column=col_spk_idx).value
                job_key = f"{current_spk_id} | {current_job_name}"
                
                path_1 = worksheet_progress.cell(row=row_idx, column=col_f1_idx).value
                path_2 = worksheet_progress.cell(row=row_idx, column=col_f2_idx).value

                # Pindahkan Data Foto ke Sheet 2 dan buat map hyperlink
                if job_key not in job_map and (path_1 or path_2):
                    # Tulis ke Sheet Foto
                    row_data_foto = [current_spk_id, current_job_name, path_1, path_2]
                    worksheet_foto.append(row_data_foto)
                    
                    # Simpan lokasi Tautan Foto 1 di Sheet Foto untuk hyperlink internal
                    job_map[job_key] = f"C{foto_sheet_row_current}"
                    foto_sheet_row_current += 1
                
                # Buat Hyperlink di Sheet 1 mengarah ke Sheet 2
                target_cell_address = job_map.get(job_key)

                if col_idx == col_f1_idx and path_1:
                    cell_p.value = "Lihat Foto 1"
                    if target_cell_address:
                        cell_p.hyperlink = f"#'Foto Dokumentasi'!{target_cell_address}"
                        cell_p.font = blue_link_font
                        cell_p.alignment = align_center

                elif col_idx == col_f2_idx and path_2:
                    cell_p.value = "Lihat Foto 2"
                    if target_cell_address:
                        # Link Foto 2 mengarah ke kolom D
                        cell_p.hyperlink = f"#'Foto Dokumentasi'!{target_cell_address.replace('C', 'D')}"
                        cell_p.font = blue_link_font
                        cell_p.alignment = align_center

        # Styling Sheet Foto
        for col_num in range(1, worksheet_foto.max_column + 1):
            cell_f = worksheet_foto.cell(row=1, column=col_num)
            cell_f.fill = header_fill
            cell_f.font = header_font
            cell_f.alignment = align_center
            cell_f.border = border_standard

        for row_idx in range(2, worksheet_foto.max_row + 1):
            for col_idx in range(1, worksheet_foto.max_column + 1):
                cell_f = worksheet_foto.cell(row=row_idx, column=col_idx)
                cell_f.border = border_standard
                cell_f.alignment = Alignment(vertical="center")
                if col_idx >= 3 and cell_f.value: # Kolom Tautan Foto
                    cell_f.font = blue_link_font
                    cell_f.alignment = align_center

        # Auto-adjust lebar kolom kedua sheet
        for worksheet in [worksheet_progress, worksheet_foto]:
            for col in worksheet.columns:
                max_length = 0
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except: pass
                worksheet.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

    return output.getvalue()

# ==========================================
# 🚀 TAMPILAN UTAMA
# ==========================================
st.title("📊 Sistem Laporan Progress Pekerjaan Infrastruktur")
st.markdown("---")

menu = st.sidebar.selectbox("Pilih Menu", ["Dashboard Progress", "Input Progress Mingguan", "Kelola Master SPK"])

if st.sidebar.button("🚪 Logout"):
    st.session_state["authenticated"] = False
    st.rerun()

# ------------------------------------------
# MENU 1: DASHBOARD PROGRESS
# ------------------------------------------
if menu == "Dashboard Progress":
    tab1, tab2 = st.tabs(["📌 Status Terkini", "📜 Riwayat"])
    
    with tab1:
        with get_db_connection() as conn:
            query = """
                SELECT id as real_id, waktu_input as [Waktu Input Terbaru], no_spk as [Nomor SPK], 
                       kontraktor as [Nama Kontraktor], jenis_pekerjaan as [Jenis Pekerjaan], 
                       progress_minggu_ini as [Progress Saat Ini (%)], foto_1 as [Foto 1], foto_2 as [Foto 2], 
                       catatan as [Catatan]
                FROM laporan_mingguan
                ORDER BY waktu_input DESC
            """
            df_view = pd.read_sql_query(query, conn)

        if df_view.empty:
            st.info("💡 Belum ada laporan minggu ini.")
        else:
            # Drop kolom ID internal untuk tampilan
            st.dataframe(
                df_view.drop(columns=['real_id'], errors='ignore'),
                use_container_width=True,
                column_config={
                    "Foto 1": st.column_config.ImageColumn("Pratinjau Foto 1"),
                    "Foto 2": st.column_config.ImageColumn("Pratinjau Foto 2")
                }
            )
            
            st.markdown("---")
            st.subheader("📥 Export Laporan ke Excel")
            st.info("File Excel akan berisi link yang jika diklik pada 'Lihat Foto 1/2' akan pindah ke sheet 'Foto Dokumentasi'.")
            
            excel_bytes = generate_excel_interaktif(df_view)
            st.download_button(
                label="📥 Download Excel (.xlsx)",
                data=excel_bytes,
                file_name="Laporan_Progress_Interaktif.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with tab2:
        with get_db_connection() as conn:
            df_history = pd.read_sql_query("SELECT * FROM history_progress ORDER BY waktu_input DESC", conn)
        st.dataframe(df_history, use_container_width=True)

# ------------------------------------------
# MENU 2: INPUT PROGRESS
# ------------------------------------------
elif menu == "Input Progress Mingguan":
    st.subheader("📝 Form Update Progress & Foto Mingguan")
    
    with get_db_connection() as conn:
        spk_list = pd.read_sql_query("SELECT id, no_spk || ' - ' || jenis_pekerjaan as display FROM master_spk", conn)

    if spk_list.empty:
        st.warning("⚠️ Daftarkan SPK terlebih dahulu.")
    else:
        selected_spk_text = st.selectbox("Pilih SPK/Pekerjaan", spk_list['display'])
        selected_spk_id = spk_list[spk_list['display'] == selected_spk_text]['id'].values[0]

        with get_db_connection() as conn:
            spk_data = pd.read_sql_query("SELECT * FROM master_spk WHERE id=?", conn, params=(int(selected_spk_id),)).iloc[0]

        st.info(f"Kontraktor: **{spk_data['kontraktor']}**")

        with st.form("form_progress"):
            prog_lalu = 0.0
            # Ambil progress terakhir jika ada
            with get_db_connection() as conn:
                existing_prog = pd.read_sql_query("SELECT progress_minggu_ini FROM laporan_mingguan WHERE no_spk=? AND jenis_pekerjaan=?", conn, params=(spk_data['no_spk'], spk_data['jenis_pekerjaan']))
                if not existing_prog.empty:
                    prog_lalu = existing_prog.iloc[0]['progress_minggu_ini']

            prog_ini = st.number_input(f"Progress Minggu Ini (%) - (Lalu: {prog_lalu}%)", min_value=0.0, max_value=100.0, value=prog_lalu)
            catatan = st.text_area("Catatan/Kendala")
            
            f1 = st.file_uploader("Upload Foto 1", type=["jpg", "png"])
            f2 = st.file_uploader("Upload Foto 2", type=["jpg", "png"])
            
            if st.form_submit_button("Simpan Laporan"):
                # Simpan Foto
                p1, p2 = "", ""
                spk_sniz = sanitize_filename(spk_data['no_spk'])
                
                if f1:
                    p1 = os.path.join(UPLOAD_DIR, f"{spk_sniz}_minggu_{len(df_history)}_f1.jpg")
                    with open(p1, "wb") as f: f.write(f1.getbuffer())
                if f2:
                    p2 = os.path.join(UPLOAD_DIR, f"{spk_sniz}_minggu_{len(df_history)}_f2.jpg")
                    with open(p2, "wb") as f: f.write(f2.getbuffer())
                
                # Masukkan ke Laporan & History
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM laporan_mingguan WHERE no_spk=? AND jenis_pekerjaan=?", (spk_data['no_spk'], spk_data['jenis_pekerjaan']))
                    cursor.execute("""
                        INSERT INTO laporan_mingguan (no_spk, jenis_pekerjaan, kontraktor, unit, jumlah, nilai_pekerjaan, progress_minggu_lalu, progress_minggu_ini, catatan, foto_1, foto_2)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?)""", 
                        (spk_data['no_spk'], spk_data['jenis_pekerjaan'], spk_data['kontraktor'], spk_data['unit'], 1, 0, prog_lalu, prog_ini, catatan, p1, p2))
                    conn.commit()
                st.success("Laporan mingguan disimpan.")
                st.rerun()

# ------------------------------------------
# MENU 3: MASTER SPK
# ------------------------------------------
elif menu == "Kelola Master SPK":
    st.subheader("📑 Kelola Master Data SPK")
    with st.form("form_master"):
        no_spk = st.text_input("No SPK")
        kontraktor = st.text_input("Nama Kontraktor")
        pekerjaan = st.text_input("Jenis Pekerjaan")
        unit = st.text_input("Unit")
        if st.form_submit_button("Tambah Master SPK"):
            if not no_spk or not kontraktor or not pekerjaan: st.error("Lengkapi form.")
            else:
                try:
                    with get_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("INSERT INTO master_spk (no_spk, kontraktor, jenis_pekerjaan, unit) VALUES (?,?,?,?)", (no_spk, kontraktor, pekerjaan, unit))
                        conn.commit()
                    st.success("Master SPK ditambahkan.")
                    st.rerun()
                except Exception as e: st.error(f"Gagal/SPK sudah ada: {e}")
    
    with get_connection() as conn:
        df_master = pd.read_sql_query("SELECT * FROM master_spk", conn)
    st.dataframe(df_master, use_container_width=True)
