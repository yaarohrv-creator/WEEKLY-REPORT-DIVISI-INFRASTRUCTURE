import os
import io
import re
import time
import requests
import psycopg2
import pandas as pd
import streamlit as st
import cloudinary
import cloudinary.uploader

# Modul untuk styling dan export Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- KONFIGURASI CLOUDINARY ---
try:
    cloudinary.config(
        cloud_name=st.secrets["cloudinary"]["cloud_name"],
        api_key=st.secrets["cloudinary"]["api_key"],
        api_secret=st.secrets["cloudinary"]["api_secret"]
    )
except Exception as e:
    st.error(f"Gagal memuat konfigurasi Cloudinary dari st.secrets: {e}")

# --- LIBRARY MEMPROSES GAMBAR ---
try:
    from PIL import Image as PILImage
    from openpyxl.drawing.image import Image as OpenPyXLImage
    has_pil = True
except ImportError:
    has_pil = False

# Konfigurasi Halaman Streamlit
st.set_page_config(page_title="Sistem Progress Proyek", layout="wide")

# Folder Penyimpanan Sementara (Opsional)
UPLOAD_DIR = "uploads"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# ---------------------------------------------------------
# FUNGSI HELPER: UPLOAD FOTO KE CLOUDINARY
# ---------------------------------------------------------
def upload_foto_to_cloudinary(uploaded_file):
    if uploaded_file is not None:
        try:
            response = cloudinary.uploader.upload(
                uploaded_file,
                folder="progress_proyek"
            )
            return response.get("secure_url")
        except Exception as e:
            st.error(f"Gagal mengunggah foto ke Cloudinary: {e}")
            return None
    return None

# ---------------------------------------------------------
# AUTENTIKASI PASSWORD
# ---------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

def check_password():
    password_benar = st.secrets.get("APP_PASSWORD", "123456")
    if st.session_state.get("password_input") == password_benar:
        st.session_state["authenticated"] = True
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
# FUNGSIONALITAS DATABASE POSTGRESQL
# ---------------------------------------------------------
def get_db_connection():
    return psycopg2.connect(st.secrets["postgres"]["url"])

def init_db():
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        # 1. TABEL MASTER_SPK
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS master_spk (
                id SERIAL PRIMARY KEY,
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

        # 2. TABEL LAPORAN_MINGGUAN
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS laporan_mingguan (
                id SERIAL PRIMARY KEY,
                waktu_input TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
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
                id SERIAL PRIMARY KEY,
                waktu_input TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
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
    except Exception as e:
        st.error(f"Gagal melakukan inisialisasi Database: {e}")

init_db()

# ---------------------------------------------------------
# FUNGSI EXPORT EXCEL
# ---------------------------------------------------------
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

        foto_row_idx = 2
        
        for index, row in df_excel.iterrows():
            judul_gabungan = f"SPK: {row.get('Nomor SPK', '')}\n\nPekerjaan: {row.get('Jenis Pekerjaan', '')}"
            cell_j = worksheet_foto.cell(row=foto_row_idx, column=1, value=judul_gabungan)
            cell_j.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
            cell_j.border = border_standard

            worksheet_foto.row_dimensions[foto_row_idx].height = 200

            def insert_image_visual_resized(path_or_url, ws, current_row, current_col, target_col_width):
                cell_p = ws.cell(row=current_row, column=current_col)
                cell_p.border = border_standard
                
                if path_or_url and has_pil:
                    try:
                        if str(path_or_url).startswith("http"):
                            res = requests.get(path_or_url, timeout=10)
                            pil_img = PILImage.open(io.BytesIO(res.content))
                        elif os.path.exists(str(path_or_url)):
                            pil_img = PILImage.open(path_or_url)
                        else:
                            cell_p.value = "Foto Tidak Ditemukan"
                            cell_p.alignment = align_center
                            return

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
                        cell_p.value = f"Error load gambar: {e}"
                        cell_p.alignment = align_center

            insert_image_visual_resized(row.get('Foto 1'), worksheet_foto, foto_row_idx, 2, LEBAR_KOLOM_FOTO)
            insert_image_visual_resized(row.get('Foto 2'), worksheet_foto, foto_row_idx, 3, LEBAR_KOLOM_FOTO)

            foto_row_idx += 1

    return output.getvalue()

# ---------------------------------------------------------
# TAMPILAN UTAMA APLIKASI STREAMLIT
# ---------------------------------------------------------
st.title("📋 Sistem Laporan Progress Pekerjaan")
st.write("Selamat datang! Silakan gunakan menu di bawah untuk mengelola data proyek.")

# Tampilkan data ringkasan jika ada
try:
    conn = get_db_connection()
    df_spk = pd.read_sql_query("SELECT * FROM master_spk", conn)
    conn.close()
    
    st.subheader("Data Master SPK")
    if not df_spk.empty:
        st.dataframe(df_spk, use_container_width=True)
    else:
        st.info("Belum ada data Master SPK di database.")
except Exception as e:
    st.warning(f"Belum dapat menampilkan data: {e}")
