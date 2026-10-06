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
from datetime import datetime

# Modul untuk styling dan export Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Konfigurasi Cloudinary dari Secrets Streamlit Cloud
cloudinary.config(
    cloud_name=st.secrets["cloudinary"]["cloud_name"],
    api_key=st.secrets["cloudinary"]["api_key"],
    api_secret=st.secrets["cloudinary"]["api_secret"]
)

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

# Folder Penyimpanan Sementara/Uploads jika dibutuhkan
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
    user_input = st.session_state.get("password_input", "")
    
    if user_input == password_benar:
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
# FUNGSIONALITAS DATABASE POSTGRESQL
# ---------------------------------------------------------
def get_db_connection():
    return psycopg2.connect(st.secrets["postgres"]["url"])

def init_db():
    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                # 1. TABEL MASTER_SPK
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS master_spk (
                        id SERIAL PRIMARY KEY,
                        no_spk TEXT,
                        kontraktor TEXT,
                        jenis_pekerjaan TEXT,
                        unit TEXT,
                        lokasi TEXT,
                        jumlah INTEGER DEFAULT 1,
                        nilai_spk_utama REAL DEFAULT 0,
                        nilai_pekerjaan REAL DEFAULT 0,
                        catatan TEXT,
                        UNIQUE(no_spk, jenis_pekerjaan)
                    );
                ''')
                cursor.execute('ALTER TABLE master_spk ADD COLUMN IF NOT EXISTS nilai_spk_utama REAL DEFAULT 0;')
                cursor.execute('ALTER TABLE master_spk ADD COLUMN IF NOT EXISTS lokasi TEXT;')

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
                    );
                ''')
                cursor.execute('ALTER TABLE laporan_mingguan ADD COLUMN IF NOT EXISTS tanggal DATE;')

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
                        foto_2 TEXT,
                        tanggal DATE
                    );
                ''')
                cursor.execute('ALTER TABLE history_progress ADD COLUMN IF NOT EXISTS tanggal DATE;')
                conn.commit()
    except Exception as e:
        st.error(f"⚠️ Gagal inisialisasi database: {e}")

init_db()

# ---------------------------------------------------------
# FUNGSI UPLOAD KE CLOUDINARY
# ---------------------------------------------------------
def upload_to_cloudinary(uploaded_file):
    if uploaded_file is None:
        return None
    try:
        response = cloudinary.uploader.upload(
            uploaded_file,
            folder="laporan_progress_proyek"
        )
        return response.get("secure_url")
    except Exception as e:
        st.error(f"Gagal mengunggah foto ke Cloudinary: {e}")
        return None

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
        
        if 'Dokumentasi' not in df_progress.columns:
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

                header_name = worksheet_progress.cell(row=1, column=col_idx).value
                if header_name in ['Nilai SPK Utama (Rp)', 'Nilai Pekerjaan (Rp)']:
                    cell.number_format = '"Rp "#,##0.00'
                elif header_name and '%' in str(header_name):
                    cell.number_format = '0.00"%"'

        for col in worksheet_progress.columns:
            header_name = col[0].value
            if header_name != 'Dokumentasi':
                max_len = max(len(str(cell.value or '')) for cell in col)
                worksheet_progress.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 12)
            else:
                worksheet_progress.column_dimensions[get_column_letter(col[0].column)].width = 18

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
        
        for idx_row, row in df_excel.iterrows():
            no_spk_val = row.get('Nomor SPK', '') or row.get('no_spk', '')
            jenis_val = row.get('Jenis Pekerjaan', '') or row.get('jenis_pekerjaan', '')
            
            judul_gabungan = f"SPK: {no_spk_val}\n\nPekerjaan: {jenis_val}"
            cell_j = worksheet_foto.cell(row=foto_row_idx, column=1, value=judul_gabungan)
            cell_j.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
            cell_j.border = border_standard
            
            job_map_targets[idx_row + 2] = foto_row_idx
            worksheet_foto.row_dimensions[foto_row_idx].height = 250

            def insert_image_visual_resized(path_or_url, ws, current_row, current_col, target_col_width):
                cell_p = ws.cell(row=current_row, column=current_col)
                cell_p.border = border_standard
                
                if path_or_url and str(path_or_url).strip() != "" and str(path_or_url) != "None" and has_pil:
                    try:
                        if str(path_or_url).startswith("http://") or str(path_or_url).startswith("https://"):
                            resp = requests.get(path_or_url, stream=True)
                            if resp.status_code == 200:
                                pil_img = PILImage.open(io.BytesIO(resp.content))
                            else:
                                raise Exception("Gagal mengunduh foto dari URL")
                        elif os.path.exists(str(path_or_url)):
                            pil_img = PILImage.open(path_or_url)
                        else:
                            raise Exception("File lokal tidak ditemukan")

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
                else:
                    cell_p.value = "Foto tidak tersedia"
                    cell_p.alignment = align_center

            insert_image_visual_resized(row.get('Pratinjau Foto 1'), worksheet_foto, foto_row_idx, 2, LEBAR_KOLOM_FOTO)
            insert_image_visual_resized(row.get('Pratinjau Foto 2'), worksheet_foto, foto_row_idx, 3, LEBAR_KOLOM_FOTO)

            foto_row_idx += 1

        try:
            doc_col_idx = df_progress.columns.get_loc('Dokumentasi') + 1
            for p_row_idx in range(2, worksheet_progress.max_row + 1):
                if p_row_idx in job_map_targets:
                    target_photo_row = job_map_targets[p_row_idx]
                    cell_link = worksheet_progress.cell(row=p_row_idx, column=doc_col_idx, value="Lihat Foto")
                    cell_link.hyperlink = f"#'Foto Dokumentasi'!A{target_photo_row}"
                    cell_link.font = blue_link_font
                    cell_link.alignment = align_center
        except Exception:
            pass

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
        "🏝 Laporan Progress Bangka", 
        "🏖️ Laporan Progress Belitung", 
        "📋 Semua Progress Proyek", 
        "📜 Riwayat / History Update"
    ])

    query_view = """
        WITH spk_totals AS (
            SELECT 
                REGEXP_REPLACE(LOWER(TRIM(no_spk)), '\\s+', ' ', 'g') AS clean_no_spk,
                SUM(COALESCE(nilai_pekerjaan, 0)) AS calculated_nilai_spk_utama
            FROM master_spk
            GROUP BY REGEXP_REPLACE(LOWER(TRIM(no_spk)), '\\s+', ' ', 'g')
        ),
        latest_laporan AS (
            SELECT DISTINCT ON (
                REGEXP_REPLACE(LOWER(TRIM(no_spk)), '\\s+', ' ', 'g'), 
                REGEXP_REPLACE(LOWER(TRIM(jenis_pekerjaan)), '\\s+', ' ', 'g')
            )
                id,
                no_spk,
                jenis_pekerjaan,
                progress_minggu_lalu,
                progress_minggu_ini,
                catatan,
                foto_1,
                foto_2,
                waktu_input,
                tanggal,
                REGEXP_REPLACE(LOWER(TRIM(no_spk)), '\\s+', ' ', 'g') AS clean_no_spk,
                REGEXP_REPLACE(LOWER(TRIM(jenis_pekerjaan)), '\\s+', ' ', 'g') AS clean_jenis_pekerjaan
            FROM laporan_mingguan
            ORDER BY 
                REGEXP_REPLACE(LOWER(TRIM(no_spk)), '\\s+', ' ', 'g'),
                REGEXP_REPLACE(LOWER(TRIM(jenis_pekerjaan)), '\\s+', ' ', 'g'),
                id DESC
        )
        SELECT 
            m.id AS master_id,
            l.id AS real_id,
            m.no_spk AS no_spk,
            m.kontraktor AS kontraktor,
            COALESCE(t.calculated_nilai_spk_utama, m.nilai_spk_utama, 0) AS nilai_spk_utama,
            m.unit AS unit,
            COALESCE(m.lokasi, '-') AS lokasi,
            m.jenis_pekerjaan AS jenis_pekerjaan,
            CAST(COALESCE(m.jumlah, 1) AS INTEGER) AS jumlah,
            COALESCE(m.nilai_pekerjaan, 0) AS nilai_pekerjaan,
            COALESCE(l.progress_minggu_lalu, 0) AS progress_minggu_lalu,
            COALESCE(l.progress_minggu_ini, 0) AS progress_minggu_ini,
            (COALESCE(l.progress_minggu_ini, 0) - COALESCE(l.progress_minggu_lalu, 0)) AS selisih_varian,
            COALESCE(l.tanggal, l.waktu_input::date) AS tanggal_update,
            l.catatan AS catatan,
            l.foto_1 AS foto_1,
            l.foto_2 AS foto_2
        FROM master_spk m
        LEFT JOIN spk_totals t
            ON REGEXP_REPLACE(LOWER(TRIM(m.no_spk)), '\\s+', ' ', 'g') = t.clean_no_spk
        LEFT JOIN latest_laporan l 
            ON REGEXP_REPLACE(LOWER(TRIM(m.no_spk)), '\\s+', ' ', 'g') = l.clean_no_spk
           AND REGEXP_REPLACE(LOWER(TRIM(m.jenis_pekerjaan)), '\\s+', ' ', 'g') = l.clean_jenis_pekerjaan
        ORDER BY 
            m.lokasi ASC,
            m.no_spk ASC,
            m.id ASC
    """

    with get_db_connection() as conn:
        try:
            df_all = pd.read_sql_query(query_view, conn)
        except Exception as e:
            st.error(f"Error membaca data dashboard: {e}")
            df_all = pd.DataFrame()

    def format_grouped_dashboard_df(df_input):
        if df_input.empty:
            return pd.DataFrame()

        df_sorted = df_input.sort_values(
            by=['master_id', 'lokasi', 'no_spk'],
            ascending=[True, True, True]
        ).reset_index(drop=True)

        df_formatted = pd.DataFrame()
        df_formatted['real_id'] = df_sorted['real_id']
        
        nomor_spk_list, kontraktor_list, nilai_spk_list, unit_list = [], [], [], []
        last_spk = None

        for _, row in df_sorted.iterrows():
            current_spk = row['no_spk']
            if current_spk != last_spk:
                nomor_spk_list.append(current_spk)
                kontraktor_list.append(row['kontraktor'])
                nilai_spk_list.append(row['nilai_spk_utama'])
                unit_list.append(row['unit'])
                last_spk = current_spk
            else:
                nomor_spk_list.append("")
                kontraktor_list.append("")
                nilai_spk_list.append(None)
                unit_list.append("")

        df_formatted['Nomor SPK'] = nomor_spk_list
        df_formatted['Kontraktor'] = kontraktor_list
        df_formatted['Nilai SPK Utama (Rp)'] = nilai_spk_list
        df_formatted['Unit / Wilayah'] = unit_list
        df_formatted['Lokasi'] = df_sorted['lokasi']
        df_formatted['Jenis Pekerjaan'] = df_sorted['jenis_pekerjaan']
        df_formatted['Jumlah'] = df_sorted['jumlah']
        df_formatted['Nilai Pekerjaan (Rp)'] = df_sorted['nilai_pekerjaan']
        df_formatted['Progress Minggu Lalu (%)'] = df_sorted['progress_minggu_lalu']
        df_formatted['Progress Minggu Ini (%)'] = df_sorted['progress_minggu_ini']
        df_formatted['Selisih / Varian (%)'] = df_sorted['selisih_varian']
        df_formatted['Tanggal Update'] = pd.to_datetime(df_sorted['tanggal_update'], errors='coerce').dt.date
        df_formatted['Catatan Pekerjaan Terbaru'] = df_sorted['catatan'].fillna('').replace(['nan', 'None', 'NaN'], '')
        df_formatted['Pratinjau Foto 1'] = df_sorted['foto_1']
        df_formatted['Pratinjau Foto 2'] = df_sorted['foto_2']

        return df_formatted

    def render_dashboard_table(df_raw, tab_key_prefix):
        if df_raw.empty:
            st.info("💡 Belum ada data progress untuk wilayah/kategori ini.")
            return

        df_display = format_grouped_dashboard_df(df_raw)

        st.caption("💡 **Tampilan Grouped Excel**: Anda dapat mengedit Tanggal Update, Progress, dan Catatan langsung di tabel bawah ini.")

        editor_key = f"editor_{tab_key_prefix}"
        edited_df = st.data_editor(
            df_display,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "real_id": None,
                "Nomor SPK": st.column_config.TextColumn("Nomor SPK", disabled=True),
                "Kontraktor": st.column_config.TextColumn("Kontraktor", disabled=True),
                "Nilai SPK Utama (Rp)": st.column_config.NumberColumn("Nilai SPK Utama (Rp)", format="Rp %',.2f", disabled=True),
                "Unit / Wilayah": st.column_config.TextColumn("Unit / Wilayah", disabled=True),
                "Lokasi": st.column_config.TextColumn("Lokasi", disabled=True),
                "Jenis Pekerjaan": st.column_config.TextColumn("Jenis Pekerjaan", disabled=True),
                "Jumlah": st.column_config.NumberColumn("Jumlah", format="%d", disabled=True),
                "Nilai Pekerjaan (Rp)": st.column_config.NumberColumn("Nilai Pekerjaan (Rp)", format="Rp %',.2f", disabled=True),
                "Progress Minggu Lalu (%)": st.column_config.NumberColumn("Progress Minggu Lalu (%)", format="%.2f %%", min_value=0.0, max_value=100.0, step=0.5),
                "Progress Minggu Ini (%)": st.column_config.NumberColumn("Progress Minggu Ini (%)", format="%.2f %%", min_value=0.0, max_value=100.0, step=0.5),
                "Selisih / Varian (%)": st.column_config.NumberColumn("Selisih / Varian (%)", format="%.2f %%", disabled=True),
                "Tanggal Update": st.column_config.DateColumn("Tanggal Update", format="DD/MM/YYYY"),
                "Catatan Pekerjaan Terbaru": st.column_config.TextColumn("Catatan Pekerjaan Terbaru"),
                "Pratinjau Foto 1": st.column_config.ImageColumn("Pratinjau Foto 1"),
                "Pratinjau Foto 2": st.column_config.ImageColumn("Pratinjau Foto 2"),
            },
            key=editor_key
        )

        if st.button("💾 Simpan Perubahan Data", key=f"btn_save_{tab_key_prefix}", type="primary"):
            try:
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    current_spk, current_kontraktor, current_unit, current_lokasi = "", "", "", ""

                    for idx, row in edited_df.iterrows():
                        if str(row.get('Nomor SPK', '') or '').strip() != "":
                            current_spk = str(row['Nomor SPK']).strip()
                        if str(row.get('Kontraktor', '') or '').strip() != "":
                            current_kontraktor = str(row['Kontraktor']).strip()
                        if str(row.get('Unit / Wilayah', '') or '').strip() != "":
                            current_unit = str(row['Unit / Wilayah']).strip()

                        real_id = row.get('real_id')
                        prog_lalu = float(row.get('Progress Minggu Lalu (%)', 0.0) or 0.0)
                        prog_ini = float(row.get('Progress Minggu Ini (%)', 0.0) or 0.0)
                        tgl_val = row.get('Tanggal Update')
                        tgl_str = tgl_val.strftime('%Y-%m-%d') if pd.notnull(tgl_val) and str(tgl_val).strip() != "" else None
                        catatan_val = str(row.get('Catatan Pekerjaan Terbaru', '') or '').strip()
                        jenis_pekerjaan = str(row.get('Jenis Pekerjaan', '') or '').strip()

                        if pd.notna(real_id) and str(real_id).strip() != "":
                            cursor.execute("""
                                UPDATE laporan_mingguan
                                SET progress_minggu_lalu = %s,
                                    progress_minggu_ini = %s,
                                    tanggal = %s,
                                    catatan = %s
                                WHERE id = %s
                            """, (prog_lalu, prog_ini, tgl_str, catatan_val, int(real_id)))
                        elif jenis_pekerjaan != "":
                            cursor.execute("""
                                INSERT INTO laporan_mingguan (
                                    no_spk, jenis_pekerjaan, kontraktor, unit, jumlah, nilai_pekerjaan,
                                    progress_minggu_lalu, progress_minggu_ini, tanggal, catatan
                                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                            """, (
                                current_spk, jenis_pekerjaan, current_kontraktor, current_unit,
                                int(row.get('Jumlah', 1) or 1), float(row.get('Nilai Pekerjaan (Rp)', 0.0) or 0.0),
                                prog_lalu, prog_ini, tgl_str, catatan_val
                            ))

                    conn.commit()

                st.success("✅ Perubahan berhasil disimpan!")
                st.rerun()
            except Exception as e:
                st.error(f"⚠️ Gagal menyimpan perubahan: {e}")

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

    with tab_bangka:
        st.subheader("📍 Laporan Progress Proyek - Wilayah Bangka")
        if not df_all.empty:
            df_bangka = df_all[df_all['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False)]
            render_dashboard_table(df_bangka, "bangka")

    with tab_belitung:
        st.subheader("📍 Laporan Progress Proyek - Wilayah Belitung")
        if not df_all.empty:
            pola_belitung = 'BELITUNG|BLT|BPSL|BPRE|BPT'
            df_belitung = df_all[
                df_all['unit'].astype(str).str.contains(pola_belitung, case=False, na=False) |
                (~df_all['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False))
            ]
            render_dashboard_table(df_belitung, "belitung")

    with tab_semua:
        st.subheader("🌐 Semua Laporan Progress Proyek")
        render_dashboard_table(df_all, "semua")

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
                        cursor.execute("DELETE FROM history_progress WHERE id = %s", (id_pilihan,))
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
# MENU 2: INPUT PROGRESS MINGGUAN
# ---------------------------------------------------------
elif menu == MENU_INPUT:
    st.title("📝 Input Laporan Progress Mingguan Berdasarkan SPK")
    st.markdown("---")

    with get_db_connection() as conn:
        df_master_all = pd.read_sql_query("SELECT * FROM master_spk", conn)

    def render_input_form(df_master_wilayah, tab_key_prefix):
        if df_master_wilayah.empty:
            st.info("💡 Belum ada data master pekerjaan terdaftar untuk wilayah ini.")
            return

        list_spk_unique = sorted(df_master_wilayah['no_spk'].dropna().unique().tolist())
        selected_spk_no = st.selectbox(
            "Pilih Nomor SPK:",
            list_spk_unique,
            key=f"select_spk_no_{tab_key_prefix}"
        )

        df_spk_filtered = df_master_wilayah[df_master_wilayah['no_spk'] == selected_spk_no]
        list_pekerjaan = df_spk_filtered['jenis_pekerjaan'].dropna().unique().tolist()
        selected_pekerjaan = st.selectbox(
            "Pilih Jenis Pekerjaan (Tergroup berdasarkan SPK):",
            list_pekerjaan,
            key=f"select_pekerjaan_{tab_key_prefix}"
        )

        spk_data_selected = df_spk_filtered[df_spk_filtered['jenis_pekerjaan'] == selected_pekerjaan].iloc[0]

        # Detail Ringkasan SPK
        val_num = float(spk_data_selected['nilai_pekerjaan'] or 0.0)
        nilai_formatted = f"Rp {val_num:,.2f}"

        st.info(f"""
        📌 **Informasi SPK Dipilih:**
        - **Kontraktor**: {spk_data_selected.get('kontraktor', '-')}
        - **Unit/Wilayah**: {spk_data_selected.get('unit', '-')}
        - **Lokasi**: {spk_data_selected.get('lokasi', '-')}
        - **Nilai Pekerjaan**: {nilai_formatted}
        """)

        # Ambil progress terakhir jika ada
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT progress_minggu_ini, foto_1, foto_2 
                FROM laporan_mingguan 
                WHERE TRIM(LOWER(no_spk)) = TRIM(LOWER(%s)) 
                  AND TRIM(LOWER(jenis_pekerjaan)) = TRIM(LOWER(%s))
                ORDER BY id DESC LIMIT 1
            """, (selected_spk_no, selected_pekerjaan))
            last_record = cursor.fetchone()

        last_progress = float(last_record[0]) if last_record and last_record[0] is not None else 0.0
        old_foto1 = last_record[1] if last_record else None
        old_foto2 = last_record[2] if last_record else None

        with st.form(key=f"form_input_{tab_key_prefix}"):
            col_tgl, col_lalu = st.columns(2)
            with col_tgl:
                tgl_input = st.date_input("Tanggal Update Laporan:", datetime.today())
            with col_lalu:
                prog_lalu = st.number_input("Progress Minggu Lalu (%)", min_value=0.0, max_value=100.0, value=last_progress, step=0.5)

            prog_ini = st.number_input("Progress Minggu Ini (%)", min_value=0.0, max_value=100.0, value=last_progress, step=0.5)
            catatan = st.text_area("Catatan/Keterangan Progress Baru:", placeholder="Masukkan detail perkembangan di lapangan...")

            st.write("📸 **Upload Foto Dokumentasi (Opsional - Cloudinary):**")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                file_foto1 = st.file_uploader("Upload Foto 1", type=["jpg", "jpeg", "png"], key=f"f1_{tab_key_prefix}")
            with col_f2:
                file_foto2 = st.file_uploader("Upload Foto 2", type=["jpg", "jpeg", "png"], key=f"f2_{tab_key_prefix}")

            submit_btn = st.form_submit_button("🚀 Kirim Laporan Progress")

        if submit_btn:
            url_foto1 = upload_to_cloudinary(file_foto1) if file_foto1 else old_foto1
            url_foto2 = upload_to_cloudinary(file_foto2) if file_foto2 else old_foto2
            selisih_progres = prog_ini - prog_lalu

            try:
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    
                    # 1. Simpan ke Laporan Utama
                    cursor.execute("""
                        INSERT INTO laporan_mingguan (
                            no_spk, jenis_pekerjaan, kontraktor, unit, jumlah, nilai_pekerjaan,
                            progress_minggu_lalu, progress_minggu_ini, catatan, foto_1, foto_2, tanggal
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (
                        selected_spk_no, selected_pekerjaan, spk_data_selected.get('kontraktor'),
                        spk_data_selected.get('unit'), int(spk_data_selected.get('jumlah', 1) or 1),
                        val_num, prog_lalu, prog_ini, catatan, url_foto1, url_foto2, tgl_input.strftime('%Y-%m-%d')
                    ))

                    # 2. Simpan ke History Progress
                    cursor.execute("""
                        INSERT INTO history_progress (
                            no_spk, jenis_pekerjaan, kontraktor, unit,
                            progress_minggu_lalu, progress_minggu_ini, progres_penambahan, catatan, foto_1, foto_2, tanggal
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (
                        selected_spk_no, selected_pekerjaan, spk_data_selected.get('kontraktor'),
                        spk_data_selected.get('unit'), prog_lalu, prog_ini, selisih_progres, catatan, url_foto1, url_foto2, tgl_input.strftime('%Y-%m-%d')
                    ))

                    conn.commit()

                st.success(f"✅ Laporan Progress untuk SPK '{selected_spk_no}' ({selected_pekerjaan}) berhasil dikirim!")
                st.rerun()
            except Exception as e:
                st.error(f"⚠️ Gagal menyimpan laporan: {e}")

    tab_i_bangka, tab_i_belitung = st.tabs([
        "🏝️️ Input Progress Bangka", 
        "🏖️ Input Progress Belitung"
    ])

    with tab_i_bangka:
        if not df_master_all.empty:
            df_m_bangka = df_master_all[df_master_all['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False)]
            render_input_form(df_m_bangka, "in_bangka")

    with tab_i_belitung:
        if not df_master_all.empty:
            pola_belitung = 'BELITUNG|BLT|BPSL|BPRE|BPT'
            df_m_belitung = df_master_all[
                df_master_all['unit'].astype(str).str.contains(pola_belitung, case=False, na=False) |
                (~df_master_all['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False))
            ]
            render_input_form(df_m_belitung, "in_belitung")

# ---------------------------------------------------------
# MENU 3: KELOLA MASTER SPK
# ---------------------------------------------------------
elif menu == MENU_MASTER:
    st.title("⚙️ Kelola Master SPK & Pekerjaan")
    st.markdown("---")

    with get_db_connection() as conn:
        df_master = pd.read_sql_query("SELECT * FROM master_spk ORDER BY id ASC", conn)

    st.subheader("📋 Daftar Master SPK Aktif")
    edited_master = st.data_editor(
        df_master,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "id": None,
            "no_spk": st.column_config.TextColumn("Nomor SPK", required=True),
            "kontraktor": st.column_config.TextColumn("Nama Kontraktor"),
            "unit": st.column_config.TextColumn("Unit / Wilayah (e.g. BANGKA / BELITUNG)"),
            "lokasi": st.column_config.TextColumn("Lokasi Detail"),
            "jenis_pekerjaan": st.column_config.TextColumn("Jenis Pekerjaan", required=True),
            "jumlah": st.column_config.NumberColumn("Jumlah Unit", min_value=1, step=1),
            "nilai_spk_utama": st.column_config.NumberColumn("Nilai SPK Utama (Rp)", format="Rp %',.2f"),
            "nilai_pekerjaan": st.column_config.NumberColumn("Nilai Pekerjaan (Rp)", format="Rp %',.2f"),
            "catatan": st.column_config.TextColumn("Catatan Master"),
        },
        key="master_spk_editor"
    )

    if st.button("💾 Simpan Perubahan Master SPK", type="primary"):
        try:
            with get_db_connection() as conn:
                cursor = conn.cursor()
                # Opsi paling stabil: Simpan ulang/update master data
                for idx, row in edited_master.iterrows():
                    m_id = row.get("id")
                    no_spk = str(row.get("no_spk", "") or "").strip()
                    jenis_pek = str(row.get("jenis_pekerjaan", "") or "").strip()

                    if no_spk and jenis_pek:
                        if pd.notna(m_id) and str(m_id).strip() != "":
                            cursor.execute("""
                                UPDATE master_spk SET 
                                    no_spk = %s, kontraktor = %s, jenis_pekerjaan = %s, unit = %s, 
                                    lokasi = %s, jumlah = %s, nilai_spk_utama = %s, nilai_pekerjaan = %s, catatan = %s
                                WHERE id = %s
                            """, (
                                no_spk, str(row.get("kontraktor", "")), jenis_pek, str(row.get("unit", "")),
                                str(row.get("lokasi", "")), int(row.get("jumlah", 1) or 1),
                                float(row.get("nilai_spk_utama", 0.0) or 0.0), float(row.get("nilai_pekerjaan", 0.0) or 0.0),
                                str(row.get("catatan", "")), int(m_id)
                            ))
                        else:
                            cursor.execute("""
                                INSERT INTO master_spk (no_spk, kontraktor, jenis_pekerjaan, unit, lokasi, jumlah, nilai_spk_utama, nilai_pekerjaan, catatan)
                                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                                ON CONFLICT (no_spk, jenis_pekerjaan) DO NOTHING
                            """, (
                                no_spk, str(row.get("kontraktor", "")), jenis_pek, str(row.get("unit", "")),
                                str(row.get("lokasi", "")), int(row.get("jumlah", 1) or 1),
                                float(row.get("nilai_spk_utama", 0.0) or 0.0), float(row.get("nilai_pekerjaan", 0.0) or 0.0),
                                str(row.get("catatan", ""))
                            ))
                conn.commit()
            st.success("✅ Master SPK berhasil diperbarui!")
            st.rerun()
        except Exception as e:
            st.error(f"⚠️ Gagal menyimpan master SPK: {e}")
