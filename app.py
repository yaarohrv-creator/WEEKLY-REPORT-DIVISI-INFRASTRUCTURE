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
                
                # AUTO MIGRATION: Pastikan kolom lokasi & nilai_spk_utama ada
                cursor.execute('''
                    ALTER TABLE master_spk 
                    ADD COLUMN IF NOT EXISTS nilai_spk_utama REAL DEFAULT 0;
                ''')
                cursor.execute('''
                    ALTER TABLE master_spk 
                    ADD COLUMN IF NOT EXISTS lokasi TEXT;
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
                        foto_2 TEXT
                    );
                ''')
                cursor.execute('ALTER TABLE history_progress ADD COLUMN IF NOT EXISTS tanggal DATE;')
                conn.commit()
    except Exception as e:
        st.error(f"⚠️ Gagal inisialisasi database: {e}")

init_db()

# ==========================================
# FUNGSI EXPORT EXCEL
# ==========================================
def generate_excel_full_feature(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_excel = df.drop(columns=['real_id'], errors='ignore').copy()
        
        # Buat dataframe untuk sheet pertama (Laporan Progress)
        df_progress = df_excel.drop(
            columns=['Foto 1', 'Foto 2', 'Pratinjau Foto 1', 'Pratinjau Foto 2'], 
            errors='ignore'
        ).copy()
        
        # Sisipkan kolom Dokumentasi di paling kanan jika belum ada
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

        # Styling Header Sheet 1
        for col_num in range(1, worksheet_progress.max_column + 1):
            cell = worksheet_progress.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_standard

        # Styling Isi Sheet 1 & Formatting Rupiah/Desimal di Excel
        for row_idx in range(2, worksheet_progress.max_row + 1):
            for col_idx in range(1, worksheet_progress.max_column + 1):
                cell = worksheet_progress.cell(row=row_idx, column=col_idx)
                cell.border = border_standard
                cell.alignment = Alignment(vertical="center")

                # Ambil Nama Header untuk format nilai Rupiah & Persentase
                header_name = worksheet_progress.cell(row=1, column=col_idx).value
                if header_name in ['Nilai SPK Utama (Rp)', 'Nilai Pekerjaan (Rp)']:
                    # Format Rupiah dengan desimal (contoh: Rp 1.500.000,00)
                    cell.number_format = '"Rp "#,##0.00'
                elif '%' in str(header_name):
                    cell.number_format = '0.00"%"'

        # Lebar Kolom
        for col in worksheet_progress.columns:
            header_name = col[0].value
            if header_name != 'Dokumentasi':
                max_len = max(len(str(cell.value or '')) for cell in col)
                worksheet_progress.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 12)
            else:
                worksheet_progress.column_dimensions[get_column_letter(col[0].column)].width = 18

        # --- SHEET 2: FOTO DOKUMENTASI ---
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
        
        # Proses baris gambar di sheet Foto Dokumentasi
        for idx_row, row in df_excel.iterrows():
            no_spk_val = row.get('Nomor SPK', '') or row.get('no_spk', '')
            jenis_val = row.get('Jenis Pekerjaan', '') or row.get('jenis_pekerjaan', '')
            
            judul_gabungan = f"SPK: {no_spk_val}\n\nPekerjaan: {jenis_val}"
            cell_j = worksheet_foto.cell(row=foto_row_idx, column=1, value=judul_gabungan)
            cell_j.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
            cell_j.border = border_standard
            
            # Simpan pemetaan indeks baris (baris Excel ke-2 dst berurutan dengan row index)
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
                        cell_p.value = f"Eror load gambar: {e}"
                        cell_p.alignment = align_center
                else:
                    cell_p.value = "Foto tidak tersedia"
                    cell_p.alignment = align_center

            insert_image_visual_resized(row.get('Pratinjau Foto 1'), worksheet_foto, foto_row_idx, 2, LEBAR_KOLOM_FOTO)
            insert_image_visual_resized(row.get('Pratinjau Foto 2'), worksheet_foto, foto_row_idx, 3, LEBAR_KOLOM_FOTO)

            foto_row_idx += 1

        # --- BUAT HYPERLINK DOKUMENTASI PADA SHEET 1 ---
        try:
            doc_col_idx = df_progress.columns.get_loc('Dokumentasi') + 1
            for p_row_idx in range(2, worksheet_progress.max_row + 1):
                if p_row_idx in job_map_targets:
                    target_photo_row = job_map_targets[p_row_idx]
                    cell_link = worksheet_progress.cell(row=p_row_idx, column=doc_col_idx, value="Lihat Foto")
                    cell_link.hyperlink = f"#'Foto Dokumentasi'!A{target_photo_row}"
                    cell_link.font = blue_link_font
                    cell_link.alignment = align_center
        except Exception as e:
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

    # 1. Query terdistribusi menyertakan COALESCE untuk tanggal
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
                tanggal,  -- Kolom tanggal laporan
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
            COALESCE(l.tanggal, l.waktu_input::date) AS tanggal_update,  -- Ambil tanggal
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

    # 2. Format Grouped Excel dengan Tanggal Update
    def format_grouped_dashboard_df(df_input):
        if df_input.empty:
            return pd.DataFrame()

        df_sorted = df_input.sort_values(
            by=['master_id', 'lokasi', 'no_spk'],
            ascending=[True, True, True]
        ).reset_index(drop=True)

        df_formatted = pd.DataFrame()
        df_formatted['real_id'] = df_sorted['real_id']
        
        nomor_spk_list = []
        kontraktor_list = []
        nilai_spk_list = []
        unit_list = []

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
        
        # Format Tanggal Update
        df_formatted['Tanggal Update'] = pd.to_datetime(df_sorted['tanggal_update'], errors='coerce').dt.date
        
        df_formatted['Catatan Pekerjaan Terbaru'] = df_sorted['catatan']
        df_formatted['Pratinjau Foto 1'] = df_sorted['foto_1']
        df_formatted['Pratinjau Foto 2'] = df_sorted['foto_2']

        return df_formatted

    # 3. Render Tabel Dashboard dengan DateColumn
    def render_dashboard_table(df_raw, tab_key_prefix):
        if df_raw.empty:
            st.info("💡 Belum ada data progress untuk wilayah/kategori ini.")
            return

        df_display = format_grouped_dashboard_df(df_raw)

        st.caption("💡 **Tampilan Grouped Excel**: Nilai SPK Utama & Kontraktor hanya muncul di baris pertama tiap SPK. Anda dapat mengedit Tanggal Update, Progress, dan Catatan langsung di bawah ini.")

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
                "Progress Minggu Lalu (%)": st.column_config.NumberColumn(
                    "Progress Minggu Lalu (%)", 
                    format="%.2f %%", 
                    min_value=0.0, 
                    max_value=100.0, 
                    step=0.5
                ),
                "Progress Minggu Ini (%)": st.column_config.NumberColumn("Progress Minggu Ini (%)", format="%.2f %%"),
                "Selisih / Varian (%)": st.column_config.NumberColumn("Selisih / Varian (%)", format="%.2f %%", disabled=True),
                
                # --- KOLOM TANGGAL (INTERAKTIF & BISA DIEDIT VIA CALENDAR) ---
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
                    
                    # Variabel penampung header SPK untuk mengisi baris baru (forward-fill)
                    current_spk = ""
                    current_kontraktor = ""
                    current_unit = ""
                    current_lokasi = ""

                    for idx, row in edited_df.iterrows():
                        if str(row.get('Nomor SPK', '') or '').strip() != "":
                            current_spk = str(row['Nomor SPK']).strip()
                        if str(row.get('Kontraktor', '') or '').strip() != "":
                            current_kontraktor = str(row['Kontraktor']).strip()
                        if str(row.get('Unit / Wilayah', '') or '').strip() != "":
                            current_unit = str(row['Unit / Wilayah']).strip()
                        if str(row.get('Lokasi', '') or '').strip() != "":
                            current_lokasi = str(row['Lokasi']).strip()

                        real_id = row.get('real_id')
                        prog_lalu = float(row.get('Progress Minggu Lalu (%)', 0.0) or 0.0)
                        prog_ini = float(row.get('Progress Minggu Ini (%)', 0.0) or 0.0)
                        tgl_val = row.get('Tanggal Update')
                        tgl_str = tgl_val.strftime('%Y-%m-%d') if pd.notnull(tgl_val) and str(tgl_val).strip() != "" else None
                        catatan_val = str(row.get('Catatan Pekerjaan Terbaru', '') or '').strip()
                        jenis_pekerjaan = str(row.get('Jenis Pekerjaan', '') or '').strip()

                        # Jika data laporan sudah ada sebelumnya (UPDATE)
                        if pd.notna(real_id) and str(real_id).strip() != "":
                            cursor.execute("""
                                UPDATE laporan_mingguan
                                SET progress_minggu_lalu = %s,
                                    progress_minggu_ini = %s,
                                    tanggal = %s,
                                    catatan = %s
                                WHERE id = %s
                            """, (
                                prog_lalu,
                                prog_ini,
                                tgl_str,
                                catatan_val,
                                int(real_id)
                            ))
                        # Jika baris laporan belum pernah diinput di database (INSERT baris baru)
                        elif jenis_pekerjaan != "":
                            cursor.execute("""
                                INSERT INTO laporan_mingguan (
                                    no_spk, jenis_pekerjaan, kontraktor, unit, jumlah, nilai_pekerjaan,
                                    progress_minggu_lalu, progress_minggu_ini, tanggal, catatan
                                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                            """, (
                                current_spk,
                                jenis_pekerjaan,
                                current_kontraktor,
                                current_unit,
                                int(row.get('Jumlah', 1) or 1),
                                float(row.get('Nilai Pekerjaan (Rp)', 0.0) or 0.0),
                                prog_lalu,
                                prog_ini,
                                tgl_str,
                                catatan_val
                            ))

                    conn.commit()

                st.success("✅ Perubahan progress minggu lalu, minggu ini, dan tanggal berhasil disimpan!")
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

    # --- TAB BANGKA ---
    with tab_bangka:
        st.subheader("📍 Laporan Progress Proyek - Wilayah Bangka")
        if not df_all.empty:
            df_bangka = df_all[df_all['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False)]
            render_dashboard_table(df_bangka, "bangka")

    # --- TAB BELITUNG ---
    with tab_belitung:
        st.subheader("📍 Laporan Progress Proyek - Wilayah Belitung")
        if not df_all.empty:
            pola_belitung = 'BELITUNG|BLT|BPSL|BPRE|BPT'
            df_belitung = df_all[
                df_all['unit'].astype(str).str.contains(pola_belitung, case=False, na=False) |
                (~df_all['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False))
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

# =========================================================
# MENU 2: INPUT PROGRESS
# =========================================================
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

        try:
            val_num = float(spk_data_selected['nilai_pekerjaan'])
            nilai_formatted = f"Rp {val_num:,.2f}"
        except (ValueError, TypeError):
            nilai_formatted = "-"

        st.markdown(f"""
📌 **Detail SPK Dipilih:**
* **Nomor SPK:** {spk_data_selected['no_spk']}
* **Kontraktor:** {spk_data_selected.get('kontraktor', '-')}
* **Jenis Pekerjaan:** {spk_data_selected['jenis_pekerjaan']}
* **Unit / Wilayah:** {spk_data_selected.get('unit', '-')}
* **Lokasi:** {spk_data_selected.get('lokasi', '-')}
* **Nilai Pekerjaan:** {nilai_formatted}
""")

        # Mengambil data laporan terakhir dari database untuk inisialisasi default input
        with get_db_connection() as conn:
            df_last_reports = pd.read_sql_query(
                """
                SELECT * FROM laporan_mingguan 
                WHERE REGEXP_REPLACE(LOWER(TRIM(no_spk)), '\\s+', ' ', 'g') = REGEXP_REPLACE(LOWER(TRIM(%s)), '\\s+', ' ', 'g')
                  AND REGEXP_REPLACE(LOWER(TRIM(jenis_pekerjaan)), '\\s+', ' ', 'g') = REGEXP_REPLACE(LOWER(TRIM(%s)), '\\s+', ' ', 'g')
                ORDER BY id DESC LIMIT 1
                """, 
                conn, 
                params=(selected_spk_no, selected_pekerjaan)
            )

        default_prog_lalu = 0.0
        default_prog_ini = 0.0
        default_catatan = ""
        foto_1_existing = None
        foto_2_existing = None

        if not df_last_reports.empty:
            last_row = df_last_reports.iloc[0]
            default_prog_lalu = float(last_row.get('progress_minggu_lalu', 0.0) or 0.0)
            default_prog_ini = float(last_row.get('progress_minggu_ini', 0.0) or 0.0)
            default_catatan = str(last_row.get('catatan', '') or '')
            foto_1_existing = last_row.get('foto_1')
            foto_2_existing = last_row.get('foto_2')

        with st.form(f"form_input_progress_{tab_key_prefix}"):
            st.subheader("📝 Form Update Progress Minggu Ini")
            
            col_in1, col_in2, col_in3 = st.columns(3)
            with col_in1:
                prog_lalu = st.number_input(
                    "Progress Minggu Lalu (%)", 
                    min_value=0.0, 
                    max_value=100.0, 
                    value=default_prog_lalu, 
                    step=0.5,
                    key=f"in_prog_lalu_{tab_key_prefix}"
                )
            with col_in2:
                prog_ini = st.number_input(
                    "Progress Minggu Ini (%)", 
                    min_value=0.0, 
                    max_value=100.0, 
                    value=default_prog_ini, 
                    step=0.5,
                    key=f"in_prog_ini_{tab_key_prefix}"
                )
            with col_in3:
                tgl_update = st.date_input(
                    "Tanggal Update Progress:", 
                    value=datetime.today(),
                    key=f"in_tgl_{tab_key_prefix}"
                )

            catatan_input = st.text_area(
                "Catatan / Kendala / Keterangan Progress:", 
                value=default_catatan,
                key=f"in_catatan_{tab_key_prefix}"
            )

            st.markdown("---")
            st.subheader("📷 Upload Foto Dokumentasi Tambahan/Baru")
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                uploaded_foto1 = st.file_uploader(
                    "Upload Foto Dokumentasi 1 (Opsional):", 
                    type=["jpg", "jpeg", "png"],
                    key=f"file_foto1_{tab_key_prefix}"
                )
            with col_f2:
                uploaded_foto2 = st.file_uploader(
                    "Upload Foto Dokumentasi 2 (Opsional):", 
                    type=["jpg", "jpeg", "png"],
                    key=f"file_foto2_{tab_key_prefix}"
                )

            submit_btn = st.form_submit_button("🚀 Kirim Laporan Progress", type="primary")

        if submit_btn:
            try:
                url_foto_1 = foto_1_existing
                url_foto_2 = foto_2_existing

                # Handling Upload Foto ke Cloudinary jika ada file baru diunggah
                if uploaded_foto1 is not None:
                    res1 = cloudinary.uploader.upload(uploaded_foto1, folder="progress_proyek")
                    url_foto_1 = res1.get("secure_url")

                if uploaded_foto2 is not None:
                    res2 = cloudinary.uploader.upload(uploaded_foto2, folder="progress_proyek")
                    url_foto_2 = res2.get("secure_url")

                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    tgl_str = tgl_update.strftime('%Y-%m-%d')
                    
                    # 1. Simpan ke Tabel Laporan Mingguan
                    cursor.execute("""
                        INSERT INTO laporan_mingguan (
                            no_spk, jenis_pekerjaan, kontraktor, unit, jumlah, nilai_pekerjaan,
                            progress_minggu_lalu, progress_minggu_ini, tanggal, catatan, foto_1, foto_2
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (
                        spk_data_selected['no_spk'],
                        spk_data_selected['jenis_pekerjaan'],
                        spk_data_selected.get('kontraktor', ''),
                        spk_data_selected.get('unit', ''),
                        int(spk_data_selected.get('jumlah', 1) or 1),
                        float(spk_data_selected.get('nilai_pekerjaan', 0.0) or 0.0),
                        prog_lalu,
                        prog_ini,
                        tgl_str,
                        catatan_input,
                        url_foto_1,
                        url_foto_2
                    ))

                    # 2. Simpan Log ke History Progress
                    penambahan = prog_ini - prog_lalu
                    cursor.execute("""
                        INSERT INTO history_progress (
                            no_spk, jenis_pekerjaan, kontraktor, unit,
                            progress_minggu_lalu, progress_minggu_ini, progres_penambahan,
                            tanggal, catatan, foto_1, foto_2
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (
                        spk_data_selected['no_spk'],
                        spk_data_selected['jenis_pekerjaan'],
                        spk_data_selected.get('kontraktor', ''),
                        spk_data_selected.get('unit', ''),
                        prog_lalu,
                        prog_ini,
                        penambahan,
                        tgl_str,
                        catatan_input,
                        url_foto_1,
                        url_foto_2
                    ))

                    conn.commit()

                st.success("✅ Laporan Progress berhasil disimpan!")
                st.rerun()

            except Exception as e:
                st.error(f"⚠️ Gagal menyimpan laporan: {e}")

        # --- SECTION PRATINJAU FOTO DOKUMENTASI TERAKHIR ---
        st.markdown("---")
        st.subheader("🖼️️ Pratinjau Foto Dokumentasi Terakhir Dipasang")
        if foto_1_existing or foto_2_existing:
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                if foto_1_existing:
                    st.image(foto_1_existing, caption="Foto Dokumentasi 1 Terakhir", use_column_width=True)
                else:
                    st.info("Foto Dokumentasi 1 belum ada.")
            with col_p2:
                if foto_2_existing:
                    st.image(foto_2_existing, caption="Foto Dokumentasi 2 Terakhir", use_column_width=True)
                else:
                    st.info("Foto Dokumentasi 2 belum ada.")
        else:
            st.info("Belum ada foto dokumentasi yang diunggah untuk SPK ini.")

    # Tab Bangka Form Input
    with tab_i_bangka:
        st.subheader("📍 Input Progress - Wilayah Bangka")
        df_master_bka = df_master_all[df_master_all['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False)]
        render_input_form(df_master_bka, "input_bangka")

    # Tab Belitung Form Input
    with tab_i_belitung:
        st.subheader("📍 Input Progress - Wilayah Belitung")
        pola_belitung = 'BELITUNG|BLT|BPSL|BPRE|BPT'
        df_master_blt = df_master_all[
            df_master_all['unit'].astype(str).str.contains(pola_belitung, case=False, na=False) |
            (~df_master_all['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False))
        ]
        render_input_form(df_master_blt, "input_belitung")

# =========================================================
# MENU 3: KELOLA MASTER SPK
# =========================================================
elif menu == MENU_MASTER:
    st.title("⚙️ Kelola Master SPK & Jenis Pekerjaan")
    st.markdown("---")

    with get_db_connection() as conn:
        df_master = pd.read_sql_query("SELECT * FROM master_spk ORDER BY id DESC", conn)

    tab_m1, tab_m2 = st.tabs(["📋 Data Master SPK", "➕ Tambah Master SPK Baru"])

    with tab_m1:
        st.subheader("Data Master SPK Terdaftar")
        if not df_master.empty:
            edited_master = st.data_editor(
                df_master,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                column_config={
                    "id": None,
                    "nilai_spk_utama": st.column_config.NumberColumn("Nilai SPK Utama (Rp)", format="Rp %',.2f"),
                    "nilai_pekerjaan": st.column_config.NumberColumn("Nilai Pekerjaan (Rp)", format="Rp %',.2f"),
                    "jumlah": st.column_config.NumberColumn("Jumlah", format="%d")
                },
                key="editor_master_spk"
            )

            col_m_btn1, col_m_btn2 = st.columns([2, 1])
            with col_m_btn1:
                if st.button("💾 Simpan Perubahan Master", type="primary", key="btn_save_master"):
                    try:
                        with get_db_connection() as conn:
                            cursor = conn.cursor()
                            for _, row in edited_master.iterrows():
                                m_id = row.get('id')
                                if pd.notna(m_id):
                                    cursor.execute("""
                                        UPDATE master_spk 
                                        SET no_spk = %s, kontraktor = %s, jenis_pekerjaan = %s,
                                            unit = %s, lokasi = %s, jumlah = %s, nilai_spk_utama = %s,
                                            nilai_pekerjaan = %s, catatan = %s
                                        WHERE id = %s
                                    """, (
                                        row.get('no_spk'), row.get('kontraktor'), row.get('jenis_pekerjaan'),
                                        row.get('unit'), row.get('lokasi'), int(row.get('jumlah', 1) or 1),
                                        float(row.get('nilai_spk_utama', 0.0) or 0.0),
                                        float(row.get('nilai_pekerjaan', 0.0) or 0.0),
                                        row.get('catatan'), int(m_id)
                                    ))
                            conn.commit()
                        st.success("✅ Data Master SPK berhasil diperbarui!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"⚠️️ Gagal memperbarui Master SPK: {e}")

            with col_m_btn2:
                id_to_del = st.selectbox("Pilih ID Master untuk Dihapus:", df_master['id'].tolist(), key="sel_del_master")
                if st.button("🗑️ Hapus Master ID", type="secondary", key="btn_del_master"):
                    try:
                        with get_db_connection() as conn:
                            cursor = conn.cursor()
                            cursor.execute("DELETE FROM master_spk WHERE id = %s", (id_to_del,))
                            conn.commit()
                        st.success(f"Master SPK ID {id_to_del} berhasil dihapus!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal menghapus master: {e}")
        else:
            st.info("Belum ada data Master SPK.")

    with tab_m2:
        st.subheader("➕ Tambah Data Master SPK Baru")
        with st.form("form_add_master"):
            c_m1, c_m2 = st.columns(2)
            with c_m1:
                add_no_spk = st.text_input("Nomor SPK:")
                add_kontraktor = st.text_input("Kontraktor:")
                add_unit = st.text_input("Unit / Wilayah (contoh: BANGKA / BELITUNG):")
                add_lokasi = st.text_input("Lokasi:")
            with c_m2:
                add_jenis_pekerjaan = st.text_input("Jenis Pekerjaan:")
                add_jumlah = st.number_input("Jumlah:", min_value=1, value=1, step=1)
                add_nilai_spk_utama = st.number_input("Nilai SPK Utama (Rp):", min_value=0.0, value=0.0, step=1000000.0)
                add_nilai_pekerjaan = st.number_input("Nilai Pekerjaan (Rp):", min_value=0.0, value=0.0, step=500000.0)

            add_catatan = st.text_area("Catatan Master SPK:")
            submit_master = st.form_submit_button("➕ Tambah Master SPK", type="primary")

        if submit_master:
            if not add_no_spk or not add_jenis_pekerjaan:
                st.warning("⚠️ Nomor SPK dan Jenis Pekerjaan wajib diisi!")
            else:
                try:
                    with get_db_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("""
                            INSERT INTO master_spk (
                                no_spk, kontraktor, jenis_pekerjaan, unit, lokasi,
                                jumlah, nilai_spk_utama, nilai_pekerjaan, catatan
                            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """, (
                            add_no_spk.strip(), add_kontraktor.strip(), add_jenis_pekerjaan.strip(),
                            add_unit.strip(), add_lokasi.strip(), add_jumlah,
                            add_nilai_spk_utama, add_nilai_pekerjaan, add_catatan.strip()
                        ))
                        conn.commit()
                    st.success("✅ Berhasil menambahkan Master SPK baru!")
                    st.rerun()
                except Exception as e:
                    st.error(f"⚠️ Gagal menambahkan Master SPK (Kemungkinan Nomor SPK + Jenis Pekerjaan sudah terdaftar): {e}")
