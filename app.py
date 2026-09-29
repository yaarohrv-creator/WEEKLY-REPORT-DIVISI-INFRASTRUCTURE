import os
import io
import psycopg2
import pandas as pd
import streamlit as st
import cloudinary
import cloudinary.uploader

# Konfigurasi Cloudinary
cloudinary.config(
    cloud_name=st.secrets["cloudinary"]["cloud_name"],
    api_key=st.secrets["cloudinary"]["api_key"],
    api_secret=st.secrets["cloudinary"]["api_secret"]
)

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from PIL import Image as PILImage
    from openpyxl.drawing.image import Image as OpenPyXLImage
    has_pil = True
except ImportError:
    st.error("⚠️ Library 'Pillow' belum terinstal. Silakan instal: pip install Pillow")
    has_pil = False

# Konfigurasi Halaman Streamlit
st.set_page_config(page_title="Sistem Progress Proyek", layout="wide")

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
        st.error("🔑 Password salah!")

if not st.session_state["authenticated"]:
    st.title("🔒 Akses Terbatas - Laporan Progress Proyek")
    st.text_input("Password Akses:", type="password", key="password_input", on_change=check_password)
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
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS master_spk (
                        id SERIAL PRIMARY KEY,
                        no_spk TEXT,
                        kontraktor TEXT,
                        jenis_pekerjaan TEXT,
                        unit TEXT,
                        jumlah INTEGER DEFAULT 1,
                        nilai_pekerjaan REAL DEFAULT 0,
                        catatan TEXT,
                        UNIQUE(no_spk, jenis_pekerjaan)
                    );
                ''')

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
                conn.commit()
    except Exception as e:
        st.error(f"⚠️ Gagal inisialisasi database: {e}")

init_db()

# ---------------------------------------------------------
# EXPORT EXCEL FUNCTION
# ---------------------------------------------------------
def generate_excel_full_feature(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_excel = df.drop(columns=['real_id'], errors='ignore').copy()
        df_progress = df_excel.drop(columns=['Foto 1', 'Foto 2', 'Pratinjau Foto 1', 'Pratinjau Foto 2'], errors='ignore').copy()
        
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
            cell.alignment = align_center
            cell.border = border_standard

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

            def insert_image_visual(path, ws, current_row, current_col, target_col_width):
                cell_p = ws.cell(row=current_row, column=current_col)
                cell_p.border = border_standard
                
                if path and str(path).startswith("http"):
                    cell_p.value = f"Link: {path}"
                    cell_p.font = blue_link_font
                    cell_p.hyperlink = str(path)
                    cell_p.alignment = align_center
                elif path and os.path.exists(str(path)) and has_pil:
                    try:
                        pil_img = PILImage.open(path)
                        orig_w, orig_h = pil_img.size
                        target_width_px = int((target_col_width * 7.5) - 5)
                        target_height_px = int((orig_h / orig_w) * target_width_px)
                        
                        pil_img_resized = pil_img.resize((target_width_px, target_height_px), PILImage.Resampling.LANCZOS)
                        img_buffer = io.BytesIO()
                        pil_img_resized.save(img_buffer, format='JPEG')
                        img_buffer.seek(0)
                        
                        opx_img = OpenPyXLImage(img_buffer)
                        col_letter = get_column_letter(current_col)
                        ws.add_image(opx_img, f'{col_letter}{current_row}')
                    except Exception as e:
                        cell_p.value = f"Error: {e}"
                else:
                    cell_p.value = "Foto tidak tersedia"
                    cell_p.alignment = align_center

            insert_image_visual(row.get('Pratinjau Foto 1'), worksheet_foto, foto_row_idx, 2, LEBAR_KOLOM_FOTO)
            insert_image_visual(row.get('Pratinjau Foto 2'), worksheet_foto, foto_row_idx, 3, LEBAR_KOLOM_FOTO)

            foto_row_idx += 1

        try:
            no_col_idx = df_progress.columns.get_loc('No') + 1
            doc_col_idx = df_progress.columns.get_loc('Dokumentasi') + 1
            for p_row_idx in range(2, worksheet_progress.max_row + 1):
                no_value = worksheet_progress.cell(row=p_row_idx, column=no_col_idx).value
                if no_value in job_map_targets:
                    target_photo_row = job_map_targets[no_value]
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

# =========================================================
# 1. HALAMAN DASHBOARD PROGRESS (ISOLASI DENGAN IF)
# =========================================================
if menu == MENU_DASHBOARD:
    st.title("📊 WEEKLY REPORT DIVISI INFRASTRUCTURE")

    tab_bangka, tab_belitung, tab_semua, tab_history = st.tabs([
        "🌴 Laporan Progress Bangka", 
        "⛵ Laporan Progress Belitung", 
        "📋 Semua Progress Proyek", 
        "📜 Riwayat / History Update"
    ])

    query_view = """
        SELECT 
            id AS real_id,
            waktu_input AS "Waktu Input Terbaru",
            no_spk AS "Nomor SPK",
            kontraktor AS "Nama Kontraktor",
            jenis_pekerjaan AS "Jenis Pekerjaan",
            unit AS "Unit Proyek",
            CAST(COALESCE(jumlah, 1) AS INTEGER) AS "Jumlah",
            nilai_pekerjaan AS "Nilai Kontrak (Rp)",
            progress_minggu_lalu AS "Progress Minggu Lalu (%)",
            progress_minggu_ini AS "Progress Minggu Ini (%)",
            (COALESCE(progress_minggu_ini, 0) - COALESCE(progress_minggu_lalu, 0)) AS "Selisih / Varian (%)",
            catatan AS "Catatan Pekerjaan Terbaru",
            foto_1 AS "Pratinjau Foto 1",
            foto_2 AS "Pratinjau Foto 2"
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
            'Nilai Kontrak (Rp)', 'Progress Minggu Lalu (%)', 'Progress Minggu Ini (%)',
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
                "Nilai Kontrak (Rp)": st.column_config.NumberColumn("Nilai Kontrak (Rp)", format="Rp %d"),
                "Progress Minggu Lalu (%)": st.column_config.NumberColumn("Progress Minggu Lalu (%)", format="%.2f %%"),
                "Progress Minggu Ini (%)": st.column_config.NumberColumn("Progress Minggu Ini (%)", format="%.2f %%"),
                "Selisih / Varian (%)": st.column_config.NumberColumn("Selisih / Varian (%)", format="%.2f %%"),
                "Pratinjau Foto 1": st.column_config.ImageColumn("Pratinjau Foto 1"),
                "Pratinjau Foto 2": st.column_config.ImageColumn("Pratinjau Foto 2"),
            },
            key=editor_key
        )

        if not df_data.empty:
            if st.button("💾 Simpan Perubahan Data", key=f"btn_save_{tab_key_prefix}"):
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    for idx, row in edited_df.iterrows():
                        if idx < len(df_display):
                            real_id = df_display.iloc[idx].get('real_id')
                            if pd.notna(real_id):
                                cursor.execute("""
                                    UPDATE laporan_mingguan
                                    SET progress_minggu_ini = %s, catatan = %s
                                    WHERE id = %s
                                """, (row.get('Progress Minggu Ini (%)'), row.get('Catatan Pekerjaan Terbaru'), int(real_id)))
                    conn.commit()
                st.success("✅ Perubahan data berhasil disimpan!")
                st.rerun()

            st.markdown("---")
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
        st.subheader("🌐 Semua Laporan Progress Proyek - Wilayah Bangka")
        if not df_all.empty:
            df_bangka = df_all[df_all['Unit Proyek'].astype(str).str.contains('BANGKA|BKA', case=False, na=False)]
            render_dashboard_table(df_bangka, "bangka")
        else:
            render_dashboard_table(pd.DataFrame(), "bangka")

    with tab_belitung:
        st.subheader("🌐 Semua Laporan Progress Proyek - Wilayah Belitung")
        if not df_all.empty:
            pola_belitung = 'BELITUNG|BLT|BPSL|BPRE|BPT'
            df_belitung = df_all[
                df_all['Unit Proyek'].astype(str).str.contains(pola_belitung, case=False, na=False) |
                (~df_all['Unit Proyek'].astype(str).str.contains('BANGKA|BKA', case=False, na=False))
            ]
            render_dashboard_table(df_belitung, "belitung")
        else:
            render_dashboard_table(pd.DataFrame(), "belitung")

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

# =========================================================
# 2. HALAMAN INPUT PROGRESS MINGGUAN (ISOLASI DENGAN ELIF)
# =========================================================
elif menu == MENU_INPUT:
    st.title("📝 Input Laporan Progress Mingguan Berdasarkan SPK")
    st.markdown("---")

    with get_db_connection() as conn:
        df_master_all = pd.read_sql_query("SELECT * FROM master_spk", conn)

    tab_i_bangka, tab_i_belitung = st.tabs([
        "🌴 Input Progress Bangka", 
        "⛵ Input Progress Belitung"
    ])

    def render_input_form(df_master_wilayah, tab_key_prefix):
        if df_master_wilayah.empty:
            st.info("💡 Belum ada data master pekerjaan terdaftar untuk wilayah ini. Silakan daftarkan SPK di menu Kelola Master SPK.")
            return

        list_spk_unique = sorted(df_master_wilayah['no_spk'].unique().tolist())
        selected_spk_no = st.selectbox("Pilih Nomor SPK:", list_spk_unique, key=f"select_spk_no_{tab_key_prefix}")

        df_spk_filtered = df_master_wilayah[df_master_wilayah['no_spk'] == selected_spk_no]
        list_pekerjaan = df_spk_filtered['jenis_pekerjaan'].unique().tolist()
        selected_pekerjaan = st.selectbox("Pilih Jenis Pekerjaan:", list_pekerjaan, key=f"select_pekerjaan_{tab_key_prefix}")

        spk_data_selected = df_spk_filtered[df_spk_filtered['jenis_pekerjaan'] == selected_pekerjaan].iloc[0]

        st.info(f"""📌 **Detail SPK:**
*   **Nomor SPK:** {spk_data_selected['no_spk']}
*   **Kontraktor:** {spk_data_selected['kontraktor']}
*   **Jenis Pekerjaan:** {spk_data_selected['jenis_pekerjaan']}
*   **Unit/Wilayah:** {spk_data_selected['unit']}
""")

        prog_terakhir = 0.0
        catatan_terakhir = ""
        existing_foto_1, existing_foto_2 = None, None

        with get_db_connection() as conn:
            query_last = "SELECT progress_minggu_ini, catatan, foto_1, foto_2 FROM laporan_mingguan WHERE no_spk=%s AND jenis_pekerjaan=%s"
            existing_prog_df = pd.read_sql_query(query_last, conn, params=(spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan']))
            
            if not existing_prog_df.empty:
                prog_terakhir = float(existing_prog_df.iloc[0]['progress_minggu_ini'] or 0.0)
                catatan_terakhir = existing_prog_df.iloc[0]['catatan'] or ""
                existing_foto_1 = existing_prog_df.iloc[0]['foto_1']
                existing_foto_2 = existing_prog_df.iloc[0]['foto_2']

        with st.form(f"form_input_week_{tab_key_prefix}", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                prog_ini = st.number_input(f"Progress Minggu Ini (%) (Lalu: {prog_terakhir:.2f}%)", min_value=prog_terakhir, max_value=100.0, value=prog_terakhir, step=0.01)
                catatan_lap = st.text_area("Catatan Pekerjaan", value=catatan_terakhir)
            
            with col2:
                f_upload_1 = st.file_uploader("Upload Foto 1", type=["jpg", "jpeg", "png"], key=f"f1_{tab_key_prefix}")
                f_upload_2 = st.file_uploader("Upload Foto 2", type=["jpg", "jpeg", "png"], key=f"f2_{tab_key_prefix}")

            if st.form_submit_button("💾 Simpan Laporan"):
                path_f1_final, path_f2_final = existing_foto_1, existing_foto_2

                if f_upload_1:
                    path_f1_final = cloudinary.uploader.upload(f_upload_1).get("secure_url")
                if f_upload_2:
                    path_f2_final = cloudinary.uploader.upload(f_upload_2).get("secure_url")

                penambahan_week = prog_ini - prog_terakhir

                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("SELECT id FROM laporan_mingguan WHERE no_spk=%s AND jenis_pekerjaan=%s", (spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan']))
                    row_found = cursor.fetchone()

                    if row_found:
                        cursor.execute("""
                            UPDATE laporan_mingguan
                            SET waktu_input = CURRENT_TIMESTAMP, progress_minggu_lalu = %s, progress_minggu_ini = %s, catatan = %s, foto_1 = %s, foto_2 = %s
                            WHERE id = %s
                        """, (prog_terakhir, prog_ini, catatan_lap, path_f1_final, path_f2_final, row_found[0]))
                    else:
                        cursor.execute("""
                            INSERT INTO laporan_mingguan (no_spk, jenis_pekerjaan, kontraktor, unit, jumlah, nilai_pekerjaan, progress_minggu_lalu, progress_minggu_ini, catatan, foto_1, foto_2)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """, (spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan'], spk_data_selected['kontraktor'], spk_data_selected['unit'], spk_data_selected['jumlah'], spk_data_selected['nilai_pekerjaan'], prog_terakhir, prog_ini, catatan_lap, path_f1_final, path_f2_final))

                    cursor.execute("""
                        INSERT INTO history_progress (no_spk, jenis_pekerjaan, kontraktor, unit, progress_minggu_lalu, progress_minggu_ini, progres_penambahan, catatan, foto_1, foto_2)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan'], spk_data_selected['kontraktor'], spk_data_selected['unit'], prog_terakhir, prog_ini, penambahan_week, catatan_lap, path_f1_final, path_f2_final))

                    conn.commit()

                st.success("✅ Laporan minggu ini berhasil disimpan!")
                st.rerun()

    with tab_i_bangka:
        if not df_master_all.empty:
            render_input_form(df_master_all[df_master_all['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False)], "in_bangka")
        else:
            render_input_form(pd.DataFrame(), "in_bangka")

    with tab_i_belitung:
        if not df_master_all.empty:
            pola_belitung = 'BELITUNG|BLT|BPSL|BPRE|BPT'
            render_input_form(df_master_all[df_master_all['unit'].astype(str).str.contains(pola_belitung, case=False, na=False) | (~df_master_all['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False))], "in_belitung")
        else:
            render_input_form(pd.DataFrame(), "in_belitung")

# =========================================================
# 3. HALAMAN KELOLA MASTER SPK (ISOLASI DENGAN ELIF)
# =========================================================
elif menu == MENU_MASTER:
    st.title("⚙️ Kelola Master Data Pekerjaan / SPK")
    st.markdown("---")

    tab_m_bangka, tab_m_belitung, tab_m_tambah = st.tabs([
        "🌴 Master Data Bangka",
        "⛵ Master Data Belitung",
        "➕ Tambah SPK / Pekerjaan Baru"
    ])

    with get_db_connection() as conn:
        try:
            df_master = pd.read_sql_query("SELECT * FROM master_spk ORDER BY id ASC", conn)
        except Exception:
            df_master = pd.DataFrame()

    def render_master_table(df_m_data, tab_key):
        if df_m_data.empty:
            st.info(f"Belum ada data SPK untuk wilayah ini. Silakan tambah data di tab 'Tambah SPK / Pekerjaan Baru'.")
            return

        edited_master = st.data_editor(
            df_m_data,
            use_container_width=True,
            hide_index=True,
            key=f"editor_master_{tab_key}",
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "jumlah": st.column_config.NumberColumn("Jumlah", format="%d"),
                "nilai_pekerjaan": st.column_config.NumberColumn("Nilai Pekerjaan (Rp)", format="Rp %d"),
            }
        )

        col_m_save, col_m_del = st.columns([1, 1])
        with col_m_save:
            if st.button("💾 Simpan Perubahan Master", key=f"btn_save_m_{tab_key}"):
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    for idx, row in edited_master.iterrows():
                        cursor.execute("""
                            UPDATE master_spk
                            SET no_spk=%s, kontraktor=%s, jenis_pekerjaan=%s, unit=%s, jumlah=%s, nilai_pekerjaan=%s, catatan=%s
                            WHERE id=%s
                        """, (
                            row['no_spk'], row['kontraktor'], row['jenis_pekerjaan'],
                            row['unit'], row['jumlah'], row['nilai_pekerjaan'],
                            row['catatan'], row['id']
                        ))
                    conn.commit()
                st.success("✅ Data master berhasil diperbarui!")
                st.rerun()

        with col_m_del:
            id_to_delete = st.selectbox("Pilih ID Master SPK untuk dihapus:", df_m_data['id'].tolist(), key=f"sel_m_del_{tab_key}")
            if st.button("🗑️ Hapus Master SPK Dipilih", key=f"btn_del_m_{tab_key}"):
                with get_db_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM master_spk WHERE id=%s", (id_to_delete,))
                    conn.commit()
                st.success(f"Master SPK ID {id_to_delete} berhasil dihapus!")
                st.rerun()

    with tab_m_bangka:
        st.subheader("📋 Master Data Wilayah Bangka")
        if not df_master.empty:
            df_mb = df_master[df_master['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False)]
            render_master_table(df_mb, "bangka")
        else:
            render_master_table(pd.DataFrame(), "bangka")

    with tab_m_belitung:
        st.subheader("📋 Master Data Wilayah Belitung")
        if not df_master.empty:
            pola_belitung = 'BELITUNG|BLT|BPSL|BPRE|BPT'
            df_mbl = df_master[
                df_master['unit'].astype(str).str.contains(pola_belitung, case=False, na=False) |
                (~df_master['unit'].astype(str).str.contains('BANGKA|BKA', case=False, na=False))
            ]
            render_master_table(df_mbl, "belitung")
        else:
            render_master_table(pd.DataFrame(), "belitung")

    with tab_m_tambah:
        st.subheader("➕ Tambah Master SPK Baru")
        with st.form("form_add_master", clear_on_submit=True):
            col_a, col_b = st.columns(2)
            with col_a:
                no_spk = st.text_input("Nomor SPK*")
                kontraktor = st.text_input("Nama Kontraktor*")
                jenis_pekerjaan = st.text_input("Jenis Pekerjaan*")
            with col_b:
                unit = st.text_input("Unit / Wilayah (contoh: BANGKA / BELITUNG)*")
                jumlah = st.number_input("Jumlah Unit", min_value=1, value=1, step=1)
                nilai_pekerjaan = st.number_input("Nilai Pekerjaan (Rp)", min_value=0.0, value=0.0, step=1000.0)
            
            catatan = st.text_area("Catatan Tambahan")
            btn_add = st.form_submit_button("💾 Simpan Master SPK")

            if btn_add:
                if not no_spk or not jenis_pekerjaan or not kontraktor:
                    st.error("⚠️ Nomor SPK, Kontraktor, dan Jenis Pekerjaan wajib diisi!")
                else:
                    try:
                        with get_db_connection() as conn:
                            cursor = conn.cursor()
                            cursor.execute("""
                                INSERT INTO master_spk (no_spk, kontraktor, jenis_pekerjaan, unit, jumlah, nilai_pekerjaan, catatan)
                                VALUES (%s, %s, %s, %s, %s, %s, %s)
                            """, (no_spk, kontraktor, jenis_pekerjaan, unit, jumlah, nilai_pekerjaan, catatan))
                            conn.commit()
                        st.success("✅ Master SPK berhasil ditambahkan!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal menyimpan data master: {e}")
