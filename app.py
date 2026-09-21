import streamlit as st
import pandas as pd
import sqlite3
import os
import re
import io
from datetime import datetime

# ---------------------------------------------------------
# CONFIGURASI HALAMAN & KONSTANTA
# ---------------------------------------------------------
st.set_page_config(
    page_title="Weekly Report Divisi Infrastructure",
    page_icon="🏗️",
    layout="wide"
)

DB_FILE = "database_progress.db"
UPLOAD_DIR = "uploads"

if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# Menu Navigasi
MENU_DASHBOARD = "Dashboard Progress"
MENU_INPUT = "Input Progress Mingguan"
MENU_MASTER = "Kelola Master SPK"

# ---------------------------------------------------------
# HELPER FUNCTIONS & DATABASE SETUP
# ---------------------------------------------------------
def get_db_connection():
    return sqlite3.connect(DB_FILE)

def sanitize_filename(filename):
    return re.sub(r'[^a-zA-Z0-9_-]', '_', str(filename))

def init_db():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        # Tabel Master SPK
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS master_spk (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                no_spk TEXT NOT NULL,
                jenis_pekerjaan TEXT NOT NULL,
                kontraktor TEXT,
                unit TEXT,
                jumlah INTEGER,
                nilai_pekerjaan REAL,
                UNIQUE(no_spk, jenis_pekerjaan)
            )
        """)
        # Tabel Laporan Mingguan Utama
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS laporan_mingguan (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                no_spk TEXT NOT NULL,
                jenis_pekerjaan TEXT NOT NULL,
                kontraktor TEXT,
                unit TEXT,
                jumlah INTEGER,
                nilai_pekerjaan REAL,
                progress_minggu_lalu REAL DEFAULT 0,
                progress_minggu_ini REAL DEFAULT 0,
                catatan TEXT,
                foto_1 TEXT,
                foto_2 TEXT,
                waktu_input TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(no_spk, jenis_pekerjaan)
            )
        """)
        # Tabel Histori Progress
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS history_progress (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
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
                waktu_input TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()

init_db()

# ---------------------------------------------------------
# HELPER EXPORT EXCEL (AMAN DARI ERROR FOTO 1)
# ---------------------------------------------------------
def generate_excel_download(df_export):
    output = io.BytesIO()
    df_temp = df_export.copy()
    
    # Amankan penanganan kolom foto_1 & foto_2 jika ada di DataFrame
    for col in ['foto_1', 'foto_2']:
        if col in df_temp.columns:
            df_temp[col] = df_temp[col].apply(
                lambda x: "Ada Foto" if pd.notna(x) and str(x).strip() != "" else "Tidak Ada"
            )

    rename_dict = {
        'no': 'No',
        'no_spk': 'Nomor SPK',
        'kontraktor': 'Nama Kontraktor',
        'jenis_pekerjaan': 'Jenis Pekerjaan',
        'unit': 'Unit Proyek',
        'jumlah': 'Jumlah',
        'nilai_pekerjaan': 'Nilai Kontrak (Rp)',
        'progress_minggu_lalu': 'Progress Minggu Lalu (%)',
        'progress_minggu_ini': 'Progress Minggu Ini (%)',
        'progres_penambahan': 'Penambahan Progress (%)',
        'catatan': 'Catatan / Kendala',
        'foto_1': 'Status Foto 1',
        'foto_2': 'Status Foto 2',
        'waktu_input': 'Waktu Update'
    }
    
    df_temp = df_temp.rename(columns={k: v for k, v in rename_dict.items() if k in df_temp.columns})

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_temp.to_excel(writer, index=False, sheet_name='Laporan Progress')
        
    return output.getvalue()

# ---------------------------------------------------------
# SIDEBAR NAVIGATION
# ---------------------------------------------------------
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/2/23/Sinarmas_logo.svg", width=180)
st.sidebar.title("Navigasi")
menu = st.sidebar.selectbox("Pilih Menu", [MENU_DASHBOARD, MENU_INPUT, MENU_MASTER])

st.sidebar.markdown("---")
if st.sidebar.button("🚪 Logout"):
    st.sidebar.info("Logout berhasil.")

# ---------------------------------------------------------
# MENU 1: DASHBOARD PROGRESS
# ---------------------------------------------------------
if menu == MENU_DASHBOARD:
    st.title("📊 WEEKLY REPORT DIVISI INFRASTRUCTURE")
    st.markdown("---")

    with get_db_connection() as conn:
        df_laporan_all = pd.read_sql_query("SELECT * FROM laporan_mingguan", conn)
        df_history_all = pd.read_sql_query("SELECT * FROM history_progress ORDER BY id DESC", conn)

    tab_d_bangka, tab_d_belitung, tab_d_all, tab_d_hist = st.tabs([
        "🏝️ Laporan Progress Bangka",
        "🏖️ Laporan Progress Belitung",
        "📋 Semua Progress Proyek",
        "📜 Riwayat / History Update"
    ])

    def render_dashboard_table(df_data, title_prefix):
        st.subheader(f"📍 {title_prefix}")
        if df_data.empty:
            st.info("Belum ada data laporan progress untuk kategori ini.")
            return

        df_display = df_data.copy()
        if 'id' in df_display.columns:
            df_display = df_display.drop(columns=['id'])

        # Display Dataframe / Editor
        edited_df = st.data_editor(
            df_display,
            num_rows="dynamic",
            use_container_width=True,
            key=f"editor_{title_prefix}"
        )

        if st.button("💾 Simpan Perubahan Data", key=f"btn_save_{title_prefix}"):
            st.success("Perubahan berhasil diperbarui.")

        # --- SECTION EXPORT & DOWNLOAD LAPORAN ---
        st.markdown("### 📥 Export & Download Laporan")
        try:
            excel_bytes = generate_excel_download(df_data)
            st.download_button(
                label=f"📊 Download Laporan {title_prefix} (Excel .xlsx)",
                data=excel_bytes,
                file_name=f"Laporan_{sanitize_filename(title_prefix)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{title_prefix}"
            )
        except Exception as e:
            st.error(f"Gagal memproses file Excel: {e}")

    with tab_d_bangka:
        df_bka = df_laporan_all[df_laporan_all['unit'].str.contains('BANGKA|BKA', case=False, na=False)]
        render_dashboard_table(df_bka, "Laporan Progress Proyek - Wilayah Bangka")

    with tab_d_belitung:
        df_blt = df_laporan_all[df_laporan_all['unit'].str.contains('BELITUNG|BLT', case=False, na=False)]
        render_dashboard_table(df_blt, "Laporan Progress Proyek - Wilayah Belitung")

    with tab_d_all:
        render_dashboard_table(df_laporan_all, "Semua Progress Proyek Infrastructure")

    with tab_d_hist:
        st.subheader("📜 Riwayat Update Progress")
        if df_history_all.empty:
            st.info("Belum ada riwayat histori.")
        else:
            st.dataframe(df_history_all, use_container_width=True)

# ---------------------------------------------------------
# MENU 2: INPUT PROGRESS (TAB WILAYAH & GROUPING SPK)
# ---------------------------------------------------------
elif menu == MENU_INPUT:
    st.title("📝 Input Laporan Progress Mingguan Berdasarkan SPK")
    st.markdown("---")

    with get_db_connection() as conn:
        df_master_all = pd.read_sql_query("SELECT * FROM master_spk", conn)

    if df_master_all.empty:
        st.warning("⚠️ Belum ada data Master SPK. Harap daftarkan SPK terlebih dahulu di menu Kelola Master SPK.")
    else:
        tab_i_bangka, tab_i_belitung = st.tabs([
            "🏝️ Input Progress Bangka", 
            "🏖️ Input Progress Belitung"
        ])

        # Helper Render Form Input dengan Grouping SPK
        def render_input_form(df_master_wilayah, tab_key_prefix):
            if df_master_wilayah.empty:
                st.info("💡 Belum ada data master pekerjaan terdaftar untuk wilayah ini.")
                return

            # Step 1: Dropdown Pilihan Nomor SPK
            list_spk_unique = sorted(df_master_wilayah['no_spk'].unique().tolist())
            selected_spk_no = st.selectbox(
                "Pilih Nomor SPK:",
                list_spk_unique,
                key=f"select_spk_no_{tab_key_prefix}"
            )

            # Filter data master berdasarkan Nomor SPK
            df_spk_filtered = df_master_wilayah[df_master_wilayah['no_spk'] == selected_spk_no]

            # Step 2: Dropdown Jenis Pekerjaan Ter-group dalam SPK tersebut
            list_pekerjaan = df_spk_filtered['jenis_pekerjaan'].unique().tolist()
            selected_pekerjaan = st.selectbox(
                "Pilih Jenis Pekerjaan (Tergroup berdasarkan SPK):",
                list_pekerjaan,
                key=f"select_pekerjaan_{tab_key_prefix}"
            )

            # Ambil detail baris SPK
            spk_data_selected = df_spk_filtered[df_spk_filtered['jenis_pekerjaan'] == selected_pekerjaan].iloc[0]

            # Format angka Nilai Kontrak
            try:
                val_num = float(spk_data_selected['nilai_pekerjaan'])
                nilai_formatted = f"Rp {val_num:,.2f}"
            except (ValueError, TypeError):
                nilai_formatted = "-"

            # Ringkasan Detail SPK
            st.info(f"""📌 **Detail SPK Dipilih:** 
*   **Nomor SPK:** {spk_data_selected['no_spk']}
*   **Kontraktor:** {spk_data_selected['kontraktor']}
*   **Jenis Pekerjaan:** {spk_data_selected['jenis_pekerjaan']}
*   **Unit/Wilayah:** {spk_data_selected['unit']}
*   **Nilai Kontrak:** {nilai_formatted}
""")

            # Ambil progress minggu lalu dari db laporan_mingguan
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

            # Display Foto Terakhir jika ada
            if existing_foto_1 or existing_foto_2:
                st.markdown("**📸 Pratinjau Foto Dokumentasi Terakhir:**")
                c_img1, c_img2 = st.columns(2)
                with c_img1:
                    if existing_foto_1 and os.path.exists(str(existing_foto_1)):
                        st.image(existing_foto_1, caption="Foto Dokumentasi 1", use_container_width=True)
                with c_img2:
                    if existing_foto_2 and os.path.exists(str(existing_foto_2)):
                        st.image(existing_foto_2, caption="Foto Dokumentasi 2", use_container_width=True)

            # --- FORM INPUT ---
            with st.form(f"form_input_week_{tab_key_prefix}", clear_on_submit=True):
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("📝 Progress Minggu Ini")
                    prog_ini = st.number_input(
                        f"Progress Akumulatif Minggu Ini (%) - (Hingga Minggu Lalu: {prog_terakhir:.2f}%)", 
                        min_value=prog_terakhir,
                        max_value=100.0, 
                        value=prog_terakhir,
                        step=0.1
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
                        st.error("⚠️ Progress minggu ini tidak boleh lebih kecil dari minggu lalu!")
                    else:
                        import time
                        ts = int(time.time())
                        spk_fniz = sanitize_filename(spk_data_selected['no_spk'])
                        
                        if f_upload_1:
                            path_f1_final = os.path.join(UPLOAD_DIR, f"{spk_fniz}_f1_{ts}.jpg")
                            with open(path_f1_final, "wb") as f: f.write(f_upload_1.getbuffer())

                        if f_upload_2:
                            path_f2_final = os.path.join(UPLOAD_DIR, f"{spk_fniz}_f2_{ts}.jpg")
                            with open(path_f2_final, "wb") as f: f.write(f_upload_2.getbuffer())
                        
                        penambahan_week = prog_ini - prog_terakhir

                        with get_db_connection() as conn:
                            cursor = conn.cursor()
                            cursor.execute("DELETE FROM laporan_mingguan WHERE no_spk=? AND jenis_pekerjaan=?", (spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan']))
                            cursor.execute("""
                                INSERT INTO laporan_mingguan (
                                    no_spk, jenis_pekerjaan, kontraktor, unit, jumlah, nilai_pekerjaan, 
                                    progress_minggu_lalu, progress_minggu_ini, catatan, foto_1, foto_2, waktu_input
                                ) VALUES (?,?,?,?,?,?,?,?,?,?,?, CURRENT_TIMESTAMP)""", 
                                (spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan'], spk_data_selected['kontraktor'], spk_data_selected['unit'], spk_data_selected['jumlah'], spk_data_selected['nilai_pekerjaan'], prog_terakhir, prog_ini, catatan_lap, path_f1_final, path_f2_final))
                            
                            cursor.execute("""
                                INSERT INTO history_progress (
                                    no_spk, jenis_pekerjaan, kontraktor, unit, 
                                    progress_minggu_lalu, progress_minggu_ini, progres_penambahan, catatan, foto_1, foto_2, waktu_input
                                ) VALUES (?,?,?,?,?,?,?,?,?,?, CURRENT_TIMESTAMP)""",
                                (spk_data_selected['no_spk'], spk_data_selected['jenis_pekerjaan'], spk_data_selected['kontraktor'], spk_data_selected['unit'], prog_terakhir, prog_ini, penambahan_week, catatan_lap, path_f1_final, path_f2_final))
                            
                            conn.commit()
                        st.success(f"✅ Laporan mingguan untuk SPK '{selected_spk_no}' - '{selected_pekerjaan}' berhasil disimpan!")
                        st.rerun()

        # --- TAB INPUT BANGKA ---
        with tab_i_bangka:
            st.subheader("🏝️ Input Progress - Wilayah Bangka")
            df_bangka_master = df_master_all[df_master_all['unit'].str.contains('BANGKA|BKA', case=False, na=False)]
            render_input_form(df_bangka_master, "bangka")

        # --- TAB INPUT BELITUNG ---
        with tab_i_belitung:
            st.subheader("🏖️ Input Progress - Wilayah Belitung")
            df_belitung_master = df_master_all[df_master_all['unit'].str.contains('BELITUNG|BLT', case=False, na=False)]
            render_input_form(df_belitung_master, "belitung")

# ---------------------------------------------------------
# MENU 3: KELOLA MASTER SPK
# ---------------------------------------------------------
elif menu == MENU_MASTER:
    st.title("⚙️ Kelola Master Data SPK & Pekerjaan")
    st.markdown("---")

    st.subheader("➕ Tambah Master SPK Baru")
    with st.form("form_add_master", clear_on_submit=True):
        m_no_spk = st.text_input("Nomor SPK")
        m_jenis = st.text_input("Jenis Pekerjaan")
        m_kontraktor = st.text_input("Nama Kontraktor")
        m_unit = st.selectbox("Unit / Wilayah Proyek", ["UNIT BANGKA", "UNIT BELITUNG"])
        m_jumlah = st.number_input("Jumlah Unit/Item", min_value=1, value=1)
        m_nilai = st.number_input("Nilai Kontrak Pekerjaan (Rp)", min_value=0.0, step=1000000.0)

        sub_master = st.form_submit_button("➕ Tambahkan Ke Master Data")

        if sub_master:
            if not m_no_spk or not m_jenis:
                st.error("⚠️ Nomor SPK dan Jenis Pekerjaan wajib diisi!")
            else:
                try:
                    with get_db_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("""
                            INSERT INTO master_spk (no_spk, jenis_pekerjaan, kontraktor, unit, jumlah, nilai_pekerjaan)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (m_no_spk, m_jenis, m_kontraktor, m_unit, m_jumlah, m_nilai))
                        conn.commit()
                    st.success("✅ Master Data SPK Berhasil Ditambahkan!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("⚠️ Kombinasi Nomor SPK dan Jenis Pekerjaan ini sudah ada di database.")

    st.markdown("---")
    st.subheader("📋 Daftar Master SPK Terdaftar")
    with get_db_connection() as conn:
        df_master_view = pd.read_sql_query("SELECT * FROM master_spk", conn)

    if not df_master_view.empty:
        st.dataframe(df_master_view, use_container_width=True)
    else:
        st.info("Belum ada data master SPK.")
