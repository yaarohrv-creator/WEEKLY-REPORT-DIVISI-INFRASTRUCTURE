import os
import io
import sqlite3
import pandas as pd
import streamlit as st

# Modul untuk styling dan menyisipkan gambar ke Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Konfigurasi Halaman Streamlit
st.set_page_config(page_title="Sistem Progress Proyek", layout="wide")

# ---------------------------------------------------------
# FUNGSIONALITAS DATABASE SQLITE
# ---------------------------------------------------------
def get_connection():
    conn = sqlite3.connect('proyek.db')
    cursor = conn.cursor()

    # 1. Tabel Master SPK (Data Induk Pekerjaan/Proyek)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS master_spk (
            no_spk TEXT PRIMARY KEY,
            jenis_pekerjaan TEXT,
            kontraktor TEXT,
            unit TEXT,
            nilai_kontrak REAL
        )
    ''')

    # 2. Tabel Laporan Progress Mingguan
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS laporan_mingguan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            waktu_input DATETIME DEFAULT CURRENT_TIMESTAMP,
            no_spk TEXT,
            jenis_pekerjaan TEXT,
            kontraktor TEXT,
            unit TEXT,
            nilai_kontrak REAL,
            progress_minggu_lalu REAL,
            progress_minggu_ini REAL,
            catatan TEXT
        )
    ''')

    # 3. Tabel Dokumentasi Foto
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS dokumentasi_foto (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tipe_dok TEXT,
            unit TEXT,
            nama_file TEXT
        )
    ''')

    conn.commit()
    return conn

# Inisialisasi koneksi database
conn = get_connection()

# ---------------------------------------------------------
# NAVIGASI SIDEBAR
# ---------------------------------------------------------
st.sidebar.title("Navigasi")
menu = st.sidebar.selectbox("Pilih Menu", [
    "Dashboard Progress", 
    "Input Progress Mingguan",
    "Kelola Master SPK", 
    "Upload Dokumentasi Foto"
])

# ---------------------------------------------------------
# MENU 1: DASHBOARD PROGRESS
# ---------------------------------------------------------
if menu == "Dashboard Progress":
    st.title("📊 Dashboard Progress Proyek")

    query_view = """
        SELECT 
            id AS real_id,
            waktu_input AS [Waktu Input],
            no_spk AS [Nomor SPK],
            jenis_pekerjaan AS [Jenis Pekerjaan],
            kontraktor AS [Nama Kontraktor],
            unit AS [Unit Proyek],
            nilai_kontrak AS [Nilai Kontrak],
            progress_minggu_lalu AS [Progress Minggu Lalu (%)],
            progress_minggu_ini AS [Progress Minggu Ini (%)],
            (COALESCE(progress_minggu_ini, 0) - COALESCE(progress_minggu_lalu, 0)) AS [Selisih / Varian (%)],
            catatan AS [Catatan Pekerjaan]
        FROM laporan_mingguan
        ORDER BY id ASC
    """

    try:
        df_view = pd.read_sql_query(query_view, conn)
    except Exception:
        df_view = pd.DataFrame()

    if df_view.empty:
        df_view = pd.DataFrame(columns=[
            'No', 'real_id', 'Waktu Input', 'Nomor SPK', 'Jenis Pekerjaan',
            'Nama Kontraktor', 'Unit Proyek', 'Nilai Kontrak',
            'Progress Minggu Lalu (%)', 'Progress Minggu Ini (%)',
            'Selisih / Varian (%)', 'Catatan Pekerjaan'
        ])
        st.info("💡 Belum ada data progress. Silakan kelola Master SPK atau input progress mingguan.")
    else:
        if 'No' not in df_view.columns:
            df_view.insert(0, 'No', range(1, len(df_view) + 1))

    edited_df = st.data_editor(
        df_view,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="editor_dashboard"
    )

    if st.button("💾 Simpan Perubahan & Hapus Data"):
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
                        jenis_pekerjaan = ?,
                        kontraktor = ?,
                        unit = ?,
                        nilai_kontrak = ?,
                        progress_minggu_lalu = ?,
                        progress_minggu_ini = ?,
                        catatan = ?
                    WHERE id = ?
                """, (
                    row.get('Nomor SPK'),
                    row.get('Jenis Pekerjaan'),
                    row.get('Nama Kontraktor'),
                    row.get('Unit Proyek'),
                    row.get('Nilai Kontrak'),
                    row.get('Progress Minggu Lalu (%)'),
                    row.get('Progress Minggu Ini (%)'),
                    row.get('Catatan Pekerjaan'),
                    real_id
                ))

        conn.commit()
        st.success("Perubahan data berhasil disimpan!")
        st.rerun()

    st.markdown("---")
    st.subheader("📥 Export & Download Laporan")

    def generate_excel():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_excel = df_view.drop(columns=['real_id'], errors='ignore')
            df_excel.to_excel(writer, index=False, sheet_name='Laporan Progress')
            
            worksheet = writer.sheets['Laporan Progress']
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
            )

            for col_num in range(1, len(df_excel.columns) + 1):
                cell = worksheet.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

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
        excel_data = generate_excel()
        st.download_button(
            label="📥 Download Laporan (Excel)",
            data=excel_data,
            file_name="Laporan_Progress_Proyek.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Gagal memproses file Excel: {e}")

# ---------------------------------------------------------
# MENU 2: KELOLA MASTER SPK (DATA INDUK)
# ---------------------------------------------------------
elif menu == "Kelola Master SPK":
    st.title("📑 Kelola Master SPK Proyek")
    st.write("Daftarkan Nomor SPK beserta informasi induknya di sini agar dapat digunakan pada pembaruan progress mingguan.")

    with st.form("form_master_spk"):
        c1, c2 = st.columns(2)
        with c1:
            no_spk = st.text_input("Nomor SPK (Unique ID)", placeholder="Contoh: SPK/INFRA/2026/001")
            jenis_pekerjaan = st.text_input("Jenis Pekerjaan", placeholder="Contoh: Pekerjaan Paving & Drainase")
            kontraktor = st.text_input("Nama Kontraktor", placeholder="Contoh: PT Utama Karya")
        with c2:
            unit = st.text_input("Unit Proyek", placeholder="Contoh: Kawasan Industri - Blok C")
            nilai_kontrak = st.number_input("Nilai Kontrak (Rp)", min_value=0.0, step=1000000.0, format="%.2f")

        submit_master = st.form_submit_button("➕ Tambah Ke Master SPK")

        if submit_master:
            if not no_spk or not kontraktor:
                st.warning("Nomor SPK dan Nama Kontraktor wajib diisi.")
            else:
                try:
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO master_spk (no_spk, jenis_pekerjaan, kontraktor, unit, nilai_kontrak)
                        VALUES (?, ?, ?, ?, ?)
                    """, (no_spk, jenis_pekerjaan, kontraktor, unit, nilai_kontrak))
                    conn.commit()
                    st.success(f"✅ Master SPK '{no_spk}' berhasil ditambahkan!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("⚠️ Nomor SPK ini sudah terdaftar di database!")

    st.markdown("---")
    st.subheader("📋 Daftar Master SPK")
    try:
        df_master = pd.read_sql_query("SELECT no_spk AS [Nomor SPK], jenis_pekerjaan AS [Jenis Pekerjaan], kontraktor AS [Nama Kontraktor], unit AS [Unit Proyek], nilai_kontrak AS [Nilai Kontrak] FROM master_spk", conn)
        st.dataframe(df_master, use_container_width=True)
    except Exception:
        st.info("Belum ada data Master SPK.")

# ---------------------------------------------------------
# MENU 3: INPUT PROGRESS MINGGUAN (OTOMATIS BERDASARKAN MASTER SPK)
# ---------------------------------------------------------
elif menu == "Input Progress Mingguan":
    st.title("📝 Input Progress Mingguan Berdasarkan SPK")

    # Ambil daftar SPK dari Master
    try:
        master_list = pd.read_sql_query("SELECT no_spk FROM master_spk", conn)['no_spk'].tolist()
    except Exception:
        master_list = []

    if not master_list:
        st.warning("⚠️ Belum ada data Master SPK. Harap daftarkan SPK di menu 'Kelola Master SPK' terlebih dahulu.")
    else:
        selected_spk = st.selectbox("Pilih Nomor SPK", master_list)

        # Ambil Detail Data dari Master SPK
        spk_detail = pd.read_sql_query("SELECT * FROM master_spk WHERE no_spk = ?", conn, params=(selected_spk,)).iloc[0]

        # Ambil Progress Terakhir dari Laporan Mingguan jika ada
        last_progress_query = """
            SELECT progress_minggu_ini 
            FROM laporan_mingguan 
            WHERE no_spk = ? 
            ORDER BY id DESC LIMIT 1
        """
        last_progress_df = pd.read_sql_query(last_progress_query, conn, params=(selected_spk,))
        
        default_progress_lalu = 0.0
        if not last_progress_df.empty and pd.notna(last_progress_df.iloc[0]['progress_minggu_ini']):
            default_progress_lalu = float(last_progress_df.iloc[0]['progress_minggu_ini'])

        st.info(f"**Detail SPK Dipilih:** {spk_detail['jenis_pekerjaan']} | **Kontraktor:** {spk_detail['kontraktor']} | **Unit:** {spk_detail['unit']}")

        with st.form("form_update_progress_mingguan"):
            col1, col2 = st.columns(2)

            with col1:
                st.text_input("Jenis Pekerjaan", value=spk_detail['jenis_pekerjaan'], disabled=True)
                st.text_input("Nama Kontraktor", value=spk_detail['kontraktor'], disabled=True)
                st.text_input("Unit Proyek", value=spk_detail['unit'], disabled=True)

            with col2:
                st.number_input("Nilai Kontrak (Rp)", value=float(spk_detail['nilai_kontrak']), disabled=True)
                prog_lalu = st.number_input("Progress Minggu Lalu (%)", value=default_progress_lalu, min_value=0.0, max_value=100.0)
                prog_ini = st.number_input("Progress Minggu Ini (%)", min_value=0.0, max_value=100.0, step=0.1)
                catatan = st.text_area("Catatan Pekerjaan Minggu Ini", placeholder="Masukkan kendala / progres pekerjaan...")

            submit_progress = st.form_submit_button("💾 Simpan Progress Minggu Ini")

            if submit_progress:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO laporan_mingguan (
                        no_spk, jenis_pekerjaan, kontraktor, unit, nilai_kontrak,
                        progress_minggu_lalu, progress_minggu_ini, catatan
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    selected_spk,
                    spk_detail['jenis_pekerjaan'],
                    spk_detail['kontraktor'],
                    spk_detail['unit'],
                    spk_detail['nilai_kontrak'],
                    prog_lalu,
                    prog_ini,
                    catatan
                ))
                conn.commit()
                st.success(f"✅ Progress mingguan untuk SPK '{selected_spk}' berhasil diperbarui!")

# ---------------------------------------------------------
# MENU 4: UPLOAD DOKUMENTASI FOTO
# ---------------------------------------------------------
elif menu == "Upload Dokumentasi Foto":
    st.title("📷 Upload Dokumentasi Foto Pekerjaan")

    with st.form("form_upload_foto"):
        tipe_dok = st.text_input("Tipe Dokumentasi (misal: Progress Atap, Pondasi)")
        unit_proyek = st.text_input("Nama Unit Proyek (Harus sama persis dengan di laporan)")
        uploaded_files = st.file_uploader("Pilih Foto Dokumentasi", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)
        
        submit_foto = st.form_submit_button("📷 Simpan Semua Foto")

        if submit_foto:
            if not unit_proyek or not uploaded_files:
                st.warning("Mohon isi Nama Unit Proyek dan unggah minimal satu foto.")
            else:
                cursor = conn.cursor()
                for file in uploaded_files:
                    file_path = os.path.join(".", file.name)
                    with open(file_path, "wb") as f:
                        f.write(file.getbuffer())

                    cursor.execute("""
                        INSERT INTO dokumentasi_foto (tipe_dok, unit, nama_file)
                        VALUES (?, ?, ?)
                    """, (tipe_dok, unit_proyek, file.name))

                conn.commit()
                st.success("📷 Foto berhasil diunggah dan disimpan ke database!")
