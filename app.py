import os
import io
import sqlite3
import pandas as pd
import streamlit as st

# Modul untuk styling dan menyisipkan gambar ke Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image

# Konfigurasi Halaman Streamlit
st.set_page_config(page_title="Sistem Progress Proyek", layout="wide")

# ---------------------------------------------------------
# FUNGSIONALITAS DATABASE SQLITE
# ---------------------------------------------------------
def get_connection():
    conn = sqlite3.connect('proyek.db')
    cursor = conn.cursor()

    # Tabel Laporan Utama
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS laporan_mingguan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            waktu_input DATETIME DEFAULT CURRENT_TIMESTAMP,
            jenis_pekerjaan TEXT,
            no_spk TEXT,
            kontraktor TEXT,
            unit TEXT,
            nilai_kontrak REAL,
            progress_minggu_lalu REAL,
            progress_minggu_ini REAL,
            catatan TEXT
        )
    ''')

    # Tabel Dokumentasi Foto
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
menu = st.sidebar.selectbox("Pilih Menu", ["Dashboard Progress", "Upload Dokumentasi Foto"])

# ---------------------------------------------------------
# MENU 1: DASHBOARD PROGRESS
# ---------------------------------------------------------
if menu == "Dashboard Progress":
    st.title("📊 Dashboard Progress Proyek")

    # Query membaca SELURUH kolom secara lengkap
    query_view = """
        SELECT 
            id AS real_id,
            waktu_input AS [Waktu Input],
            jenis_pekerjaan AS [Jenis Pekerjaan],
            no_spk AS [Nomor SPK],
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

    # Jika tabel masih kosong, siapkan struktur kolom lengkap
    if df_view.empty:
        df_view = pd.DataFrame(columns=[
            'No', 'real_id', 'Waktu Input', 'Jenis Pekerjaan', 'Nomor SPK',
            'Nama Kontraktor', 'Unit Proyek', 'Nilai Kontrak',
            'Progress Minggu Lalu (%)', 'Progress Minggu Ini (%)',
            'Selisih / Varian (%)', 'Catatan Pekerjaan'
        ])
        st.info("💡 Belum ada data progress. Tabel di bawah siap menerima data baru.")
    else:
        if 'No' not in df_view.columns:
            df_view.insert(0, 'No', range(1, len(df_view) + 1))

    # Tampilkan Tabel Data Interaktif
    edited_df = st.data_editor(
        df_view,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="editor_dashboard"
    )

    if st.button("💾 Simpan Perubahan & Hapus Data"):
        cursor = conn.cursor()
        
        # Ambil ID yang ada di tabel editor saat ini
        current_ids = [row['real_id'] for idx, row in edited_df.iterrows() if pd.notna(row.get('real_id'))]

        # 1. Hapus data dari database jika barisnya dihapus dari tabel editor
        if current_ids:
            format_strings = ','.join(['?'] * len(current_ids))
            cursor.execute(f"DELETE FROM laporan_mingguan WHERE id NOT IN ({format_strings})", current_ids)
        else:
            cursor.execute("DELETE FROM laporan_mingguan")

        # 2. Update data lama & Insert data baru yang diinput di tabel
        for idx, row in edited_df.iterrows():
            real_id = row.get('real_id')
            
            if pd.notna(real_id) and real_id != "":
                # Update baris lama
                cursor.execute("""
                    UPDATE laporan_mingguan
                    SET jenis_pekerjaan = ?,
                        no_spk = ?,
                        kontraktor = ?,
                        unit = ?,
                        nilai_kontrak = ?,
                        progress_minggu_lalu = ?,
                        progress_minggu_ini = ?,
                        catatan = ?
                    WHERE id = ?
                """, (
                    row.get('Jenis Pekerjaan'),
                    row.get('Nomor SPK'),
                    row.get('Nama Kontraktor'),
                    row.get('Unit Proyek'),
                    row.get('Nilai Kontrak'),
                    row.get('Progress Minggu Lalu (%)'),
                    row.get('Progress Minggu Ini (%)'),
                    row.get('Catatan Pekerjaan'),
                    real_id
                ))
            else:
                # Insert baris baru jika ada data baru diisi
                if pd.notna(row.get('Nama Kontraktor')) or pd.notna(row.get('Unit Proyek')):
                    cursor.execute("""
                        INSERT INTO laporan_mingguan (
                            jenis_pekerjaan, no_spk, kontraktor, unit, 
                            nilai_kontrak, progress_minggu_lalu, progress_minggu_ini, catatan
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        row.get('Jenis Pekerjaan'),
                        row.get('Nomor SPK'),
                        row.get('Nama Kontraktor'),
                        row.get('Unit Proyek'),
                        row.get('Nilai Kontrak'),
                        row.get('Progress Minggu Lalu (%)'),
                        row.get('Progress Minggu Ini (%)'),
                        row.get('Catatan Pekerjaan')
                    ))

        conn.commit()
        st.success("Perubahan data berhasil disimpan ke database!")
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
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
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
# MENU 2: UPLOAD DOKUMENTASI FOTO
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
                st.success("Foto berhasil diunggah dan disimpan ke database!")
