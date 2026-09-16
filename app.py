import os
import io
import sqlite3
import pandas as pd
import streamlit as st

# Modul untuk styling dan menyisipkan gambar ke Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image

st.set_page_config(page_title="Sistem Progress Proyek", layout="wide")

def get_connection():
    conn = sqlite3.connect('proyek.db')
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS laporan_mingguan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            jenis_pekerjaan TEXT,
            no_spk TEXT,
            kontraktor TEXT,
            unit TEXT,
            nilai_kontrak REAL,
            progress_minggu_lalu REAL,
            progress_minggu_ini REAL,
            selisih_progress REAL,
            catatan TEXT
        )
    ''')
    
    # Tambah kolom baru jika database sudah ada sebelumnya
    try:
        cursor.execute("ALTER TABLE laporan_mingguan ADD COLUMN nilai_kontrak REAL DEFAULT 0")
    except sqlite3.OperationalError:
        pass

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

# --- SIDEBAR NAVIGATION ---
# Menampilkan Logo Perusahaan di bagian atas sidebar
if os.path.exists("logo.png"):
    st.sidebar.image("logo.png", use_container_width=True)

st.sidebar.title("Navigasi")
menu = st.sidebar.selectbox("Pilih Menu", ["Dashboard Progress", "Input Progress Mingguan", "Upload Dokumentasi Foto"])

if menu == "Dashboard Progress":
    st.header("📊 WEEKLY REPORT DIVISI INFRASTRUCTURE")

    conn = get_connection()
    
    # 1. Query ambil data asli
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
            selisih_progress AS [Selisih / Varian (%)],
            catatan AS [Catatan Pekerjaan]
        FROM laporan_mingguan
        ORDER BY real_id ASC
    """
    df_view = pd.read_sql_query(query_view, conn)
    
    if df_view.empty:
        df_view = pd.DataFrame(columns=[
            'No', 'real_id', 'Waktu Input', 'Jenis Pekerjaan', 'Nomor SPK', 
            'Nama Kontraktor', 'Unit Proyek', 'Nilai Kontrak', 
            'Progress Minggu Lalu (%)', 'Progress Minggu Ini (%)', 
            'Selisih / Varian (%)', 'Catatan Pekerjaan'
        ])
        st.info("💡 Belum ada data progress. Tabel di bawah siap menerima data baru.")
    else:
        # 2. Buat nomor urut tampilan yang otomatis menyesuaikan jumlah data (1, 2, 3, ...)
        df_view.insert(0, 'No', range(1, len(df_view) + 1))

    # 3. Tampilkan tabel dengan menyembunyikan 'real_id' agar tidak membingungkan
    edited_df = st.data_editor(
        df_view, 
        num_rows="dynamic", 
        key="data_editor",
        column_config={"real_id": None} # Menyembunyikan ID asli database dari layar
    )
    
    if not df_view.empty:
        if st.button("💾 Simpan Perubahan"):
            cursor = conn.cursor()
            id_tersisa = edited_df['real_id'].dropna().tolist()
            
            if id_tersisa:
                query_delete = f"DELETE FROM laporan_mingguan WHERE id NOT IN ({','.join(['?']*len(id_tersisa))})"
                cursor.execute(query_delete, id_tersisa)
            else:
                cursor.execute("DELETE FROM laporan_mingguan")

            for idx, row in edited_df.iterrows():
                jenis_pekerjaan_val = "" if pd.isna(row['Jenis Pekerjaan']) or str(row['Jenis Pekerjaan']) == "None" else str(row['Jenis Pekerjaan'])
                no_spk_val = "" if pd.isna(row['Nomor SPK']) or str(row['Nomor SPK']) == "None" else str(row['Nomor SPK'])
                nilai_kontrak_val = 0 if pd.isna(row['Nilai Kontrak']) else float(row['Nilai Kontrak'])
                
                cursor.execute("""
                    UPDATE laporan_mingguan 
                    SET jenis_pekerjaan = ?, no_spk = ?, kontraktor = ?, unit = ?, nilai_kontrak = ?,
                        progress_minggu_lalu = ?, progress_minggu_ini = ?, 
                        selisih_progress = ?, catatan = ?
                    WHERE id = ?
                """, (
                    jenis_pekerjaan_val,
                    no_spk_val,
                    row['Nama Kontraktor'], 
                    row['Unit Proyek'], 
                    nilai_kontrak_val,
                    row['Progress Minggu Lalu (%)'], 
                    row['Progress Minggu Ini (%)'], 
                    (row['Progress Minggu Ini (%)'] - row['Progress Minggu Lalu (%)']),
                    row['Catatan Pekerjaan'], 
                    row['real_id']
                ))
                
            conn.commit()
            st.success("✅ Perubahan berhasil disimpan!")
            st.rerun()

    st.markdown("---")
    st.subheader("📥 Export & Download Laporan")

    buffer = io.BytesIO()
    try:
        query_excel = """
            SELECT 
                l.waktu_input AS [Waktu Input],
                l.jenis_pekerjaan AS [Jenis Pekerjaan],
                l.no_spk AS [Nomor SPK],
                l.kontraktor AS [Nama Kontraktor],
                l.unit AS [Unit Proyek],
                l.nilai_kontrak AS [Nilai Kontrak],
                l.progress_minggu_lalu AS [Progress Minggu Lalu (%)],
                l.progress_minggu_ini AS [Progress Minggu Ini (%)],
                l.selisih_progress AS [Selisih / Varian (%)],
                l.catatan AS [Catatan Pekerjaan],
                d.nama_file AS [Dokumentasi Foto]
            FROM laporan_mingguan l
            LEFT JOIN dokumentasi_foto d ON l.unit = d.unit
            ORDER BY l.id DESC
        """
        df_excel = pd.read_sql_query(query_excel, conn)

        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_excel.to_excel(writer, index=False, sheet_name='Laporan Progress')
            
            workbook = writer.book
            worksheet = writer.sheets['Laporan Progress']
            
            thin_border = Border(
                left=Side(style='thin', color='000000'),
                right=Side(style='thin', color='000000'),
                top=Side(style='thin', color='000000'),
                bottom=Side(style='thin', color='000000')
            )
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            
            for col_num in range(1, len(df_excel.columns) + 1):
                cell = worksheet.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for row_idx, row in enumerate(df_excel.itertuples(), start=2):
                for col_idx in range(1, len(df_excel.columns) + 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="center")
                    
                    # Format Rupiah pada Kolom Nilai Kontrak (Kolom Ke-6)
                    if col_idx == 6:
                        cell.number_format = '#,##0'

                nama_foto = getattr(row, 'Dokumentasi_Foto', getattr(row, 'nama_file', None))
                if isinstance(nama_foto, str) and nama_foto.strip() and os.path.exists(nama_foto):
                    img = Image(nama_foto)
                    img.width = 80
                    img.height = 80
                    
                    worksheet.row_dimensions[row_idx].height = 65
                    col_letter = get_column_letter(len(df_excel.columns))
                    worksheet.cell(row=row_idx, column=len(df_excel.columns)).value = ""
                    worksheet.add_image(img, f"{col_letter}{row_idx}")

            for col in worksheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                worksheet.column_dimensions[col_letter].width = max(max_len + 4, 18)

        st.download_button(
            label="📥 Download Laporan (Excel + Foto)",
            data=buffer.getvalue(),
            file_name="Laporan_Progress_Lengkap.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.warning(f"Gagal memproses gambar ke Excel: {e}")
            
    else:
        st.info("Belum ada riwayat laporan tersimpan di dalam database.")
        
    conn.close()

elif menu == "Input Progress Mingguan":
    st.header("📝 Form Input Progress Pekerjaan")

    with st.form(key="form_progress"):
        col_spk1, col_spk2 = st.columns(2)
        with col_spk1:
            jenis_pekerjaan = st.text_input("Jenis Pekerjaan")
        with col_spk2:
            no_spk = st.text_input("Nomor SPK")

        col_k1, col_k2, col_k3 = st.columns(3)
        with col_k1:
            kontraktor = st.text_input("Nama Kontraktor")
        with col_k2:
            unit = st.text_input("Unit Proyek")
        with col_k3:
            nilai_kontrak = st.number_input("Nilai Kontrak (Rp)", min_value=0.0, step=100000.0)
        
        col1, col2 = st.columns(2)
        with col1:
            progress_lalu = st.number_input("Progress Minggu Lalu (%)", min_value=0.0, max_value=100.0, step=0.1)
        with col2:
            progress_ini = st.number_input("Progress Minggu Ini (%)", min_value=0.0, max_value=100.0, step=0.1)
            
        selisih = progress_ini - progress_lalu
        st.info(f"📊 **Selisih / Varian Progress:** {selisih:+.2f}%")
        
        catatan = st.text_area("Catatan Pekerjaan")
        submit_button = st.form_submit_button(label="Simpan Data")

    if submit_button:
        if kontraktor and unit:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO laporan_mingguan (jenis_pekerjaan, no_spk, kontraktor, unit, nilai_kontrak, progress_minggu_lalu, progress_minggu_ini, selisih_progress, catatan)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (jenis_pekerjaan, no_spk, kontraktor, unit, nilai_kontrak, progress_lalu, progress_ini, selisih, catatan))
            conn.commit()
            conn.close()
            st.success("✅ Data progress berhasil disimpan!")
        else:
            st.error("⚠️ Mohon isi Nama Kontraktor dan Unit Proyek.")

elif menu == "Upload Dokumentasi Foto":
    st.subheader("📷 Upload Dokumentasi Foto")
    tipe_dok = st.selectbox("Tipe Dokumentasi", ["Dokumentasi BINE", "Dokumentasi BPRE", "Dokumentasi BPTE", "Dokumentasi BMSE", "Dokumentasi BPAE", "Dokumentasi LWUE", "Dokumentasi LWSE", "Dokumentasi BPRM", "Dokumentasi LWSM"])
    unit_foto = st.text_input("Unit / Lokasi Proyek")
    uploaded_files = st.file_uploader("Pilih Foto Pekerjaan", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    
    if uploaded_files:
        cols = st.columns(3)
        for idx, file in enumerate(uploaded_files):
            with cols[idx % 3]:
                st.image(file, caption=file.name, use_container_width=True)
        
        if st.button("Simpan Semua Foto"):
            conn = get_connection()
            cursor = conn.cursor()
            for file in uploaded_files:
                cursor.execute(
                    "INSERT INTO dokumentasi_foto (tipe_dok, unit, nama_file) VALUES (?, ?, ?)",
                    (tipe_dok, unit_foto, file.name)
                )
            conn.commit()
            conn.close()
            st.success("Foto berhasil tersimpan!")