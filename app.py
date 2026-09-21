import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import os

# ==========================================
# CONFIG & PAGE SETUP
# ==========================================
st.set_page_config(
    page_title="Sistem Laporan Progress Pekerjaan",
    page_icon="📊",
    layout="wide"
)

# Folder penyimpanan foto dokumentasi
UPLOAD_DIR = "uploads_dokumentasi"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# ==========================================
# INISIALISASI SESSION STATE / DATASTORE
# ==========================================
if "data_progress" not in st.session_state:
    # Sample data awal
    st.session_state["data_progress"] = pd.DataFrame([
        {
            "real_id": 1,
            "No": 1,
            "Nomor SPK": "052/PSM/2BPRE/BPS/LIJKTO/INF/VII/2025",
            "Nama Kontraktor": "PT Butun Bintana BPRE",
            "Jenis Pekerjaan": "RENOVASI ATAP RUMAH G1 NO 02 TAHUN 1999 BUTUN",
            "Unit Proyek": "Unit 1",
            "Target Progress (%)": 100,
            "Realizasi Progress (%)": 100,
            "Selisih / Varian (%)": 0,
            "Link Path Foto 1": f"{UPLOAD_DIR}/RENOVASI_ATAP_RUMAH_G1_NO_02_TAHUN_1999_BUTUN_BPRE_Foto1.jpg",
            "Link Path Foto 2": f"{UPLOAD_DIR}/RENOVASI_ATAP_RUMAH_G1_NO_02_TAHUN_1999_BUTUN_BPRE_Foto2.jpg",
            "Catatan Pekerjaan Terbaru": "Pekerjaan konstruksi atap telah selesai 100%."
        }
    ])

# ==========================================
# FUNGSI GENERATE EXCEL DENGAN 2 SHEET
# ==========================================
def generate_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # ---------------------------------------------------------
        # SHEET 1: LAPORAN PROGRESS
        # ---------------------------------------------------------
        df_excel = df.drop(columns=['real_id'], errors='ignore').copy()
        df_progress = df_excel.copy()
        
        df_progress.to_excel(writer, index=False, sheet_name='Laporan Progress')
        worksheet1 = writer.sheets['Laporan Progress']
        
        # ---------------------------------------------------------
        # SHEET 2: FOTO DOKUMENTASI
        # ---------------------------------------------------------
        cols_foto = [
            'No', 'Nomor SPK', 'Nama Kontraktor', 'Jenis Pekerjaan', 
            'Unit Proyek', 'Link Path Foto 1', 'Link Path Foto 2', 'Catatan Pekerjaan Terbaru'
        ]
        cols_exist = [c for c in cols_foto if c in df_excel.columns]
        df_foto = df_excel[cols_exist].copy()
        
        df_foto.to_excel(writer, index=False, sheet_name='Foto Dokumentasi')
        worksheet2 = writer.sheets['Foto Dokumentasi']

        # ---------------------------------------------------------
        # STYLING & FORMATTING DENGAN OPENPYXL
        # ---------------------------------------------------------
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        link_font = Font(name="Calibri", size=11, color="0563C1", underline="single")
        
        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )

        for sheet in [worksheet1, worksheet2]:
            # Format Header
            for col_num in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin_border

            # Format Data & Pemendekan Hyperlink Foto
            for row_idx in range(2, sheet.max_row + 1):
                for col_idx in range(1, sheet.max_column + 1):
                    cell = sheet.cell(row=row_idx, column=col_idx)
                    cell.border = thin_border
                    cell.alignment = Alignment(vertical="center")
                    
                    header_name = str(sheet.cell(row=1, column=col_idx).value)
                    
                    # Konversi Teks Path Foto yang Panjang menjadi Link Ringkas
                    if "Link Path Foto" in header_name and cell.value:
                        path_str = str(cell.value).strip()
                        if path_str and path_str.lower() != 'none':
                            foto_num = "1" if "1" in header_name else "2"
                            # Mengubah teks isi sel menjadi pendek
                            cell.value = f"Lihat Foto {foto_num}"
                            
                            # Membuat Tautan/Hyperlink
                            cell.hyperlink = path_str
                            cell.font = link_font
                            cell.alignment = Alignment(horizontal="center", vertical="center")

            # Mengatur Lebar Kolom Otomatis
            for col in sheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                sheet.column_dimensions[col_letter].width = max(max_len + 3, 15)

    return output.getvalue()


# ==========================================
# TAMPILAN UTAMA STREAMLIT
# ==========================================
st.title("📊 Laporan Progress & Dokumentasi Pekerjaan")
st.markdown("---")

# Sidebar Menu Navigation
menu = st.sidebar.selectbox("Navigasi Menu", ["Lihat Data & Download", "Tambah Progress Baru"])

# ------------------------------------------
# MENU 1: LIHAT DATA & DOWNLOAD EXCEL
# ------------------------------------------
if menu == "Lihat Data & Download":
    st.subheader("📋 Ringkasan Data Progress Pekerjaan")
    
    df_current = st.session_state["data_progress"]
    
    if df_current.empty:
        st.warning("Belum ada data progress pekerjaan.")
    else:
        # Tampilkan Data di Streamlit
        st.dataframe(
            df_current.drop(columns=["real_id"], errors="ignore"),
            use_container_width=True
        )
        
        st.markdown("---")
        st.subheader("📥 Export Laporan ke Excel")
        st.info("File Excel akan berisi 2 Sheet: **Laporan Progress** dan **Foto Dokumentasi**, dengan link foto yang rapi (format ringkas).")
        
        excel_bytes = generate_excel(df_current)
        
        st.download_button(
            label="💾 Download File Excel (.xlsx)",
            data=excel_bytes,
            file_name="Laporan_Progress_dan_Foto_Dokumentasi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ------------------------------------------
# MENU 2: TAMBAH DATA PROGRESS BARU
# ------------------------------------------
elif menu == "Tambah Progress Baru":
    st.subheader("➕ Form Input Progress & Foto Dokumentasi")
    
    with st.form(key="form_progress", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            no_spk = st.text_input("Nomor SPK", value="052/PSM/2BPRE/BPS/LIJKTO/INF/VII/2025")
            kontraktor = st.text_input("Nama Kontraktor", value="PT Butun Bintana BPRE")
            jenis_pekerjaan = st.text_input("Jenis Pekerjaan", value="RENOVASI ATAP RUMAH G1 NO 02 TAHUN 1999 BUTUN")
            unit_proyek = st.text_input("Unit Proyek", value="Unit 1")
            
        with col2:
            target_prog = st.number_input("Target Progress (%)", min_value=0.0, max_value=100.0, value=100.0)
            realisasi_prog = st.number_input("Realisasi Progress (%)", min_value=0.0, max_value=100.0, value=100.0)
            catatan = st.text_area("Catatan Pekerjaan Terbaru", value="Pekerjaan selesai sesuai spesifikasi.")
            
        st.markdown("**Upload Foto Dokumentasi (Opsional):**")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            file_foto1 = st.file_uploader("Upload Foto Dokumentasi 1", type=["jpg", "jpeg", "png"])
        with col_f2:
            file_foto2 = st.file_uploader("Upload Foto Dokumentasi 2", type=["jpg", "jpeg", "png"])
            
        submit_button = st.form_submit_button(label="Simpan Data Progress")
        
    if submit_button:
        # Simpan file foto jika ada
        path_foto1 = ""
        path_foto2 = ""
        
        if file_foto1 is not None:
            filename1 = f"{UPLOAD_DIR}/{jenis_pekerjaan.replace(' ', '_')}_Foto1_{file_foto1.name}"
            with open(filename1, "wb") as f:
                f.write(file_foto1.getbuffer())
            path_foto1 = filename1
            
        if file_foto2 is not None:
            filename2 = f"{UPLOAD_DIR}/{jenis_pekerjaan.replace(' ', '_')}_Foto2_{file_foto2.name}"
            with open(filename2, "wb") as f:
                f.write(file_foto2.getbuffer())
            path_foto2 = filename2
            
        # Hitung nomor dan selisih
        df_old = st.session_state["data_progress"]
        new_no = len(df_old) + 1
        selisih = realisasi_prog - target_prog
        
        new_entry = {
            "real_id": new_no,
            "No": new_no,
            "Nomor SPK": no_spk,
            "Nama Kontraktor": kontraktor,
            "Jenis Pekerjaan": jenis_pekerjaan,
            "Unit Proyek": unit_proyek,
            "Target Progress (%)": target_prog,
            "Realizasi Progress (%)": realisasi_prog,
            "Selisih / Varian (%)": selisih,
            "Link Path Foto 1": path_foto1,
            "Link Path Foto 2": path_foto2,
            "Catatan Pekerjaan Terbaru": catatan
        }
        
        st.session_state["data_progress"] = pd.concat(
            [df_old, pd.DataFrame([new_entry])], 
            ignore_index=True
        )
        
        st.success("Data progress dan foto dokumentasi berhasil disimpan!")
