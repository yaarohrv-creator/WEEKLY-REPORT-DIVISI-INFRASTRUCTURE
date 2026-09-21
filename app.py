import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image as OpenPyXLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import os
from PIL import Image as PILImage
from pathlib import Path

# ==========================================
# CONFIG & PAGE SETUP
# ==========================================
st.set_page_config(
    page_title="Sistem Laporan Progress Pekerjaan",
    page_icon="📊",
    layout="wide"
)

# Direktori utama foto
BASE_UPLOAD_DIR = "uploads_dokumentasi"
if not os.path.exists(BASE_UPLOAD_DIR):
    os.makedirs(BASE_UPLOAD_DIR)

# ==========================================
# DATABASE SIMULATION (SESSION STATE)
# ==========================================
if "data_progress" not in st.session_state:
    # Buat dummy data dengan path yang terstruktur berdasarkan Nama Pekerjaan
    job_name_1 = "RENOVASI ATAP RUMAH G2 05"
    
    # Path folder khusus pekerjaan ini
    job_folder_1 = f"{BASE_UPLOAD_DIR}/{job_name_1.replace(' ', '_')}"
    if not os.path.exists(job_folder_1):
        os.makedirs(job_folder_1)

    # Path file foto (asumsi file dummy .jpg ada di folder tersebut untuk tes)
    dummy_f1 = f"{job_folder_1}/Foto1_Sample.jpg"
    dummy_f2 = f"{job_folder_1}/Foto2_Sample.jpg"
    
    # Buat file placeholder kosong jika belum ada (untuk simulasi)
    if not os.path.exists(dummy_f1): Path(dummy_f1).touch()
    if not os.path.exists(dummy_f2): Path(dummy_f2).touch()

    st.session_state["data_progress"] = pd.DataFrame([
        {
            "real_id": 1,
            "No": 1,
            "Nomor SPK": "001/SPK/INF/2026",
            "Nama Pekerjaan": job_name_1,
            "Target (%)": 100,
            "Realisasi (%)": 85,
            "Link Path Foto 1": dummy_f1,
            "Link Path Foto 2": dummy_f2,
        }
    ])

# ==========================================
# FUNGSI EXPORT EXCEL (2 SHEET + HYPERLINK TO IMAGE SHEET)
# ==========================================
def generate_excel_with_image_links(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # ---------------------------------------------------------
        # Persiapan Data
        # ---------------------------------------------------------
        df_excel = df.drop(columns=['real_id'], errors='ignore').copy()
        
        # ---------------------------------------------------------
        # SHEET 2: FOTO DOKUMENTASI (LAYOUT SEPERTI CONTOH USER)
        # ---------------------------------------------------------
        # Buat sheet kosong
        workbook = writer.book
        worksheet2 = workbook.create_sheet(title='Foto Dokumentasi')
        
        # Style Judul
        title_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        title_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        center_align = Alignment(horizontal="center", vertical="center")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # Set Lebar Kolom (A & B untuk Foto)
        worksheet2.column_dimensions['A'].width = 60
        worksheet2.column_dimensions['B'].width = 60

        current_row = 1
        
        # Dictionary untuk menyimpan lokasi (Cell ID) dari Foto 1 setiap baris data
        # Key: Nomor SPK, Value: Alamat Cell Judul Foto 1 di Sheet 2
        photo_locations_map = {}

        # Loop setiap baris data untuk mengisi Sheet Foto Dokumentasi
        for index, row in df_excel.iterrows():
            job_title = row.get('Nama Pekerjaan', f'Pekerjaan {index+1}')
            spk_id = row.get('Nomor SPK', f'SPK_{index}')
            
            p1_path = row.get('Link Path Foto 1', '')
            p2_path = row.get('Link Path Foto 2', '')

            # Simpan lokasi awal pekerjaan ini di Sheet Foto
            # Kita gunakan Cell A di baris judul sebagai target link
            photo_locations_map[spk_id] = f"A{current_row}"

            # 1. Baris Judul (Merged Cells A-B)
            worksheet2.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
            title_cell = worksheet2.cell(row=current_row, column=1, value=job_title)
            title_cell.fill = title_fill
            title_cell.font = title_font
            title_cell.alignment = center_align
            title_cell.border = border
            
            current_row += 1

            # 2. Baris Foto (Penyisipan Gambar Fisik)
            # Set Tinggi Baris agar foto terlihat
            worksheet2.row_dimensions[current_row].height = 300
            
            col_index = 1
            for p_path in [p1_path, p2_path]:
                cell = worksheet2.cell(row=current_row, column=col_index)
                cell.border = border # Beri border kosong
                
                if p_path and os.path.exists(str(p_path)):
                    try:
                        # Load & Resize Gambar
                        pil_img = PILImage.open(p_path)
                        original_width, original_height = pil_img.size
                        
                        # Resize proporsional
                        target_height = 390 # Sedikit kurang dari tinggi baris Excel
                        ratio = target_height / original_height
                        target_width = int(original_width * ratio)
                        
                        # Batasi lebar maksimal agar tidak overlap
                        if target_width > 420:
                             target_width = 420
                             ratio = target_width / original_width
                             target_height = int(original_height * ratio)

                        pil_img = pil_img.resize((target_width, target_height), PILImage.Resampling.LANCZOS)
                        
                        # Ubah ke format OpenPyXL Image
                        img_byte_arr = io.BytesIO()
                        pil_img.save(img_byte_arr, format=pil_img.format if pil_img.format else 'JPEG')
                        img_byte_arr.seek(0)
                        
                        opx_img = OpenPyXLImage(img_byte_arr)
                        
                        # Sisipkan ke Excel
                        col_letter = get_column_letter(col_index)
                        worksheet2.add_image(opx_img, f"{col_letter}{current_row}")
                        
                    except Exception as e:
                        cell.value = f"Gagal memuat foto"
                        cell.alignment = center_align
                else:
                    cell.value = "Foto Tidak Tersedia"
                    cell.alignment = center_align
                    
                col_index += 1
                
            current_row += 2 # Beri jarak antar pekerjaan

        # ---------------------------------------------------------
        # SHEET 1: LAPORAN PROGRESS (DENGAN INTERNAL HYPERLINK)
        # ---------------------------------------------------------
        # Ganti teks path yang panjang menjadi "Lihat Foto" di DataFrame copy
        df_progress = df_excel.copy()
        
        # Menulis data utama (Header & Teks) ke sheet 1
        df_progress.to_excel(writer, index=False, sheet_name='Laporan Progress')
        worksheet1 = writer.sheets['Laporan Progress']

        # Styling Sheet 1
        header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        header_font = Font(bold=True)
        link_font = Font(color="0563C1", underline="single") # Biru Standar Hyperlink
        
        # Indeks Kolom Foto (1-based)
        col_f1_idx = df_progress.columns.get_loc("Link Path Foto 1") + 1
        col_f2_idx = df_progress.columns.get_loc("Link Path Foto 2") + 1
        col_spk_idx = df_progress.columns.get_loc("Nomor SPK") + 1

        # Format Data & Pembuatan Hyperlink Internal
        for row_idx in range(1, worksheet1.max_row + 1):
            for col_idx in range(1, worksheet1.max_column + 1):
                cell = worksheet1.cell(row=row_idx, column=col_idx)
                cell.border = border
                
                # Format Header
                if row_idx == 1:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", wrap_text=True)
                    continue

                # Ambil Nomor SPK baris ini sebagai kunci map
                current_spk_id = worksheet1.cell(row=row_idx, column=col_spk_idx).value
                target_cell_address = photo_locations_map.get(current_spk_id)

                # Cek jika kolom adalah Foto 1 atau Foto 2
                if col_idx == col_f1_idx or col_idx == col_f2_idx:
                    path_val = cell.value
                    
                    if path_val and path_val != "":
                        # 1. Ubah teks isi sel menjadi pendek (Tipe data jadi teks biasa)
                        foto_num = "1" if col_idx == col_f1_idx else "2"
                        cell.value = f"Lihat Foto {foto_num}"
                        
                        # 2. Membuat Tautan/Hyperlink INTERNAL ke Sheet 2
                        if target_cell_address:
                            # Formula internal link: #SheetName!CellAddress
                            # Teks akan berubah warna dan bergaris bawah otomatis di Excel
                            cell.hyperlink = f"#'Foto Dokumentasi'!{target_cell_address}"
                            cell.font = link_font
                            cell.alignment = center_align
                    else:
                        cell.value = "-"
                        cell.alignment = center_align

        # Mengatur Lebar Kolom Otomatis di Sheet 1
        for col in worksheet1.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            worksheet1.column_dimensions[col_letter].width = max(max_len + 3, 15)

    return output.getvalue()


# ==========================================
# MAIN APP TAMPILAN
# ==========================================
st.title("📊 Laporan Progress & Dokumentasi Foto Interaktif")
st.markdown("---")

df_current = st.session_state["data_progress"]

if df_current.empty:
    st.info("Belum ada data progress.")
else:
    st.subheader("📋 Ringkasan Laporan")
    st.dataframe(
        df_current.drop(columns=["real_id"], errors="ignore"),
        use_container_width=True
    )
    
    st.markdown("---")
    st.subheader("📥 Export ke Excel")
    st.info("""File Excel akan berisi 2 Sheet:
1. **Laporan Progress:** Kolom foto berisi link pendek biru yang jika diklik akan pindah ke Sheet 2.
2. **Foto Dokumentasi:** Layout khusus berisi foto fisik yang tersusun rapi per pekerjaan (sesuai contoh).""")
    
    # Tombol Download
    if st.download_button(
        label="💾 Download File Excel (Interaktif)",
        data=generate_excel_with_image_links(df_current),
        file_name="Laporan_Progress_dan_Foto_Dokumentasi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ):
        st.success("File berhasil dibuat. Pastikan Anda memiliki foto fisik di direktori 'uploads_dokumentasi' agar gambar muncul di Excel.")
