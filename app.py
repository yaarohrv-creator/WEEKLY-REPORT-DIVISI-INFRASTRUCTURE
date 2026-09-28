import os
import io
import re
import pandas as pd
import streamlit as st
from sqlalchemy import create_engine
import cloudinary
import cloudinary.uploader

# Modul untuk styling dan export Excel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =========================================================
# KONFIGURASI HALAMAN STREAMLIT
# =========================================================
st.set_page_config(
    page_title="Laporan Mingguan & Monitoring SPK",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# KONFIGURASI DATABASE POSTGRESQL (NEON) & CLOUDINARY
# =========================================================
# Mengambil credentials dari st.secrets (.streamlit/secrets.toml)
db_url = st.secrets["DB_URL"]
engine = create_engine(db_url)

cloudinary.config(
    cloud_name=st.secrets["CLOUD_NAME"],
    api_key=st.secrets["API_KEY"],
    api_secret=st.secrets["API_SECRET"],
    secure=True
)

def get_db_connection():
    return engine.connect()

# =========================================================
# HELPER FUNCTIONS
# =========================================================
def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "", str(filename))

def load_data_spk():
    conn = get_db_connection()
    try:
        df = pd.read_sql("SELECT * FROM data_spk ORDER BY no_spk ASC", conn)
    finally:
        conn.close()
    return df

def load_data_laporan():
    conn = get_db_connection()
    try:
        df = pd.read_sql("SELECT * FROM laporan_mingguan ORDER BY id DESC", conn)
    finally:
        conn.close()
    return df

# =========================================================
# FORM INPUT PROGRESS MINGGUAN
# =========================================================
def render_input_form():
    st.header("📝 Input Progress / Laporan Mingguan")
    
    df_spk = load_data_spk()
    if df_spk.empty:
        st.warning("Data SPK masih kosong. Silakan lengkapi Data SPK terlebih dahulu.")
        return

    spk_list = df_spk["no_spk"].tolist()
    
    with st.form("form_laporan", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            selected_spk = st.selectbox("Pilih No SPK", spk_list)
            minggu_ke = st.number_input("Minggu Ke-", min_value=1, max_value=100, step=1)
            tgl_laporan = st.date_input("Tanggal Laporan")
            progress_real = st.number_input("Progress Realisasi (%)", min_value=0.0, max_value=100.0, step=0.1)
            progress_rencana = st.number_input("Progress Rencana (%)", min_value=0.0, max_value=100.0, step=0.1)
        
        with col2:
            uraian_pekerjaan = st.text_area("Uraian Pekerjaan / Catatan Progress")
            f_upload_1 = st.file_uploader("Upload Foto Dokumentasi 1", type=["jpg", "jpeg", "png"])
            f_upload_2 = st.file_uploader("Upload Foto Dokumentasi 2", type=["jpg", "jpeg", "png"])

        submitted = st.form_submit_button("Simpan Laporan")

    if submitted:
        url_f1 = None
        url_f2 = None

        # Upload Gambar ke Cloudinary
        if f_upload_1:
            with st.spinner("Uploading Foto 1 ke Cloudinary..."):
                res1 = cloudinary.uploader.upload(f_upload_1, folder="weekly_reports")
                url_f1 = res1.get("secure_url")

        if f_upload_2:
            with st.spinner("Uploading Foto 2 ke Cloudinary..."):
                res2 = cloudinary.uploader.upload(f_upload_2, folder="weekly_reports")
                url_f2 = res2.get("secure_url")

        # Simpan ke PostgreSQL Neon
        conn = get_db_connection()
        query = """
            INSERT INTO laporan_mingguan 
            (no_spk, minggu_ke, tgl_laporan, progress_real, progress_rencana, uraian_pekerjaan, foto_1, foto_2)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        """
        try:
            conn.exec_driver_sql(query, (
                selected_spk, int(minggu_ke), str(tgl_laporan), 
                float(progress_real), float(progress_rencana), 
                uraian_pekerjaan, url_f1, url_f2
            ))
            conn.commit()
            st.success("Data laporan berhasil disimpan ke PostgreSQL!")
        except Exception as e:
            st.error(f"Gagal menyimpan data: {e}")
        finally:
            conn.close()

# =========================================================
# MONITORING & RIWAYAT DATA
# =========================================================
def render_monitoring():
    st.header("📊 Riwayat & Monitoring Progress")
    
    df_laporan = load_data_laporan()
    if df_laporan.empty:
        st.info("Belum ada data laporan yang diinput.")
        return

    st.dataframe(df_laporan, use_container_width=True)

    # Opsi Hapus Data Laporan
    st.subheader("🗑️ Hapus Laporan")
    report_ids = df_laporan["id"].tolist()
    id_to_delete = st.selectbox("Pilih ID Laporan yang ingin dihapus", report_ids)
    
    if st.button("Hapus Laporan"):
        conn = get_db_connection()
        try:
            conn.exec_driver_sql("DELETE FROM laporan_mingguan WHERE id = %s", (int(id_to_delete),))
            conn.commit()
            st.success(f"Laporan ID {id_to_delete} berhasil dihapus.")
            st.rerun()
        except Exception as e:
            st.error(f"Gagal menghapus laporan: {e}")
        finally:
            conn.close()

# =========================================================
# MAIN APP NAVIGATION
# =========================================================
def main():
    st.sidebar.title("Navigasi Sistem")
    menu = st.sidebar.radio("Pilih Menu", ["Input Progress", "Monitoring & Data"])

    if menu == "Input Progress":
        render_input_form()
    elif menu == "Monitoring & Data":
        render_monitoring()

if __name__ == "__main__":
    main()
