import io
import streamlit as st
import pandas as pd
import numpy as np
from parser_engine import (
    parse_bank_kalsel,
    generate_form_pdf,
    generate_form_excel,
    sort_resume_chronological,
    MONTH_NAMES_ID
)

# ==============================================================================
# 1. KONFIGURASI HALAMAN & TEMA LIGHT MODE
# ==============================================================================
st.set_page_config(
    page_title="AREKO Bank Kalsel - Let’s Make Recap Less Recap-y",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inisialisasi Session State
if "form_ver" not in st.session_state:
    st.session_state["form_ver"] = 0
if "resume_list" not in st.session_state:
    st.session_state["resume_list"] = None

def reset_all():
    st.session_state["resume_list"] = None
    st.session_state["form_ver"] += 1

# Custom Light Theme CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    html, body, [class*="css"], [data-testid="stAppViewContainer"], [data-testid="stApp"] {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        background-color: #ffffff !important;
        color: #1e293b !important;
    }

    [data-testid="stSidebar"] {
        background-color: #f8fafc !important;
        border-right: 1px solid #e2e8f0;
    }

    .sidebar-section-title {
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #64748b;
        margin: 16px 0 8px 0;
        padding-bottom: 4px;
        border-bottom: 1px solid #e2e8f0;
    }

    /* Empty State Box */
    .empty-state-box {
        background-color: #ffffff;
        border: 1px dashed #cbd5e1;
        border-radius: 12px;
        padding: 50px 24px;
        text-align: center !important;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        margin-top: 20px;
        width: 100%;
    }
    .empty-state-icon {
        font-size: 2.8rem;
        margin-bottom: 12px;
        text-align: center !important;
    }
    .empty-state-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 8px;
        text-align: center !important;
    }
    .empty-state-desc {
        font-size: 0.88rem;
        color: #64748b;
        max-width: 540px;
        margin: 0 auto !important;
        line-height: 1.6;
        text-align: center !important;
    }

    /* Download Cards */
    .dl-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 18px 20px;
        margin-bottom: 12px;
        min-height: 90px;
    }
    .dl-card-title {
        font-size: 15px;
        font-weight: 600;
        color: #0f172a;
        margin-bottom: 4px;
    }
    .dl-card-desc {
        font-size: 12.5px;
        color: #64748b;
        line-height: 1.4;
    }

    /* Clean Buttons */
    div.stDownloadButton > button:first-child {
        border-radius: 8px !important;
        font-weight: 500 !important;
        border: 1px solid #cbd5e1 !important;
        background-color: #ffffff !important;
        color: #1e293b !important;
        width: 100% !important;
        padding: 10px 16px !important;
        transition: all 0.15s ease-in-out !important;
    }
    div.stDownloadButton > button:first-child:hover {
        background-color: #f8fafc !important;
        border-color: #94a3b8 !important;
        color: #0f172a !important;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. SIDEBAR - FORM INPUT DATA NASABAH
# ==============================================================================
ver = st.session_state.get("form_ver", 0)

with st.sidebar:
    st.markdown("### ⚙️ Parameter Dokumen")

    st.markdown('<div class="sidebar-section-title">Informasi Nasabah</div>', unsafe_allow_html=True)
    cabang = st.text_input("Cabang", key=f"cabang_{ver}", placeholder="Contoh: Banjarmasin")
    nama_cust = st.text_input("Nama Cust", key=f"cust_{ver}", placeholder="Nama lengkap customer")
    no_rek = st.text_input("Nomor Rekening", key=f"rek_{ver}", placeholder="Contoh: 7300157275")
    nama_bank = st.text_input("Nama Bank", value="BANK KALSEL", disabled=True, key=f"bank_{ver}")

    pemegang_rek = st.text_input("Nama Pemegang Rekening", key=f"pemegang_{ver}", placeholder="Sesuai buku tabungan")

    st.markdown('<div class="sidebar-section-title">Pengesahan & Catatan</div>', unsafe_allow_html=True)
    nama_so = st.text_input("Nama SO (Sales Officer)", key=f"so_{ver}", placeholder="Nama Sales Officer")
    nama_oh = st.text_input("Operation Head", key=f"oh_{ver}", placeholder="Nama Operation Head")
    note_oh = st.text_area("Note OH", value="-", key=f"note_{ver}", placeholder="Catatan dari Operation Head...")

    st.markdown("---")
    st.button("🔄 Reset Form & Data", on_click=reset_all, use_container_width=True)

    header_input = {
        "cabang": cabang,
        "nama_cust": nama_cust,
        "no_rekening": no_rek,
        "nama_bank": nama_bank,
        "nama_pemegang_rek": pemegang_rek,
        "nama_so": nama_so,
        "nama_oh": nama_oh,
    }

# ==============================================================================
# 3. KONTEN UTAMA: HEADER & FILE UPLOADER
# ==============================================================================
st.markdown("""
<div style="margin-bottom: 24px;">
    <h1 style="font-size: 2.2rem; font-weight: 800; color: #0f172a; margin: 0; letter-spacing: -0.02em;">AREKO - Bank Kalsel</h1>
    <p style="font-size: 1.05rem; font-weight: 500; color: #64748b; margin-top: 4px; margin-bottom: 0;">Let’s Make Recap Less Recap-y • Bank Kalsel</p>
</div>
""", unsafe_allow_html=True)

st.markdown("#### 📂 Berkas Rekening Koran Bank Kalsel")
uploaded_files = st.file_uploader(
    "Pilih 1 hingga 3 Berkas Rekening Koran Bank Kalsel (Format PDF)",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"uploader_{ver}",
    help="Unggah dokumen PDF mutasi rekening Bank Kalsel untuk dianalisis dan direkapitulasi secara otomatis."
)

if uploaded_files:
    col_btn, col_rst = st.columns([2, 1])
    with col_btn:
        proses_clicked = st.button("🚀 Ekstraksi & Rekap Data", type="primary", use_container_width=True)
    with col_rst:
        st.button("🔄 Reset", on_click=reset_all, use_container_width=True)

    if proses_clicked:
        resume_list = []
        progress_bar = st.progress(0, text="Memulai ekstraksi...")

        for i, file_obj in enumerate(uploaded_files):
            progress_bar.progress(
                (i + 1) / len(uploaded_files),
                text=f"Menganalisis & mengekstrak berkas: {file_obj.name}...",
            )
            extracted = parse_bank_kalsel(file_obj, filename=file_obj.name)
            resume_list.append(extracted)

        st.session_state["resume_list"] = resume_list
        progress_bar.empty()
        st.toast("Ekstraksi data mutasi Bank Kalsel berhasil selesai!", icon="✅")
        st.rerun()

# ==============================================================================
# 4. TABEL & UNDUH DOKUMEN (HANYA MUNCUL JIKA SUDAH ADA DATA / EKSTRAKSI)
# ==============================================================================
if st.session_state.get("resume_list"):
    res_list = st.session_state["resume_list"]
    sorted_data = sort_resume_chronological(res_list)

    st.markdown("<br>", unsafe_allow_html=True)

    preview_rows = []
    for r in sorted_data:
        preview_rows.append({
            "Bulan": r.get("bulan", "-"),
            "Bank": "BANK KALSEL",
            "Freq Debet": r.get("freq_db", 0),
            "Freq Kredit": r.get("freq_cr", 0),
            "Total Debet (Rp)": f"{r.get('mutasi_db', 0):,.2f}",
            "Total Kredit (Rp)": f"{r.get('mutasi_cr', 0):,.2f}",
            "Saldo Tertinggi (Rp)": f"{r.get('saldo_max', 0):,.2f}",
            "Saldo Rata-Rata (Rp)": f"{r.get('saldo_avg', 0):,.2f}",
            "Saldo Terendah (Rp)": f"{r.get('saldo_min', 0):,.2f}",
        })

    n = len(sorted_data)
    if n > 0:
        avg_f_db = int(round(sum(r.get("freq_db", 0) for r in sorted_data) / n))
        avg_f_cr = int(round(sum(r.get("freq_cr", 0) for r in sorted_data) / n))
        avg_m_db = sum(r.get("mutasi_db", 0.0) for r in sorted_data) / n
        avg_m_cr = sum(r.get("mutasi_cr", 0.0) for r in sorted_data) / n
        avg_s_max = sum(r.get("saldo_max", 0.0) for r in sorted_data) / n
        avg_s_avg = sum(r.get("saldo_avg", 0.0) for r in sorted_data) / n
        avg_s_min = sum(r.get("saldo_min", 0.0) for r in sorted_data) / n

        # Baris Rata-Rata di akhir tabel
        preview_rows.append({
            "Bulan": "Rata-Rata",
            "Bank": "BANK KALSEL",
            "Freq Debet": avg_f_db,
            "Freq Kredit": avg_f_cr,
            "Total Debet (Rp)": f"{avg_m_db:,.2f}",
            "Total Kredit (Rp)": f"{avg_m_cr:,.2f}",
            "Saldo Tertinggi (Rp)": f"{avg_s_max:,.2f}",
            "Saldo Rata-Rata (Rp)": f"{avg_s_avg:,.2f}",
            "Saldo Terendah (Rp)": f"{avg_s_min:,.2f}",
        })

    # 1. Tabel Native Streamlit dengan Highlight Warna pada Baris Rata-Rata
    df_preview = pd.DataFrame(preview_rows)

    def highlight_rata_rata(row):
        if str(row.get("Bulan", "")).strip().lower() == "rata-rata":
            return ["background-color: #E2EFDA; color: #0f5132; font-weight: 700;"] * len(row)
        return [""] * len(row)

    styled_df = df_preview.style.apply(highlight_rata_rata, axis=1)

    st.markdown("#### 📋 Rekapitulasi Mutasi Rekening Bank Kalsel")
    st.dataframe(styled_df, use_container_width=True, hide_index=True)

    # 2. Section Unduh Hasil
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### 📥 Unduh Dokumen Form Validasi Bank Kalsel")
    col_dl1, col_dl2 = st.columns(2)

    # Generate Buffer File
    pdf_bytes = generate_form_pdf(
        None,
        header_input,
        res_list,
        note_oh=note_oh,
        nama_so=nama_so,
        nama_oh=nama_oh,
    )

    xlsx_bytes = generate_form_excel(
        None,
        header_input,
        res_list,
        note_oh=note_oh,
        nama_so=nama_so,
        nama_oh=nama_oh,
    )

    with col_dl1:
        st.markdown(
            """
            <div class="dl-card">
                <div class="dl-card-title">📄 Dokumen PDF Resmi Bank Kalsel</div>
                <div class="dl-card-desc">Format standar A4 siap cetak dengan tanda tangan SO & Operation Head.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.download_button(
            label="Unduh PDF Form Validasi Bank Kalsel",
            data=pdf_bytes,
            file_name="FORM_VALIDASI_MUTASI_REKENING_BANK_KALSEL.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with col_dl2:
        st.markdown(
            """
            <div class="dl-card">
                <div class="dl-card-title">📊 Dokumen Excel Spreadsheet Bank Kalsel</div>
                <div class="dl-card-desc">Format lembar kerja dinamis lengkap dengan rincian transaksi per bulan.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.download_button(
            label="Unduh Excel Form Validasi Bank Kalsel (.xlsx)",
            data=xlsx_bytes,
            file_name="FORM_VALIDASI_MUTASI_REKENING_BANK_KALSEL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

else:
    # KONDISI AWAL (EMPTY STATE)
    st.markdown("""
    <div class="empty-state-box" style="text-align: center; display: flex; flex-direction: column; align-items: center; justify-content: center;">
        <div class="empty-state-icon" style="text-align: center;">📄</div>
        <div class="empty-state-title" style="text-align: center;">Belum Ada Rekening Koran Bank Kalsel yang Diproses</div>
        <p class="empty-state-desc" style="text-align: center; margin: 0 auto;">
            Silakan unggah berkas PDF rekening koran Bank Kalsel di atas, lengkapi parameter dokumen di sidebar, lalu klik <b>Ekstraksi & Rekap Data</b> untuk menampilkan tabel ringkasan mutasi dan mengunduh formulir validasi.
        </p>
    </div>
    """, unsafe_allow_html=True)
