import io
import os
import re
import shutil
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageEnhance, ImageFilter
import pdfplumber
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

MONTH_NAMES_ID = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

MONTH_MAP = {
    "JAN": "Januari", "FEB": "Februari", "MAR": "Maret", "APR": "April",
    "MAY": "Mei", "MEI": "Mei", "JUN": "Juni", "JUL": "Juli",
    "AUG": "Agustus", "AGU": "Agustus", "SEP": "September", "OCT": "Oktober",
    "OKT": "Oktober", "NOV": "November", "DEC": "Desember", "DES": "Desember"
}

def get_tesseract_cmd():
    """Mencari path binary tesseract pada sistem."""
    tess_path = shutil.which("tesseract")
    if not tess_path:
        for p in ["/opt/homebrew/bin/tesseract", "/usr/local/bin/tesseract", "/usr/bin/tesseract"]:
            if os.path.exists(p):
                return p
    return tess_path

def preprocess_image(img):
    """Meningkatkan kontras dan ketajaman teks angka untuk OCR."""
    gray = img.convert("L")
    enhancer = ImageEnhance.Contrast(gray)
    contrasted = enhancer.enhance(2.0)
    return contrasted.filter(ImageFilter.SHARPEN)

def detect_kalsel_month(full_text, filename=""):
    """Mendeteksi nama bulan transaksi Bank Kalsel dari teks atau nama berkas."""
    txt = (full_text + " " + filename).upper()
    
    # Format Bank Kalsel: Periode Transaksi 01 Juli 2026 s/d 31 Juli 2026
    pm = re.search(r"Periode\s+Transaksi\s*\d{1,2}\s+([A-Za-z]+)\s+\d{4}\s*s/d\s*\d{1,2}\s+([A-Za-z]+)\s+\d{4}", full_text, re.IGNORECASE)
    if pm:
        raw_m = pm.group(1).capitalize()
        if raw_m in MONTH_NAMES_ID:
            return raw_m
        for k, v in MONTH_MAP.items():
            if k in raw_m.upper():
                return v

    for m in MONTH_NAMES_ID:
        if m.upper() in txt:
            return m

    for k, v in MONTH_MAP.items():
        if re.search(r"\b" + k + r"\b", txt):
            return v

    return "Juli"

def clean_kalsel_val(val_str):
    """Membersihkan string nominal mata uang Bank Kalsel."""
    if not val_str:
        return 0.0
    s = str(val_str)
    s = re.sub(r"[\ue088\ue092]", "", s)
    s = re.sub(r"IDR", "", s, flags=re.IGNORECASE)
    s = re.sub(r"Rp\.", "", s, flags=re.IGNORECASE)
    s = re.sub(r"Rp", "", s, flags=re.IGNORECASE)
    s = re.sub(r"[-–\s]", "", s)
    s = s.replace(",", "")
    if s:
        try:
            return float(s)
        except ValueError:
            return 0.0
    return 0.0

def convert_pdf_to_images(pdf_input, dpi=300):
    """
    Mengonversi PDF ke daftar image PIL.
    Mendukung pypdfium2 dan pdf2image sebagai fallback.
    """
    if isinstance(pdf_input, (str, os.PathLike)):
        with open(pdf_input, "rb") as f:
            pdf_bytes = f.read()
    elif isinstance(pdf_input, io.BytesIO):
        pdf_bytes = pdf_input.getvalue()
    elif hasattr(pdf_input, "read"):
        current_pos = pdf_input.tell() if hasattr(pdf_input, "tell") else 0
        pdf_bytes = pdf_input.read()
        if hasattr(pdf_input, "seek"):
            pdf_input.seek(current_pos)
    else:
        pdf_bytes = bytes(pdf_input)

    images = []
    try:
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(pdf_bytes)
        scale = dpi / 72.0
        for page in pdf:
            pil_image = page.render(scale=scale).to_pil()
            images.append(pil_image)
        if images:
            return images
    except Exception:
        pass

    try:
        from pdf2image import convert_from_bytes
        poppler_path = None
        for p in ["/opt/homebrew/bin", "/usr/local/bin", "/usr/bin"]:
            if os.path.exists(os.path.join(p, "pdftoppm")):
                poppler_path = p
                break
        images = convert_from_bytes(pdf_bytes, dpi=dpi, poppler_path=poppler_path)
        if images:
            return images
    except Exception:
        pass

    return images

def extract_ocr_text(images):
    """Ekstraksi teks dari gambar menggunakan Tesseract OCR lokal."""
    import pytesseract
    tess_cmd = get_tesseract_cmd()
    if tess_cmd:
        pytesseract.pytesseract.tesseract_cmd = tess_cmd

    lines = []
    for img in images:
        proc_img = preprocess_image(img)
        ocr_str = pytesseract.image_to_string(proc_img, config="--oem 3 --psm 6")
        for line in ocr_str.splitlines():
            line_str = line.strip()
            if line_str:
                lines.append(line_str)
    return lines

def extract_pdf_digital_text(pdf_input):
    """Ekstraksi teks vektor digital langsung dari dokumen PDF."""
    if isinstance(pdf_input, (str, os.PathLike)):
        with open(pdf_input, "rb") as f:
            pdf_bytes = f.read()
    elif isinstance(pdf_input, io.BytesIO):
        pdf_bytes = pdf_input.getvalue()
    elif hasattr(pdf_input, "read"):
        current_pos = pdf_input.tell() if hasattr(pdf_input, "tell") else 0
        pdf_bytes = pdf_input.read()
        if hasattr(pdf_input, "seek"):
            pdf_input.seek(current_pos)
    else:
        pdf_bytes = bytes(pdf_input)

    text_pages = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                txt = page.extract_text()
                if txt:
                    text_pages.append(txt)
    except Exception:
        pass

    return "\n".join(text_pages)

def parse_bank_kalsel(pdf_input, filename=""):
    """
    Parser mutasi rekening koran Bank Kalsel dengan dukungan ekstraksi tabel
    digital native serta fallback OCR presisi tinggi.
    """
    print(f"\n[+] Memproses berkas Bank Kalsel: {filename or 'PDF Upload'} ...")

    if isinstance(pdf_input, (str, os.PathLike)):
        with open(pdf_input, "rb") as f:
            pdf_bytes = f.read()
    elif isinstance(pdf_input, io.BytesIO):
        pdf_bytes = pdf_input.getvalue()
    elif hasattr(pdf_input, "read"):
        current_pos = pdf_input.tell() if hasattr(pdf_input, "tell") else 0
        pdf_bytes = pdf_input.read()
        if hasattr(pdf_input, "seek"):
            pdf_input.seek(current_pos)
    else:
        pdf_bytes = bytes(pdf_input)

    digital_full_text = extract_pdf_digital_text(pdf_bytes)
    bulan_nama = detect_kalsel_month(digital_full_text, filename)

    # Baca ringkasan resmi dari header
    m_sa = re.search(r"Saldo\s+Awal\s+Rp\.\s*([\d,]+\.\d{2})", digital_full_text, re.IGNORECASE)
    m_sk = re.search(r"Saldo\s+Akhir\s+Rp\.\s*([\d,]+\.\d{2})", digital_full_text, re.IGNORECASE)
    official_sa = float(m_sa.group(1).replace(",", "")) if m_sa else 0.0
    official_sk = float(m_sk.group(1).replace(",", "")) if m_sk else 0.0

    transactions = []
    saldo_list = []
    freq_db = 0
    freq_cr = 0
    mutasi_db = 0.0
    mutasi_cr = 0.0
    tx_no = 1

    # 1. Metode Utama: Ekstraksi Tabel Digital pdfplumber
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for pno, p in enumerate(pdf.pages):
                tables = p.extract_tables()
                for tbl in tables:
                    if not tbl:
                        continue
                    header = [str(c or "").strip() for c in tbl[0]]
                    if "Tanggal" in header and any(x in header for x in ["Debet", "Kredit", "Saldo"]):
                        data_rows = tbl[1:]
                    elif len(header) >= 6 and any("IDR" in str(c) or "Rp" in str(c) for c in header):
                        data_rows = tbl
                    else:
                        continue

                    for r in data_rows:
                        if len(r) < 6:
                            continue
                        tgl_raw = str(r[0] or "").strip()
                        if "Tanggal" in tgl_raw:
                            continue

                        v_db = clean_kalsel_val(r[3])
                        v_cr = clean_kalsel_val(r[4])
                        v_bal = clean_kalsel_val(r[5])

                        if v_db == 0.0 and v_cr == 0.0 and v_bal == 0.0:
                            continue

                        # Format tanggal ringkas (DD/MM)
                        m_dt = re.search(r"(\d{2})/(\d{2})/\d{4}", tgl_raw)
                        short_dt = f"{m_dt.group(1)}/{m_dt.group(2)}" if m_dt else tgl_raw.splitlines()[0]

                        ref_val = str(r[1] or "").strip()
                        ket_val = " ".join(str(r[2] or "").replace("\ue088", "-").replace("\ue092", ":").split())
                        full_desc = f"{ref_val} - {ket_val}" if ref_val else ket_val

                        if v_db > 0:
                            freq_db += 1
                            mutasi_db += v_db
                        if v_cr > 0:
                            freq_cr += 1
                            mutasi_cr += v_cr
                        if v_bal > 0:
                            saldo_list.append(v_bal)

                        transactions.append({
                            "no": tx_no,
                            "tanggal": short_dt,
                            "uraian": full_desc,
                            "nominal_db": v_db,
                            "nominal_cr": v_cr,
                            "saldo": v_bal,
                        })
                        tx_no += 1
    except Exception as e:
        print(f"  [!] Peringatan ekstraksi tabel: {e}")

    # 2. Metode Cadangan: OCR jika tabel digital kosong
    if not transactions:
        print("  [i] Tabel digital tidak ditemukan, beralih ke OCR fallback...")
        images = convert_pdf_to_images(pdf_bytes, dpi=300)
        ocr_lines = extract_ocr_text(images)
        full_ocr = "\n".join(ocr_lines)
        if bulan_nama == "Juli":
            bulan_nama = detect_kalsel_month(full_ocr, filename)

        m_sa = re.search(r"Saldo\s+Awal[^\d]*?([\d,]+\.\d{2})", full_ocr, re.IGNORECASE)
        m_sk = re.search(r"Saldo\s+Akhir[^\d]*?([\d,]+\.\d{2})", full_ocr, re.IGNORECASE)
        if m_sa: official_sa = float(m_sa.group(1).replace(",", ""))
        if m_sk: official_sk = float(m_sk.group(1).replace(",", ""))

        for line in ocr_lines:
            m = re.search(r"(\d{2}/\d{2}/\d{4}).*?(-IDR\s*Rp\.\s*[\d,]+|IDR\s*Rp\.\s*[\d,]+).*?IDR\s*Rp\.\s*([\d,]+\.\d{2})", line)
            if m:
                short_dt = m.group(1)[:5]
                raw_amt = m.group(2)
                raw_bal = m.group(3)
                v_bal = float(raw_bal.replace(",", ""))
                is_db = "-" in raw_amt
                v_amt = clean_kalsel_val(raw_amt)

                if is_db:
                    freq_db += 1
                    mutasi_db += v_amt
                    v_db = v_amt
                    v_cr = 0.0
                else:
                    freq_cr += 1
                    mutasi_cr += v_amt
                    v_db = 0.0
                    v_cr = v_amt

                saldo_list.append(v_bal)
                transactions.append({
                    "no": tx_no,
                    "tanggal": short_dt,
                    "uraian": line[:50],
                    "nominal_db": v_db,
                    "nominal_cr": v_cr,
                    "saldo": v_bal,
                })
                tx_no += 1

    saldo_awal = official_sa if official_sa > 0 else (transactions[0]["saldo"] + transactions[0]["nominal_db"] - transactions[0]["nominal_cr"] if transactions else 0.0)
    saldo_akhir = official_sk if official_sk > 0 else (transactions[-1]["saldo"] if transactions else 0.0)

    valid_saldos = [s for s in saldo_list if s > 0]
    if valid_saldos:
        saldo_max = max(valid_saldos)
        saldo_min = min(valid_saldos)
        saldo_avg = float(np.mean(valid_saldos))
    else:
        saldo_max = max(saldo_awal, saldo_akhir)
        saldo_min = min(saldo_awal, saldo_akhir)
        saldo_avg = (saldo_awal + saldo_akhir) / 2.0 if (saldo_awal > 0 or saldo_akhir > 0) else 0.0

    print(f"  [✓] Sukses {bulan_nama}: Debet={freq_db}x (Rp {mutasi_db:,.2f}) | Kredit={freq_cr}x (Rp {mutasi_cr:,.2f})")
    print(f"  [✓] Saldo {bulan_nama}: Max=Rp {saldo_max:,.2f} | Avg=Rp {saldo_avg:,.2f} | Min=Rp {saldo_min:,.2f} | Total Baris={len(transactions)}")

    return {
        "bulan": bulan_nama,
        "freq_db": freq_db,
        "freq_cr": freq_cr,
        "mutasi_db": mutasi_db,
        "mutasi_cr": mutasi_cr,
        "saldo_awal": saldo_awal,
        "saldo_akhir": saldo_akhir,
        "saldo_max": saldo_max,
        "saldo_avg": saldo_avg,
        "saldo_min": saldo_min,
        "transactions": transactions,
    }

def sort_resume_chronological(resume_list):
    """Mengurutkan daftar rekap bulanan secara kronologis."""
    def get_month_index(item):
        b = item.get("bulan", "")
        for idx, m_name in enumerate(MONTH_NAMES_ID):
            if m_name.lower() in b.lower():
                return idx
        return 99
    return sorted(resume_list, key=get_month_index)

def generate_form_pdf(output_target, header_info, resume_list, note_oh="", nama_so="", nama_oh=""):
    """
    Menghasilkan dokumen PDF Formulir Validasi Mutasi Rekening resmi Bank Kalsel.
    Mendukung output ke path file (string) maupun in-memory BytesIO (return bytes).
    """
    sorted_resume = sort_resume_chronological(resume_list)

    is_buffer = isinstance(output_target, io.BytesIO) or output_target is None
    buffer = output_target if isinstance(output_target, io.BytesIO) else (io.BytesIO() if is_buffer else None)
    target = buffer if is_buffer else output_target

    doc = SimpleDocTemplate(
        target,
        pagesize=portrait(A4),
        rightMargin=18,
        leftMargin=18,
        topMargin=25,
        bottomMargin=25,
    )
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=12,
        alignment=1,
        spaceAfter=15,
    )
    story.append(Paragraph("<b>FORM VALIDASI MUTASI REKENING</b>", title_style))

    header_rows = [
        ["Cabang", ":", header_info.get("cabang", "")],
        ["Nama Cust", ":", header_info.get("nama_cust", "")],
        ["", "", ""],
        ["Nomor Rekening", ":", header_info.get("no_rekening", "")],
        ["Nama Bank", ":", header_info.get("nama_bank", "BANK KALSEL")],
        ["Nama Pemegang Rekening", ":", header_info.get("nama_pemegang_rek", "")],
    ]
    t_header = Table(header_rows, colWidths=[150, 15, 390], rowHeights=13)
    t_header.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(t_header)
    story.append(Spacer(1, 10))

    col_widths = [75, 55, 40, 40, 80, 80, 60, 60, 60]
    table_data = [
        [
            Paragraph("<para align='center'><b>Bulan</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Bank</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Freq<br/>Debet</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Freq<br/>Kredit</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Total Debet<br/>(Rp.)</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Total Kredit<br/>(Rp.)</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Saldo<br/>Tertinggi</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Saldo<br/>Rata-Rata</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Saldo<br/>Terendah</b></para>", styles["Normal"]),
        ]
    ]

    for r in sorted_resume:
        table_data.append([
            r.get("bulan", "-"),
            "BANK KALSEL",
            str(r.get("freq_db", 0)),
            str(r.get("freq_cr", 0)),
            f"{r.get('mutasi_db', 0):,.2f}",
            f"{r.get('mutasi_cr', 0):,.2f}",
            f"{r.get('saldo_max', 0):,.2f}",
            f"{r.get('saldo_avg', 0):,.2f}",
            f"{r.get('saldo_min', 0):,.2f}",
        ])

    n = len(sorted_resume)
    if n > 0:
        avg_f_db = int(round(sum(r.get("freq_db", 0) for r in sorted_resume) / n))
        avg_f_cr = int(round(sum(r.get("freq_cr", 0) for r in sorted_resume) / n))
        avg_m_db = sum(r.get("mutasi_db", 0.0) for r in sorted_resume) / n
        avg_m_cr = sum(r.get("mutasi_cr", 0.0) for r in sorted_resume) / n
        avg_s_max = sum(r.get("saldo_max", 0.0) for r in sorted_resume) / n
        avg_s_avg = sum(r.get("saldo_avg", 0.0) for r in sorted_resume) / n
        avg_s_min = sum(r.get("saldo_min", 0.0) for r in sorted_resume) / n

        table_data.append([
            "Rata - Rata",
            "BANK KALSEL",
            str(avg_f_db),
            str(avg_f_cr),
            f"{avg_m_db:,.2f}",
            f"{avg_m_cr:,.2f}",
            f"{avg_s_max:,.2f}",
            f"{avg_s_avg:,.2f}",
            f"{avg_s_min:,.2f}",
        ])

    t_resume = Table(table_data, colWidths=col_widths)
    t_style = [
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E1F2")),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("ALIGN", (2, 0), (3, -1), "CENTER"),
        ("ALIGN", (4, 0), (-1, -1), "RIGHT"),
    ]

    if n > 0:
        last_row = len(table_data) - 1
        t_style.extend([
            ("BACKGROUND", (0, last_row), (-1, last_row), colors.HexColor("#E2EFDA")),
            ("FONTNAME", (0, last_row), (-1, last_row), "Helvetica-Bold"),
        ])

    t_resume.setStyle(TableStyle(t_style))
    story.append(t_resume)
    story.append(Spacer(1, 15))

    oh_text = note_oh if note_oh and note_oh.strip() else "-"
    note_data = [
        [Paragraph("<b>Note OH :</b>", styles["Normal"])],
        [Paragraph(f"{oh_text}", styles["Normal"])],
    ]
    t_note = Table(note_data, colWidths=[555])
    t_note.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(t_note)
    story.append(Spacer(1, 20))

    so_disp = nama_so if nama_so and nama_so.strip() else "( .................................... )"
    oh_disp = nama_oh if nama_oh and nama_oh.strip() else "( .................................... )"

    sign_data = [
        [
            Paragraph("<para align='center'>Yang Mengajukan,</para>", styles["Normal"]),
            Paragraph("<para align='center'>Yang Menyetujui,</para>", styles["Normal"]),
        ],
        ["", ""],
        ["", ""],
        [
            Paragraph(f"<para align='center'><b>{so_disp}</b><br/>Sales Officer</para>", styles["Normal"]),
            Paragraph(f"<para align='center'><b>{oh_disp}</b><br/>Operation Head</para>", styles["Normal"]),
        ],
    ]
    t_sign = Table(sign_data, colWidths=[277, 278], rowHeights=[15, 20, 20, 25])
    t_sign.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    story.append(t_sign)

    doc.build(story)
    print("  [✓] Form PDF Bank Kalsel selesai dibuat.")

    if is_buffer:
        buffer.seek(0)
        return buffer.getvalue()
    return None

def generate_form_excel(output_target, header_info, resume_list, note_oh="", nama_so="", nama_oh=""):
    """
    Menghasilkan dokumen Excel Spreadsheet Formulir Validasi Mutasi Rekening resmi Bank Kalsel.
    Lengkap dengan rincian transaksi per bulan di kolom sebelah kanan.
    """
    sorted_resume = sort_resume_chronological(resume_list)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Validasi Mutasi"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Arial", size=11, bold=True)
    font_sec = Font(name="Arial", size=10, bold=True)
    font_bold = Font(name="Arial", size=9, bold=True)
    font_norm = Font(name="Arial", size=9)

    fill_green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_blue = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    fill_grey = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
    fill_soft_blue = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

    thin = Side(border_style="thin", color="000000")
    box_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # 1. Header Identitas Nasabah
    ws.merge_cells("C1:G1")
    ws["C1"] = "FORM VALIDASI MUTASI REKENING"
    ws["C1"].font = font_title
    ws["C1"].alignment = Alignment(horizontal="center", vertical="center")

    headers = [
        (3, "Cabang", header_info.get("cabang", "")),
        (4, "Nama Cust", header_info.get("nama_cust", "")),
        (6, "Nomor Rekening", header_info.get("no_rekening", "")),
        (7, "Nama Bank", header_info.get("nama_bank", "BANK KALSEL")),
        (8, "Nama Pemegang Rekening", header_info.get("nama_pemegang_rek", "")),
    ]
    for r_idx, lbl, val in headers:
        ws.cell(row=r_idx, column=1, value=lbl).font = font_norm
        ws.cell(row=r_idx, column=2, value=":").font = font_norm
        ws.cell(row=r_idx, column=3, value=val).font = font_norm

    # 2. Tabel Resume Mutasi
    tbl_headers = [
        "Bulan", "Bank", "Freq Debet", "Freq Kredit", "Total Debet (Rp.)",
        "Total Kredit (Rp.)", "Saldo Tertinggi", "Saldo Rata-Rata", "Saldo Terendah"
    ]
    start_r = 10
    for col_i, h in enumerate(tbl_headers, start=1):
        c = ws.cell(row=start_r, column=col_i, value=h)
        c.font = font_bold
        c.fill = fill_blue
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = box_border

    curr_r = start_r + 1
    for r in sorted_resume:
        row_vals = [
            r.get("bulan", "-"),
            "BANK KALSEL",
            r.get("freq_db", 0),
            r.get("freq_cr", 0),
            r.get("mutasi_db", 0.0),
            r.get("mutasi_cr", 0.0),
            r.get("saldo_max", 0.0),
            r.get("saldo_avg", 0.0),
            r.get("saldo_min", 0.0),
        ]
        for col_i, val in enumerate(row_vals, start=1):
            c = ws.cell(row=curr_r, column=col_i, value=val)
            c.font = font_norm
            c.border = box_border
            if col_i in [1, 2]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i in [3, 4]:
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.number_format = "#,##0"
            else:
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = "#,##0.00"
        curr_r += 1

    n = len(sorted_resume)
    if n > 0:
        avg_f_db = int(round(sum(r.get("freq_db", 0) for r in sorted_resume) / n))
        avg_f_cr = int(round(sum(r.get("freq_cr", 0) for r in sorted_resume) / n))
        avg_m_db = sum(r.get("mutasi_db", 0.0) for r in sorted_resume) / n
        avg_m_cr = sum(r.get("mutasi_cr", 0.0) for r in sorted_resume) / n
        avg_s_max = sum(r.get("saldo_max", 0.0) for r in sorted_resume) / n
        avg_s_avg = sum(r.get("saldo_avg", 0.0) for r in sorted_resume) / n
        avg_s_min = sum(r.get("saldo_min", 0.0) for r in sorted_resume) / n

        avg_vals = [
            "Rata - Rata",
            "BANK KALSEL",
            avg_f_db,
            avg_f_cr,
            avg_m_db,
            avg_m_cr,
            avg_s_max,
            avg_s_avg,
            avg_s_min,
        ]
        for col_i, val in enumerate(avg_vals, start=1):
            c = ws.cell(row=curr_r, column=col_i, value=val)
            c.font = font_bold
            c.fill = fill_green
            c.border = box_border
            if col_i in [1, 2]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_i in [3, 4]:
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.number_format = "#,##0"
            else:
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = "#,##0.00"
        curr_r += 1

    # 3. Note OH & Signatures
    curr_r += 1
    ws.cell(row=curr_r, column=1, value="Note OH :").font = font_bold
    curr_r += 1

    oh_text = note_oh if note_oh and note_oh.strip() else "-"
    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r + 2, end_column=9)
    note_cell = ws.cell(row=curr_r, column=1, value=oh_text)
    note_cell.font = font_norm
    note_cell.alignment = Alignment(vertical="top", wrap_text=True)

    for ro in range(curr_r, curr_r + 3):
        for co in range(1, 10):
            ws.cell(row=ro, column=co).border = box_border

    curr_r += 4
    ws.merge_cells(start_row=curr_r, start_column=2, end_row=curr_r, end_column=4)
    c_peng = ws.cell(row=curr_r, column=2, value="Yang Mengajukan,")
    c_peng.font = font_norm
    c_peng.alignment = Alignment(horizontal="center")

    ws.merge_cells(start_row=curr_r, start_column=6, end_row=curr_r, end_column=8)
    c_setuju = ws.cell(row=curr_r, column=6, value="Yang Menyetujui,")
    c_setuju.font = font_norm
    c_setuju.alignment = Alignment(horizontal="center")

    curr_r += 4
    so_disp = nama_so if nama_so and nama_so.strip() else "( .................................... )"
    oh_disp = nama_oh if nama_oh and nama_oh.strip() else "( .................................... )"

    ws.merge_cells(start_row=curr_r, start_column=2, end_row=curr_r, end_column=4)
    c_so = ws.cell(row=curr_r, column=2, value=so_disp)
    c_so.font = font_bold
    c_so.alignment = Alignment(horizontal="center")

    ws.merge_cells(start_row=curr_r, start_column=6, end_row=curr_r, end_column=8)
    c_oh = ws.cell(row=curr_r, column=6, value=oh_disp)
    c_oh.font = font_bold
    c_oh.alignment = Alignment(horizontal="center")

    curr_r += 1
    ws.merge_cells(start_row=curr_r, start_column=2, end_row=curr_r, end_column=4)
    c_so_lbl = ws.cell(row=curr_r, column=2, value="Sales Officer")
    c_so_lbl.font = font_norm
    c_so_lbl.alignment = Alignment(horizontal="center")

    ws.merge_cells(start_row=curr_r, start_column=6, end_row=curr_r, end_column=8)
    c_oh_lbl = ws.cell(row=curr_r, column=6, value="Operation Head")
    c_oh_lbl.font = font_norm
    c_oh_lbl.alignment = Alignment(horizontal="center")

    # 4. Kolom Rincian Transaksi per Bulan (Kolom K ke kanan)
    k_start = 11
    ws.cell(row=1, column=k_start, value="RINCIAN TRANSAKSI PER BULAN").font = font_title

    cur_k_col = k_start
    for r in sorted_resume:
        m_name = r.get("bulan", "Bulan")
        txs = r.get("transactions", [])

        ws.cell(row=3, column=cur_k_col, value=f"Bulan: {m_name} (BANK KALSEL)").font = font_sec

        headers_tx = ["No", "Tanggal", "Uraian Transaksi", "Debet", "Kredit", "Saldo"]
        fills_tx = [fill_yellow, fill_yellow, fill_yellow, fill_yellow, fill_blue, fill_grey]

        for ci, (ht, fl) in enumerate(zip(headers_tx, fills_tx)):
            cell_h = ws.cell(row=4, column=cur_k_col + ci, value=ht)
            cell_h.font = font_bold
            cell_h.fill = fl
            cell_h.alignment = Alignment(horizontal="center", vertical="center")
            cell_h.border = box_border

        tx_r = 5
        for tx in txs:
            c_no = ws.cell(row=tx_r, column=cur_k_col, value=tx.get("no", tx_r - 4))
            c_no.alignment = Alignment(horizontal="center")
            c_no.font = font_norm
            c_no.border = box_border

            c_tgl = ws.cell(row=tx_r, column=cur_k_col + 1, value=tx.get("tanggal", ""))
            c_tgl.alignment = Alignment(horizontal="center")
            c_tgl.font = font_norm
            c_tgl.border = box_border

            c_ur = ws.cell(row=tx_r, column=cur_k_col + 2, value=tx.get("uraian", ""))
            c_ur.font = font_norm
            c_ur.border = box_border

            c_db = ws.cell(row=tx_r, column=cur_k_col + 3, value=tx.get("nominal_db", 0.0))
            c_db.number_format = "#,##0.00"
            c_db.font = font_norm
            c_db.border = box_border

            c_cr = ws.cell(row=tx_r, column=cur_k_col + 4, value=tx.get("nominal_cr", 0.0))
            c_cr.number_format = "#,##0.00"
            c_cr.font = font_norm
            c_cr.border = box_border

            c_bal = ws.cell(row=tx_r, column=cur_k_col + 5, value=tx.get("saldo", 0.0))
            c_bal.number_format = "#,##0.00"
            c_bal.font = font_norm
            c_bal.border = box_border

            tx_r += 1

        cur_k_col += 7

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

    is_buffer = isinstance(output_target, io.BytesIO) or output_target is None
    if is_buffer:
        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)
        return bio.getvalue()
    else:
        wb.save(output_target)
        return None
