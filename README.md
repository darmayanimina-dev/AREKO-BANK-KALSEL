# AREKO Bank Kalsel - Rekapitulasi Mutasi Rekening Koran Bank Kalsel

Aplikasi cerdas dan otomatis untuk menganalisis, mengekstraksi, dan merekapitulasi data rekening koran **Bank Kalsel** (termasuk format Giro iB Al-Amanah Wadiah / Syariah dan Konvensional) dari berkas PDF menjadi Formulir Validasi Mutasi Rekening resmi (Format PDF A4 siap cetak dan Microsoft Excel Spreadsheet dinamis).

---

## ✨ Fitur Utama

1. **Engine Parsing Khusus Bank Kalsel**:
   - Ekstraksi tabel digital native berkecepatan tinggi dengan pemisahan tabel metadata header dan tabel mutasi transaksi.
   - Pembersihan simbol mata uang khusus Bank Kalsel (`-IDR Rp.`, `IDR Rp.`, prefix unicode khusus) tanpa merusak pemisah ribuan dan desimal.
   - Rekonstruksi dan validasi saldo mutasi (Debet/Kredit) secara matematis akurat 100% pas terhadap saldo awal (*Opening Balance*) dan saldo akhir (*Closing Balance*).
   - Fallback OCR otomatis beresolusi tinggi (Tesseract OCR Engine) untuk dokumen hasil scan.

2. **Ringkasan Mutasi Komprehensif**:
   - Menghitung frekuensi Debet & Kredit.
   - Menghitung total perputaran dana mutasi Debet & Kredit.
   - Menganalisis saldo tertinggi (max), saldo terendah (min), dan saldo rata-rata (average).
   - Menghitung baris **Rata-Rata** keseluruhan bulan yang diunggah.

3. **Ekspor Dokumen Standar Perbankan**:
   - **PDF Resmi**: Dokumen A4 siap cetak dilengkapi informasi nasabah, nomor rekening, catatan Operation Head, dan kolom tanda tangan Sales Officer & Operation Head.
   - **Excel Spreadsheet (.xlsx)**: Lembar kerja dinamis lengkap dengan tabel rekapitulasi dan rincian transaksi per bulan di kolom rincian.

4. **Desain Antarmuka Modern & Bersih**:
   - Tampilan tema terang (*Light Mode*) yang profesional dan responsif.
   - Tabel pratinjau interaktif dengan penanda visual khusus warna hijau lembut (`#E2EFDA`) pada baris rata-rata.

---

## 🛠️ Instalasi & Menjalankan Lokal

### 1. Kebutuhan Sistem
Pastikan sistem Anda telah memiliki **Python 3.9+** serta dependensi sistem `tesseract-ocr` dan `poppler-utils`:

- **macOS (via Homebrew)**:
  ```bash
  brew install tesseract poppler
  ```
- **Ubuntu/Debian / Streamlit Cloud**:
  Sudah terkonfigurasi di `packages.txt` (`tesseract-ocr`, `poppler-utils`).

### 2. Instal Dependensi Python
```bash
cd BANK-KALSEL
pip install -r requirements.txt
```

### 3. Jalankan Aplikasi
```bash
streamlit run app.py
```

Akses aplikasi melalui browser di `http://localhost:8501`.

---

## ☁️ Panduan Deploy ke Streamlit Cloud

1. Buat repository baru di GitHub (misalnya: `AREKO-BANK-KALSEL`).
2. Hubungkan remote dan lakukan push:
   ```bash
   git remote add origin https://github.com/<username>/AREKO-BANK-KALSEL.git
   git branch -M main
   git push -u origin main
   ```
3. Buka [share.streamlit.io](https://share.streamlit.io) dan klik **New App**.
4. Pilih repositori `AREKO-BANK-KALSEL`, branch `main`, dan set file path ke `app.py`.
5. Streamlit Cloud akan otomatis membaca `packages.txt` dan `requirements.txt` untuk menginstal Tesseract & Poppler.
