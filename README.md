# 📦 INDOARSIP
### Sistem Otomatis Penamaan Arsip Digital

Aplikasi web untuk rename file arsip secara batch berdasarkan data referensi Excel. Dibangun dengan Streamlit.

---

## ✨ Fitur

- Upload file arsip dalam format **ZIP** atau **multiple files**
- Upload file **Excel referensi** sebagai acuan penamaan
- **Matching otomatis** berdasarkan kode numerik di nama file
- Preview hasil rename sebelum diproses
- Download hasil sebagai **ZIP** atau **file individual**
- Laporan Excel untuk file yang tidak cocok

---

## 🚀 Cara Install & Run

### 1. Clone repository
```bash
git clone https://github.com/DylaNSWER06/indoarsip_rename.git
cd indoarsip_rename
```

### 2. Buat virtual environment
```bash
python -m venv venv
```

### 3. Aktifkan virtual environment
```bash
# Windows
venv\Scripts\activate

# Mac/Linux
source venv/bin/activate
```

### 4. Install dependencies
```bash
pip install -r requirements.txt
```

### 5. Jalankan aplikasi
```bash
streamlit run app.py
```

Aplikasi akan terbuka otomatis di browser di `http://localhost:8501`

---

## 📖 Cara Pakai

1. **Tab 1 - Upload & Validasi**
   - Pilih tipe upload (ZIP atau multiple files)
   - Upload file arsip
   - Upload file Excel referensi
   - Isi nama kolom referensi di Excel
   - Klik **Validasi & Cek Arsip**

2. **Tab 2 - Preview & Proses Rename**
   - Cek preview nama file lama → nama file baru
   - Klik **Mulai Proses Rename Arsip**
   - Download hasil sebagai ZIP atau file individual

---

## 📁 Struktur File

```
indoarsip/
├── app.py              # File utama aplikasi
├── requirements.txt    # Daftar dependencies
├── README.md           # Dokumentasi ini
└── .gitignore          # File yang diabaikan Git
```

---

## 🔧 Requirements

- Python 3.8+
- Streamlit
- Pandas
- Openpyxl

---

## 🏢 Tentang

**INDOARSIP** - Layanan Penyimpanan & Manajemen Arsip Profesional
