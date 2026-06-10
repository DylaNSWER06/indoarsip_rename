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

Ada 2 cara untuk mendapatkan file project ini:

---

### 🅰️ Opsi 1: Menggunakan Git (punya GitHub)

#### 1. Clone repository
```bash
git clone https://github.com/DylaNSWER06/indoarsip_rename.git
```
Lalu buka folder `indoarsip_rename` di VS Code.

---

### 🅱️ Opsi 2: Download ZIP (tanpa GitHub)

1. Buka [https://github.com/DylaNSWER06/indoarsip_rename](https://github.com/DylaNSWER06/indoarsip_rename)
2. Klik tombol hijau **"Code"**
3. Klik **"Download ZIP"**
4. Extract ZIP nya
5. Buka folder hasil extract di VS Code

---

### Langkah selanjutnya di Terminal VS Code (sama untuk kedua opsi) 

#### 1. Buat virtual environment
```bash
python -m venv venv
```

#### 2. Aktifkan virtual environment
```bash
# Windows
venv\Scripts\activate

# Mac/Linux
source venv/bin/activate
```

#### 3. Install dependencies
```bash
pip install -r requirements.txt
```

#### 4. Jalankan aplikasi
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