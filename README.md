# XAMPP & IIS Directory Scanner

Aplikasi desktop (Windows) berbasis Python + [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) untuk mengaudit direktori server lokal — memindai folder `htdocs` (XAMPP) dan `wwwroot` (IIS), mengidentifikasi project/aplikasi web di dalamnya, lalu meng-export hasilnya ke file Excel (`.xlsx`) yang sudah rapi dan siap dipakai untuk dokumentasi/laporan.

---

## ✨ Fitur

- **Auto-detect path default** saat aplikasi dibuka:
  - `C:\xampp\htdocs`
  - `D:\xampp\htdocs`
  - `D:\xampp8\htdocs`
  - `C:\xampp8\htdocs`
  - `C:\inetpub\wwwroot`
- **Tambah path manual** lewat dialog pilih folder.
- **Deteksi framework/tech stack** otomatis per project (lihat tabel di bawah).
- **Optimasi performa**: folder berat seperti `node_modules`, `vendor`, dan `.git` tidak ditelusuri isinya (hanya dihitung sebagai 1 subfolder), sehingga scan tetap cepat walau ada banyak dependency.
- **Scan berjalan di background thread** — GUI tidak freeze, progress bar & log console update realtime.
- **Preview hasil** dalam tabel (Treeview), dengan klik ganda untuk langsung membuka folder project di File Explorer.
- **Export ke Excel (.xlsx)**:
  - Header berwarna biru gelap, teks putih bold
  - Lebar kolom auto-fit menyesuaikan isi
  - Dialog "Save As" untuk memilih lokasi & nama file

---

## 🧩 Deteksi Framework / Tech Stack

| Penanda yang Ditemukan                          | Hasil Deteksi                     |
|--------------------------------------------------|------------------------------------|
| `wp-config.php`                                  | WordPress                          |
| `artisan`                                        | Laravel                            |
| `composer.json` (berisi `laravel/framework`)     | Laravel                            |
| `composer.json` (tanpa `laravel/framework`)      | PHP (Composer)                     |
| `package.json` (berisi `next`)                   | Node.js (Next.js)                  |
| `package.json` (berisi `react`)                  | Node.js (React)                    |
| `package.json` (berisi `vue`)                    | Node.js (Vue)                      |
| `package.json` (berisi `@angular/core`)          | Node.js (Angular)                  |
| `package.json` (lainnya)                         | Node.js                            |
| `web.config`                                     | ASP.NET / IIS Configured           |
| `index.php`                                      | Native PHP                         |
| `index.html`                                     | Native HTML                        |
| Tidak ada file penanda di atas                   | Folder / Asset Biasa               |

---

## 📋 Requirements

- **Windows 7 SP1** atau lebih baru (Windows 10/11 direkomendasikan)
- **Python 3.8+**
  > ⚠️ Untuk **Windows 7**, gunakan **Python 3.8.10** — versi resmi terakhir yang masih mendukung Windows 7 dan menyediakan installer biner. Python 3.9 ke atas **tidak bisa diinstall** di Windows 7.
  > Download: https://www.python.org/downloads/release/python-3810/

---

## 🚀 Instalasi

1. Install dependensi Python:
   ```bash
   pip install customtkinter pandas openpyxl
   ```

2. Jalankan aplikasi:
   ```bash
   python xampp_iis_scanner.py
   ```

---

## 📦 Build Menjadi Aplikasi Standalone (.exe)

Kalau ingin dibagikan tanpa perlu install Python di komputer tujuan, aplikasi ini bisa di-build jadi `.exe` pakai PyInstaller.

> ⚠️ Build **harus dijalankan di Windows** (bukan Linux/Mac), karena PyInstaller tidak melakukan cross-compile — `.exe` Windows hanya bisa dihasilkan saat script berjalan di Windows.

**Cara build:**

1. Pastikan `build_exe.py` ada di folder yang sama dengan `xampp_iis_scanner.py`.
2. Jalankan salah satu:
   ```bash
   python build_exe.py
   ```
   atau cukup **double-click `build.bat`**.
3. Script akan otomatis: cek/install dependensi (termasuk PyInstaller), mendeteksi lokasi folder `customtkinter`, lalu menjalankan PyInstaller dengan konfigurasi yang benar.
4. Hasil build ada di:
   ```
   dist/XAMPP_IIS_Scanner/XAMPP_IIS_Scanner.exe
   ```
5. Untuk didistribusikan, salin **seluruh folder** `dist/XAMPP_IIS_Scanner/` (bukan cuma file `.exe`-nya) — folder ini berisi semua dependency yang dibutuhkan agar `.exe` bisa jalan tanpa Python terinstall.

---

## 📁 Struktur Project

```
.
├── xampp_iis_scanner.py   # Source code utama aplikasi
├── build_exe.py           # Script otomatis untuk build ke .exe
├── build.bat              # Wrapper (double-click untuk build di Windows)
└── README.md              # Dokumentasi ini
```

---

## 📊 Format Export Excel

| No | Nama Project | Sumber Server | Full Path | Framework/Tech | Ukuran (MB) | Terakhir Dimodifikasi | Jumlah File |
|----|---------------|----------------|-----------|------------------|--------------|-------------------------|--------------|

---

## 🛠️ Cara Kerja Singkat

1. Aplikasi memeriksa path default & path manual yang dicentang pengguna.
2. Untuk setiap subfolder di dalamnya (dianggap 1 project), aplikasi:
   - Mendeteksi framework/tech stack berdasarkan file penanda.
   - Menghitung ukuran total, jumlah file, jumlah subfolder, dan tanggal modifikasi terakhir (sambil melewati `node_modules`, `vendor`, `.git`).
3. Hasil ditampilkan realtime di log & tabel preview.
4. Pengguna dapat meng-export seluruh hasil ke file `.xlsx` yang sudah diformat.

---

## 📝 Catatan

- Fitur "buka folder di Explorer" (klik ganda pada baris hasil) hanya berfungsi di Windows karena menggunakan `os.startfile()`.
- Path server (`C:\xampp\...`, `C:\inetpub\...`) adalah path khas Windows, sehingga aplikasi ini ditujukan untuk dijalankan di Windows.
