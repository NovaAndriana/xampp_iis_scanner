"""
=========================================================================
 build_exe.py
=========================================================================
Script otomatis untuk build "XAMPP & IIS Directory Scanner" menjadi
aplikasi standalone (.exe di Windows) menggunakan PyInstaller.

PENTING: Jalankan script ini di WINDOWS (di komputer yang sama dengan
target akhir aplikasi), bukan di Linux/Mac. PyInstaller tidak bisa
cross-compile -> .exe Windows harus di-build sambil berjalan di Windows.

Cara pakai:
    1. Taruh file ini di folder yang sama dengan xampp_iis_scanner.py
    2. Buka Command Prompt di folder tersebut
    3. Jalankan:  python build_exe.py
    4. Tunggu sampai selesai, lalu cek folder:
       dist/XAMPP_IIS_Scanner/XAMPP_IIS_Scanner.exe

Catatan: proses ini menghasilkan folder "dist/XAMPP_IIS_Scanner/" yang
berisi .exe + semua dependency-nya. Folder ini harus disalin utuh
(bukan cuma file .exe-nya saja) kalau mau dipindahkan ke komputer lain.
=========================================================================
"""

import os
import sys
import subprocess
import importlib.util

SCRIPT_NAME = "xampp_iis_scanner.py"
APP_NAME = "XAMPP_IIS_Scanner"

# (nama_pip, nama_import) - beberapa package punya nama berbeda antara
# yang dipakai untuk "pip install" dan yang dipakai untuk "import"
REQUIRED_PACKAGES = [
    ("customtkinter", "customtkinter"),
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("pyinstaller", "PyInstaller"),
]


def ensure_installed(pip_name: str, import_name: str):
    if importlib.util.find_spec(import_name) is None:
        print(f"[INFO] Menginstall '{pip_name}' ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
    else:
        print(f"[OK] '{pip_name}' sudah terinstall.")


def get_customtkinter_data_path() -> str:
    import customtkinter
    return os.path.dirname(customtkinter.__file__)


def main():
    print("=" * 60)
    print(" Build XAMPP & IIS Directory Scanner -> Aplikasi Standalone")
    print("=" * 60)

    if not os.path.isfile(SCRIPT_NAME):
        print(f"\n[ERROR] File '{SCRIPT_NAME}' tidak ditemukan di folder ini.")
        print("Pastikan build_exe.py berada di folder yang sama dengan "
              f"'{SCRIPT_NAME}'.")
        sys.exit(1)

    print("\n[1/3] Memeriksa dependensi...")
    for pip_name, import_name in REQUIRED_PACKAGES:
        ensure_installed(pip_name, import_name)

    print("\n[2/3] Menyiapkan konfigurasi build...")
    ctk_data_path = get_customtkinter_data_path()
    print(f"      Lokasi data customtkinter: {ctk_data_path}")

    # PyInstaller pakai ';' sebagai separator --add-data di Windows,
    # dan ':' di Linux/Mac.
    sep = ";" if os.name == "nt" else ":"
    add_data_arg = f"{ctk_data_path}{sep}customtkinter/"

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onedir",          # wajib untuk customtkinter (lihat catatan di atas)
        "--windowed",         # tanpa jendela console hitam saat aplikasi jalan
        "--name", APP_NAME,
        "--add-data", add_data_arg,
        SCRIPT_NAME,
    ]

    print("\n[3/3] Menjalankan PyInstaller...")
    print("      Perintah:", " ".join(cmd))
    print()

    subprocess.check_call(cmd)

    exe_name = f"{APP_NAME}.exe" if os.name == "nt" else APP_NAME
    result_path = os.path.join("dist", APP_NAME, exe_name)

    print("\n" + "=" * 60)
    print(" BUILD SELESAI")
    print("=" * 60)
    print(f"Aplikasi ada di: {result_path}")
    print(f"Salin SELURUH folder 'dist/{APP_NAME}/' (bukan cuma file "
          f"{exe_name}) kalau mau dipindahkan/dijalankan di komputer lain.")


if __name__ == "__main__":
    main()
