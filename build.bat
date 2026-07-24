@echo off
REM =========================================================
REM  Build XAMPP & IIS Directory Scanner menjadi file .exe
REM  Tinggal double-click file ini di Windows.
REM =========================================================
title Build XAMPP ^& IIS Directory Scanner

python build_exe.py

echo.
echo Tekan tombol apa saja untuk menutup jendela ini...
pause >nul
