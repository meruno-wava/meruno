@echo off
chcp 65001 >nul
cls
echo.
echo  ========================================================
echo   DAIRY MERUNO ^| Sistem Informasi Arsip RS
echo   http://localhost:8080
echo  ========================================================
echo.

cd /d "%~dp0"

REM Cari Python yang tersedia
set PY=
python --version >nul 2>&1
if not errorlevel 1 set PY=python

if "%PY%"=="" (
    py --version >nul 2>&1
    if not errorlevel 1 set PY=py
)

if "%PY%"=="" (
    echo [ERROR] Python tidak ditemukan. Install dari https://python.org
    pause & exit /b
)

echo  Python: %PY%
echo.

REM Install dependencies MENGGUNAKAN python -m pip (bukan pip langsung)
echo [1/2] Menginstal dependencies...
%PY% -m pip install flask openpyxl pandas pdfplumber --quiet
if errorlevel 1 (
    echo        Coba dengan --user...
    %PY% -m pip install flask openpyxl pandas pdfplumber --user --quiet
)
echo        selesai.
echo.

echo [2/2] Menjalankan server Flask...
echo       Tekan Ctrl+C untuk berhenti.
echo.

%PY% app.py

pause
