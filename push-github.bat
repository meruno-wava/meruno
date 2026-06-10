@echo off
cd /d "%~dp0"

git --version >nul 2>&1
if errorlevel 1 (
    echo Git tidak ditemukan. Install dari https://git-scm.com
    pause
    exit /b
)

if not exist ".git" (
    git init
    git branch -M main
)

git remote remove origin 2>nul
git remote add origin https://github.com/JIAkbar/meruno.git

git add .
git commit -m "update: multi-doc archive, delete endpoint, list panel Excel & PDF"
git push -u origin main

echo.
echo Selesai! Cek https://github.com/JIAkbar/meruno
pause
