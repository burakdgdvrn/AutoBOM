@echo off
chcp 65001 >nul
echo.
echo ============================================
echo   PDF'den Excel'e Tablo Cikarma Sistemi
echo   (Fabrication + Erection + Cut Pipe Length)
echo ============================================
echo.
echo Sunucu baslatiliyor...
echo Lutfen tarayicidan frontend/ERP Sistemi.html sayfasini acin.
echo.
call venv\Scripts\activate.bat
set PYTHONIOENCODING=utf-8
python backend\app.py
echo.
echo Sunucu kapandi.
pause
