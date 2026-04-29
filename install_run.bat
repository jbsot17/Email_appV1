@echo off
echo ========================================
echo  Email Masivo Sender
echo  Creador de Borradores Gmail
echo ========================================
echo.
cd /d "%~dp0"
pip install -r requirements.txt
echo.
echo Instalacion completada!
echo.
python app.py
pause