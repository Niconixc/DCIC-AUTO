@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo WMS DCIC - Instalando dependencias...
echo ============================================================

pip install customtkinter pdfplumber selenium webdriver-manager >nul 2>&1

echo Iniciando interfaz grafica...
python "%~dp0wms_dcic_gui.py"

pause
