@echo off
chcp 65001 > nul
title DCIC - Instalador
color 0A

echo ============================================================
echo       DCIC - SISTEMA DE DESPACHOS - INSTALADOR
echo ============================================================
echo.
echo Este script instalará todas las dependencias necesarias.
echo.
echo Requisitos previos:
echo   - Python 3.10+ instalado (python.org)
echo   - Google Chrome instalado
echo.
pause

echo.
echo [1/4] Verificando Python...
python --version
if errorlevel 1 (
    echo.
    echo ERROR: Python no está instalado o no está en el PATH.
    echo.
    echo Por favor:
    echo   1. Descarga Python desde https://python.org
    echo   2. Durante la instalación, marca "Add Python to PATH"
    echo   3. Ejecuta este instalador nuevamente
    echo.
    pause
    exit /b 1
)

echo.
echo [2/4] Actualizando pip...
python -m pip install --upgrade pip

echo.
echo [3/4] Instalando dependencias de Python...
pip install -r requirements.txt

echo.
echo [4/4] Verificando instalación...
python -c "import customtkinter; import pdfplumber; import selenium; print('OK - Dependencias instaladas correctamente')"

echo.
echo ============================================================
echo       INSTALACION COMPLETADA
echo ============================================================
echo.
echo Para ejecutar la aplicación:
echo   - Doble clic en "ejecutar.bat"
echo   - O ejecuta: python wms_dcic_gui.py
echo.
echo NOTA: Si los PDFs de Mercadolibre son imágenes, necesitas
echo instalar Tesseract OCR (opcional):
echo   winget install UB-Mannheim.TesseractOCR
echo.
pause
