@echo off
title CertGen - Instalación
color 0A
echo.
echo  ============================================
echo    CertGen JetSMART - Instalación inicial
echo  ============================================
echo.

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado.
    echo  1. Instala Python desde https://python.org
    echo  2. Marcá "Add Python to PATH"
    echo  3. Volvé a ejecutar este archivo
    pause
    exit /b 1
)

:: Instalar dependencias
echo  Instalando dependencias Python...
python -m pip install flask openpyxl python-docx --quiet
echo  OK.

:: Verificar LibreOffice
if not exist "C:\Program Files\LibreOffice\program\soffice.exe" (
    if not exist "C:\Program Files (x86)\LibreOffice\program\soffice.exe" (
        echo.
        echo  [AVISO] LibreOffice no encontrado.
        echo  Para generar PDFs, instalalo desde:
        echo  https://www.libreoffice.org
        echo.
    )
)

echo.
echo  ¡Todo listo! Iniciando CertGen...
timeout /t 2 >nul
call INICIAR.bat
