@echo off
title CertGen - JetSMART
color 0A

echo.
echo  ========================================
echo    CertGen JetSMART - Iniciando...
echo  ========================================
echo.

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no encontrado.
    echo  Instala Python desde https://python.org
    echo  y asegurate de marcar "Add Python to PATH"
    pause
    exit /b 1
)

:: Instalar dependencias
echo  Instalando/verificando dependencias...
python -m pip install flask openpyxl python-docx --quiet --disable-pip-version-check 2>nul
echo  Dependencias OK.
echo.

:: Ir al directorio del bat
cd /d "%~dp0"

:: Abrir navegador luego de 4 segundos
start /b cmd /c "timeout /t 4 >nul && start http://localhost:5000"

echo  Servidor en http://localhost:5000
echo  Deja esta ventana abierta. Para cerrar, presiona CTRL+C
echo.

:: Iniciar Flask
python app.py
pause
