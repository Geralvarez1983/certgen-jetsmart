@echo off
title CertGen - Actualizando...
color 0A
echo.
echo  ========================================
echo    CertGen - Buscando actualizaciones...
echo  ========================================
echo.
cd /d "%~dp0"

:: Verificar si git está instalado
git --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Git no está instalado.
    echo  Descargalo desde: https://git-scm.com
    pause
    exit /b 1
)

echo  Actualizando desde GitHub...
git pull origin main
echo.
echo  ¡App actualizada! Iniciando...
echo.
timeout /t 2 >nul
call INICIAR.bat
