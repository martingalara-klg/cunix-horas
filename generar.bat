@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo ERROR: no se encontro Python en esta computadora.
    echo.
    echo   Esta herramienta necesita Python 3 para funcionar.
    echo.
    echo   Que hacer:
    echo     1. Instalar Python desde https://www.python.org/downloads/
    echo     2. En el instalador, TILDAR "Add python.exe to PATH" antes de instalar.
    echo     3. Cerrar esta ventana, abrir una nueva y volver a hacer doble clic aqui.
    echo.
    echo   Si Python ya esta instalado, es que no quedo en el PATH:
    echo   reinstalalo tildando esa opcion, o pedile ayuda a quien mantiene la herramienta.
    echo.
    pause
    exit /b 1
)

if "%~1"=="" (
    set /p MES="Mes a generar (AAAA-MM, ej 2025-10): "
) else (
    set MES=%~1
)
python -m cunix_horas %MES%
echo.
pause
