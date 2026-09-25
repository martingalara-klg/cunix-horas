@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
if "%~1"=="" (
    set /p MES="Mes a generar (AAAA-MM, ej 2025-10): "
) else (
    set MES=%~1
)
python -m cunix_horas %MES%
echo.
pause
