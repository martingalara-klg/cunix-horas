@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

setlocal enabledelayedexpansion

rem `where python` no alcanza: Windows 10/11 trae un stub de la Microsoft Store
rem en WindowsApps que aparece en el PATH aunque Python no este instalado. Con
rem `where` la verificacion daba exito y al usuario se le abria la tienda.
rem Por eso: se descartan los stubs de WindowsApps y se comprueba que el
rem ejecutable realmente corra.
set "PYTHON_EXE="
for /f "delims=" %%P in ('where python 2^>nul') do (
    if not defined PYTHON_EXE (
        echo %%P | find /i "\WindowsApps\" >nul
        if errorlevel 1 (
            "%%P" -c "import sys" >nul 2>nul
            rem Con comillas: la ruta puede tener espacios, como Program Files.
            if not errorlevel 1 set PYTHON_EXE="%%P"
        )
    )
)

rem Ultimo intento: el lanzador `py`, que se instala con Python y no depende
rem de que python.exe haya quedado en el PATH.
if not defined PYTHON_EXE (
    py -3 -c "import sys" >nul 2>nul
    if not errorlevel 1 set "PYTHON_EXE=py -3"
)

if not defined PYTHON_EXE (
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
    echo   Si Python ya esta instalado, es que no quedo en el PATH, o lo unico
    echo   que hay es el atajo a la Microsoft Store, que no sirve:
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
%PYTHON_EXE% -m cunix_horas %MES%
echo.
pause
endlocal
