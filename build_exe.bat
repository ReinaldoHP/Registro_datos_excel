:: ==============================================================================
:: Autor: Reinaldo Hurtado
:: Proyecto: Constructor de Ejecutable
:: Año: 2026
:: ------------------------------------------------------------------------------
:: "sistema": "Control de Órdenes de Servicio"
:: "version": "1.0"
:: "desarrollador": "Reinaldo Hurtado"
:: ==============================================================================
@echo off
cd /d "%~dp0"

rem Usar el Python del entorno virtual si existe; si no, el de PATH.
set "PY=python"
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
echo Usando Python: %PY%
echo.

echo Instalando requerimientos...
"%PY%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: no se pudieron instalar los requerimientos.
    pause
    exit /b 1
)
echo.

echo Cerrando el ejecutable anterior si está en ejecución...
powershell -NoProfile -Command "Get-Process -Name 'auditor_facturas*' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue"
echo.

echo Actualizando el número de versión (parche)...
"%PY%" scripts\incrementar_version.py
if errorlevel 1 (
    echo ERROR: no se pudo actualizar la versión.
    pause
    exit /b 1
)
echo.

echo Construyendo el ejecutable de auditor_facturas...
"%PY%" -m PyInstaller auditor_facturas.spec
if errorlevel 1 (
    echo ERROR: falló la construcción con PyInstaller.
    pause
    exit /b 1
)
echo Proceso finalizado. Puedes encontrar tu nuevo ejecutable en la carpeta 'dist'.
pause
