@echo off
cd /d "%~dp0"
set "APP_EXE=%~dp0dist\auditor_facturas v2.0.13.exe"
if exist "%APP_EXE%" (
    start "" "%APP_EXE%"
    exit /b 0
)

rem Sin ejecutable: usar el Python del entorno virtual si existe; si no, el de PATH.
set "PYW=pythonw"
set "PY=python"
if exist "%~dp0.venv\Scripts\pythonw.exe" set "PYW=%~dp0.venv\Scripts\pythonw.exe"
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"

start "" "%PYW%" "%~dp0auditor_facturas.py"
if %errorlevel% neq 0 (
    echo No se pudo iniciar con pythonw, intentando con python...
    "%PY%" "%~dp0auditor_facturas.py"
)
