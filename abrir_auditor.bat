@echo off
cd /d "%~dp0"
set "APP_EXE=%~dp0dist\auditor_facturas v2.0.12.exe"
if exist "%APP_EXE%" (
    start "" "%APP_EXE%"
    exit /b 0
)
start "" "C:\Users\reina\AppData\Local\Python\bin\pythonw.exe" "%~dp0auditor_facturas.py"
if %errorlevel% neq 0 (
    echo No se pudo iniciar con pythonw.exe, intentando con python.exe...
    "C:\Users\reina\AppData\Local\Python\bin\python.exe" "%~dp0auditor_facturas.py"
)
