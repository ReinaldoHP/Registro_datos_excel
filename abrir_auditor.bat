@echo off
cd /d "%~dp0"
start "" "C:\Users\reina\AppData\Local\Python\bin\pythonw.exe" "%~dp0auditor_facturas.py"
if %errorlevel% neq 0 (
    echo No se pudo iniciar con pythonw.exe, intentando con python.exe...
    "C:\Users\reina\AppData\Local\Python\bin\python.exe" "%~dp0auditor_facturas.py"
)
