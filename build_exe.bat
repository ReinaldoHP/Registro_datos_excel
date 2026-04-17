:: ==============================================================================
:: Autor: Reinaldo Hurtado
:: Proyecto: Constructor de Ejecutable
:: Año: 2026
:: ------------------------------------------------------------------------------
:: "sistema": "Control de Órdenes de Servicio"
:: "version": "1.0"
:: "desarrollador": "Reinaldo Hurtado"
:: ==============================================================================
echo Instalando requerimientos de interfaz...
C:\Users\reina\AppData\Local\Python\bin\python.exe -m pip install customtkinter
echo.
echo Construyendo el ejecutable de auditor_facturas...
C:\Users\reina\AppData\Local\Python\bin\python.exe -m PyInstaller auditor_facturas.spec
echo Proceso finalizado. Puedes encontrar tu nuevo ejecutable en la carpeta 'dist'.
pause
