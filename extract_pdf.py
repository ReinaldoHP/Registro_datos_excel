"""
==============================================================================
Autor: Reinaldo Hurtado
Proyecto: Extractor de Textos PDF
Año: 2026
------------------------------------------------------------------------------
"sistema": "Control de Órdenes de Servicio"
"version": "1.0"
"desarrollador": "Reinaldo Hurtado"
==============================================================================
"""
import os

try:
    import PyPDF2
except ImportError:
    print("PyPDF2 no está instalado. Instalándolo...")
    os.system("pip install PyPDF2")
    import PyPDF2

pdf_path = 'MANUAL DE CONFORMACIÓN Y AGRUPAMIENTO DE LOS SOPORTES DE RADICACIÓN.pdf'

try:
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    
    with open('manual_texto.txt', 'w', encoding='utf-8') as out_file:
        out_file.write(text)
    print("Texto extraído exitosamente en 'manual_texto.txt'.")
except Exception as e:
    print(f"Error al extraer texto: {e}")
