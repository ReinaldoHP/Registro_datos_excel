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
from pathlib import Path

try:
    import PyPDF2
except ImportError:
    print("PyPDF2 no está instalado. Instalándolo...")
    os.system("pip install PyPDF2")
    import PyPDF2

project_root = Path(__file__).resolve().parent.parent
references_dir = project_root / "docs" / "referencias"
pdf_path = references_dir / "MANUAL DE CONFORMACIÓN Y AGRUPAMIENTO DE LOS SOPORTES DE RADICACIÓN.pdf"
text_path = references_dir / "manual_texto.txt"

try:
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    
    with open(text_path, 'w', encoding='utf-8') as out_file:
        out_file.write(text)
    print(f"Texto extraído exitosamente en '{text_path}'.")
except Exception as e:
    print(f"Error al extraer texto: {e}")
