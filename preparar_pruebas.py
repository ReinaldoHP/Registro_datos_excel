import pandas as pd
from pathlib import Path
import os

def crear_pdf_falso(ruta):
    Path(ruta).parent.mkdir(parents=True, exist_ok=True)
    with open(ruta, 'w') as f:
        f.write("%PDF-1.4 Mock Content")

def preparar_entorno():
    base_test = Path("d:/Registro_datos_excel/PRUEBAS_AUDITOR")
    base_test.mkdir(parents=True, exist_ok=True)
    
    config = {
        "100": {"tipo": "Normal", "pdfs": 4, "subfolder": ""},
        "200_AZUL": {"tipo": "Azul", "pdfs": 4, "subfolder": ""},
        "300": {"tipo": "Policia", "pdfs": 4, "subfolder": "POLICIA"},
        "400": {"tipo": "Incompleto", "pdfs": 2, "subfolder": ""},
        "500": {"tipo": "Duplicado1", "pdfs": 4, "subfolder": "ZONA_NORTE"},
        "500_DUPLICADO": {"tipo": "Duplicado2", "pdfs": 4, "subfolder": "ZONA_SUR"}
    }
    
    # Crear carpetas y PDFs
    for name, data in config.items():
        folder_path = base_test / data["subfolder"] / name
        for i in range(data["pdfs"]):
            crear_pdf_falso(folder_path / f"soporte_{i+1}.pdf")
    
    # Crear Excel de Prueba
    facturas = [
        "HSVE000100", 
        "HSVE000200", 
        "HSVE000300", 
        "HSVE000400", 
        "HSVE000500", 
        "HSVE000999" # No existe
    ]
    
    df = pd.DataFrame({"SFANUMFAC": facturas})
    excel_path = base_test / "excel_de_prueba.xlsx"
    df.to_excel(excel_path, index=False)
    
    print(f"Entorno de pruebas creado en: {base_test}")
    print(f"Excel creado en: {excel_path}")
    return excel_path, base_test

if __name__ == "__main__":
    preparar_entorno()
