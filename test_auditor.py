from unittest.mock import patch
import tkinter as tk
from auditor_facturas import InvoiceAuditor
import preparar_pruebas
from pathlib import Path
import pandas as pd

def run_tests():
    print("=== INICIANDO PRUEBAS DEL AUDITOR ===")
    
    # 1. Preparar el entorno de pruebas
    excel_path, search_root = preparar_pruebas.preparar_entorno()
    print(f"Entorno preparado. Excel: {excel_path}, Root: {search_root}")

    # 2. Instanciar la app y ocultarla
    app = InvoiceAuditor()
    app.withdraw()

    # 3. Configurar variables
    app.xl_file_var.set(str(excel_path))
    app.search_root_var.set(str(search_root))

    # --- TEST 1: COLSANITAS (PREPAGADA) ---
    print("\n--- TEST 1: COLSANITAS (PREPAGADA) ---")
    app.empresa_var.set("COLSANITAS (PREPAGADA)")
    colsanitas_save = excel_path.parent / "excel_colsanitas_out.xlsx"
    app.save_path_var.set(str(colsanitas_save))

    with patch('tkinter.messagebox.showinfo') as mock_info, \
         patch('tkinter.messagebox.showerror') as mock_error, \
         patch('tkinter.messagebox.showwarning') as mock_warning:
        app.audit_process()
        
    df_col = pd.read_excel(colsanitas_save)
    print(df_col[['SFANUMFAC', 'RESULTADO_AUDITORIA']])
    
    # Comprobar resultados esperados para Colsanitas
    res_2127892 = df_col.loc[df_col['SFANUMFAC'] == 'HSVE2127892', 'RESULTADO_AUDITORIA'].values[0]
    res_2127893 = df_col.loc[df_col['SFANUMFAC'] == 'HSVE2127893', 'RESULTADO_AUDITORIA'].values[0]
    res_2127894 = df_col.loc[df_col['SFANUMFAC'] == 'HSVE2127894', 'RESULTADO_AUDITORIA'].values[0]
    res_212789 = df_col.loc[df_col['SFANUMFAC'] == 'HSVE212789', 'RESULTADO_AUDITORIA'].values[0]
    
    assert res_2127892 == "SIN RADICAR", f"Esperado SIN RADICAR para 2127892, obtenido: {res_2127892}"
    assert res_2127893 == "FALTAN SOPORTES", f"Esperado FALTAN SOPORTES para 2127893, obtenido: {res_2127893}"
    assert res_2127894 == "NO CARPETA", f"Esperado NO CARPETA para 2127894, obtenido: {res_2127894}"
    assert res_212789 == "NO CARPETA", f"Esperado NO CARPETA para 212789 (parcial), obtenido: {res_212789}"
    print("OK - TEST 1 (COLSANITAS) COMPLETADO CON EXITO")

    # --- TEST 2: GENERAL Y POLICIA ---
    print("\n--- TEST 2: GENERAL Y POLICIA ---")
    app.empresa_var.set("General")
    general_save = excel_path.parent / "excel_general_out.xlsx"
    app.save_path_var.set(str(general_save))

    with patch('tkinter.messagebox.showinfo') as mock_info, \
         patch('tkinter.messagebox.showerror') as mock_error, \
         patch('tkinter.messagebox.showwarning') as mock_warning:
        app.audit_process()

    df_gen = pd.read_excel(general_save)
    print(df_gen[['SFANUMFAC', 'RESULTADO_AUDITORIA']])
    
    # Comprobar resultados esperados para Policia
    res_300 = df_gen.loc[df_gen['SFANUMFAC'] == 'HSVE000300', 'RESULTADO_AUDITORIA'].values[0]
    assert res_300 == "SIN RADICAR", f"Esperado SIN RADICAR para HSVE000300 (Policia), obtenido: {res_300}"
    print("OK - TEST 2 (POLICIA) COMPLETADO CON EXITO")

    # --- TEST 3: RE-AUDITAR SIN DUPLICAR COLUMNAS ---
    print("\n--- TEST 3: RE-AUDITAR SIN DUPLICAR COLUMNAS ---")
    # Usar el archivo general_save como entrada y salida
    app.xl_file_var.set(str(general_save))
    app.save_path_var.set(str(general_save))
    
    with patch('tkinter.messagebox.showinfo') as mock_info, \
         patch('tkinter.messagebox.showerror') as mock_error, \
         patch('tkinter.messagebox.showwarning') as mock_warning:
        app.audit_process()
        
    df_re = pd.read_excel(general_save)
    cols = list(df_re.columns)
    print("Columnas en el archivo re-auditado:", cols)
    
    assert cols.count("RESULTADO_AUDITORIA") == 1, f"Se esperaba 1 columna RESULTADO_AUDITORIA, se encontraron: {cols.count('RESULTADO_AUDITORIA')}"
    print("OK - TEST 3 (SIN DUPLICADOS) COMPLETADO CON EXITO")
    
    print("\n=== TODOS LOS TEST PASARON EXITOSAMENTE ===")

if __name__ == "__main__":
    run_tests()
