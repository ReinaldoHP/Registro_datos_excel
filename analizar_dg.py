import pandas as pd
from pathlib import Path

def analizar_reporte():
    ruta_archivo = Path('D:/Registro_datos_excel/DG Report generao marzo.xlsx')
    ruta_salida = Path('D:/Registro_datos_excel/Analisis_Empresas_Marzo.xlsx')

    print(f"Leyendo el archivo: {ruta_archivo.name} ...")
    
    try:
        # Leemos el Excel. Usamos usecols="B,C" o leemos todo
        # La columna B es la 1, la C es la 2 (indexadas desde 0)
        df = pd.read_excel(ruta_archivo, engine='openpyxl')
        
        # Asumiendo que la columna B (índice 1) es Empresa y C (índice 2) es Régimen
        col_empresa = df.columns[1]
        col_regimen = df.columns[2]
        
        print(f"Analizando Columna B: '{col_empresa}' y Columna C: '{col_regimen}'")
        
        # Limpiar datos (quitar espacios en blanco al inicio/final y pasar a mayúsculas para unificar)
        df[col_empresa] = df[col_empresa].astype(str).str.strip().str.upper()
        df[col_regimen] = df[col_regimen].astype(str).str.strip().str.upper()
        
        # Filtrar valores nulos o "NAN"
        df_filtrado = df[(df[col_empresa] != 'NAN') & (df[col_regimen] != 'NAN')]
        
        # Agrupar por Empresa y Régimen y contar
        resumen = df_filtrado.groupby([col_empresa, col_regimen]).size().reset_index(name='Cantidad_Facturas_o_Registros')
        
        # Ordenar por empresa
        resumen = resumen.sort_values(by=col_empresa)
        
        # Guardar en un nuevo Excel
        resumen.to_excel(ruta_salida, index=False)
        print(f"\n¡Análisis completado!")
        print(f"Se ha guardado el resultado en: {ruta_salida.name}")
        
        # Mostrar un pequeño resumen en pantalla
        print("\n=== RESUMEN POR RÉGIMEN ===")
        print(df_filtrado[col_regimen].value_counts())
        
        print("\n=== TOP 10 EMPRESAS CON MÁS REGISTROS ===")
        print(df_filtrado[col_empresa].value_counts().head(10))

    except Exception as e:
        print(f"Ocurrió un error al procesar el archivo: {e}")

if __name__ == "__main__":
    analizar_reporte()
