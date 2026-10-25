from pathlib import Path
import pandas as pd

from auditor_facturas import InvoiceAuditor


def test_actualizar_excel_con_nuevos_sin_perder_datos(tmp_path):
    base_file = tmp_path / 'excel_base.xlsx'
    nuevo_file = tmp_path / 'excel_nuevo.xlsx'
    salida_file = tmp_path / 'excel_mezclado.xlsx'

    df_base = pd.DataFrame([
        {'SFANUMFAC': '1001', 'TOTAL_FACTURADO': 100, 'ESTADO': 'ANTIGUO'},
        {'SFANUMFAC': '1002', 'TOTAL_FACTURADO': 200, 'ESTADO': 'ANTIGUO'},
    ])
    df_nuevo = pd.DataFrame([
        {'SFANUMFAC': '1002', 'TOTAL_FACTURADO': 250, 'ESTADO': 'NUEVO'},
        {'SFANUMFAC': '1003', 'TOTAL_FACTURADO': 300, 'ESTADO': 'NUEVO'},
    ])

    df_base.to_excel(base_file, index=False)
    df_nuevo.to_excel(nuevo_file, index=False)

    resultado = InvoiceAuditor.actualizar_excel_con_nuevos(
        excel_antiguo=base_file,
        excel_nuevo=nuevo_file,
        clave='SFANUMFAC',
        destino=salida_file,
    )

    assert resultado == salida_file

    df_resultado = pd.read_excel(salida_file)
    sfanums = [str(value).strip() for value in df_resultado['SFANUMFAC']]
    assert sfanums == ['1001', '1002', '1003']
    assert df_resultado.loc[df_resultado['SFANUMFAC'] == 1002, 'TOTAL_FACTURADO'].iat[0] == 200
    assert df_resultado.loc[df_resultado['SFANUMFAC'] == 1002, 'ESTADO'].iat[0] == 'ANTIGUO'
    assert df_resultado.loc[df_resultado['SFANUMFAC'] == 1003, 'ESTADO'].iat[0] == 'NUEVO'
