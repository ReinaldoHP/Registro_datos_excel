import tempfile
import unittest
from pathlib import Path

import pandas as pd

from auditor_facturas import InvoiceAuditor


class MergeExcelTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.tmp_path = Path(self.temp_dir.name)

    def test_actualizar_excel_con_nuevos_sin_perder_datos(self):
        base_file = self.tmp_path / 'excel_base.xlsx'
        nuevo_file = self.tmp_path / 'excel_nuevo.xlsx'
        salida_file = self.tmp_path / 'excel_mezclado.xlsx'

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

        self.assertEqual(resultado, salida_file)

        df_resultado = pd.read_excel(salida_file)
        sfanums = [str(value).strip() for value in df_resultado['SFANUMFAC']]
        self.assertEqual(sfanums, ['1001', '1002', '1003'])
        fila_1002 = df_resultado['SFANUMFAC'] == 1002
        fila_1003 = df_resultado['SFANUMFAC'] == 1003
        self.assertEqual(df_resultado.loc[fila_1002, 'TOTAL_FACTURADO'].iat[0], 200)
        self.assertEqual(df_resultado.loc[fila_1002, 'ESTADO'].iat[0], 'ANTIGUO')
        self.assertEqual(df_resultado.loc[fila_1003, 'ESTADO'].iat[0], 'NUEVO')


if __name__ == '__main__':
    unittest.main()
