"""Pruebas de integración del flujo completo de auditoría (GUI oculta).

Crea el entorno de pruebas (carpetas, PDFs falsos y Excel) en un directorio
temporal mediante tests/preparar_pruebas.py, instancia la aplicación Tk oculta
y ejecuta audit_process() con los cuadros de diálogo parcheados.
"""
import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

import auditor_facturas
from auditor_facturas import InvoiceAuditor
from tests import preparar_pruebas


class AuditorIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp_dir = tempfile.TemporaryDirectory()
        cls.excel_path, cls.search_root = preparar_pruebas.preparar_entorno(
            Path(cls.temp_dir.name) / "PRUEBAS_AUDITOR"
        )

        # Evitar leer/escribir el auditor_config.json real del usuario.
        cls._patchers = [
            patch.object(InvoiceAuditor, "load_config", lambda self: None),
            patch.object(InvoiceAuditor, "save_config", lambda self: None),
        ]
        for p in cls._patchers:
            p.start()

        try:
            cls.app = InvoiceAuditor()
        except tk.TclError as exc:
            cls._stop_patchers()
            cls.temp_dir.cleanup()
            raise unittest.SkipTest(f"No hay entorno gráfico disponible: {exc}")
        cls.app.withdraw()

    @classmethod
    def _stop_patchers(cls):
        for p in reversed(cls._patchers):
            p.stop()

    @classmethod
    def tearDownClass(cls):
        try:
            cls.app.destroy()
        except tk.TclError:
            pass
        cls._stop_patchers()
        cls.temp_dir.cleanup()

    def run_audit(self, empresa, entrada, salida):
        app = self.app
        app.xl_file_var.set(str(entrada))
        app.search_root_var.set(str(self.search_root))
        app.empresa_var.set(empresa)
        app.save_path_var.set(str(salida))

        with patch("tkinter.messagebox.showinfo"), \
             patch("tkinter.messagebox.showerror") as mock_error, \
             patch("tkinter.messagebox.showwarning") as mock_warning:
            app.audit_process()

        self.assertFalse(mock_error.called, f"showerror llamado: {mock_error.call_args_list}")
        self.assertFalse(mock_warning.called, f"showwarning llamado: {mock_warning.call_args_list}")
        self.assertTrue(Path(salida).exists(), f"No se generó el archivo de salida {salida}")
        return pd.read_excel(salida)

    @staticmethod
    def resultado(df, factura):
        valores = df.loc[df["SFANUMFAC"] == factura, "RESULTADO_AUDITORIA"].values
        if len(valores) != 1:
            raise AssertionError(f"Se esperaba 1 fila para {factura}, hay {len(valores)}")
        return valores[0]

    def test_colsanitas_prepagada(self):
        salida = Path(self.temp_dir.name) / "excel_colsanitas_out.xlsx"
        df = self.run_audit("COLSANITAS (PREPAGADA)", self.excel_path, salida)

        self.assertEqual(self.resultado(df, "HSVE2127892"), "SIN RADICAR")
        self.assertEqual(self.resultado(df, "HSVE2127893"), "FALTAN SOPORTES")
        self.assertEqual(self.resultado(df, "HSVE2127894"), "NO CARPETA")
        # Coincidencia parcial (212789 es prefijo de 2127892/3) no debe contar.
        self.assertEqual(self.resultado(df, "HSVE212789"), "NO CARPETA")

    def test_general_policia(self):
        salida = Path(self.temp_dir.name) / "excel_general_out.xlsx"
        df = self.run_audit("General", self.excel_path, salida)

        # POLICÍA va dentro de CONSULTA EXTERNA y no exige 7 dígitos.
        self.assertEqual(self.resultado(df, "HSVE000300"), "SIN RADICAR")
        # Carpeta con letras fuera de PENDIENTE.
        self.assertEqual(self.resultado(df, "HSVE000200"), "UBICACIÓN EQUIVOCADA")
        # Carpeta con letras dentro de PENDIENTE: estado azul con la descripción.
        self.assertEqual(self.resultado(df, "HSVE000600"), "PTE LAB")
        # Mismo número en dos carpetas.
        self.assertTrue(self.resultado(df, "HSVE000500").startswith("DUPLICADO (2): "))

    def test_reauditar_sin_duplicar_columnas(self):
        salida = Path(self.temp_dir.name) / "excel_reauditado_out.xlsx"
        self.run_audit("General", self.excel_path, salida)
        # Usar la salida como entrada y salida a la vez.
        df = self.run_audit("General", salida, salida)

        self.assertEqual(list(df.columns).count("RESULTADO_AUDITORIA"), 1, list(df.columns))


if __name__ == "__main__":
    unittest.main()
