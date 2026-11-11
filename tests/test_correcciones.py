import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

import auditor_facturas
from auditor_facturas import InvoiceAuditor


class InvoiceMatchingTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.root = Path(self.temp_dir.name)

    def build_index(self, ids):
        # build_native_index no usa estado de la ventana, se puede llamar sin Tk.
        return InvoiceAuditor.build_native_index(InvoiceAuditor, self.root, ids)

    def test_short_invoice_does_not_match_longer_number(self):
        (self.root / "CONSULTA EXTERNA" / "2127892").mkdir(parents=True)

        index = self.build_index({"212789", "2127892"})

        self.assertEqual(index["212789"], [])
        self.assertEqual(len(index["2127892"]), 1)

    def test_zip_with_nit_and_invoice_is_found(self):
        (self.root / "SANITAS").mkdir()
        (self.root / "SANITAS" / "800218979_2175737.zip").write_bytes(b"")

        index = self.build_index({"2175737"})

        self.assertEqual(index["2175737"], [self.root / "SANITAS" / "800218979_2175737.zip"])

    def test_pending_folder_with_text_is_found(self):
        (self.root / "PENDIENTE" / "2127892 FALTA FIRMA").mkdir(parents=True)

        index = self.build_index({"2127892"})

        self.assertEqual(len(index["2127892"]), 1)

    def test_name_has_invoice_ignores_leading_zeros(self):
        self.assertTrue(InvoiceAuditor._name_has_invoice("2127892_SOP_1", "02127892"))
        self.assertFalse(InvoiceAuditor._name_has_invoice("21278921_SOP_1", "2127892"))


SAMPLE_FEV = Path(__file__).resolve().parent.parent / "2197025" / "FEV_800218979_2197025.pdf"


@unittest.skipUnless(auditor_facturas.PyPDF2 and SAMPLE_FEV.exists(), "Requiere PyPDF2 y el FEV de ejemplo")
class FevInvoiceNumberTests(unittest.TestCase):
    """El número impreso en el FEV debe corresponder con el de la carpeta."""

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.root = Path(self.temp_dir.name)
        # Instancia sin ventana: estos métodos no usan estado de la interfaz.
        self.auditor = object.__new__(InvoiceAuditor)

    def make_folder(self, name):
        folder = self.root / name
        folder.mkdir()
        (folder / "FEV_800218979_2197025.pdf").write_bytes(SAMPLE_FEV.read_bytes())
        return folder

    def test_matching_folder_is_ok(self):
        folder = self.make_folder("800218979_2197025")
        self.assertIsNone(self.auditor._fev_invoice_mismatch(folder, "2197025"))

    def test_folder_with_other_invoice_is_reported(self):
        folder = self.make_folder("800218979_2197026")
        self.assertEqual(self.auditor._fev_invoice_mismatch(folder, "2197026"), "2197025")

    def test_zip_with_other_invoice_is_reported(self):
        zip_path = self.root / "800218979_2197026.zip"
        with zipfile.ZipFile(zip_path, "w") as archive:
            archive.write(SAMPLE_FEV, "FEV_800218979_2197026.pdf")
        self.assertEqual(self.auditor._fev_invoice_mismatch(zip_path, "2197026"), "2197025")

    def test_folder_without_fev_is_not_reported(self):
        folder = self.root / "800218979_2197025"
        folder.mkdir()
        self.assertIsNone(self.auditor._fev_invoice_mismatch(folder, "2197025"))


class DuplicateMessageTests(unittest.TestCase):
    def test_lists_relative_locations(self):
        root = Path("D:/Soportes")
        msg = InvoiceAuditor._duplicate_message(
            [root / "CONSULTA EXTERNA" / "2177633", root / "PENDIENTE" / "2177633 PTE"], root)
        self.assertEqual(
            msg,
            f"DUPLICADO (2): {Path('CONSULTA EXTERNA/2177633')} | {Path('PENDIENTE/2177633 PTE')}",
        )


class AppDirTests(unittest.TestCase):
    def test_frozen_app_uses_executable_folder(self):
        fake_exe = Path(tempfile.gettempdir()) / "dist" / "auditor.exe"
        with patch.object(sys, "frozen", True, create=True), \
                patch.object(sys, "executable", str(fake_exe)):
            self.assertEqual(auditor_facturas._app_dir(), fake_exe.resolve().parent)

    def test_script_uses_source_folder(self):
        self.assertEqual(auditor_facturas._app_dir(), Path(auditor_facturas.__file__).resolve().parent)


if __name__ == "__main__":
    unittest.main()
