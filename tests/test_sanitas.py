import tempfile
import unittest
import zipfile
from pathlib import Path

from auditor_facturas import InvoiceAuditor


class SanitasValidationTests(unittest.TestCase):
    invoice_id = "2175737"

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.zip_path = Path(self.temp_dir.name) / f"800218979_{self.invoice_id}.zip"

    def write_zip(self, members):
        with zipfile.ZipFile(self.zip_path, "w") as archive:
            for name, contents in members:
                archive.writestr(name, contents)

    def test_flat_zip_with_only_pdfs_is_valid(self):
        self.write_zip([
            ("FEV_800218979_2175737.pdf", b"pdf"),
            ("DQX_800218979_2175737.pdf", b"pdf"),
        ])

        self.assertTrue(InvoiceAuditor._validate_sanitas_zip(self.zip_path, self.invoice_id))

    def test_zip_with_nested_pdf_is_invalid(self):
        self.write_zip([("SUBCARPETA/FEV_800218979_2175737.pdf", b"pdf")])

        self.assertFalse(InvoiceAuditor._validate_sanitas_zip(self.zip_path, self.invoice_id))

    def test_zip_with_directory_entry_is_invalid(self):
        self.write_zip([
            ("FEV_800218979_2175737.pdf", b"pdf"),
            ("SUBCARPETA/", b""),
        ])

        self.assertFalse(InvoiceAuditor._validate_sanitas_zip(self.zip_path, self.invoice_id))

    def test_zip_with_non_pdf_file_is_invalid(self):
        self.write_zip([
            ("FEV_800218979_2175737.pdf", b"pdf"),
            ("factura.xml", b"xml"),
        ])

        self.assertFalse(InvoiceAuditor._validate_sanitas_zip(self.zip_path, self.invoice_id))

    def test_zip_name_must_match_invoice(self):
        self.write_zip([("FEV_800218979_2175737.pdf", b"pdf")])

        self.assertFalse(InvoiceAuditor._validate_sanitas_zip(self.zip_path, "2175738"))

    def test_zip_must_be_named_with_nit_and_invoice(self):
        self.write_zip([("FEV_800218979_2175737.pdf", b"pdf")])
        incorrectly_named_path = self.zip_path.with_name(f"{self.invoice_id}.zip")
        self.zip_path.rename(incorrectly_named_path)

        self.assertFalse(
            InvoiceAuditor._validate_sanitas_zip(incorrectly_named_path, self.invoice_id)
        )

    def test_location_must_be_under_sanitas_consulta_externa(self):
        valid_path = Path("D:/Hospital/SANITAS/CONSULTA EXTERNA/SUBSIDIADO/800218979_2175737.zip")
        invalid_path = Path("D:/Hospital/COLSANITAS/CONSULTA EXTERNA/SUBSIDIADO/800218979_2175737.zip")

        self.assertFalse(InvoiceAuditor._is_wrong_sanitas_location(valid_path))
        self.assertTrue(InvoiceAuditor._is_wrong_sanitas_location(invalid_path))


if __name__ == "__main__":
    unittest.main()
