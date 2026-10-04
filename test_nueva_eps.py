import unittest
from pathlib import Path

from auditor_facturas import InvoiceAuditor


class NuevaEpsValidationTests(unittest.TestCase):
    invoice_id = "2197025"
    nit = "800218979"

    def test_required_supports_and_four_counted_pdfs_are_valid(self):
        result = InvoiceAuditor._validate_nueva_eps_pdfs(
            [
                f"FEV_{self.nit}_{self.invoice_id}.pdf",
                f"OPF_{self.nit}_{self.invoice_id}.pdf",
                f"PDE_{self.nit}_{self.invoice_id}.pdf",
                f"HEV_{self.nit}_{self.invoice_id}.pdf",
                f"LDP_{self.nit}_{self.invoice_id}.pdf",
            ],
            self.invoice_id,
        )

        self.assertEqual(result, (4, 0, set(), 4))

    def test_ldp_is_required_but_does_not_count_toward_four(self):
        result = InvoiceAuditor._validate_nueva_eps_pdfs(
            [
                f"FEV_{self.nit}_{self.invoice_id}.pdf",
                f"OPF_{self.nit}_{self.invoice_id}.pdf",
                f"PDE_{self.nit}_{self.invoice_id}.pdf",
                f"LDP_{self.nit}_{self.invoice_id}.pdf",
            ],
            self.invoice_id,
        )

        self.assertEqual(result, (3, 0, set(), 4))

    def test_missing_required_type_is_reported(self):
        result = InvoiceAuditor._validate_nueva_eps_pdfs(
            [
                f"FEV_{self.nit}_{self.invoice_id}.pdf",
                f"PDE_{self.nit}_{self.invoice_id}.pdf",
                f"HEV_{self.nit}_{self.invoice_id}.pdf",
                f"PDX_{self.nit}_{self.invoice_id}.pdf",
                f"RAN_{self.nit}_{self.invoice_id}.pdf",
                f"LDP_{self.nit}_{self.invoice_id}.pdf",
            ],
            self.invoice_id,
        )

        self.assertEqual(result, (5, 0, {"OPF"}, 4))

    def test_unlisted_support_type_is_invalid(self):
        result = InvoiceAuditor._validate_nueva_eps_pdfs(
            [
                f"FEV_{self.nit}_{self.invoice_id}.pdf",
                f"XYZ_{self.nit}_{self.invoice_id}.pdf",
            ],
            self.invoice_id,
        )

        self.assertEqual(result, (1, 1, {"OPF", "PDE", "LDP"}, 4))

    def test_wrong_invoice_number_is_invalid(self):
        result = InvoiceAuditor._validate_nueva_eps_pdfs(
            [
                f"FEV_{self.nit}_9999999.pdf",
                f"OPF_{self.nit}_{self.invoice_id}.pdf",
                f"PDE_{self.nit}_{self.invoice_id}.pdf",
                f"HEV_{self.nit}_{self.invoice_id}.pdf",
                f"LDP_{self.nit}_{self.invoice_id}.pdf",
            ],
            self.invoice_id,
        )

        self.assertEqual(result, (3, 1, {"FEV"}, 4))

    def test_duplicate_support_is_invalid_and_does_not_count_twice(self):
        filename = f"FEV_{self.nit}_{self.invoice_id}.pdf"
        result = InvoiceAuditor._validate_nueva_eps_pdfs(
            [
                filename,
                filename,
                f"OPF_{self.nit}_{self.invoice_id}.pdf",
                f"PDE_{self.nit}_{self.invoice_id}.pdf",
                f"LDP_{self.nit}_{self.invoice_id}.pdf",
            ],
            self.invoice_id,
        )

        self.assertEqual(result, (3, 1, set(), 4))

    def test_pending_folder_in_pending_location_uses_pending_rule(self):
        pending_path = Path("C:/Soportes/PENDIENTE/2175724 PTE LAB")

        self.assertFalse(InvoiceAuditor._is_wrong_invoice_location(pending_path))
        self.assertEqual(
            InvoiceAuditor._pending_invoice_description(pending_path, "2175724"),
            "PTE LAB",
        )

    def test_pending_named_folder_outside_pending_location_is_wrong(self):
        misplaced_path = Path("C:/Soportes/CONSULTA EXTERNA/2175724 PTE LAB")

        self.assertTrue(InvoiceAuditor._is_wrong_invoice_location(misplaced_path))

    def test_regular_invoice_folder_requires_consulta_externa_location(self):
        correct_path = Path("C:/Soportes/CONSULTA EXTERNA/2175724")
        misplaced_path = Path("C:/Soportes/PENDIENTE/2175724")

        self.assertFalse(InvoiceAuditor._is_wrong_invoice_location(correct_path))
        self.assertTrue(InvoiceAuditor._is_wrong_invoice_location(misplaced_path))

    def test_duplicate_exact_invoice_folders_are_detected_as_ambiguous(self):
        first_copy = Path("C:/Soportes/CONSULTA EXTERNA/2177633")
        second_copy = Path("D:/Hospital/Nueva carpeta/6-JUN/2177633")
        unrelated_prefix = Path("C:/Soportes/CONSULTA EXTERNA/2177633 PTE")

        exact_matches = InvoiceAuditor._get_exact_invoice_folders(
            [first_copy, second_copy, unrelated_prefix],
            "2177633",
        )

        self.assertEqual(exact_matches, [first_copy, second_copy])


if __name__ == "__main__":
    unittest.main()
