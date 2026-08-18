import unittest

import wb_barcode_gui as app


class BarcodeTests(unittest.TestCase):
    def test_checksum_and_normalization(self):
        self.assertEqual(app.ean13_checksum("460123456789"), "3")
        self.assertEqual(app.normalize_ean13("460123456789"), "4601234567893")

    def test_gtin_with_leading_zero_converts_to_ean13(self):
        self.assertTrue(app.is_gtin("04601234567893"))
        self.assertEqual(app.normalize_ean13("04601234567893"), "4601234567893")

    def test_invalid_values_are_rejected(self):
        for value in ("123", "4601234567894"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                app.normalize_ean13(value)

    def test_preview_encoding_has_standard_module_count(self):
        self.assertEqual(len(app.ean13_modules("4601234567893")), 95)


if __name__ == "__main__":
    unittest.main()
