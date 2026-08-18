import unittest

import wb_barcode_gui as app


class BrandingTests(unittest.TestCase):
    def test_general_product_name_covers_labels_and_barcodes(self):
        self.assertIn("этикеток", app.APP_NAME.lower())
        self.assertIn("штрихкоды", app.APP_NAME.lower())
        self.assertNotIn("wb", app.APP_NAME.lower())

    def test_application_icon_exists(self):
        self.assertTrue(app.resource_path(app.APP_ICON).is_file())


class MouseWheelTests(unittest.TestCase):
    def test_wheel_always_changes_numeric_value_by_tenth(self):
        self.assertEqual(app.numeric_wheel_value(36, 120), 36.1)
        self.assertEqual(app.numeric_wheel_value(36, -120), 35.9)
        self.assertEqual(app.numeric_wheel_value(7, 120), 7.1)


if __name__ == "__main__":
    unittest.main()
