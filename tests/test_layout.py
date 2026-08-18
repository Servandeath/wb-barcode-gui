import unittest

import wb_barcode_gui as app


class LayoutTests(unittest.TestCase):
    def test_action_buttons_reflow_at_panel_breakpoints(self):
        self.assertEqual(app.action_button_columns(900), 5)
        self.assertEqual(app.action_button_columns(600), 3)
        self.assertEqual(app.action_button_columns(400), 2)

    def test_preview_safe_area_is_inside_page_boundary(self):
        self.assertGreater(app.PREVIEW_SAFE_MARGIN_MM, 0)
        self.assertLess(app.PREVIEW_SAFE_MARGIN_MM * 2, app.LABEL_H_MM)


if __name__ == "__main__":
    unittest.main()
