import unittest

import wb_barcode_gui as app


class ResponsivePreviewTests(unittest.TestCase):
    def test_preview_uses_full_scale_when_space_is_available(self):
        self.assertEqual(app.preview_pixels_per_mm(None), 11)
        self.assertEqual(app.preview_pixels_per_mm(654), 11)

    def test_preview_scales_down_to_fit_narrow_pane(self):
        pane_width = 300
        scale = app.preview_pixels_per_mm(pane_width)
        rendered_width = app.LABEL_W_MM * scale
        self.assertLessEqual(rendered_width + 16, pane_width)

    def test_preview_never_scales_above_design_size(self):
        self.assertLessEqual(app.preview_pixels_per_mm(2000), 11)


if __name__ == "__main__":
    unittest.main()
