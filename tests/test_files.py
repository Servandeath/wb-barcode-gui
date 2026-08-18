import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook
from pypdf import PdfReader

import wb_barcode_gui as app


class FileTests(unittest.TestCase):
    def test_template_can_be_read(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "template.xlsx"
            app.make_template(str(path))
            rows = app.read_excel_rows(str(path))
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]["Баркод"], "2000000000244")
            self.assertEqual(load_workbook(path).active["F2"].number_format, "@")

    def test_pdf_page_size_and_count(self):
        data = app.render_pdf(
            [dict(app.App.TEST_ROW)], dict(app.DEFAULT_SETTINGS), app.register_pdf_font()
        )
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "labels.pdf"
            path.write_bytes(data)
            page = PdfReader(path).pages[0]
            self.assertAlmostEqual(float(page.mediabox.width), 58 * 72 / 25.4, places=2)
            self.assertAlmostEqual(float(page.mediabox.height), 40 * 72 / 25.4, places=2)

    def test_unique_path_does_not_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            original = Path(tmp) / "label.pdf"
            original.touch()
            self.assertEqual(app.unique_path(original).name, "label (1).pdf")


if __name__ == "__main__":
    unittest.main()
