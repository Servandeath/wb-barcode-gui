import json
import tempfile
import unittest
from pathlib import Path

import wb_barcode_gui as app


class SettingsTests(unittest.TestCase):
    def setUp(self):
        self.original_settings = app.SETTINGS_FILE
        self.original_legacy = app.LEGACY_SETTINGS_FILE

    def tearDown(self):
        app.SETTINGS_FILE = self.original_settings
        app.LEGACY_SETTINGS_FILE = self.original_legacy

    def test_save_creates_user_settings_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            app.SETTINGS_FILE = Path(tmp) / "nested" / "settings.json"
            app.LEGACY_SETTINGS_FILE = Path(tmp) / "legacy.json"
            app.save_settings({"font_size": 8})
            self.assertTrue(app.SETTINGS_FILE.is_file())
            self.assertEqual(app.load_settings()["font_size"], 8)

    def test_legacy_settings_are_loaded_when_user_file_is_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            app.SETTINGS_FILE = Path(tmp) / "new" / "settings.json"
            app.LEGACY_SETTINGS_FILE = Path(tmp) / "legacy.json"
            app.LEGACY_SETTINGS_FILE.write_text(
                json.dumps({"font_size": 9}), encoding="utf-8"
            )
            self.assertEqual(app.load_settings()["font_size"], 9)


if __name__ == "__main__":
    unittest.main()
