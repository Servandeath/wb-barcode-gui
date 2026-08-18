# Windows build and release

Install build dependencies and create the standalone application:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements-build.txt
pyinstaller --noconfirm --clean --onefile --windowed `
  --name LabelMaster `
  --icon assets/app_icon.ico `
  --version-file build/version_info.txt `
  --add-data "assets;assets" `
  wb_barcode_gui.py
```

The executable is written to `dist/LabelMaster.exe`. Pushing a version tag such
as `v1.1.0` runs the same build on GitHub Actions and attaches the executable to
the corresponding GitHub Release.
