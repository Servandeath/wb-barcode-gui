# Testing

Create the local environment and run the test suite on Windows:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements-dev.txt
python -m unittest discover -s tests -v
```

The suite covers EAN-13 validation, GTIN conversion, Excel input and exclusions,
unique output paths, template generation, and the physical PDF page size.

GitHub Actions runs the same suite on Python 3.11 and 3.13 for every push and
pull request.
