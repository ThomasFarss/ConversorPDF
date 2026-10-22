name: Instalador Windows
on: [push, workflow_dispatch]
jobs:
  build:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: '3.12', cache: pip }
      - run: pip install -r requirements.txt pyinstaller
      - run: pyinstaller --noconfirm --onefile --windowed --name ConversorPDF --icon assets/icon.ico --add-data "assets/icon.ico;assets" --collect-all customtkinter app.py
      - run: iscc installer.iss
      - uses: actions/upload-artifact@v4
        with: { name: ConversorPDF_Setup, path: Output/ConversorPDF_Setup.exe }
