# File Converter

A small Windows-friendly Python project for converting files into Microsoft Office formats.

The app includes a custom File Converter icon, a modern interface, and a Light/Dark theme toggle.
It also shows interaction popups, a step-by-step conversion timeline, a determinate progress bar, and a download-ready screen when conversion finishes.

## What It Converts

- CSV to Excel `.xlsx`
- TSV to Excel `.xlsx`
- JSON to Excel `.xlsx`
- PDF to Word `.docx`
- TXT or Markdown to Word `.docx`

PDF conversion extracts readable text. Scanned or image-only PDFs need OCR first.

## Setup

```powershell
cd "$env:USERPROFILE\Downloads\FILE CONVERTER"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .
```

The app uses Python's standard library. If you later install `pypdf`, PDF text extraction may work better on more complex PDFs, but it is not required for the app to start.

## Start The Window

Double-click:

```text
START FILE CONVERTER.bat
```

Then:

1. Pick the conversion type.
2. Choose a single file or a folder of files.
3. Click the plus box in **Upload**, then click **Submit**.
4. Choose a save folder.
5. Choose whether to rename the output file.
6. Click **Convert**.
7. When conversion completes, use the **Download Ready** window to open the file, open its folder, or start another conversion.

Use the theme button in the top-right corner to switch between Light and Dark mode.

The default output name uses the original filename and adds the output type, for example:

```text
Sales Report (Excel).xlsx
Contract (Word).docx
```

Use **Cancel** to stop a folder conversion between files. If a single file is already being written, it may finish before the cancellation can interrupt it.

## Build A Shareable EXE

Double-click:

```text
BUILD EXE.bat
```

The first build creates a local `.build-venv` folder and installs PyInstaller. When it finishes, share this file:

```text
dist\File Converter.exe
```

Notes:

- The `.exe` may take a few seconds to open because it is a one-file bundle.
- Windows may show a SmartScreen warning for unsigned apps. Choose **More info** and **Run anyway** if you trust the file.
- Some antivirus tools are cautious with newly built PyInstaller apps. Signing the app with a code-signing certificate is the professional fix.

## Command Line Examples

CSV to Excel:

```powershell
python -m csv_to_excel "data\sales.csv" -o "reports\sales.xlsx"
```

Folder of CSV files to Excel:

```powershell
python -m csv_to_excel "data" --batch -o "reports"
```

PDF to Word:

```powershell
python -m csv_to_excel.office "docs\contract.pdf" -o "docs\contract.docx"
```

TXT or Markdown to Word:

```powershell
python -m csv_to_excel.office "notes\summary.txt" -o "notes\summary.docx"
```

## Run Tests

```powershell
python -B -m unittest discover -s tests
```
