# File Converter

A small Windows-friendly Python project for converting files into Microsoft Office formats.

The app includes a custom File Converter icon, icon-led controls, home-screen navigation, a modern interface, and a Light/Dark theme toggle.
It also shows interaction popups, a step-by-step conversion timeline, a determinate progress bar, and a download-ready screen when conversion finishes.
It includes the FC Text Editor for text, data, web, code, script, and config files before conversion.

## What It Converts

- CSV to Excel `.xlsx`
- TSV to Excel `.xlsx`
- JSON to Excel `.xlsx`
- PDF to Word `.docx`
- PDF to HTML `.html`
- EPUB to PDF `.pdf`
- ZIP archive extraction to a folder
- Folder compression to ZIP `.zip`
- Editable text/code/data files to Word `.docx`
- Word files `.doc`, `.docx`, `.docm`, `.dot`, `.dotx`, `.dotm`, `.rtf` to PDF or `.docx`
- Excel files `.xls`, `.xlsx`, `.xlsm`, `.xlsb`, `.xlt`, `.xltx`, `.xltm` to PDF or `.xlsx`
- PowerPoint files `.ppt`, `.pptx`, `.pptm`, `.pps`, `.ppsx`, `.ppsm`, `.pot`, `.potx`, `.potm` to PDF or `.pptx`
- Audio files `.mp3`, `.wav`, `.m4a`, `.aac`, `.flac`, `.ogg`, `.wma` to MP3, WAV, M4A, FLAC, or OGG
- Video files `.mp4`, `.mov`, `.mkv`, `.avi`, `.webm`, `.wmv`, `.mpeg`, `.mpg` to MP4, MOV, MKV, WEBM, or MP3 audio

PDF to Word extracts readable text. PDF to HTML preserves the original PDF layout by creating an HTML viewer and copying the source PDF into a matching assets folder, so formatting, spacing, fonts, icons, and images stay visually faithful to the uploaded document.
Scanned PDFs without selectable text may still need OCR first before PDF to Word can produce editable text.
EPUB conversion extracts readable chapter text and writes it into a simple PDF.
Office-to-PDF and legacy Office conversions need Microsoft Office installed on Windows.
Audio and video conversions need FFmpeg installed and available in `PATH`.

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

1. Pick **Convert from** and **Convert to**.
2. Choose a single file or a folder of files.
3. Click the plus box under the file type selection, then click **Submit**.
4. Choose a save folder.
5. Use **FC Text Editor** in the navigation to edit the current upload, or choose another editable file.
6. Choose whether to rename the output file.
7. Click **Convert**.
8. When conversion completes, use the **Download Ready** window to open the file, open its folder, or start another conversion.

Use the theme button in the top-right corner to switch between Light and Dark mode.
The editor toolbar can insert linked media references, code blocks, validation snippets, and security snippets into editable text-based files.
Binary files such as PDFs, Office files, and media files are converted by the converter, but they are not rewritten directly inside the text editor.

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
- Media conversion works on computers that have FFmpeg installed.
- Office-to-PDF and legacy Office conversions work on computers that have Microsoft Office installed.

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

PDF to HTML:

```powershell
python -m csv_to_excel.office "docs\contract.pdf" -o "docs\contract.html"
```

TXT or Markdown to Word:

```powershell
python -m csv_to_excel.office "notes\summary.txt" -o "notes\summary.docx"
```

## Run Tests

```powershell
python -B -m unittest discover -s tests
```
