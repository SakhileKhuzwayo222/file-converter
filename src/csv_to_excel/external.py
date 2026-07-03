from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

from .converter import ConversionError, ensure_output_path


WORD_EXTENSIONS = (".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf")
EXCEL_EXTENSIONS = (".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm")
POWERPOINT_EXTENSIONS = (".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm", ".pot", ".potx", ".potm")
AUDIO_EXTENSIONS = (".mp3", ".wav", ".m4a", ".aac", ".flac", ".ogg", ".wma")
VIDEO_EXTENSIONS = (".mp4", ".mov", ".mkv", ".avi", ".webm", ".wmv", ".mpeg", ".mpg")


def check_source_file(source: Path, suffixes: tuple[str, ...], label: str) -> None:
    if not source.exists():
        raise ConversionError(f"{label} file not found: {source}")
    if not source.is_file():
        raise ConversionError(f"Input must be a {label} file: {source}")
    if source.suffix.lower() not in suffixes:
        expected = ", ".join(suffixes)
        raise ConversionError(f"Input file must end with {expected}: {source}")


def prepare_destination(source: Path, output_path: Path | str | None, output_extension: str, overwrite: bool) -> Path:
    destination = Path(output_path) if output_path else source.with_suffix(output_extension)
    ensure_output_path(destination, overwrite)
    if overwrite and destination.exists():
        destination.unlink()
    return destination


def powershell_quote(value: Path | str) -> str:
    return "'" + str(value).replace("'", "''") + "'"


def run_powershell(script: str, tool_label: str) -> None:
    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        raise ConversionError(f"{tool_label} conversion needs Windows PowerShell, but it was not found.")

    result = subprocess.run(
        [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        text=True,
    )
    if result.returncode:
        details = (result.stderr or result.stdout).strip()
        message = details if details else f"{tool_label} could not complete the conversion."
        raise ConversionError(message)


def convert_word_with_office(
    input_path: Path | str,
    output_path: Path | str | None,
    output_extension: str,
    overwrite: bool,
) -> Path:
    source = Path(input_path)
    check_source_file(source, WORD_EXTENSIONS, "Word")
    destination = prepare_destination(source, output_path, output_extension, overwrite)
    file_format = 17 if output_extension == ".pdf" else 16
    script = f"""
$ErrorActionPreference = 'Stop'
$inputPath = {powershell_quote(source)}
$outputPath = {powershell_quote(destination)}
$fileFormat = {file_format}
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {{
    $document = $word.Documents.Open($inputPath)
    try {{
        $document.SaveAs2([ref]$outputPath, [ref]$fileFormat)
    }} finally {{
        $document.Close($false)
    }}
}} finally {{
    $word.Quit()
}}
"""
    run_powershell(script, "Microsoft Word")
    return destination


def convert_excel_with_office(
    input_path: Path | str,
    output_path: Path | str | None,
    output_extension: str,
    overwrite: bool,
) -> Path:
    source = Path(input_path)
    check_source_file(source, EXCEL_EXTENSIONS, "Excel")
    destination = prepare_destination(source, output_path, output_extension, overwrite)
    script = f"""
$ErrorActionPreference = 'Stop'
$inputPath = {powershell_quote(source)}
$outputPath = {powershell_quote(destination)}
$outputExtension = {powershell_quote(output_extension)}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {{
    $workbook = $excel.Workbooks.Open($inputPath)
    try {{
        if ($outputExtension -eq '.pdf') {{
            $workbook.ExportAsFixedFormat(0, $outputPath)
        }} else {{
            $workbook.SaveAs($outputPath, 51)
        }}
    }} finally {{
        $workbook.Close($false)
    }}
}} finally {{
    $excel.Quit()
}}
"""
    run_powershell(script, "Microsoft Excel")
    return destination


def convert_powerpoint_with_office(
    input_path: Path | str,
    output_path: Path | str | None,
    output_extension: str,
    overwrite: bool,
) -> Path:
    source = Path(input_path)
    check_source_file(source, POWERPOINT_EXTENSIONS, "PowerPoint")
    destination = prepare_destination(source, output_path, output_extension, overwrite)
    file_format = 32 if output_extension == ".pdf" else 24
    script = f"""
$ErrorActionPreference = 'Stop'
$inputPath = {powershell_quote(source)}
$outputPath = {powershell_quote(destination)}
$fileFormat = {file_format}
$powerpoint = New-Object -ComObject PowerPoint.Application
try {{
    $presentation = $powerpoint.Presentations.Open($inputPath, $true, $false, $false)
    try {{
        $presentation.SaveAs($outputPath, $fileFormat)
    }} finally {{
        $presentation.Close()
    }}
}} finally {{
    $powerpoint.Quit()
}}
"""
    run_powershell(script, "Microsoft PowerPoint")
    return destination


def convert_media_with_ffmpeg(
    input_path: Path | str,
    output_path: Path | str | None,
    output_extension: str,
    valid_suffixes: tuple[str, ...],
    overwrite: bool,
) -> Path:
    source = Path(input_path)
    check_source_file(source, valid_suffixes, "media")
    destination = prepare_destination(source, output_path, output_extension, overwrite)
    ffmpeg = shutil.which("ffmpeg")
    if not ffmpeg:
        raise ConversionError(
            "Media conversion needs FFmpeg. Install FFmpeg and make sure ffmpeg.exe is available in PATH."
        )

    command = [ffmpeg, "-hide_banner", "-loglevel", "error", "-y", "-i", str(source), str(destination)]
    result = subprocess.run(
        command,
        capture_output=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        text=True,
    )
    if result.returncode:
        details = (result.stderr or result.stdout).strip()
        raise ConversionError(details or "FFmpeg could not complete the media conversion.")
    return destination
