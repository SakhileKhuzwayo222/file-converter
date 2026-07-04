from __future__ import annotations

import base64
import html
import math
import os
import struct
import threading
import tkinter as tk
import zlib
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .archive import ZIP_EXTENSIONS, create_zip_archive, extract_zip_archive
from .converter import (
    ConversionError,
    convert_csv_to_excel,
    convert_json_to_excel,
    convert_tsv_to_excel,
    files_in_directory,
)
from .external import (
    AUDIO_EXTENSIONS,
    EXCEL_EXTENSIONS,
    POWERPOINT_EXTENSIONS,
    VIDEO_EXTENSIONS,
    WORD_EXTENSIONS,
    convert_excel_with_office,
    convert_media_with_ffmpeg,
    convert_powerpoint_with_office,
    convert_word_with_office,
)
from .office import EPUB_EXTENSIONS, TEXT_DOCUMENT_EXTENSIONS, convert_epub_to_pdf, convert_pdf_to_word, convert_text_to_word


ConverterFunction = Callable[[Path, Path | None, str, str | None, bool], Path]
INVALID_FILENAME_CHARS = '<>:"/\\|?*'
EDITABLE_FORMAT_GROUPS = (
    ("Text documents", (".txt", ".text", ".md", ".markdown", ".rst", ".log")),
    ("Data files", (".csv", ".tsv", ".json", ".jsonl", ".xml", ".yaml", ".yml", ".toml")),
    ("Config files", (".ini", ".cfg", ".conf", ".properties")),
    ("Web files", (".html", ".htm", ".css", ".js", ".ts", ".jsx", ".tsx")),
    ("Code and scripts", (".py", ".java", ".cs", ".c", ".cpp", ".h", ".hpp", ".php", ".rb", ".go", ".rs")),
    ("Shell and query files", (".swift", ".kt", ".kts", ".sql", ".ps1", ".bat", ".cmd", ".sh")),
)
TEXT_EDIT_EXTENSIONS = TEXT_DOCUMENT_EXTENSIONS
TEXT_EDIT_SUMMARY = "text, Markdown, data, web, code, script, and config files"


@dataclass(frozen=True)
class AppTheme:
    background: str
    panel: str
    panel_alt: str
    entry: str
    text: str
    muted: str
    border: str
    accent: str
    accent_hover: str
    accent_text: str
    danger: str
    danger_hover: str
    log_background: str
    log_text: str
    selection: str


THEMES = {
    "light": AppTheme(
        background="#f4f7fb",
        panel="#ffffff",
        panel_alt="#f0fdfa",
        entry="#ffffff",
        text="#0f172a",
        muted="#64748b",
        border="#d9e4ec",
        accent="#0d9488",
        accent_hover="#115e59",
        accent_text="#ffffff",
        danger="#e11d48",
        danger_hover="#be123c",
        log_background="#111827",
        log_text="#dbeafe",
        selection="#ccfbf1",
    ),
    "dark": AppTheme(
        background="#121212",
        panel="#1e2027",
        panel_alt="#182d2b",
        entry="#151820",
        text="#f8fafc",
        muted="#a3aebc",
        border="#343945",
        accent="#2dd4bf",
        accent_hover="#5eead4",
        accent_text="#06201d",
        danger="#fb7185",
        danger_hover="#fda4af",
        log_background="#090b10",
        log_text="#c7d2fe",
        selection="#164e63",
    ),
}


@dataclass(frozen=True)
class ConversionSpec:
    from_label: str
    to_label: str
    input_label: str
    input_extensions: tuple[str, ...]
    output_extension: str
    output_description: str
    output_tag: str
    supports_encoding: bool
    supports_delimiter: bool
    convert: ConverterFunction
    source_kind: str = "file"

    @property
    def name(self) -> str:
        return f"{self.from_label} to {self.to_label}"


def csv_to_excel(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_csv_to_excel(
        input_path,
        output_path,
        encoding=encoding,
        delimiter=delimiter,
        overwrite=overwrite,
    )


def tsv_to_excel(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_tsv_to_excel(input_path, output_path, encoding=encoding, overwrite=overwrite)


def json_to_excel(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_json_to_excel(input_path, output_path, encoding=encoding, overwrite=overwrite)


def pdf_to_word(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_pdf_to_word(input_path, output_path, overwrite=overwrite)


def text_to_word(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_text_to_word(input_path, output_path, encoding=encoding, overwrite=overwrite)


def epub_to_pdf(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_epub_to_pdf(input_path, output_path, overwrite=overwrite)


def zip_to_folder(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return extract_zip_archive(input_path, output_path, overwrite=overwrite)


def folder_to_zip(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return create_zip_archive(input_path, output_path, overwrite=overwrite)


def word_to_pdf(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_word_with_office(input_path, output_path, ".pdf", overwrite)


def word_to_docx(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_word_with_office(input_path, output_path, ".docx", overwrite)


def excel_to_pdf(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_excel_with_office(input_path, output_path, ".pdf", overwrite)


def excel_to_xlsx(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_excel_with_office(input_path, output_path, ".xlsx", overwrite)


def powerpoint_to_pdf(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_powerpoint_with_office(input_path, output_path, ".pdf", overwrite)


def powerpoint_to_pptx(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_powerpoint_with_office(input_path, output_path, ".pptx", overwrite)


def audio_to_mp3(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_media_with_ffmpeg(input_path, output_path, ".mp3", AUDIO_EXTENSIONS, overwrite)


def audio_to_wav(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_media_with_ffmpeg(input_path, output_path, ".wav", AUDIO_EXTENSIONS, overwrite)


def audio_to_m4a(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_media_with_ffmpeg(input_path, output_path, ".m4a", AUDIO_EXTENSIONS, overwrite)


def audio_to_flac(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_media_with_ffmpeg(input_path, output_path, ".flac", AUDIO_EXTENSIONS, overwrite)


def audio_to_ogg(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_media_with_ffmpeg(input_path, output_path, ".ogg", AUDIO_EXTENSIONS, overwrite)


def video_to_mp4(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_media_with_ffmpeg(input_path, output_path, ".mp4", VIDEO_EXTENSIONS, overwrite)


def video_to_mov(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_media_with_ffmpeg(input_path, output_path, ".mov", VIDEO_EXTENSIONS, overwrite)


def video_to_mkv(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_media_with_ffmpeg(input_path, output_path, ".mkv", VIDEO_EXTENSIONS, overwrite)


def video_to_webm(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_media_with_ffmpeg(input_path, output_path, ".webm", VIDEO_EXTENSIONS, overwrite)


def video_to_mp3(input_path: Path, output_path: Path | None, encoding: str, delimiter: str | None, overwrite: bool) -> Path:
    return convert_media_with_ffmpeg(input_path, output_path, ".mp3", VIDEO_EXTENSIONS, overwrite)


CONVERSIONS = [
    ConversionSpec(
        from_label="CSV (.csv)",
        to_label="Excel Workbook (.xlsx)",
        input_label="CSV",
        input_extensions=(".csv",),
        output_extension=".xlsx",
        output_description="Excel workbook",
        output_tag="Excel",
        supports_encoding=True,
        supports_delimiter=True,
        convert=csv_to_excel,
    ),
    ConversionSpec(
        from_label="TSV (.tsv)",
        to_label="Excel Workbook (.xlsx)",
        input_label="TSV",
        input_extensions=(".tsv",),
        output_extension=".xlsx",
        output_description="Excel workbook",
        output_tag="Excel",
        supports_encoding=True,
        supports_delimiter=False,
        convert=tsv_to_excel,
    ),
    ConversionSpec(
        from_label="JSON (.json)",
        to_label="Excel Workbook (.xlsx)",
        input_label="JSON",
        input_extensions=(".json",),
        output_extension=".xlsx",
        output_description="Excel workbook",
        output_tag="Excel",
        supports_encoding=True,
        supports_delimiter=False,
        convert=json_to_excel,
    ),
    ConversionSpec(
        from_label="PDF (.pdf)",
        to_label="Word Document (.docx)",
        input_label="PDF",
        input_extensions=(".pdf",),
        output_extension=".docx",
        output_description="Word document",
        output_tag="Word",
        supports_encoding=False,
        supports_delimiter=False,
        convert=pdf_to_word,
    ),
    ConversionSpec(
        from_label="EPUB (.epub)",
        to_label="PDF Document (.pdf)",
        input_label="EPUB",
        input_extensions=EPUB_EXTENSIONS,
        output_extension=".pdf",
        output_description="PDF document",
        output_tag="PDF",
        supports_encoding=False,
        supports_delimiter=False,
        convert=epub_to_pdf,
    ),
    ConversionSpec(
        from_label="ZIP archive (.zip)",
        to_label="Extracted folder",
        input_label="ZIP",
        input_extensions=ZIP_EXTENSIONS,
        output_extension="",
        output_description="extracted folder",
        output_tag="Extracted",
        supports_encoding=False,
        supports_delimiter=False,
        convert=zip_to_folder,
    ),
    ConversionSpec(
        from_label="Folder",
        to_label="ZIP Archive (.zip)",
        input_label="folder",
        input_extensions=(),
        output_extension=".zip",
        output_description="ZIP archive",
        output_tag="ZIP",
        supports_encoding=False,
        supports_delimiter=False,
        convert=folder_to_zip,
        source_kind="folder",
    ),
    ConversionSpec(
        from_label="Editable text files",
        to_label="Word Document (.docx)",
        input_label="editable text",
        input_extensions=TEXT_DOCUMENT_EXTENSIONS,
        output_extension=".docx",
        output_description="Word document",
        output_tag="Word",
        supports_encoding=True,
        supports_delimiter=False,
        convert=text_to_word,
    ),
    ConversionSpec(
        from_label="Word files (.docx, .doc, .docm, .rtf)",
        to_label="PDF Document (.pdf)",
        input_label="Word",
        input_extensions=WORD_EXTENSIONS,
        output_extension=".pdf",
        output_description="PDF document",
        output_tag="PDF",
        supports_encoding=False,
        supports_delimiter=False,
        convert=word_to_pdf,
    ),
    ConversionSpec(
        from_label="Word files (.docx, .doc, .docm, .rtf)",
        to_label="Word Document (.docx)",
        input_label="Word",
        input_extensions=WORD_EXTENSIONS,
        output_extension=".docx",
        output_description="Word document",
        output_tag="Word",
        supports_encoding=False,
        supports_delimiter=False,
        convert=word_to_docx,
    ),
    ConversionSpec(
        from_label="Excel files (.xlsx, .xls, .xlsm, .xlsb)",
        to_label="PDF Document (.pdf)",
        input_label="Excel",
        input_extensions=EXCEL_EXTENSIONS,
        output_extension=".pdf",
        output_description="PDF document",
        output_tag="PDF",
        supports_encoding=False,
        supports_delimiter=False,
        convert=excel_to_pdf,
    ),
    ConversionSpec(
        from_label="Excel files (.xlsx, .xls, .xlsm, .xlsb)",
        to_label="Excel Workbook (.xlsx)",
        input_label="Excel",
        input_extensions=EXCEL_EXTENSIONS,
        output_extension=".xlsx",
        output_description="Excel workbook",
        output_tag="Excel",
        supports_encoding=False,
        supports_delimiter=False,
        convert=excel_to_xlsx,
    ),
    ConversionSpec(
        from_label="PowerPoint files (.pptx, .ppt, .pptm, .ppsx)",
        to_label="PDF Document (.pdf)",
        input_label="PowerPoint",
        input_extensions=POWERPOINT_EXTENSIONS,
        output_extension=".pdf",
        output_description="PDF document",
        output_tag="PDF",
        supports_encoding=False,
        supports_delimiter=False,
        convert=powerpoint_to_pdf,
    ),
    ConversionSpec(
        from_label="PowerPoint files (.pptx, .ppt, .pptm, .ppsx)",
        to_label="PowerPoint Presentation (.pptx)",
        input_label="PowerPoint",
        input_extensions=POWERPOINT_EXTENSIONS,
        output_extension=".pptx",
        output_description="PowerPoint presentation",
        output_tag="PowerPoint",
        supports_encoding=False,
        supports_delimiter=False,
        convert=powerpoint_to_pptx,
    ),
    ConversionSpec(
        from_label="Audio files (.mp3, .wav, .m4a, .flac)",
        to_label="MP3 Audio (.mp3)",
        input_label="audio",
        input_extensions=AUDIO_EXTENSIONS,
        output_extension=".mp3",
        output_description="MP3 audio",
        output_tag="MP3",
        supports_encoding=False,
        supports_delimiter=False,
        convert=audio_to_mp3,
    ),
    ConversionSpec(
        from_label="Audio files (.mp3, .wav, .m4a, .flac)",
        to_label="WAV Audio (.wav)",
        input_label="audio",
        input_extensions=AUDIO_EXTENSIONS,
        output_extension=".wav",
        output_description="WAV audio",
        output_tag="WAV",
        supports_encoding=False,
        supports_delimiter=False,
        convert=audio_to_wav,
    ),
    ConversionSpec(
        from_label="Audio files (.mp3, .wav, .m4a, .flac)",
        to_label="M4A Audio (.m4a)",
        input_label="audio",
        input_extensions=AUDIO_EXTENSIONS,
        output_extension=".m4a",
        output_description="M4A audio",
        output_tag="M4A",
        supports_encoding=False,
        supports_delimiter=False,
        convert=audio_to_m4a,
    ),
    ConversionSpec(
        from_label="Audio files (.mp3, .wav, .m4a, .flac)",
        to_label="FLAC Audio (.flac)",
        input_label="audio",
        input_extensions=AUDIO_EXTENSIONS,
        output_extension=".flac",
        output_description="FLAC audio",
        output_tag="FLAC",
        supports_encoding=False,
        supports_delimiter=False,
        convert=audio_to_flac,
    ),
    ConversionSpec(
        from_label="Audio files (.mp3, .wav, .m4a, .flac)",
        to_label="OGG Audio (.ogg)",
        input_label="audio",
        input_extensions=AUDIO_EXTENSIONS,
        output_extension=".ogg",
        output_description="OGG audio",
        output_tag="OGG",
        supports_encoding=False,
        supports_delimiter=False,
        convert=audio_to_ogg,
    ),
    ConversionSpec(
        from_label="Video files (.mp4, .mov, .mkv, .avi)",
        to_label="MP4 Video (.mp4)",
        input_label="video",
        input_extensions=VIDEO_EXTENSIONS,
        output_extension=".mp4",
        output_description="MP4 video",
        output_tag="MP4",
        supports_encoding=False,
        supports_delimiter=False,
        convert=video_to_mp4,
    ),
    ConversionSpec(
        from_label="Video files (.mp4, .mov, .mkv, .avi)",
        to_label="MOV Video (.mov)",
        input_label="video",
        input_extensions=VIDEO_EXTENSIONS,
        output_extension=".mov",
        output_description="MOV video",
        output_tag="MOV",
        supports_encoding=False,
        supports_delimiter=False,
        convert=video_to_mov,
    ),
    ConversionSpec(
        from_label="Video files (.mp4, .mov, .mkv, .avi)",
        to_label="MKV Video (.mkv)",
        input_label="video",
        input_extensions=VIDEO_EXTENSIONS,
        output_extension=".mkv",
        output_description="MKV video",
        output_tag="MKV",
        supports_encoding=False,
        supports_delimiter=False,
        convert=video_to_mkv,
    ),
    ConversionSpec(
        from_label="Video files (.mp4, .mov, .mkv, .avi)",
        to_label="WEBM Video (.webm)",
        input_label="video",
        input_extensions=VIDEO_EXTENSIONS,
        output_extension=".webm",
        output_description="WEBM video",
        output_tag="WEBM",
        supports_encoding=False,
        supports_delimiter=False,
        convert=video_to_webm,
    ),
    ConversionSpec(
        from_label="Video files (.mp4, .mov, .mkv, .avi)",
        to_label="MP3 Audio (.mp3)",
        input_label="video",
        input_extensions=VIDEO_EXTENSIONS,
        output_extension=".mp3",
        output_description="MP3 audio",
        output_tag="MP3",
        supports_encoding=False,
        supports_delimiter=False,
        convert=video_to_mp3,
    ),
]

FROM_FORMATS = list(dict.fromkeys(conversion.from_label for conversion in CONVERSIONS))
SUPPORTED_OPEN_EXTENSIONS = tuple(
    sorted(
        {
            extension
            for conversion in CONVERSIONS
            for extension in (*conversion.input_extensions, conversion.output_extension)
            if extension.startswith(".")
        }
    )
)


FC_ICON_SOURCE = (
    ("rounded_rect", 2.5, 2.0, 19.0, 20.0, 3.2, "fc_dark"),
    ("rounded_rect", 2.0, 2.0, 18.8, 18.8, 3.0, "accent"),
    ("fill_rect", 17.2, 6.0, 3.6, 14.8, "fc_shadow"),
    ("fill_rect", 4.7, 17.2, 14.3, 3.6, "fc_dark"),
    ("fill_rect", 6.0, 6.0, 2.0, 12.0, "white"),
    ("fill_rect", 6.0, 6.0, 7.4, 2.0, "white"),
    ("fill_rect", 6.0, 11.0, 6.0, 2.0, "white"),
    ("fill_rect", 14.6, 6.0, 5.0, 2.0, "white"),
    ("fill_rect", 12.6, 8.0, 2.0, 8.0, "white"),
    ("fill_rect", 14.6, 16.0, 5.0, 2.0, "white"),
)


ICON_SOURCES = {
    "home": (
        ("polyline", ((3.5, 10.5), (12.0, 4.0), (20.5, 10.5)), 1.8, "ink"),
        ("line", 5.5, 10.0, 5.5, 20.0, 1.8, "ink"),
        ("line", 18.5, 10.0, 18.5, 20.0, 1.8, "ink"),
        ("line", 5.5, 20.0, 10.0, 20.0, 1.8, "ink"),
        ("line", 14.0, 20.0, 18.5, 20.0, 1.8, "ink"),
        ("rect", 10.0, 14.0, 4.0, 6.0, 1.7, "ink"),
    ),
    "convert": (
        ("line", 5.0, 7.0, 18.0, 7.0, 1.8, "ink"),
        ("polyline", ((15.0, 4.0), (18.5, 7.0), (15.0, 10.0)), 1.8, "ink"),
        ("line", 19.0, 17.0, 6.0, 17.0, 1.8, "ink"),
        ("polyline", ((9.0, 14.0), (5.5, 17.0), (9.0, 20.0)), 1.8, "ink"),
    ),
    "open": (
        ("polyline", ((7.0, 4.0), (14.0, 4.0), (18.0, 8.0), (18.0, 20.0), (7.0, 20.0), (7.0, 4.0)), 1.7, "ink"),
        ("polyline", ((14.0, 4.0), (14.0, 8.0), (18.0, 8.0)), 1.5, "muted"),
        ("line", 10.0, 14.0, 19.0, 5.0, 1.8, "accent"),
        ("polyline", ((14.7, 5.0), (19.0, 5.0), (19.0, 9.3)), 1.8, "accent"),
    ),
    "edit": (
        ("polygon", ((5.0, 17.2), (6.7, 20.0), (9.8, 18.6), (18.0, 10.4), (13.4, 5.8)), "accent"),
        ("line", 13.4, 5.8, 18.0, 10.4, 1.6, "ink"),
        ("line", 6.4, 17.0, 8.9, 19.5, 1.4, "ink"),
        ("line", 4.8, 20.2, 9.0, 18.8, 1.5, "ink"),
        ("line", 15.2, 4.0, 20.0, 8.8, 2.2, "muted"),
    ),
    "download": (
        ("line", 12.0, 4.0, 12.0, 14.5, 1.9, "ink"),
        ("polyline", ((7.5, 10.5), (12.0, 15.0), (16.5, 10.5)), 1.9, "ink"),
        ("polyline", ((5.0, 17.0), (5.0, 20.0), (19.0, 20.0), (19.0, 17.0)), 1.8, "ink"),
    ),
    "upload": (
        ("line", 12.0, 15.0, 12.0, 4.5, 1.9, "ink"),
        ("polyline", ((7.5, 8.5), (12.0, 4.0), (16.5, 8.5)), 1.9, "ink"),
        ("polyline", ((5.0, 17.0), (5.0, 20.0), (19.0, 20.0), (19.0, 17.0)), 1.8, "ink"),
    ),
    "folder": (
        ("polyline", ((3.5, 7.0), (9.0, 7.0), (10.8, 9.0), (20.5, 9.0), (20.5, 19.5), (3.5, 19.5), (3.5, 7.0)), 1.7, "ink"),
        ("line", 3.5, 10.5, 20.5, 10.5, 1.4, "accent"),
    ),
    "file": (
        ("polyline", ((7.0, 3.5), (15.0, 3.5), (19.0, 7.5), (19.0, 20.5), (7.0, 20.5), (7.0, 3.5)), 1.7, "ink"),
        ("polyline", ((15.0, 3.5), (15.0, 8.0), (19.0, 8.0)), 1.4, "muted"),
        ("line", 9.5, 11.0, 16.5, 11.0, 1.3, "accent"),
        ("line", 9.5, 14.0, 16.5, 14.0, 1.3, "accent"),
        ("line", 9.5, 17.0, 14.0, 17.0, 1.3, "accent"),
    ),
    "document": (
        ("polyline", ((7.0, 3.5), (15.0, 3.5), (19.0, 7.5), (19.0, 20.5), (7.0, 20.5), (7.0, 3.5)), 1.7, "blue"),
        ("polyline", ((15.0, 3.5), (15.0, 8.0), (19.0, 8.0)), 1.4, "blue"),
        ("line", 9.5, 11.0, 16.5, 11.0, 1.3, "ink"),
        ("line", 9.5, 14.0, 16.5, 14.0, 1.3, "ink"),
        ("line", 9.5, 17.0, 14.0, 17.0, 1.3, "ink"),
    ),
    "data": (
        ("polyline", ((7.0, 3.5), (15.0, 3.5), (19.0, 7.5), (19.0, 20.5), (7.0, 20.5), (7.0, 3.5)), 1.7, "green"),
        ("line", 9.5, 10.0, 16.5, 10.0, 1.4, "green"),
        ("line", 9.5, 13.0, 16.5, 13.0, 1.4, "green"),
        ("line", 9.5, 16.0, 16.5, 16.0, 1.4, "green"),
    ),
    "audio": (
        ("line", 14.0, 5.0, 14.0, 16.2, 2.0, "pink"),
        ("line", 14.0, 5.0, 18.5, 6.2, 1.8, "pink"),
        ("circle", 9.0, 17.0, 3.0, 1.9, "pink"),
        ("circle", 15.0, 16.0, 2.5, 1.9, "pink"),
    ),
    "video": (
        ("rect", 4.0, 7.0, 11.0, 10.0, 1.8, "purple"),
        ("polygon", ((15.0, 10.0), (20.5, 7.2), (20.5, 16.8), (15.0, 14.0)), "purple"),
    ),
    "media": (
        ("rect", 4.0, 5.0, 16.0, 14.0, 1.7, "blue"),
        ("circle", 9.0, 9.0, 1.8, 1.4, "blue"),
        ("polyline", ((5.5, 17.0), (10.0, 12.5), (13.0, 15.2), (15.0, 13.0), (19.0, 17.0)), 1.7, "blue"),
    ),
    "archive": (
        ("polygon", ((12.0, 4.0), (19.0, 8.0), (12.0, 12.0), (5.0, 8.0)), "orange"),
        ("polyline", ((5.0, 12.0), (12.0, 16.0), (19.0, 12.0)), 1.8, "orange"),
        ("polyline", ((5.0, 16.0), (12.0, 20.0), (19.0, 16.0)), 1.8, "orange"),
    ),
    "code": (
        ("polyline", ((9.0, 7.0), (5.0, 12.0), (9.0, 17.0)), 1.9, "blue"),
        ("polyline", ((15.0, 7.0), (19.0, 12.0), (15.0, 17.0)), 1.9, "blue"),
        ("line", 13.0, 5.5, 11.0, 18.5, 1.7, "ink"),
    ),
    "security": (
        ("polygon", ((12.0, 3.5), (19.0, 6.5), (18.0, 13.8), (12.0, 20.5), (6.0, 13.8), (5.0, 6.5)), "green"),
        ("polyline", ((8.5, 12.0), (11.0, 14.5), (16.0, 9.5)), 1.8, "white"),
    ),
    "check": (
        ("circle", 12.0, 12.0, 8.0, 1.7, "green"),
        ("polyline", ((8.0, 12.2), (11.0, 15.0), (16.8, 9.0)), 1.9, "green"),
    ),
    "cancel": (
        ("circle", 12.0, 12.0, 8.0, 1.7, "red"),
        ("line", 8.5, 8.5, 15.5, 15.5, 1.9, "red"),
        ("line", 15.5, 8.5, 8.5, 15.5, 1.9, "red"),
    ),
    "close": (
        ("line", 7.0, 7.0, 17.0, 17.0, 1.8, "muted"),
        ("line", 17.0, 7.0, 7.0, 17.0, 1.8, "muted"),
    ),
    "save": (
        ("rect", 5.0, 4.0, 14.0, 16.0, 1.8, "ink"),
        ("line", 8.0, 4.0, 8.0, 9.0, 1.5, "ink"),
        ("line", 15.0, 4.0, 15.0, 9.0, 1.5, "ink"),
        ("rect", 8.0, 13.0, 8.0, 5.0, 1.5, "accent"),
    ),
    "theme": (
        ("circle", 12.0, 12.0, 4.0, 1.7, "orange"),
        ("line", 12.0, 3.0, 12.0, 5.0, 1.6, "orange"),
        ("line", 12.0, 19.0, 12.0, 21.0, 1.6, "orange"),
        ("line", 3.0, 12.0, 5.0, 12.0, 1.6, "orange"),
        ("line", 19.0, 12.0, 21.0, 12.0, 1.6, "orange"),
        ("line", 5.6, 5.6, 7.0, 7.0, 1.5, "orange"),
        ("line", 17.0, 17.0, 18.4, 18.4, 1.5, "orange"),
        ("line", 18.4, 5.6, 17.0, 7.0, 1.5, "orange"),
        ("line", 7.0, 17.0, 5.6, 18.4, 1.5, "orange"),
    ),
    "settings": (
        ("circle", 12.0, 12.0, 3.0, 1.8, "ink"),
        ("circle", 12.0, 12.0, 7.0, 1.6, "ink"),
        ("line", 12.0, 3.0, 12.0, 5.0, 1.7, "ink"),
        ("line", 12.0, 19.0, 12.0, 21.0, 1.7, "ink"),
        ("line", 3.0, 12.0, 5.0, 12.0, 1.7, "ink"),
        ("line", 19.0, 12.0, 21.0, 12.0, 1.7, "ink"),
        ("line", 5.7, 5.7, 7.1, 7.1, 1.7, "ink"),
        ("line", 16.9, 16.9, 18.3, 18.3, 1.7, "ink"),
        ("line", 18.3, 5.7, 16.9, 7.1, 1.7, "ink"),
        ("line", 7.1, 16.9, 5.7, 18.3, 1.7, "ink"),
    ),
}


def hex_to_rgba(color: str) -> tuple[int, int, int, int]:
    value = color.strip().lstrip("#")
    if len(value) == 3:
        value = "".join(channel * 2 for channel in value)
    return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16), 255


def token_color(token: str, palette: dict[str, str]) -> tuple[int, int, int, int]:
    return hex_to_rgba(palette.get(token, token))


def put_rgba_pixel(pixels: bytearray, width: int, height: int, x: int, y: int, color: tuple[int, int, int, int]) -> None:
    if x < 0 or y < 0 or x >= width or y >= height:
        return
    index = (y * width + x) * 4
    pixels[index : index + 4] = bytes(color)


def draw_line_icon(
    pixels: bytearray,
    width: int,
    height: int,
    unit: float,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    stroke: float,
    color: tuple[int, int, int, int],
) -> None:
    padding = stroke / 2 + 0.8
    left = max(0, int(math.floor((min(x1, x2) - padding) * unit)))
    right = min(width, int(math.ceil((max(x1, x2) + padding) * unit)))
    top = max(0, int(math.floor((min(y1, y2) - padding) * unit)))
    bottom = min(height, int(math.ceil((max(y1, y2) + padding) * unit)))
    dx = x2 - x1
    dy = y2 - y1
    length_squared = dx * dx + dy * dy
    radius = stroke / 2
    for py in range(top, bottom):
        cy = (py + 0.5) / unit
        for px in range(left, right):
            cx = (px + 0.5) / unit
            if length_squared:
                position = max(0.0, min(1.0, ((cx - x1) * dx + (cy - y1) * dy) / length_squared))
                nearest_x = x1 + position * dx
                nearest_y = y1 + position * dy
            else:
                nearest_x = x1
                nearest_y = y1
            if math.hypot(cx - nearest_x, cy - nearest_y) <= radius:
                put_rgba_pixel(pixels, width, height, px, py, color)


def draw_fill_rect_icon(
    pixels: bytearray,
    width: int,
    height: int,
    unit: float,
    x: float,
    y: float,
    rect_width: float,
    rect_height: float,
    color: tuple[int, int, int, int],
) -> None:
    left = max(0, int(math.floor(x * unit)))
    right = min(width, int(math.ceil((x + rect_width) * unit)))
    top = max(0, int(math.floor(y * unit)))
    bottom = min(height, int(math.ceil((y + rect_height) * unit)))
    for py in range(top, bottom):
        for px in range(left, right):
            put_rgba_pixel(pixels, width, height, px, py, color)


def draw_rounded_rect_icon(
    pixels: bytearray,
    width: int,
    height: int,
    unit: float,
    x: float,
    y: float,
    rect_width: float,
    rect_height: float,
    radius: float,
    color: tuple[int, int, int, int],
) -> None:
    left = max(0, int(math.floor(x * unit)))
    right = min(width, int(math.ceil((x + rect_width) * unit)))
    top = max(0, int(math.floor(y * unit)))
    bottom = min(height, int(math.ceil((y + rect_height) * unit)))
    for py in range(top, bottom):
        cy = (py + 0.5) / unit
        for px in range(left, right):
            cx = (px + 0.5) / unit
            nearest_x = min(max(cx, x + radius), x + rect_width - radius)
            nearest_y = min(max(cy, y + radius), y + rect_height - radius)
            if math.hypot(cx - nearest_x, cy - nearest_y) <= radius:
                put_rgba_pixel(pixels, width, height, px, py, color)


def draw_circle_icon(
    pixels: bytearray,
    width: int,
    height: int,
    unit: float,
    cx: float,
    cy: float,
    radius: float,
    stroke: float | None,
    color: tuple[int, int, int, int],
) -> None:
    padding = (stroke or 0) / 2 + 0.8
    left = max(0, int(math.floor((cx - radius - padding) * unit)))
    right = min(width, int(math.ceil((cx + radius + padding) * unit)))
    top = max(0, int(math.floor((cy - radius - padding) * unit)))
    bottom = min(height, int(math.ceil((cy + radius + padding) * unit)))
    for py in range(top, bottom):
        py_coord = (py + 0.5) / unit
        for px in range(left, right):
            px_coord = (px + 0.5) / unit
            distance = math.hypot(px_coord - cx, py_coord - cy)
            if stroke is None:
                hit = distance <= radius
            else:
                hit = abs(distance - radius) <= stroke / 2
            if hit:
                put_rgba_pixel(pixels, width, height, px, py, color)


def point_in_polygon(x: float, y: float, points: tuple[tuple[float, float], ...]) -> bool:
    inside = False
    previous_x, previous_y = points[-1]
    for current_x, current_y in points:
        crosses = (current_y > y) != (previous_y > y)
        if crosses:
            slope_x = (previous_x - current_x) * (y - current_y) / ((previous_y - current_y) or 1e-9) + current_x
            if x < slope_x:
                inside = not inside
        previous_x, previous_y = current_x, current_y
    return inside


def draw_polygon_icon(
    pixels: bytearray,
    width: int,
    height: int,
    unit: float,
    points: tuple[tuple[float, float], ...],
    color: tuple[int, int, int, int],
) -> None:
    left = max(0, int(math.floor(min(point[0] for point in points) * unit)))
    right = min(width, int(math.ceil(max(point[0] for point in points) * unit)))
    top = max(0, int(math.floor(min(point[1] for point in points) * unit)))
    bottom = min(height, int(math.ceil(max(point[1] for point in points) * unit)))
    for py in range(top, bottom):
        cy = (py + 0.5) / unit
        for px in range(left, right):
            cx = (px + 0.5) / unit
            if point_in_polygon(cx, cy, points):
                put_rgba_pixel(pixels, width, height, px, py, color)


def render_icon_pixels(
    commands: tuple[tuple[object, ...], ...],
    palette: dict[str, str],
    size: int,
    scale: int,
) -> tuple[int, int, bytearray]:
    high_size = size * scale
    unit = high_size / 24.0
    high_pixels = bytearray(high_size * high_size * 4)
    for command in commands:
        kind = command[0]
        if kind == "line":
            _, x1, y1, x2, y2, stroke, token = command
            draw_line_icon(high_pixels, high_size, high_size, unit, float(x1), float(y1), float(x2), float(y2), float(stroke), token_color(str(token), palette))
        elif kind == "polyline":
            _, points, stroke, token = command
            point_tuple = tuple((float(x), float(y)) for x, y in points)
            for start, end in zip(point_tuple, point_tuple[1:]):
                draw_line_icon(high_pixels, high_size, high_size, unit, start[0], start[1], end[0], end[1], float(stroke), token_color(str(token), palette))
        elif kind == "rect":
            _, x, y, rect_width, rect_height, stroke, token = command
            color = token_color(str(token), palette)
            x = float(x)
            y = float(y)
            rect_width = float(rect_width)
            rect_height = float(rect_height)
            stroke = float(stroke)
            draw_line_icon(high_pixels, high_size, high_size, unit, x, y, x + rect_width, y, stroke, color)
            draw_line_icon(high_pixels, high_size, high_size, unit, x + rect_width, y, x + rect_width, y + rect_height, stroke, color)
            draw_line_icon(high_pixels, high_size, high_size, unit, x + rect_width, y + rect_height, x, y + rect_height, stroke, color)
            draw_line_icon(high_pixels, high_size, high_size, unit, x, y + rect_height, x, y, stroke, color)
        elif kind == "fill_rect":
            _, x, y, rect_width, rect_height, token = command
            draw_fill_rect_icon(high_pixels, high_size, high_size, unit, float(x), float(y), float(rect_width), float(rect_height), token_color(str(token), palette))
        elif kind == "rounded_rect":
            _, x, y, rect_width, rect_height, radius, token = command
            draw_rounded_rect_icon(high_pixels, high_size, high_size, unit, float(x), float(y), float(rect_width), float(rect_height), float(radius), token_color(str(token), palette))
        elif kind == "circle":
            _, cx, cy, radius, stroke, token = command
            draw_circle_icon(high_pixels, high_size, high_size, unit, float(cx), float(cy), float(radius), float(stroke), token_color(str(token), palette))
        elif kind == "fill_circle":
            _, cx, cy, radius, token = command
            draw_circle_icon(high_pixels, high_size, high_size, unit, float(cx), float(cy), float(radius), None, token_color(str(token), palette))
        elif kind == "polygon":
            _, points, token = command
            point_tuple = tuple((float(x), float(y)) for x, y in points)
            draw_polygon_icon(high_pixels, high_size, high_size, unit, point_tuple, token_color(str(token), palette))

    pixels = bytearray(size * size * 4)
    samples = scale * scale
    for y in range(size):
        for x in range(size):
            red = green = blue = alpha = 0
            for sample_y in range(scale):
                for sample_x in range(scale):
                    source_index = (((y * scale + sample_y) * high_size) + (x * scale + sample_x)) * 4
                    red += high_pixels[source_index]
                    green += high_pixels[source_index + 1]
                    blue += high_pixels[source_index + 2]
                    alpha += high_pixels[source_index + 3]
            target_index = (y * size + x) * 4
            pixels[target_index : target_index + 4] = bytes((red // samples, green // samples, blue // samples, alpha // samples))
    return size, size, pixels


def png_bytes_from_rgba(width: int, height: int, pixels: bytes | bytearray) -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    raw_rows = bytearray()
    stride = width * 4
    for row in range(height):
        raw_rows.append(0)
        raw_rows.extend(pixels[row * stride : (row + 1) * stride])
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(bytes(raw_rows), 9))
        + chunk(b"IEND", b"")
    )


def inline_icon_photo(commands: tuple[tuple[object, ...], ...], palette: dict[str, str], size: int = 22, scale: int = 4) -> tk.PhotoImage:
    width, height, pixels = render_icon_pixels(commands, palette, size, scale)
    encoded = base64.b64encode(png_bytes_from_rgba(width, height, pixels)).decode("ascii")
    return tk.PhotoImage(data=encoded, format="png")


class ConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("File Converter")
        self.root.minsize(1180, 720)

        self.style = ttk.Style()
        self.theme_name = tk.StringVar(value="light")
        self.active_nav = tk.StringVar(value="home")
        self.from_format = tk.StringVar(value=CONVERSIONS[0].from_label)
        self.to_format = tk.StringVar(value=CONVERSIONS[0].to_label)
        self.mode = tk.StringVar(value="file")
        self.input_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.encoding = tk.StringVar(value="utf-8-sig")
        self.delimiter = tk.StringVar()
        self.overwrite = tk.BooleanVar(value=False)
        self.rename_output = tk.BooleanVar(value=False)
        self.custom_name = tk.StringVar()
        self.suggested_name = tk.StringVar(value="Suggested output: choose a file first")
        self.status = tk.StringVar(value="Ready")
        self.progress_value = tk.DoubleVar(value=0)
        self.progress_text = tk.StringVar(value="0%")
        self.editor_status = tk.StringVar(value="Choose an editable source file")
        self.upload_status_text = tk.StringVar(value="No file selected yet")
        self.opener_path = tk.StringVar()
        self.opener_status = tk.StringVar(value="Choose any supported file to open")

        self.input_submitted = False
        self.cancel_event = threading.Event()
        self.is_converting = False
        self.toast_window: tk.Toplevel | None = None
        self.upload_popup: tk.Toplevel | None = None
        self.upload_popup_box: tk.Frame | None = None
        self.upload_popup_plus: tk.Label | None = None
        self.upload_popup_title: ttk.Label | None = None
        self.upload_popup_hint: ttk.Label | None = None
        self.upload_popup_panels: list[tk.Frame] = []
        self.editor_popup: tk.Toplevel | None = None
        self.editor_text: tk.Text | None = None
        self.editor_path: Path | None = None
        self.editor_encoding_used = "utf-8-sig"
        self.status_popup: tk.Toplevel | None = None
        self.status_stage_canvas: tk.Canvas | None = None
        self.status_progress: ttk.Progressbar | None = None
        self.status_log: tk.Text | None = None
        self.status_popup_panels: list[tk.Frame] = []
        self.status_log_lines: list[str] = []
        self.active_step_index: int | None = None
        self.completed_step_count = 0
        self.current_step_detail = ""
        self.animation_frame = 0
        self.animation_job: str | None = None
        self.last_outputs: list[Path] = []
        self.step_titles = [
            "Input submitted",
            "Output name prepared",
            "Source read",
            "Conversion engine running",
            "Finished",
        ]

        self.tk_backgrounds: list[tuple[tk.Widget, str]] = []
        self.card_frames: list[tk.Frame] = []
        self.card_targets: dict[str, tk.Frame] = {}
        self.nav_buttons: dict[str, ttk.Button] = {}
        self.window_icon = self.create_window_icon()
        self.root.iconphoto(True, self.window_icon)
        self.ui_icons = self.create_ui_icons()
        self.logo_image: tk.PhotoImage | None = None
        self.icon_buttons: list[tuple[ttk.Button, str]] = []
        self.icon_labels: list[tuple[tk.Label, str]] = []

        self.build_layout()
        self.apply_theme()
        self.update_labels(clear_paths=False)
        self.render_timeline()

    @property
    def theme(self) -> AppTheme:
        return THEMES[self.theme_name.get()]

    @property
    def spec(self) -> ConversionSpec:
        for conversion in CONVERSIONS:
            if conversion.from_label == self.from_format.get() and conversion.to_label == self.to_format.get():
                return conversion
        matching_sources = [conversion for conversion in CONVERSIONS if conversion.from_label == self.from_format.get()]
        return matching_sources[0] if matching_sources else CONVERSIONS[0]

    def target_labels_for_source(self, source_label: str) -> list[str]:
        return [conversion.to_label for conversion in CONVERSIONS if conversion.from_label == source_label]

    def create_window_icon(self) -> tk.PhotoImage:
        return inline_icon_photo(
            FC_ICON_SOURCE,
            {
                "accent": "#14b8a6",
                "fc_dark": "#0f766e",
                "fc_shadow": "#0b4f4a",
                "white": "#ffffff",
            },
            size=128,
            scale=4,
        )

    def create_ui_icons(self) -> dict[str, tk.PhotoImage]:
        theme = self.theme
        palette = {
            "ink": theme.text,
            "muted": theme.muted,
            "accent": theme.accent,
            "white": "#ffffff",
            "blue": "#2563eb",
            "green": "#059669",
            "orange": "#f59e0b",
            "pink": "#db2777",
            "purple": "#7c3aed",
            "red": theme.danger,
        }
        return {name: inline_icon_photo(source, palette) for name, source in ICON_SOURCES.items()}

    def register_background(self, widget: tk.Widget, role: str) -> None:
        self.tk_backgrounds.append((widget, role))

    def icon_button(
        self,
        parent: tk.Widget,
        text: str,
        icon_name: str,
        style: str,
        command: Callable[[], object],
        **options: object,
    ) -> ttk.Button:
        button = ttk.Button(
            parent,
            text=text,
            image=self.ui_icons.get(icon_name),
            compound="left",
            style=style,
            command=command,
            **options,
        )
        self.icon_buttons.append((button, icon_name))
        return button

    def icon_label(self, parent: tk.Widget, icon_name: str, role: str = "panel") -> tk.Label:
        label = tk.Label(parent, image=self.ui_icons.get(icon_name), bd=0)
        self.icon_labels.append((label, icon_name))
        self.register_background(label, role)
        return label

    def refresh_icon_widgets(self) -> None:
        active_buttons: list[tuple[ttk.Button, str]] = []
        for button, icon_name in self.icon_buttons:
            try:
                if button.winfo_exists():
                    button.configure(image=self.ui_icons.get(icon_name), compound="left")
                    active_buttons.append((button, icon_name))
            except tk.TclError:
                continue
        self.icon_buttons = active_buttons

        active_labels: list[tuple[tk.Label, str]] = []
        for label, icon_name in self.icon_labels:
            try:
                if label.winfo_exists():
                    label.configure(image=self.ui_icons.get(icon_name))
                    active_labels.append((label, icon_name))
            except tk.TclError:
                continue
        self.icon_labels = active_labels

    def build_layout(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.shell = tk.Frame(self.root)
        self.shell.grid(row=0, column=0, sticky="nsew")
        self.shell.columnconfigure(0, weight=1)
        self.shell.rowconfigure(1, weight=1)
        self.register_background(self.shell, "background")

        self.build_app_bar()

        self.workspace = tk.Frame(self.shell)
        self.workspace.grid(row=1, column=0, sticky="nsew")
        self.workspace.columnconfigure(1, weight=1)
        self.workspace.rowconfigure(0, weight=1)
        self.register_background(self.workspace, "background")

        self.build_sidebar()

        self.page_host = tk.Frame(self.workspace)
        self.page_host.grid(row=0, column=1, sticky="nsew")
        self.page_host.columnconfigure(0, weight=1)
        self.page_host.rowconfigure(0, weight=1)
        self.register_background(self.page_host, "background")

        self.pages: dict[str, tk.Frame] = {}
        self.build_home_page()
        self.build_convert_page()
        self.build_opener_page()
        self.build_editor_page()
        self.build_downloads_page()
        self.build_footer()
        self.show_page("home")

    def update_scroll_region(self, event: tk.Event | None = None) -> None:
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def resize_canvas_content(self, event: tk.Event) -> None:
        content_width = max(event.width, 620)
        self.canvas.itemconfigure(self.main_window, width=content_width)
        wrap_width = max(content_width - 180, 320)
        if hasattr(self, "upload_hint"):
            self.upload_hint.configure(wraplength=wrap_width)
        if hasattr(self, "suggested_label"):
            self.suggested_label.configure(wraplength=wrap_width)

    def on_page_mousewheel(self, event: tk.Event) -> None:
        try:
            if event.widget.winfo_toplevel() != self.root:
                return
        except tk.TclError:
            return

        top, bottom = self.canvas.yview()
        if top <= 0 and bottom >= 1:
            return

        if getattr(event, "num", None) == 4:
            units = -3
        elif getattr(event, "num", None) == 5:
            units = 3
        else:
            delta = getattr(event, "delta", 0)
            if not delta:
                return
            units = -int(delta / 120)
            if units == 0:
                units = -1 if delta > 0 else 1
        self.canvas.yview_scroll(units, "units")

    def scroll_to_widget(self, widget: tk.Widget | None) -> None:
        if widget:
            widget.focus_set()

    def go_home(self) -> None:
        self.show_page("home")

    def navigate_to(self, target: str) -> None:
        self.show_page(target)
        if target == "editor":
            self.load_current_upload_in_editor(silent=True)

    def open_downloads_from_nav(self) -> None:
        self.show_page("downloads")

    def refresh_nav_buttons(self) -> None:
        for key, button in self.nav_buttons.items():
            try:
                button.configure(style="Active.Nav.TButton" if key == self.active_nav.get() else "Nav.TButton")
            except tk.TclError:
                continue

    def show_page(self, page_name: str) -> None:
        page = self.pages.get(page_name)
        if not page:
            return
        self.active_nav.set(page_name)
        page.tkraise()
        self.refresh_nav_buttons()
        self.status.set("Ready" if page_name == "home" else page_name.replace("_", " ").title())

    def make_page(self, name: str) -> tk.Frame:
        page = tk.Frame(self.page_host, padx=14, pady=14)
        page.grid(row=0, column=0, sticky="nsew")
        page.columnconfigure(0, weight=1)
        page.rowconfigure(0, weight=1)
        self.register_background(page, "background")
        self.pages[name] = page
        return page

    def make_panel(self, parent: tk.Widget, row: int, column: int, **grid_options: object) -> tk.Frame:
        panel = tk.Frame(parent, bd=0, highlightthickness=1, padx=18, pady=16)
        panel.grid(row=row, column=column, **grid_options)
        panel.columnconfigure(0, weight=1)
        self.card_frames.append(panel)
        return panel

    def build_app_bar(self) -> None:
        bar = tk.Frame(self.shell, bd=0, highlightthickness=1, padx=24, pady=14)
        bar.grid(row=0, column=0, sticky="ew")
        bar.columnconfigure(1, weight=1)
        self.card_frames.append(bar)

        self.logo_canvas = tk.Canvas(bar, width=50, height=50, highlightthickness=0, bd=0)
        self.logo_canvas.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 16))
        self.register_background(self.logo_canvas, "panel")
        ttk.Label(bar, text="File Converter", style="Title.TLabel").grid(row=0, column=1, sticky="sw")
        ttk.Label(bar, text="Conversion workspace", style="HeaderSubtitle.TLabel").grid(row=1, column=1, sticky="nw")

        search_shell = tk.Frame(bar, bd=0, highlightthickness=1, padx=14, pady=8)
        search_shell.grid(row=0, column=2, rowspan=2, sticky="ew", padx=(16, 18))
        search_shell.columnconfigure(1, weight=1)
        self.card_frames.append(search_shell)
        ttk.Label(search_shell, text="Search", style="Muted.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.search_entry = ttk.Entry(search_shell)
        self.search_entry.insert(0, "Search conversions, files, or formats...")
        self.search_entry.grid(row=0, column=1, sticky="ew")
        ttk.Label(search_shell, text="Ctrl + K", style="Muted.TLabel").grid(row=0, column=2, sticky="e", padx=(10, 0))

        self.theme_button = self.icon_button(bar, "Light", "theme", "Ghost.TButton", self.toggle_theme)
        self.theme_button.grid(row=0, column=3, rowspan=2, sticky="e", padx=(0, 10))
        self.icon_button(bar, "Settings", "settings", "Ghost.TButton", lambda: self.show_info("Settings", "More settings are coming soon.")).grid(
            row=0,
            column=4,
            rowspan=2,
            sticky="e",
        )

    def build_sidebar(self) -> None:
        sidebar = tk.Frame(self.workspace, bd=0, highlightthickness=1, padx=12, pady=18, width=210)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.columnconfigure(0, weight=1)
        sidebar.rowconfigure(6, weight=1)
        self.card_frames.append(sidebar)

        buttons = (
            ("home", "Home", "home", self.go_home),
            ("convert", "Convert", "convert", lambda: self.navigate_to("convert")),
            ("opener", "File Opener", "open", lambda: self.navigate_to("opener")),
            ("editor", "FC Text Editor", "edit", lambda: self.navigate_to("editor")),
            ("downloads", "Downloads", "download", self.open_downloads_from_nav),
        )
        for row, (key, label, icon_name, command) in enumerate(buttons):
            button = self.icon_button(sidebar, label, icon_name, "Nav.TButton", command)
            button.grid(row=row, column=0, sticky="ew", pady=(0, 8))
            self.nav_buttons[key] = button

        secure = tk.Frame(sidebar, bd=0, highlightthickness=1, padx=14, pady=14)
        secure.grid(row=7, column=0, sticky="sew")
        self.card_frames.append(secure)
        self.register_background(secure, "panel")
        self.icon_label(secure, "security").grid(row=0, column=0, sticky="w")
        ttk.Label(secure, text="Secure & private", style="Section.TLabel").grid(row=0, column=1, sticky="w", padx=(10, 0))
        ttk.Label(secure, text="Your files are processed locally on this device.", style="Muted.TLabel", wraplength=150).grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(12, 8),
        )
        ttk.Label(secure, text="Learn more", style="AccentLink.TLabel").grid(row=2, column=0, columnspan=2, sticky="w")

    def build_footer(self) -> None:
        footer = tk.Frame(self.shell, bd=0, highlightthickness=1, padx=20, pady=8)
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(1, weight=1)
        self.card_frames.append(footer)
        self.register_background(footer, "panel")
        tk.Label(footer, text="●", font=("Segoe UI", 10), fg="#22c55e").grid(row=0, column=0, sticky="w", padx=(0, 8))
        ttk.Label(footer, textvariable=self.status, style="Muted.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(footer, text="Version 2.0.0", style="Muted.TLabel").grid(row=0, column=2, padx=(0, 18))
        ttk.Label(footer, text="Check for updates", style="AccentLink.TLabel").grid(row=0, column=3, padx=(0, 18))
        ttk.Label(footer, text="Secure & private", style="Muted.TLabel").grid(row=0, column=4)

    def build_home_page(self) -> None:
        page = self.make_page("home")
        page.columnconfigure(0, weight=3)
        page.columnconfigure(1, weight=1)
        page.rowconfigure(0, weight=1)

        main = tk.Frame(page)
        main.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)
        self.register_background(main, "background")

        welcome = self.make_panel(main, 0, 0, sticky="ew", pady=(0, 12))
        welcome.columnconfigure(0, weight=1)
        ttk.Label(welcome, text="Welcome back!", style="Hero.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            welcome,
            text="Convert files quickly and securely. Everything happens on your device.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(8, 22))

        actions = tk.Frame(welcome)
        actions.grid(row=2, column=0, sticky="ew")
        for index in range(4):
            actions.columnconfigure(index, weight=1)
        self.register_background(actions, "panel")
        quick_actions = (
            ("Convert Files", "Convert files between formats in seconds.", "convert", lambda: self.show_page("convert")),
            ("Open Any File", "Preview or open supported converter files.", "open", lambda: self.show_page("opener")),
            ("Recent Downloads", "View your recently converted files.", "download", lambda: self.show_page("downloads")),
            ("Open Editor", "Edit and inspect files with the FC Text Editor.", "edit", lambda: self.show_page("editor")),
        )
        for column, (title, subtitle, icon_name, command) in enumerate(quick_actions):
            card = tk.Frame(actions, bd=0, highlightthickness=1, padx=14, pady=14, cursor="hand2")
            card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
            card.columnconfigure(1, weight=1)
            self.card_frames.append(card)
            self.icon_label(card, icon_name).grid(row=0, column=0, rowspan=2, sticky="nw", padx=(0, 10))
            ttk.Label(card, text=title, style="Section.TLabel").grid(row=0, column=1, sticky="w")
            ttk.Label(card, text=subtitle, style="Muted.TLabel", wraplength=135).grid(row=1, column=1, sticky="w", pady=(4, 0))
            for widget in (card,):
                widget.bind("<Button-1>", lambda event, action=command: action())

        recent = self.make_panel(main, 1, 0, sticky="nsew")
        recent.rowconfigure(2, weight=1)
        ttk.Label(recent, text="Recent jobs", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        self.icon_button(recent, "View all", "open", "Ghost.TButton", self.open_downloads_from_nav).grid(row=0, column=1, sticky="e")
        headers = ("File", "Conversion", "Output", "Status", "Time")
        table = tk.Frame(recent)
        table.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(14, 0))
        for column in range(len(headers)):
            table.columnconfigure(column, weight=1)
        self.register_background(table, "panel")
        for column, heading in enumerate(headers):
            ttk.Label(table, text=heading, style="FieldLabel.TLabel").grid(row=0, column=column, sticky="w", pady=(0, 8))
        rows = (
            ("Project_Proposal.docx", "Word → TXT", "C:\\Output\\Project_Proposal.txt", "Completed", "Today, 10:42 AM"),
            ("Financial_Report.pdf", "PDF → Word", "C:\\Output\\Financial_Report.docx", "Completed", "Today, 9:15 AM"),
            ("Sales_Data.csv", "CSV → Excel", "C:\\Output\\Sales_Data.xlsx", "Converting", "Today, 9:07 AM"),
            ("Meeting_Notes.docx", "Word → PDF", "C:\\Output\\Meeting_Notes.pdf", "Queued", "Today, 9:05 AM"),
            ("Deck_Presentation.pptx", "PowerPoint → PDF", "C:\\Output\\Deck_Presentation.pdf", "Failed", "Yesterday, 4:21 PM"),
        )
        for row_index, row in enumerate(rows, start=1):
            for column, value in enumerate(row):
                style = "StatusGood.TLabel" if value == "Completed" else "StatusBad.TLabel" if value == "Failed" else "Muted.TLabel"
                ttk.Label(table, text=value, style=style, wraplength=190).grid(row=row_index, column=column, sticky="w", pady=8)

        side = tk.Frame(page)
        side.grid(row=0, column=1, sticky="nsew")
        side.columnconfigure(0, weight=1)
        self.register_background(side, "background")
        self.build_supported_categories(side, 0)
        self.build_quick_stats(side, 1)

    def build_supported_categories(self, parent: tk.Widget, row: int) -> None:
        panel = self.make_panel(parent, row, 0, sticky="ew", pady=(0, 12))
        ttk.Label(panel, text="Supported categories", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(panel, text="View all", style="AccentLink.TLabel").grid(row=0, column=1, sticky="e")
        categories = (
            ("Documents", "DOCX, PDF, TXT, ODT, RTF...", "document", "12"),
            ("Data", "CSV, XLSX, JSON, XML, DBF...", "data", "8"),
            ("Audio", "MP3, WAV, AAC, FLAC, OGG...", "audio", "10"),
            ("Video", "MP4, MKV, AVI, MOV, WEBM...", "video", "9"),
        )
        for index, (title, subtitle, icon_name, count) in enumerate(categories, start=1):
            row_frame = tk.Frame(panel)
            row_frame.grid(row=index, column=0, columnspan=2, sticky="ew", pady=(14, 0))
            row_frame.columnconfigure(1, weight=1)
            self.register_background(row_frame, "panel")
            self.icon_label(row_frame, icon_name).grid(row=0, column=0, rowspan=2, padx=(0, 12), sticky="w")
            ttk.Label(row_frame, text=title, style="Card.TLabel").grid(row=0, column=1, sticky="w")
            ttk.Label(row_frame, text=subtitle, style="Muted.TLabel").grid(row=1, column=1, sticky="w")
            ttk.Label(row_frame, text=count, style="Count.TLabel").grid(row=0, column=2, rowspan=2, sticky="e")

    def build_quick_stats(self, parent: tk.Widget, row: int) -> None:
        panel = self.make_panel(parent, row, 0, sticky="ew")
        ttk.Label(panel, text="Quick stats", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        stats = (("Today's conversions", "23", "↑ 35%"), ("Files converted", "1.2 GB", "↑ 18%"), ("Success rate", "98%", "↑ 2%"), ("Batch jobs", "4", "View all"))
        for index, (label, value, delta) in enumerate(stats):
            card = tk.Frame(panel, bd=0, highlightthickness=1, padx=12, pady=12)
            card.grid(row=1 + index // 2, column=index % 2, sticky="ew", padx=(0 if index % 2 == 0 else 8, 0), pady=(14, 0))
            self.card_frames.append(card)
            ttk.Label(card, text=label, style="Muted.TLabel").grid(row=0, column=0, sticky="w")
            ttk.Label(card, text=value, style="HeroSmall.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 0))
            ttk.Label(card, text=delta, style="StatusGood.TLabel").grid(row=1, column=1, sticky="e", padx=(8, 0))

    def build_convert_page(self) -> None:
        page = self.make_page("convert")
        page.columnconfigure(0, weight=3)
        page.columnconfigure(1, weight=1)
        page.rowconfigure(0, weight=1)

        left = tk.Frame(page)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.columnconfigure(0, weight=1)
        self.register_background(left, "background")
        self.main = left
        self.build_conversion_card()
        self.build_output_card()
        self.build_action_area()

        right = tk.Frame(page)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        self.register_background(right, "background")
        queue = self.make_panel(right, 0, 0, sticky="ew", pady=(0, 12))
        ttk.Label(queue, text="Current queue", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(queue, textvariable=self.progress_text, style="Hero.TLabel").grid(row=1, column=0, sticky="w", pady=(18, 8))
        ttk.Label(queue, text="Status", style="FieldLabel.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Label(queue, textvariable=self.status, style="Muted.TLabel").grid(row=2, column=1, sticky="e")

        hints = self.make_panel(right, 1, 0, sticky="ew")
        ttk.Label(hints, text="Conversion hints", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        hint_rows = (
            ("CSV / TSV / JSON to Excel", "Convert delimited or JSON data to Excel workbooks."),
            ("PDF to Word", "Convert PDF documents to editable Word files."),
            ("EPUB to PDF", "Convert ebooks to PDF for easy sharing and printing."),
            ("ZIP tools", "Compress folders or extract ZIP archives quickly."),
            ("Audio and video", "Convert media with FFmpeg when installed."),
        )
        for index, (title, body) in enumerate(hint_rows, start=1):
            ttk.Label(hints, text=title, style="Card.TLabel").grid(row=index * 2 - 1, column=0, sticky="w", pady=(14, 0))
            ttk.Label(hints, text=body, style="Muted.TLabel", wraplength=230).grid(row=index * 2, column=0, sticky="w", pady=(2, 0))

    def build_opener_page(self) -> None:
        page = self.make_page("opener")
        page.columnconfigure(0, weight=3)
        page.columnconfigure(1, weight=1)
        page.rowconfigure(0, weight=1)

        left = tk.Frame(page)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.columnconfigure(0, weight=1)
        self.register_background(left, "background")

        opener = self.make_panel(left, 0, 0, sticky="ew")
        opener.columnconfigure(0, weight=1)
        ttk.Label(opener, text="File Opener", style="Hero.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            opener,
            text="Open supported converter files, send editable files to FC Text Editor, or start a conversion.",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(6, 16))

        drop_box = tk.Frame(opener, bd=0, highlightthickness=1, padx=22, pady=26, cursor="hand2")
        drop_box.grid(row=2, column=0, sticky="ew")
        drop_box.columnconfigure(1, weight=1)
        self.card_frames.append(drop_box)
        self.icon_label(drop_box, "open").grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 16))
        ttk.Label(drop_box, text="Choose a file to open", style="UploadTitle.TLabel").grid(row=0, column=1, sticky="sw")
        ttk.Label(
            drop_box,
            text="Documents, data files, archives, audio, video, and outputs created by File Converter.",
            style="UploadHint.TLabel",
        ).grid(row=1, column=1, sticky="nw", pady=(4, 0))
        self.icon_button(drop_box, "Choose File", "file", "Accent.TButton", self.choose_opener_file).grid(
            row=0,
            column=2,
            rowspan=2,
            sticky="e",
            padx=(18, 0),
        )
        for widget in (drop_box,):
            widget.bind("<Button-1>", lambda event: self.choose_opener_file())

        status_card = tk.Frame(opener, bd=0, highlightthickness=1, padx=14, pady=12)
        status_card.grid(row=3, column=0, sticky="ew", pady=(14, 0))
        status_card.columnconfigure(1, weight=1)
        self.card_frames.append(status_card)
        self.icon_label(status_card, "check").grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 12))
        ttk.Label(status_card, text="Selected file", style="Card.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(status_card, textvariable=self.opener_status, style="Muted.TLabel", wraplength=680).grid(
            row=1,
            column=1,
            sticky="w",
            pady=(3, 0),
        )

        action_row = tk.Frame(opener)
        action_row.grid(row=4, column=0, sticky="ew", pady=(14, 0))
        for column in range(4):
            action_row.columnconfigure(column, weight=1)
        self.register_background(action_row, "panel")
        self.icon_button(action_row, "Open File", "open", "Secondary.TButton", self.open_selected_file).grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 8),
        )
        self.icon_button(action_row, "Open Folder", "folder", "Secondary.TButton", self.open_selected_folder).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(0, 8),
        )
        self.icon_button(action_row, "Edit Text", "edit", "Secondary.TButton", self.send_opener_to_editor).grid(
            row=0,
            column=2,
            sticky="ew",
            padx=(0, 8),
        )
        self.icon_button(action_row, "Use for Convert", "convert", "Accent.TButton", self.send_opener_to_converter).grid(
            row=0,
            column=3,
            sticky="ew",
        )

        tips = self.make_panel(left, 1, 0, sticky="ew", pady=(12, 0))
        ttk.Label(tips, text="What the opener can do", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        opener_rows = (
            ("Open with Windows", "Uses the default app already installed on this computer.", "open"),
            ("Edit supported text", f"Sends {TEXT_EDIT_SUMMARY} to FC Text Editor.", "edit"),
            ("Start a conversion", "Detects the source type and fills the Convert page for you.", "convert"),
        )
        for index, (title, body, icon_name) in enumerate(opener_rows, start=1):
            row_frame = tk.Frame(tips)
            row_frame.grid(row=index, column=0, sticky="ew", pady=(14, 0))
            row_frame.columnconfigure(1, weight=1)
            self.register_background(row_frame, "panel")
            self.icon_label(row_frame, icon_name).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 12))
            ttk.Label(row_frame, text=title, style="Card.TLabel").grid(row=0, column=1, sticky="w")
            ttk.Label(row_frame, text=body, style="Muted.TLabel", wraplength=650).grid(row=1, column=1, sticky="w", pady=(2, 0))

        right = tk.Frame(page)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        self.register_background(right, "background")
        supported = self.make_panel(right, 0, 0, sticky="ew", pady=(0, 12))
        ttk.Label(supported, text="Supported opener types", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        opener_groups = (
            ("Documents", "PDF, DOC, DOCX, ODT, RTF, TXT, EPUB", "document"),
            ("Data", "CSV, TSV, JSON, XLS, XLSX", "data"),
            ("Archives", "ZIP folders and extracted outputs", "archive"),
            ("Media", "MP3, WAV, MP4, MOV, MKV, AVI", "media"),
        )
        for index, (title, body, icon_name) in enumerate(opener_groups, start=1):
            row_frame = tk.Frame(supported)
            row_frame.grid(row=index, column=0, sticky="ew", pady=(14, 0))
            row_frame.columnconfigure(1, weight=1)
            self.register_background(row_frame, "panel")
            self.icon_label(row_frame, icon_name).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 12))
            ttk.Label(row_frame, text=title, style="Card.TLabel").grid(row=0, column=1, sticky="w")
            ttk.Label(row_frame, text=body, style="Muted.TLabel", wraplength=230).grid(row=1, column=1, sticky="w", pady=(2, 0))

        self.build_quick_stats(right, 1)

    def build_editor_page(self) -> None:
        page = self.make_page("editor")
        page.columnconfigure(0, weight=3)
        page.columnconfigure(1, weight=1)
        page.rowconfigure(0, weight=1)
        left = tk.Frame(page)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(4, weight=1)
        self.register_background(left, "background")
        self.main = left
        self.build_editor_card()

        right = tk.Frame(page)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        self.register_background(right, "background")
        assistant = self.make_panel(right, 0, 0, sticky="ew", pady=(0, 12))
        ttk.Label(assistant, text="Assistant", style="AccentLink.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(assistant, text="Validation", style="Muted.TLabel").grid(row=0, column=1, padx=20)
        ttk.Label(assistant, text="Security", style="Muted.TLabel").grid(row=0, column=2, sticky="e")
        tools = (
            ("Media Tools", "Insert images, icons, video, audio, or file references.", "media"),
            ("Code Tools", "Insert code blocks and programming-document snippets.", "code"),
            ("Validation Summary", "JSON syntax, required fields, and duplicate checks.", "check"),
            ("Security Scan", "No threats detected. Files stay local.", "security"),
        )
        for index, (title, body, icon_name) in enumerate(tools, start=1):
            row_frame = tk.Frame(assistant, bd=0, highlightthickness=1, padx=12, pady=12)
            row_frame.grid(row=index, column=0, columnspan=3, sticky="ew", pady=(14, 0))
            row_frame.columnconfigure(1, weight=1)
            self.card_frames.append(row_frame)
            self.icon_label(row_frame, icon_name).grid(row=0, column=0, rowspan=2, sticky="nw", padx=(0, 10))
            ttk.Label(row_frame, text=title, style="Card.TLabel").grid(row=0, column=1, sticky="w")
            ttk.Label(row_frame, text=body, style="Muted.TLabel", wraplength=230).grid(row=1, column=1, sticky="w")

    def build_downloads_page(self) -> None:
        page = self.make_page("downloads")
        page.columnconfigure(0, weight=3)
        page.columnconfigure(1, weight=1)
        page.rowconfigure(0, weight=1)

        main = self.make_panel(page, 0, 0, sticky="nsew", padx=(0, 12))
        main.rowconfigure(3, weight=1)
        ttk.Label(main, text="Downloads", style="Hero.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(main, text="View and manage your converted output files.", style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 16))

        stats = tk.Frame(main)
        stats.grid(row=2, column=0, sticky="ew", pady=(0, 14))
        for index in range(4):
            stats.columnconfigure(index, weight=1)
        self.register_background(stats, "panel")
        for index, (label, value) in enumerate((("Converted today", "23"), ("Storage used", "1.2 GB"), ("Success rate", "98%"), ("Total conversions", "1,248"))):
            card = tk.Frame(stats, bd=0, highlightthickness=1, padx=12, pady=12)
            card.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 8, 0))
            self.card_frames.append(card)
            ttk.Label(card, text=label, style="Muted.TLabel").grid(row=0, column=0, sticky="w")
            ttk.Label(card, text=value, style="HeroSmall.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 0))

        table = tk.Frame(main)
        table.grid(row=3, column=0, sticky="nsew")
        table.columnconfigure(0, weight=2)
        for column in range(1, 6):
            table.columnconfigure(column, weight=1)
        self.register_background(table, "panel")
        for column, heading in enumerate(("Name", "Source", "Output", "Size", "Status", "Location")):
            ttk.Label(table, text=heading, style="FieldLabel.TLabel").grid(row=0, column=column, sticky="w", pady=(0, 8))
        rows = (
            ("Project_Proposal.docx", "DOCX", "TXT", "12.4 KB", "Completed", "C:\\Output\\Text"),
            ("Financial_Report.pdf", "PDF", "Word", "1.8 MB", "Completed", "C:\\Output\\Word"),
            ("Sales_Data.xlsx", "XLSX", "CSV", "245 KB", "Completed", "C:\\Output\\CSV"),
            ("Deck_Presentation.pptx", "PPTX", "PDF", "5.2 MB", "Failed", "C:\\Output\\PDF"),
        )
        for row_index, row in enumerate(rows, start=1):
            for column, value in enumerate(row):
                style = "StatusGood.TLabel" if value == "Completed" else "StatusBad.TLabel" if value == "Failed" else "Muted.TLabel"
                ttk.Label(table, text=value, style=style).grid(row=row_index, column=column, sticky="w", pady=10)

        details = self.make_panel(page, 0, 1, sticky="nsew")
        ttk.Label(details, text="Project_Proposal.docx", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(details, text="TXT • 12.4 KB", style="Muted.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 14))
        self.icon_button(details, "Open File", "open", "Secondary.TButton", lambda: self.open_path(self.last_outputs[0]) if self.last_outputs else None).grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(0, 8),
        )
        self.icon_button(details, "Open Folder", "folder", "Secondary.TButton", lambda: self.open_path(self.last_outputs[0].parent) if self.last_outputs else None).grid(
            row=3,
            column=0,
            sticky="ew",
            pady=(0, 18),
        )
        ttk.Label(details, text="Details", style="Section.TLabel").grid(row=4, column=0, sticky="w")
        for index, (label, value) in enumerate((("Source format", "DOCX"), ("Output format", "TXT"), ("Created", "Today, 10:42 AM"), ("Conversion time", "1.2s")), start=5):
            ttk.Label(details, text=label, style="FieldLabel.TLabel").grid(row=index, column=0, sticky="w", pady=(10, 0))
            ttk.Label(details, text=value, style="Muted.TLabel").grid(row=index, column=1, sticky="e", pady=(10, 0))

    def build_header(self) -> None:
        header = tk.Frame(self.main, bd=0, highlightthickness=1, padx=22, pady=18)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        header.columnconfigure(1, weight=1)
        self.card_frames.append(header)

        self.logo_canvas = tk.Canvas(header, width=54, height=54, highlightthickness=0, bd=0)
        self.logo_canvas.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 14))
        self.register_background(self.logo_canvas, "panel")

        ttk.Label(header, text="File Converter", style="Title.TLabel").grid(row=0, column=1, sticky="sw")
        ttk.Label(header, text="Conversion workspace", style="HeaderSubtitle.TLabel").grid(
            row=1,
            column=1,
            sticky="nw",
        )
        self.theme_button = self.icon_button(header, "Light", "theme", "Ghost.TButton", self.toggle_theme)
        self.theme_button.grid(row=0, column=2, rowspan=2, sticky="e")

    def build_navigation(self) -> None:
        nav = tk.Frame(self.main, bd=0, highlightthickness=1, padx=18, pady=12)
        nav.grid(row=1, column=0, sticky="ew", pady=(0, 16))
        nav.columnconfigure(6, weight=1)
        self.card_frames.append(nav)
        self.register_background(nav, "panel")

        ttk.Label(nav, text="Navigate", style="FieldLabel.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 12))
        buttons = (
            ("home", "Home", "home", self.go_home),
            ("convert", "Convert", "convert", lambda: self.navigate_to("convert")),
            ("opener", "File Opener", "open", lambda: self.navigate_to("opener")),
            ("editor", "FC Text Editor", "edit", lambda: self.navigate_to("editor")),
            ("downloads", "Downloads", "download", self.open_downloads_from_nav),
        )
        for column, (key, label, icon_name, command) in enumerate(buttons, start=1):
            button = self.icon_button(nav, label, icon_name, "Nav.TButton", command)
            button.grid(row=0, column=column, sticky="w", padx=(0, 8))
            self.nav_buttons[key] = button

    def make_card(self, row: int, title: str, icon_name: str = "file") -> tk.Frame:
        card = tk.Frame(self.main, bd=0, highlightthickness=1)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 16))
        card.columnconfigure(0, weight=1)
        self.card_frames.append(card)
        self.card_targets[title] = card

        title_row = tk.Frame(card)
        title_row.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 10))
        title_row.columnconfigure(1, weight=1)
        self.register_background(title_row, "panel")
        self.icon_label(title_row, icon_name).grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Label(title_row, text=title, style="Section.TLabel").grid(row=0, column=1, sticky="w")

        body = tk.Frame(card)
        body.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 20))
        body.columnconfigure(1, weight=1)
        self.register_background(body, "panel")
        return body

    def build_conversion_card(self) -> None:
        body = self.make_card(2, "Convert files", "convert")
        body.columnconfigure(0, weight=1)

        selector_row = tk.Frame(body)
        selector_row.grid(row=0, column=0, sticky="ew")
        selector_row.columnconfigure(0, weight=1)
        selector_row.columnconfigure(1, weight=1)
        self.register_background(selector_row, "panel")

        from_group = tk.Frame(selector_row)
        from_group.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        from_group.columnconfigure(0, weight=1)
        self.register_background(from_group, "panel")
        ttk.Label(from_group, text="Convert from", style="FieldLabel.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.from_menu = ttk.Combobox(
            from_group,
            textvariable=self.from_format,
            values=FROM_FORMATS,
            state="readonly",
        )
        self.from_menu.grid(row=1, column=0, sticky="ew")
        self.from_menu.bind("<<ComboboxSelected>>", lambda event: self.update_target_choices())

        to_group = tk.Frame(selector_row)
        to_group.grid(row=0, column=1, sticky="ew", padx=(10, 0))
        to_group.columnconfigure(0, weight=1)
        self.register_background(to_group, "panel")
        ttk.Label(to_group, text="Convert to", style="FieldLabel.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.to_menu = ttk.Combobox(
            to_group,
            textvariable=self.to_format,
            values=self.target_labels_for_source(self.from_format.get()),
            state="readonly",
        )
        self.to_menu.grid(row=1, column=0, sticky="ew")
        self.to_menu.bind("<<ComboboxSelected>>", lambda event: self.update_labels())

        mode_row = tk.Frame(body)
        mode_row.grid(row=1, column=0, sticky="w", pady=(14, 0))
        self.register_background(mode_row, "panel")
        ttk.Radiobutton(
            mode_row,
            text="Files",
            value="file",
            variable=self.mode,
            style="Card.TRadiobutton",
            command=self.update_labels,
        ).grid(row=0, column=0, padx=(0, 18), sticky="w")
        ttk.Radiobutton(
            mode_row,
            text="Folder",
            value="folder",
            variable=self.mode,
            style="Card.TRadiobutton",
            command=self.update_labels,
        ).grid(row=0, column=1, sticky="w")

        self.build_upload_section(body, 2)

    def build_upload_section(self, parent: tk.Widget, row: int) -> None:
        upload_section = tk.Frame(parent)
        upload_section.grid(row=row, column=0, sticky="ew", pady=(18, 0))
        upload_section.columnconfigure(0, weight=1)
        self.register_background(upload_section, "panel")

        ttk.Label(upload_section, text="Upload source", style="FieldLabel.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))

        self.upload_summary_box = tk.Frame(upload_section, bd=0, highlightthickness=1, padx=16, pady=14, cursor="hand2")
        self.upload_summary_box.grid(row=1, column=0, sticky="ew")
        self.upload_summary_box.columnconfigure(1, weight=1)
        self.register_background(self.upload_summary_box, "panel_alt")

        self.upload_icon_label = tk.Label(self.upload_summary_box, text="+", font=("Segoe UI", 28, "bold"), cursor="hand2")
        self.upload_icon_label.grid(row=0, column=0, rowspan=2, padx=(0, 14))
        self.register_background(self.upload_icon_label, "panel_alt")

        self.upload_title = ttk.Label(self.upload_summary_box, text="Add file", style="UploadTitle.TLabel")
        self.upload_title.grid(row=0, column=1, sticky="sw")
        self.upload_hint = ttk.Label(self.upload_summary_box, text="Open the upload popup to choose a source file.", style="UploadHint.TLabel")
        self.upload_hint.grid(row=1, column=1, sticky="nw", pady=(4, 0))

        self.upload_popup_button = self.icon_button(
            self.upload_summary_box,
            "Open Upload",
            "upload",
            "Secondary.TButton",
            self.show_upload_popup,
        )
        self.upload_popup_button.grid(row=0, column=2, rowspan=2, sticky="e", padx=(16, 0))

        for widget in (self.upload_summary_box, self.upload_icon_label, self.upload_title, self.upload_hint):
            widget.bind("<Button-1>", lambda event: self.show_upload_popup())

    def build_output_card(self) -> None:
        body = self.make_card(3, "Output settings", "folder")

        ttk.Label(body, text="Save folder", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 12), pady=(0, 10))
        self.output_entry = ttk.Entry(body, textvariable=self.output_folder)
        self.output_entry.grid(row=0, column=1, sticky="ew", pady=(0, 10))
        self.output_button = self.icon_button(body, "Browse", "folder", "Secondary.TButton", self.choose_output_folder)
        self.output_button.grid(row=0, column=2, sticky="ew", padx=(10, 0), pady=(0, 10))

        ttk.Label(body, text="Rename output file?", style="Card.TLabel").grid(row=1, column=0, sticky="w", padx=(0, 12))
        rename_row = tk.Frame(body)
        rename_row.grid(row=1, column=1, columnspan=2, sticky="w")
        self.register_background(rename_row, "panel")
        self.rename_no = ttk.Radiobutton(
            rename_row,
            text="No - use source name plus file type",
            value=False,
            variable=self.rename_output,
            style="Card.TRadiobutton",
            command=self.update_rename_controls,
        )
        self.rename_no.grid(row=0, column=0, sticky="w", padx=(0, 18))
        self.rename_yes = ttk.Radiobutton(
            rename_row,
            text="Yes - type a custom name",
            value=True,
            variable=self.rename_output,
            style="Card.TRadiobutton",
            command=self.update_rename_controls,
        )
        self.rename_yes.grid(row=0, column=1, sticky="w")

        ttk.Label(body, text="Custom name", style="Card.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 12), pady=(10, 0))
        self.custom_name_entry = ttk.Entry(body, textvariable=self.custom_name)
        self.custom_name_entry.grid(row=2, column=1, columnspan=2, sticky="ew", pady=(10, 0))

        self.suggested_label = ttk.Label(body, textvariable=self.suggested_name, style="Muted.TLabel")
        self.suggested_label.grid(row=3, column=1, columnspan=2, sticky="w", pady=(8, 0))

    def build_editor_card(self) -> None:
        body = self.make_card(4, "FC Text Editor", "edit")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)

        ttk.Label(body, textvariable=self.editor_status, style="Muted.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 10))

        toolbar = tk.Frame(body)
        toolbar.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        toolbar.columnconfigure(8, weight=1)
        self.register_background(toolbar, "panel")

        self.icon_button(toolbar, "Load Upload", "upload", "Secondary.TButton", self.load_current_upload_in_editor).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 8),
        )
        self.icon_button(toolbar, "Choose File", "folder", "Secondary.TButton", self.choose_editor_file).grid(
            row=0,
            column=1,
            sticky="w",
            padx=(0, 8),
        )
        self.icon_button(toolbar, "Save", "save", "Accent.TButton", self.save_editor_file).grid(
            row=0,
            column=2,
            sticky="w",
            padx=(0, 8),
        )
        self.icon_button(toolbar, "Media", "media", "Secondary.TButton", self.insert_media_reference).grid(
            row=0,
            column=3,
            sticky="w",
            padx=(0, 8),
        )
        self.icon_button(toolbar, "Code", "code", "Secondary.TButton", self.insert_code_block).grid(
            row=0,
            column=4,
            sticky="w",
            padx=(0, 8),
        )
        self.icon_button(toolbar, "Validation", "check", "Secondary.TButton", self.insert_validation_snippet).grid(
            row=0,
            column=5,
            sticky="w",
            padx=(0, 8),
        )
        self.icon_button(toolbar, "Security", "security", "Secondary.TButton", self.insert_security_snippet).grid(
            row=0,
            column=6,
            sticky="w",
        )

        editor_frame = tk.Frame(body, bd=0, highlightthickness=1, padx=0, pady=0)
        editor_frame.grid(row=2, column=0, sticky="nsew")
        editor_frame.columnconfigure(0, weight=1)
        editor_frame.rowconfigure(0, weight=1)
        self.register_background(editor_frame, "panel")
        self.card_frames.append(editor_frame)

        self.editor_text = tk.Text(editor_frame, wrap="none", bd=0, padx=14, pady=14, undo=True, height=16)
        self.editor_text.grid(row=0, column=0, sticky="nsew")
        editor_y = ttk.Scrollbar(editor_frame, orient="vertical", command=self.editor_text.yview)
        editor_y.grid(row=0, column=1, sticky="ns")
        editor_x = ttk.Scrollbar(editor_frame, orient="horizontal", command=self.editor_text.xview)
        editor_x.grid(row=1, column=0, sticky="ew")
        self.editor_text.configure(yscrollcommand=editor_y.set, xscrollcommand=editor_x.set)

        ttk.Label(
            body,
            text="Use Media for linked images/audio/video in editable files. Use Code, Validation, or Security to insert programming snippets.",
            style="Muted.TLabel",
        ).grid(
            row=3,
            column=0,
            sticky="w",
            pady=(10, 0),
        )

    def build_action_area(self) -> None:
        action_row = tk.Frame(self.main, bd=0, highlightthickness=1, padx=20, pady=18)
        action_row.grid(row=5, column=0, sticky="ew", pady=(0, 16))
        action_row.columnconfigure(0, weight=1)
        self.card_frames.append(action_row)

        self.convert_button = self.icon_button(action_row, "Convert", "convert", "Accent.TButton", self.start_conversion)
        self.convert_button.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.cancel_button = self.icon_button(action_row, "Cancel", "cancel", "Danger.TButton", self.request_cancel, state="disabled")
        self.cancel_button.grid(row=0, column=1, sticky="ew")

    def toggle_theme(self) -> None:
        self.theme_name.set("dark" if self.theme_name.get() == "light" else "light")
        self.apply_theme()

    def apply_theme(self) -> None:
        theme = self.theme
        old_icons = self.ui_icons
        self.ui_icons = self.create_ui_icons()
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.root.configure(bg=theme.background)
        for widget, role in self.tk_backgrounds:
            widget.configure(bg=getattr(theme, role))
            if widget in (self.upload_icon_label,):
                widget.configure(fg=theme.accent)
        for frame in self.card_frames:
            frame.configure(bg=theme.panel, highlightbackground=theme.border, highlightcolor=theme.border)
        self.upload_summary_box.configure(highlightbackground=theme.border, highlightcolor=theme.accent)

        self.style.configure(".", font=("Segoe UI", 10), background=theme.background, foreground=theme.text)
        self.style.configure("Title.TLabel", font=("Segoe UI", 24, "bold"), background=theme.panel, foreground=theme.text)
        self.style.configure("Hero.TLabel", font=("Segoe UI", 21, "bold"), background=theme.panel, foreground=theme.text)
        self.style.configure("HeroSmall.TLabel", font=("Segoe UI", 18, "bold"), background=theme.panel, foreground=theme.text)
        self.style.configure("HeaderSubtitle.TLabel", font=("Segoe UI", 10), background=theme.panel, foreground=theme.muted)
        self.style.configure("Subtitle.TLabel", font=("Segoe UI", 10), background=theme.background, foreground=theme.muted)
        self.style.configure("Section.TLabel", font=("Segoe UI", 12, "bold"), background=theme.panel, foreground=theme.text)
        self.style.configure("Card.TLabel", background=theme.panel, foreground=theme.text)
        self.style.configure("FieldLabel.TLabel", font=("Segoe UI", 9, "bold"), background=theme.panel, foreground=theme.muted)
        self.style.configure("Muted.TLabel", background=theme.panel, foreground=theme.muted)
        self.style.configure("AccentLink.TLabel", font=("Segoe UI", 9, "bold"), background=theme.panel, foreground=theme.accent)
        self.style.configure("StatusGood.TLabel", font=("Segoe UI", 9, "bold"), background=theme.panel, foreground="#059669")
        self.style.configure("StatusBad.TLabel", font=("Segoe UI", 9, "bold"), background=theme.panel, foreground="#dc2626")
        self.style.configure("Count.TLabel", font=("Segoe UI", 10, "bold"), background=theme.panel_alt, foreground=theme.muted)
        self.style.configure("UploadTitle.TLabel", font=("Segoe UI", 14, "bold"), background=theme.panel_alt, foreground=theme.text)
        self.style.configure("UploadHint.TLabel", font=("Segoe UI", 10), background=theme.panel_alt, foreground=theme.muted)

        self.style.configure(
            "TEntry",
            fieldbackground=theme.entry,
            foreground=theme.text,
            bordercolor=theme.border,
            lightcolor=theme.border,
            darkcolor=theme.border,
            insertcolor=theme.text,
            padding=10,
        )
        self.style.map(
            "TEntry",
            fieldbackground=[("disabled", theme.panel_alt), ("readonly", theme.entry)],
            foreground=[("disabled", theme.muted)],
        )
        self.style.configure(
            "TCombobox",
            fieldbackground=theme.entry,
            background=theme.panel_alt,
            foreground=theme.text,
            arrowcolor=theme.accent,
            bordercolor=theme.border,
            lightcolor=theme.border,
            darkcolor=theme.border,
            padding=10,
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", theme.entry)],
            foreground=[("readonly", theme.text)],
            selectbackground=[("readonly", theme.selection)],
            selectforeground=[("readonly", theme.text)],
        )

        self.configure_button_styles()
        self.refresh_icon_widgets()
        self.refresh_nav_buttons()
        self.style.configure("Card.TRadiobutton", background=theme.panel, foreground=theme.text)
        self.style.map("Card.TRadiobutton", background=[("active", theme.panel)], foreground=[("active", theme.text)])
        self.style.configure("Card.TCheckbutton", background=theme.panel, foreground=theme.text)
        self.style.map("Card.TCheckbutton", background=[("active", theme.panel)], foreground=[("active", theme.text)])
        self.style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor=theme.panel_alt,
            background=theme.accent,
            bordercolor=theme.border,
            lightcolor=theme.accent,
            darkcolor=theme.accent,
        )

        self.draw_logo()
        self.theme_button.configure(text="Dark" if self.theme_name.get() == "dark" else "Light")
        self.refresh_editor_popup_theme()
        self.refresh_upload_popup_theme()
        self.refresh_status_popup_theme()
        self.render_timeline()
        self.root.update_idletasks()

    def configure_button_styles(self) -> None:
        theme = self.theme
        self.style.configure(
            "Accent.TButton",
            font=("Segoe UI", 11, "bold"),
            background=theme.accent,
            foreground=theme.accent_text,
            bordercolor=theme.accent,
            focusthickness=0,
            padding=(16, 13),
        )
        self.style.map(
            "Accent.TButton",
            background=[("active", theme.accent_hover), ("disabled", theme.border)],
            foreground=[("disabled", theme.muted)],
        )
        self.style.configure(
            "Danger.TButton",
            font=("Segoe UI", 10, "bold"),
            background=theme.danger,
            foreground="#ffffff",
            bordercolor=theme.danger,
            padding=(14, 12),
        )
        self.style.map(
            "Danger.TButton",
            background=[("active", theme.danger_hover), ("disabled", theme.border)],
            foreground=[("disabled", theme.muted)],
        )
        self.style.configure(
            "Secondary.TButton",
            background=theme.panel_alt,
            foreground=theme.text,
            bordercolor=theme.border,
            padding=(14, 10),
        )
        self.style.map("Secondary.TButton", background=[("active", theme.selection)])
        self.style.configure(
            "Ghost.TButton",
            background=theme.panel,
            foreground=theme.text,
            bordercolor=theme.border,
            padding=(14, 10),
        )
        self.style.map("Ghost.TButton", background=[("active", theme.panel_alt)])
        self.style.configure(
            "Nav.TButton",
            background=theme.panel,
            foreground=theme.text,
            bordercolor=theme.border,
            padding=(12, 9),
        )
        self.style.map("Nav.TButton", background=[("active", theme.panel_alt)])
        self.style.configure(
            "Active.Nav.TButton",
            background=theme.panel_alt,
            foreground=theme.text,
            bordercolor=theme.accent,
            padding=(12, 9),
        )
        self.style.map("Active.Nav.TButton", background=[("active", theme.selection)])

    def draw_logo(self) -> None:
        theme = self.theme
        canvas = self.logo_canvas
        canvas.delete("all")
        canvas.configure(bg=theme.panel)
        self.logo_image = inline_icon_photo(
            FC_ICON_SOURCE,
            {
                "accent": "#14b8a6",
                "fc_dark": "#0f766e",
                "fc_shadow": "#0b4f4a",
                "white": "#ffffff",
            },
            size=50,
            scale=4,
        )
        canvas.create_image(25, 25, image=self.logo_image)

    def render_timeline(self, active_index: int | None = None, completed: int = 0, detail: str = "") -> None:
        self.active_step_index = active_index
        self.completed_step_count = completed
        self.current_step_detail = detail
        self.draw_status_stages()

    def opener_filetypes(self) -> list[tuple[str, str]]:
        patterns = " ".join(f"*{extension}" for extension in SUPPORTED_OPEN_EXTENSIONS)
        return [("Supported File Converter files", patterns), ("All files", "*.*")]

    def selected_opener_file(self) -> Path | None:
        path_text = self.opener_path.get().strip()
        if not path_text:
            self.show_error("Choose a file", "Choose a file before using the opener actions.")
            return None
        path = Path(path_text)
        if not path.is_file():
            self.show_error("Missing file", f"This file no longer exists:\n{path}")
            return None
        return path

    def set_opener_file(self, path: str | Path) -> None:
        selected = Path(path)
        self.opener_path.set(str(selected))
        extension = selected.suffix.lower()
        if extension in SUPPORTED_OPEN_EXTENSIONS:
            conversion = self.matching_conversion_for_extension(extension)
            if conversion:
                self.opener_status.set(f"Ready: {selected.name} can be opened or converted as {conversion.input_label}.")
            elif self.can_edit_extension(selected):
                self.opener_status.set(f"Ready: {selected.name} can be opened or edited.")
            else:
                self.opener_status.set(f"Ready: {selected.name} is supported by File Converter outputs.")
        else:
            self.opener_status.set(f"Selected: {selected.name}. Windows can try to open it with the default app.")
        self.status.set("File selected for opener")

    def choose_opener_file(self) -> None:
        if self.is_converting:
            return
        path = filedialog.askopenfilename(
            title="Choose file to open",
            filetypes=self.opener_filetypes(),
        )
        if path:
            self.set_opener_file(path)

    def open_selected_file(self) -> None:
        path = self.selected_opener_file()
        if path:
            self.open_path(path)

    def open_selected_folder(self) -> None:
        path = self.selected_opener_file()
        if path:
            self.open_path(path.parent)

    def send_opener_to_editor(self) -> None:
        path = self.selected_opener_file()
        if not path:
            return
        if not self.can_edit_extension(path):
            self.show_error("Editor unavailable", f"The FC Text Editor supports {TEXT_EDIT_SUMMARY}.")
            return

        conversion = self.matching_conversion_for_extension(path.suffix.lower())
        if conversion:
            self.mode.set("file")
            self.from_format.set(conversion.from_label)
            self.to_format.set(conversion.to_label)
            self.update_target_choices(clear_paths=False)
        self.input_path.set(str(path))
        self.input_submitted = False
        self.update_upload_display()
        self.load_editor_file(path)
        self.show_page("editor")

    def send_opener_to_converter(self) -> None:
        path = self.selected_opener_file()
        if not path:
            return
        conversion = self.matching_conversion_for_extension(path.suffix.lower())
        if not conversion:
            self.show_error("Conversion unavailable", "This file type can be opened, but it is not a supported conversion source.")
            return

        self.mode.set("file")
        self.from_format.set(conversion.from_label)
        self.to_format.set(conversion.to_label)
        self.update_target_choices(clear_paths=False)
        self.input_path.set(str(path))
        self.input_submitted = False
        self.update_upload_display()
        self.status.set("File sent to Convert")
        self.show_page("convert")

    def editable_filetypes(self) -> list[tuple[str, str]]:
        filetypes = [(label, " ".join(f"*{extension}" for extension in extensions)) for label, extensions in EDITABLE_FORMAT_GROUPS]
        all_patterns = " ".join(f"*{extension}" for extension in TEXT_EDIT_EXTENSIONS)
        return [("All editable files", all_patterns), *filetypes, ("All files", "*.*")]

    def can_edit_extension(self, path: Path) -> bool:
        return path.suffix.lower() in TEXT_EDIT_EXTENSIONS

    def read_editable_text(self, path: Path) -> str:
        encodings = (self.encoding.get().strip() or "utf-8-sig", "utf-8-sig", "utf-8", "utf-16", "cp1252", "latin-1")
        seen: set[str] = set()
        last_error: UnicodeDecodeError | None = None
        for encoding in encodings:
            if encoding in seen:
                continue
            seen.add(encoding)
            try:
                content = path.read_text(encoding=encoding)
            except UnicodeDecodeError as error:
                last_error = error
                continue
            self.editor_encoding_used = encoding
            return content
        if last_error:
            raise last_error
        return path.read_text(encoding="utf-8-sig")

    def matching_conversion_for_extension(self, extension: str) -> ConversionSpec | None:
        for conversion in CONVERSIONS:
            if extension in conversion.input_extensions:
                return conversion
        return None

    def editable_source_path(self) -> Path | None:
        if self.mode.get() == "folder":
            return None
        path_text = self.input_path.get().strip()
        if not path_text:
            return None
        path = Path(path_text)
        if not self.can_edit_extension(path):
            return None
        return path

    def update_editor_status(self) -> None:
        path_text = self.input_path.get().strip()
        if self.mode.get() == "folder":
            self.editor_status.set("Editor available for single editable files")
        elif not path_text:
            self.editor_status.set("Choose an editable source file")
        else:
            path = Path(path_text)
            if self.can_edit_extension(path):
                self.editor_status.set(f"Editable source: {path.name}")
            else:
                self.editor_status.set(f"Editor not available for {path.suffix.lower() or 'this file'} files")

    def choose_editor_file(self) -> None:
        if self.is_converting:
            return
        path_text = filedialog.askopenfilename(
            title="Choose editable file",
            filetypes=self.editable_filetypes(),
        )
        if not path_text:
            return

        path = Path(path_text)
        if not self.can_edit_extension(path):
            self.show_error("Editor unavailable", f"The editor supports {TEXT_EDIT_SUMMARY}.")
            return

        matching_conversion = self.matching_conversion_for_extension(path.suffix.lower())
        if matching_conversion:
            self.mode.set("file")
            self.from_format.set(matching_conversion.from_label)
            self.to_format.set(matching_conversion.to_label)
            self.update_target_choices(clear_paths=False)

        self.input_path.set(path_text)
        self.input_submitted = False
        self.status.set("Editable file selected")
        self.set_progress(5, "Editable file selected")
        self.render_timeline()
        self.update_upload_display()
        self.load_editor_file(path)
        self.active_nav.set("editor")
        self.refresh_nav_buttons()
        self.scroll_to_widget(self.card_targets.get("FC Text Editor"))

    def load_current_upload_in_editor(self, silent: bool = False) -> None:
        source = self.editable_source_path()
        if not source:
            if not silent:
                self.show_error("Editor unavailable", f"Choose a single editable upload first. The editor supports {TEXT_EDIT_SUMMARY}.")
            return
        self.load_editor_file(source)

    def load_editor_file(self, source: Path) -> None:
        if not self.can_edit_extension(source):
            self.show_error("Editor unavailable", f"The FC Text Editor supports {TEXT_EDIT_SUMMARY}.")
            return
        if not source.is_file():
            self.show_error("Missing source", f"This source file no longer exists:\n{source}")
            return

        try:
            content = self.read_editable_text(source)
        except UnicodeDecodeError as error:
            self.show_error("Could not read file", f"This file could not be opened as editable text:\n{error}")
            return
        except OSError as error:
            self.show_error("Could not read file", f"Windows could not open this file:\n{source}\n\n{error}")
            return

        if self.editor_text and self.editor_text.winfo_exists():
            self.editor_text.delete("1.0", "end")
            self.editor_text.insert("1.0", content)
            self.editor_text.edit_modified(False)
            self.editor_text.focus_set()
        self.editor_path = source
        self.editor_status.set(f"FC Text Editor: {source.name}")
        self.status.set(f"Editing {source.name}")

    def editor_suffix(self) -> str:
        return self.editor_path.suffix.lower() if self.editor_path else ".txt"

    def insert_at_cursor(self, value: str) -> None:
        if not self.editor_text or not self.editor_text.winfo_exists():
            return
        self.editor_text.insert("insert", value)
        self.editor_text.focus_set()
        self.editor_text.edit_modified(True)

    def relative_media_reference(self, media_path: Path) -> str:
        if self.editor_path:
            try:
                return media_path.resolve().relative_to(self.editor_path.parent.resolve()).as_posix()
            except ValueError:
                pass
        return media_path.as_posix()

    def media_markup_for(self, media_path: Path) -> str:
        suffix = media_path.suffix.lower()
        reference = self.relative_media_reference(media_path)
        name = media_path.stem.replace('"', "'") or "media"
        editor_suffix = self.editor_suffix()
        video_types = {".mp4", ".webm", ".mov", ".mkv", ".avi"}
        audio_types = {".mp3", ".wav", ".m4a", ".aac", ".flac", ".ogg", ".wma"}
        image_types = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".bmp"}

        if editor_suffix in {".html", ".htm"}:
            if suffix in video_types:
                return f'\n<video controls src="{reference}"></video>\n'
            if suffix in audio_types:
                return f'\n<audio controls src="{reference}"></audio>\n'
            if suffix in image_types:
                return f'\n<img src="{reference}" alt="{name}">\n'
            return f'\n<a href="{reference}">{media_path.name}</a>\n'
        if editor_suffix in {".md", ".markdown"}:
            if suffix in image_types:
                return f"\n![{name}]({reference})\n"
            if suffix in video_types:
                return f'\n<video controls src="{reference}"></video>\n'
            if suffix in audio_types:
                return f'\n<audio controls src="{reference}"></audio>\n'
            return f"\n[{media_path.name}]({reference})\n"
        return f"\n[media: {media_path.name}]\npath={reference}\n"

    def insert_media_reference(self) -> None:
        if not self.editor_text:
            return
        path_text = filedialog.askopenfilename(
            title="Choose media to reference",
            filetypes=[
                ("Media files", "*.png *.jpg *.jpeg *.gif *.webp *.svg *.mp4 *.webm *.mov *.mkv *.mp3 *.wav *.m4a *.ogg *.pdf"),
                ("All files", "*.*"),
            ],
        )
        if path_text:
            self.insert_at_cursor(self.media_markup_for(Path(path_text)))

    def wrap_code_for_editor(self, code: str, language: str = "text") -> str:
        suffix = self.editor_suffix()
        if suffix in {".md", ".markdown"}:
            return f"\n```{language}\n{code.rstrip()}\n```\n"
        if suffix in {".html", ".htm"}:
            escaped = html.escape(code.rstrip())
            return f"\n<pre><code>{escaped}</code></pre>\n"
        return f"\n{code.rstrip()}\n"

    def insert_code_block(self) -> None:
        snippet = "# Code block\n# Add custom document logic here.\n"
        self.insert_at_cursor(self.wrap_code_for_editor(snippet, "python"))

    def validation_snippet(self) -> tuple[str, str]:
        suffix = self.editor_suffix()
        if suffix == ".py":
            return (
                "python",
                """def validate_required_fields(record, required_fields):
    missing = [field for field in required_fields if not str(record.get(field, "")).strip()]
    if missing:
        raise ValueError(f"Missing required fields: {', '.join(missing)}")
    return True
""",
            )
        if suffix in {".js", ".ts", ".jsx", ".tsx"}:
            return (
                "javascript",
                """function validateRequiredFields(record, requiredFields) {
  const missing = requiredFields.filter((field) => !String(record[field] ?? "").trim());
  if (missing.length) throw new Error(`Missing required fields: ${missing.join(", ")}`);
  return true;
}
""",
            )
        return (
            "text",
            """Validation rule:
- Check required fields before conversion or submission.
- Reject empty, malformed, or unexpected values.
- Return a clear error message when validation fails.
""",
        )

    def security_snippet(self) -> tuple[str, str]:
        suffix = self.editor_suffix()
        if suffix == ".py":
            return (
                "python",
                """from pathlib import Path

def safe_child_path(base_dir, filename):
    base = Path(base_dir).resolve()
    target = (base / filename).resolve()
    if target != base and base not in target.parents:
        raise ValueError("Blocked unsafe path traversal.")
    return target
""",
            )
        if suffix in {".js", ".ts", ".jsx", ".tsx"}:
            return (
                "javascript",
                """function sanitizePlainText(value) {
  return String(value ?? "").replace(/[<>]/g, "").trim();
}

function assertAllowedExtension(filename, allowedExtensions) {
  const lower = String(filename).toLowerCase();
  if (!allowedExtensions.some((extension) => lower.endsWith(extension))) {
    throw new Error("Blocked unsupported file type.");
  }
}
""",
            )
        return (
            "text",
            """Security rule:
- Allow only expected file types.
- Sanitize user-provided text before saving or rendering.
- Keep output paths inside the chosen folder.
- Never run embedded code automatically.
""",
        )

    def insert_validation_snippet(self) -> None:
        language, snippet = self.validation_snippet()
        self.insert_at_cursor(self.wrap_code_for_editor(snippet, language))

    def insert_security_snippet(self) -> None:
        language, snippet = self.security_snippet()
        self.insert_at_cursor(self.wrap_code_for_editor(snippet, language))

    def show_editor_popup(self) -> None:
        self.load_current_upload_in_editor()

    def close_editor_popup(self) -> None:
        if self.editor_popup and self.editor_popup.winfo_exists():
            self.editor_popup.destroy()
        self.editor_popup = None

    def refresh_editor_popup_theme(self) -> None:
        theme = self.theme
        if self.editor_text and self.editor_text.winfo_exists():
            self.editor_text.configure(
                bg=theme.log_background,
                fg=theme.log_text,
                insertbackground=theme.log_text,
                selectbackground=theme.selection,
                selectforeground=theme.text,
            )
        if not self.editor_popup or not self.editor_popup.winfo_exists():
            return
        self.editor_popup.configure(bg=theme.background)
        for child in self.editor_popup.winfo_children():
            if isinstance(child, tk.Frame):
                child.configure(bg=theme.background)
                for nested in child.winfo_children():
                    if isinstance(nested, tk.Frame):
                        nested.configure(bg=theme.panel, highlightbackground=theme.border, highlightcolor=theme.border)
                        for label in nested.winfo_children():
                            if isinstance(label, tk.Label):
                                if label.cget("text") == "ED":
                                    label.configure(bg=theme.accent, fg=theme.accent_text)
                                else:
                                    label.configure(bg=theme.panel, fg=theme.text)

    def save_editor_file(self) -> None:
        if not self.editor_text:
            return
        if not self.editor_path:
            self.show_error("Choose a file", "Choose an editable file or load the current upload before saving.")
            return
        try:
            content = self.editor_text.get("1.0", "end-1c")
            self.editor_path.write_text(content, encoding=self.editor_encoding_used)
        except OSError as error:
            self.show_error("Could not save file", f"Windows could not save this file:\n{self.editor_path}\n\n{error}")
            return

        self.input_submitted = False
        self.status.set("File saved. Submit upload before converting.")
        self.update_upload_display()
        self.editor_status.set(f"Saved edits: {self.editor_path.name}")

    def show_upload_popup(self) -> None:
        if self.is_converting:
            return
        if self.upload_popup and self.upload_popup.winfo_exists():
            self.upload_popup.lift()
            self.upload_popup.focus_force()
            return

        theme = self.theme
        popup = tk.Toplevel(self.root)
        self.upload_popup = popup
        self.upload_popup_panels = []
        popup.title("Upload Source")
        popup.minsize(640, 500)
        popup.transient(self.root)
        popup.iconphoto(True, self.window_icon)
        popup.configure(bg=theme.background)
        popup.columnconfigure(0, weight=1)
        popup.rowconfigure(0, weight=1)
        popup.protocol("WM_DELETE_WINDOW", self.close_upload_popup)

        shell = tk.Frame(popup, bg=theme.background, padx=24, pady=22)
        shell.grid(row=0, column=0, sticky="nsew")
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        header = tk.Frame(shell, bg=theme.panel, highlightbackground=theme.border, highlightthickness=1, padx=18, pady=16)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(1, weight=1)
        self.upload_popup_panels.append(header)

        tk.Label(
            header,
            text="UP",
            bg=theme.accent,
            fg=theme.accent_text,
            font=("Segoe UI", 11, "bold"),
            padx=12,
            pady=10,
        ).grid(row=0, column=0, rowspan=2, sticky="nsw", padx=(0, 14))
        tk.Label(header, text="Upload source", bg=theme.panel, fg=theme.text, font=("Segoe UI", 18, "bold")).grid(
            row=0,
            column=1,
            sticky="w",
        )
        tk.Label(
            header,
            text="Choose the file or folder you want to convert, then submit it.",
            bg=theme.panel,
            fg=theme.muted,
            font=("Segoe UI", 10),
            wraplength=420,
            justify="left",
        ).grid(row=1, column=1, sticky="w", pady=(4, 0))

        body = tk.Frame(shell, bg=theme.panel, highlightbackground=theme.border, highlightthickness=1, padx=18, pady=18)
        body.grid(row=1, column=0, sticky="nsew", pady=(0, 14))
        body.columnconfigure(0, weight=1)
        self.upload_popup_panels.append(body)

        self.upload_popup_box = tk.Frame(body, bd=0, highlightthickness=1, padx=20, pady=34, cursor="hand2")
        self.upload_popup_box.grid(row=0, column=0, sticky="ew")
        self.upload_popup_box.columnconfigure(0, weight=1)

        self.upload_popup_plus = tk.Label(self.upload_popup_box, text="+", font=("Segoe UI", 42, "bold"), cursor="hand2")
        self.upload_popup_plus.grid(row=0, column=0, sticky="n")

        self.upload_popup_title = ttk.Label(self.upload_popup_box, text="Drag & drop files or folders here", style="UploadTitle.TLabel")
        self.upload_popup_title.grid(row=1, column=0, sticky="n", pady=(10, 0))
        self.upload_popup_hint = ttk.Label(self.upload_popup_box, text="or click to browse", style="UploadHint.TLabel")
        self.upload_popup_hint.grid(row=2, column=0, sticky="n", pady=(4, 0))

        for widget in (self.upload_popup_box, self.upload_popup_plus, self.upload_popup_title, self.upload_popup_hint):
            widget.bind("<Button-1>", lambda event: self.choose_input())

        status_card = tk.Frame(body, bd=0, highlightthickness=1, padx=14, pady=12)
        status_card.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        status_card.columnconfigure(1, weight=1)
        self.upload_popup_panels.append(status_card)
        self.icon_label(status_card, "check").grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 12))
        ttk.Label(status_card, text="Ready to upload", style="Card.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(status_card, textvariable=self.upload_status_text, style="Muted.TLabel").grid(row=1, column=1, sticky="w", pady=(3, 0))

        choose_row = tk.Frame(body)
        choose_row.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        choose_row.columnconfigure(0, weight=1)
        choose_row.columnconfigure(1, weight=1)
        self.register_background(choose_row, "panel")
        self.icon_button(choose_row, "Choose file", "file", "Secondary.TButton", self.choose_file_input).grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(0, 8),
        )
        self.icon_button(choose_row, "Choose folder", "folder", "Secondary.TButton", self.choose_folder_input).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(8, 0),
        )
        ttk.Label(
            body,
            text="Accepted source types: all supported document, data, archive, audio, and video files.",
            style="Muted.TLabel",
        ).grid(row=3, column=0, sticky="w", pady=(12, 0))

        button_row = tk.Frame(shell, bg=theme.background)
        button_row.grid(row=2, column=0, sticky="ew")
        button_row.columnconfigure(0, weight=1)
        self.icon_button(button_row, "Cancel", "close", "Ghost.TButton", self.close_upload_popup).grid(
            row=0,
            column=1,
            sticky="e",
            padx=(0, 10),
        )
        self.icon_button(button_row, "Submit Upload", "check", "Accent.TButton", self.submit_input_from_popup).grid(
            row=0,
            column=2,
            sticky="e",
        )

        self.refresh_upload_popup_theme()
        self.refresh_upload_popup_text()
        self.center_popup(popup, 640, 500)
        popup.focus_force()

    def close_upload_popup(self) -> None:
        if self.upload_popup and self.upload_popup.winfo_exists():
            self.upload_popup.destroy()
        self.upload_popup = None
        self.upload_popup_box = None
        self.upload_popup_plus = None
        self.upload_popup_title = None
        self.upload_popup_hint = None
        self.upload_popup_panels = []

    def submit_input_from_popup(self) -> None:
        if self.submit_input():
            self.close_upload_popup()

    def refresh_upload_popup_text(self) -> None:
        if not self.upload_popup or not self.upload_popup.winfo_exists():
            return
        if self.upload_popup_title and self.upload_popup_title.winfo_exists():
            self.upload_popup_title.configure(text="Drag & drop files or folders here")
        if self.upload_popup_hint and self.upload_popup_hint.winfo_exists():
            self.upload_popup_hint.configure(text="or click to browse")

    def refresh_upload_popup_theme(self) -> None:
        if not self.upload_popup or not self.upload_popup.winfo_exists():
            return

        theme = self.theme
        self.upload_popup.configure(bg=theme.background)
        children = self.upload_popup.winfo_children()
        if children and isinstance(children[0], tk.Frame):
            shell = children[0]
            shell.configure(bg=theme.background)
            for child in shell.winfo_children():
                if isinstance(child, tk.Frame) and child not in self.upload_popup_panels:
                    child.configure(bg=theme.background)
        for panel in self.upload_popup_panels:
            panel.configure(bg=theme.panel, highlightbackground=theme.border, highlightcolor=theme.border)
            for child in panel.winfo_children():
                if isinstance(child, tk.Label):
                    if child.cget("text") == "UP":
                        child.configure(bg=theme.accent, fg=theme.accent_text)
                    elif child.cget("text") == "Choose the file or folder you want to convert, then submit it.":
                        child.configure(bg=theme.panel, fg=theme.muted)
                    else:
                        child.configure(bg=theme.panel, fg=theme.text)
                elif isinstance(child, tk.Frame):
                    child.configure(bg=theme.panel_alt, highlightbackground=theme.border, highlightcolor=theme.accent)
        if self.upload_popup_box and self.upload_popup_box.winfo_exists():
            self.upload_popup_box.configure(bg=theme.panel_alt, highlightbackground=theme.border, highlightcolor=theme.accent)
        if self.upload_popup_plus and self.upload_popup_plus.winfo_exists():
            self.upload_popup_plus.configure(bg=theme.panel_alt, fg=theme.accent)

    def show_status_popup(self) -> None:
        if self.status_popup and self.status_popup.winfo_exists():
            self.status_popup.lift()
            self.status_popup.focus_force()
            self.start_status_animation()
            return

        theme = self.theme
        popup = tk.Toplevel(self.root)
        self.status_popup = popup
        self.status_popup_panels = []
        popup.title("Conversion Status")
        popup.minsize(680, 500)
        popup.transient(self.root)
        popup.iconphoto(True, self.window_icon)
        popup.configure(bg=theme.background)
        popup.columnconfigure(0, weight=1)
        popup.rowconfigure(0, weight=1)
        popup.protocol("WM_DELETE_WINDOW", self.close_status_popup)

        shell = tk.Frame(popup, bg=theme.background, padx=26, pady=24)
        shell.grid(row=0, column=0, sticky="nsew")
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(2, weight=1)

        header = tk.Frame(shell, bg=theme.panel, highlightbackground=theme.border, highlightthickness=1, padx=18, pady=16)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(1, weight=1)
        self.status_popup_panels.append(header)
        tk.Label(
            header,
            text="FC",
            bg=theme.accent,
            fg=theme.accent_text,
            font=("Segoe UI", 13, "bold"),
            padx=12,
            pady=8,
        ).grid(row=0, column=0, rowspan=2, sticky="nsw", padx=(0, 14))
        tk.Label(header, text="Live conversion", bg=theme.panel, fg=theme.text, font=("Segoe UI", 18, "bold")).grid(
            row=0,
            column=1,
            sticky="w",
        )
        tk.Label(
            header,
            textvariable=self.status,
            bg=theme.panel,
            fg=theme.muted,
            font=("Segoe UI", 10),
            wraplength=440,
            justify="left",
        ).grid(
            row=1,
            column=1,
            sticky="w",
            pady=(4, 0),
        )
        tk.Label(
            header,
            text="WORKING",
            bg=theme.panel_alt,
            fg=theme.accent,
            font=("Segoe UI", 8, "bold"),
            padx=10,
            pady=4,
        ).grid(row=0, column=2, sticky="ne", padx=(14, 0))

        stages_panel = tk.Frame(shell, bg=theme.panel, highlightbackground=theme.border, highlightthickness=1, padx=14, pady=14)
        stages_panel.grid(row=1, column=0, sticky="ew", pady=(0, 14))
        stages_panel.columnconfigure(0, weight=1)
        self.status_popup_panels.append(stages_panel)

        self.status_stage_canvas = tk.Canvas(stages_panel, height=150, highlightthickness=0, bd=0, bg=theme.panel)
        self.status_stage_canvas.grid(row=0, column=0, sticky="ew")
        self.status_stage_canvas.bind("<Configure>", lambda event: self.draw_status_stages())

        progress_row = tk.Frame(stages_panel, bg=theme.panel)
        progress_row.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        progress_row.columnconfigure(0, weight=1)
        self.status_progress = ttk.Progressbar(
            progress_row,
            mode="determinate",
            maximum=100,
            variable=self.progress_value,
            style="Accent.Horizontal.TProgressbar",
        )
        self.status_progress.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        ttk.Label(progress_row, textvariable=self.progress_text, style="Muted.TLabel").grid(row=0, column=1, sticky="e")

        log_panel = tk.Frame(shell, bg=theme.panel, highlightbackground=theme.border, highlightthickness=1, padx=14, pady=14)
        log_panel.grid(row=2, column=0, sticky="nsew", pady=(0, 14))
        log_panel.columnconfigure(0, weight=1)
        log_panel.rowconfigure(1, weight=1)
        self.status_popup_panels.append(log_panel)
        tk.Label(log_panel, text="Activity", bg=theme.panel, fg=theme.text, font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.status_log = tk.Text(log_panel, height=7, wrap="word", bd=0, padx=12, pady=10, state="disabled")
        self.status_log.grid(row=1, column=0, sticky="nsew", pady=(8, 0))

        button_row = tk.Frame(shell, bg=theme.background)
        button_row.grid(row=3, column=0, sticky="ew")
        button_row.columnconfigure(0, weight=1)
        self.icon_button(button_row, "Cancel Conversion", "cancel", "Danger.TButton", self.request_cancel).grid(
            row=0,
            column=1,
            sticky="e",
        )

        self.refresh_status_popup_theme()
        self.refresh_status_log()
        self.draw_status_stages()
        self.start_status_animation()
        self.center_popup(popup, 680, 500)

    def center_popup(self, window: tk.Toplevel, width: int, height: int) -> None:
        self.root.update_idletasks()
        window.update_idletasks()

        parent_width = max(self.root.winfo_width(), 1)
        parent_height = max(self.root.winfo_height(), 1)
        parent_x = self.root.winfo_rootx()
        parent_y = self.root.winfo_rooty()

        if parent_width <= 1 or parent_height <= 1:
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2
        else:
            x = parent_x + (parent_width - width) // 2
            y = parent_y + (parent_height - height) // 2

        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = max(0, min(x, screen_width - width))
        y = max(0, min(y, screen_height - height))
        window.geometry(f"{width}x{height}+{x}+{y}")

    def refresh_status_popup_theme(self) -> None:
        if not self.status_popup or not self.status_popup.winfo_exists():
            return

        theme = self.theme
        self.status_popup.configure(bg=theme.background)
        popup_children = self.status_popup.winfo_children()
        if popup_children and isinstance(popup_children[0], tk.Frame):
            shell = popup_children[0]
            shell.configure(bg=theme.background)
            for child in shell.winfo_children():
                if isinstance(child, tk.Frame) and child not in self.status_popup_panels:
                    child.configure(bg=theme.background)
            shell_children = shell.winfo_children()
            if shell_children and isinstance(shell_children[0], tk.Frame):
                header_labels = [child for child in shell_children[0].winfo_children() if isinstance(child, tk.Label)]
                if header_labels:
                    header_labels[0].configure(bg=theme.background, fg=theme.text)
                if len(header_labels) > 1:
                    header_labels[1].configure(bg=theme.background, fg=theme.muted)
        for panel in self.status_popup_panels:
            panel.configure(bg=theme.panel, highlightbackground=theme.border, highlightcolor=theme.border)
            for child in panel.winfo_children():
                if isinstance(child, tk.Frame):
                    child.configure(bg=theme.panel)
                elif isinstance(child, tk.Label):
                    if child.cget("text") == "FC":
                        child.configure(bg=theme.accent, fg=theme.accent_text)
                    elif child.cget("text") == "WORKING":
                        child.configure(bg=theme.panel_alt, fg=theme.accent)
                    elif child.cget("textvariable"):
                        child.configure(bg=theme.panel, fg=theme.muted)
                    else:
                        child.configure(bg=theme.panel, fg=theme.text)
        if self.status_stage_canvas and self.status_stage_canvas.winfo_exists():
            self.status_stage_canvas.configure(bg=theme.panel)
        if self.status_log and self.status_log.winfo_exists():
            self.status_log.configure(
                bg=theme.log_background,
                fg=theme.log_text,
                insertbackground=theme.log_text,
                selectbackground=theme.selection,
                selectforeground=theme.text,
            )
        self.draw_status_stages()

    def draw_status_stages(self) -> None:
        canvas = self.status_stage_canvas
        if not canvas or not canvas.winfo_exists():
            return

        theme = self.theme
        canvas.delete("all")
        width = max(canvas.winfo_width(), 540)
        line_y = 48
        start_x = 48
        end_x = width - 48
        step_count = len(self.step_titles)
        spacing = (end_x - start_x) / max(step_count - 1, 1)

        for index in range(step_count - 1):
            x1 = start_x + index * spacing
            x2 = start_x + (index + 1) * spacing
            color = theme.accent if index < self.completed_step_count else theme.border
            canvas.create_line(x1, line_y, x2, line_y, fill=color, width=4)

        for index, title in enumerate(self.step_titles):
            x = start_x + index * spacing
            is_active = self.active_step_index == index
            is_done = index < self.completed_step_count
            fill = theme.accent if is_done or is_active else theme.panel_alt
            outline = theme.accent_hover if is_active else theme.border
            text_color = theme.accent_text if is_done or is_active else theme.muted
            label_color = theme.text if is_done or is_active else theme.muted
            radius = 16

            if is_active:
                pulse = 4 + (self.animation_frame % 8)
                canvas.create_oval(x - radius - pulse, line_y - radius - pulse, x + radius + pulse, line_y + radius + pulse, outline=theme.accent, width=2)

            canvas.create_oval(x - radius, line_y - radius, x + radius, line_y + radius, fill=fill, outline=outline, width=2)
            canvas.create_text(x, line_y, text=str(index + 1), fill=text_color, font=("Segoe UI", 10, "bold"))
            canvas.create_text(x, 86, text=title, width=96, fill=label_color, font=("Segoe UI", 8, "bold" if is_active else "normal"), justify="center", anchor="n")

        if self.current_step_detail:
            canvas.create_text(
                width / 2,
                132,
                text=self.current_step_detail,
                width=width - 72,
                fill=theme.muted,
                font=("Segoe UI", 9),
                justify="center",
            )

    def start_status_animation(self) -> None:
        if self.animation_job is None:
            self.animation_job = self.root.after(120, self.animate_status_popup)

    def animate_status_popup(self) -> None:
        self.animation_job = None
        if not self.status_popup or not self.status_popup.winfo_exists() or not self.is_converting:
            return
        self.animation_frame = (self.animation_frame + 1) % 32
        self.draw_status_stages()
        self.start_status_animation()

    def stop_status_animation(self) -> None:
        if self.animation_job is not None:
            try:
                self.root.after_cancel(self.animation_job)
            except tk.TclError:
                pass
            self.animation_job = None

    def close_status_popup(self) -> None:
        self.stop_status_animation()
        self.stop_busy_progress()
        if self.status_popup and self.status_popup.winfo_exists():
            self.status_popup.destroy()
        self.status_popup = None
        self.status_stage_canvas = None
        self.status_progress = None
        self.status_log = None
        self.status_popup_panels = []

    def refresh_status_log(self) -> None:
        if not self.status_log or not self.status_log.winfo_exists():
            return
        self.status_log.configure(state="normal")
        self.status_log.delete("1.0", "end")
        self.status_log.insert("end", "\n".join(self.status_log_lines))
        self.status_log.configure(state="disabled")
        self.status_log.see("end")

    def queue_step(self, active_index: int, detail: str, completed: int | None = None) -> None:
        completed_count = active_index if completed is None else completed
        progress_points = [12, 28, 55, 82, 100]
        self.root.after(0, lambda: self.render_timeline(active_index, completed_count, detail))
        self.root.after(0, lambda: self.status.set(detail))
        self.root.after(0, lambda: self.set_progress(progress_points[active_index], detail))
        self.root.after(0, lambda: self.write_log(f"Step {active_index + 1}: {detail}"))

    def set_progress(self, value: float, detail: str | None = None) -> None:
        self.stop_busy_progress()
        bounded = max(0, min(100, value))
        self.progress_value.set(bounded)
        self.progress_text.set(f"{int(bounded)}%")
        if detail:
            self.status.set(detail)

    def queue_progress(self, value: float, detail: str | None = None) -> None:
        self.root.after(0, lambda: self.set_progress(value, detail))

    def set_busy_progress(self, detail: str) -> None:
        self.status.set(detail)
        self.progress_text.set("Working")
        if self.status_progress and self.status_progress.winfo_exists():
            self.status_progress.configure(mode="indeterminate")
            self.status_progress.start(12)

    def queue_busy_progress(self, detail: str) -> None:
        self.root.after(0, lambda: self.set_busy_progress(detail))

    def stop_busy_progress(self) -> None:
        if self.status_progress and self.status_progress.winfo_exists():
            self.status_progress.stop()
            self.status_progress.configure(mode="determinate", variable=self.progress_value)

    def show_toast(self, title: str, message: str, duration: int = 2200) -> None:
        if self.toast_window and self.toast_window.winfo_exists():
            self.toast_window.destroy()

        theme = self.theme
        toast = tk.Toplevel(self.root)
        self.toast_window = toast
        toast.title(title)
        toast.resizable(False, False)
        toast.transient(self.root)
        toast.configure(bg=theme.panel)
        toast.attributes("-topmost", True)

        frame = tk.Frame(toast, bg=theme.panel, highlightbackground=theme.border, highlightthickness=1, padx=16, pady=12)
        frame.grid(row=0, column=0, sticky="nsew")
        tk.Label(frame, text=title, bg=theme.panel, fg=theme.text, font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(frame, text=message, bg=theme.panel, fg=theme.muted, font=("Segoe UI", 9), wraplength=300, justify="left").grid(
            row=1,
            column=0,
            sticky="w",
            pady=(4, 0),
        )

        self.root.update_idletasks()
        toast.update_idletasks()
        self.center_popup(toast, toast.winfo_reqwidth(), toast.winfo_reqheight())
        toast.after(duration, toast.destroy)

    def dialog_parent(self) -> tk.Misc:
        if self.editor_popup and self.editor_popup.winfo_exists():
            return self.editor_popup
        if self.upload_popup and self.upload_popup.winfo_exists():
            return self.upload_popup
        if self.status_popup and self.status_popup.winfo_exists():
            return self.status_popup
        return self.root

    def show_error(self, title: str, message: str) -> None:
        messagebox.showerror(title, message, parent=self.dialog_parent())

    def show_info(self, title: str, message: str) -> None:
        messagebox.showinfo(title, message, parent=self.dialog_parent())

    def update_target_choices(self, clear_paths: bool = True) -> None:
        targets = self.target_labels_for_source(self.from_format.get())
        self.to_menu.configure(values=targets)
        if self.to_format.get() not in targets:
            self.to_format.set(targets[0] if targets else "")
        self.update_labels(clear_paths=clear_paths)

    def input_type_hint(self) -> str:
        if not self.spec.input_extensions:
            return "folder contents"
        if self.spec.input_extensions == TEXT_DOCUMENT_EXTENSIONS:
            return TEXT_EDIT_SUMMARY
        return ", ".join(self.spec.input_extensions)

    def uses_folder_source(self) -> bool:
        return self.spec.source_kind == "folder"

    def uses_batch_folder_mode(self) -> bool:
        return self.mode.get() == "folder" and not self.uses_folder_source()

    def input_file_phrase(self) -> str:
        label = self.spec.input_label
        article = "an" if label[:1].lower() in "aeiou" else "a"
        return f"{article} {label}"

    def update_labels(self, clear_paths: bool = True) -> None:
        spec = self.spec
        if self.uses_folder_source():
            self.mode.set("folder")
            self.upload_title.configure(text="Add folder")
            self.upload_hint.configure(text="Open the upload popup to choose a folder to compress into a ZIP archive.")
            self.rename_output.set(False)
            self.status.set("Ready for folder upload")
        elif self.mode.get() == "folder":
            self.upload_title.configure(text=f"Add {spec.input_label} folder")
            self.upload_hint.configure(text=f"Open the upload popup to choose a folder containing {self.input_type_hint()} files.")
            self.rename_output.set(False)
            self.status.set("Ready for folder upload")
        else:
            self.upload_title.configure(text=f"Add {spec.input_label} file")
            self.upload_hint.configure(text=f"Open the upload popup to choose {self.input_file_phrase()} file.")
            self.status.set("Ready")

        encoding_state = "normal" if spec.supports_encoding else "disabled"
        delimiter_state = "normal" if spec.supports_delimiter else "disabled"
        if hasattr(self, "encoding_entry"):
            self.encoding_entry.configure(state=encoding_state)
        if hasattr(self, "delimiter_entry"):
            self.delimiter_entry.configure(state=delimiter_state)
        if clear_paths:
            self.input_path.set("")
            self.output_folder.set("")
            self.custom_name.set("")
            self.input_submitted = False
            self.render_timeline()
        self.update_upload_display()
        self.update_rename_controls()

    def update_upload_display(self) -> None:
        path = self.input_path.get().strip()
        if path:
            state = "submitted" if self.input_submitted else "selected"
            self.upload_hint.configure(text=f"{state.title()}: {self.display_path(Path(path))}")
            self.upload_status_text.set(self.display_path(Path(path), max_length=78))
        elif self.uses_folder_source():
            self.upload_hint.configure(text="Open the upload popup to choose a folder to compress into a ZIP archive.")
            self.upload_status_text.set("No file selected yet")
        elif self.mode.get() == "folder":
            self.upload_hint.configure(text=f"Open the upload popup to choose a folder containing {self.input_type_hint()} files.")
            self.upload_status_text.set("No file selected yet")
        else:
            self.upload_hint.configure(text=f"Open the upload popup to choose {self.input_file_phrase()} file.")
            self.upload_status_text.set("No file selected yet")
        self.refresh_upload_popup_text()
        self.update_suggested_name()
        self.update_editor_status()

    def display_path(self, path: Path, max_length: int = 92) -> str:
        text = str(path)
        if len(text) <= max_length:
            return text
        return f"...{text[-(max_length - 3):]}"

    def update_rename_controls(self) -> None:
        folder_mode = self.mode.get() == "folder"
        if folder_mode:
            self.rename_output.set(False)
        state = "normal" if self.rename_output.get() and not folder_mode else "disabled"
        self.custom_name_entry.configure(state=state)
        self.rename_yes.configure(state="disabled" if folder_mode else "normal")
        self.update_suggested_name()

    def update_suggested_name(self) -> None:
        input_text = self.input_path.get().strip()
        if not input_text:
            self.suggested_name.set("Suggested output: choose a file first")
            return
        path = Path(input_text)
        if self.uses_folder_source():
            self.suggested_name.set(f"Suggested output: {self.default_output_name(path).name}")
            return
        if self.mode.get() == "folder":
            self.suggested_name.set(f'Folder output: each file will use "Original Name ({self.spec.output_tag}){self.spec.output_extension}"')
            return
        self.suggested_name.set(f"Suggested output: {self.default_output_name(path).name}")

    def filetypes(self) -> list[tuple[str, str]]:
        patterns = " ".join(f"*{extension}" for extension in self.spec.input_extensions)
        label = "Editable files" if self.spec.input_extensions == TEXT_DOCUMENT_EXTENSIONS else f"{self.spec.input_label.title()} files"
        return [(label, patterns), ("All files", "*.*")]

    def set_selected_input(self, path: str) -> None:
        self.input_path.set(path)
        self.input_submitted = False
        self.status.set("Upload selected. Press Submit.")
        self.set_progress(5, "Upload selected")
        self.render_timeline()
        self.update_upload_display()

    def choose_file_input(self) -> None:
        if self.is_converting:
            return
        if self.uses_folder_source():
            self.show_error("Folder required", "This conversion creates a ZIP archive, so choose a folder instead.")
            return
        self.mode.set("file")
        self.update_labels(clear_paths=False)
        path = filedialog.askopenfilename(
            title=f"Choose {self.spec.input_label} file",
            filetypes=self.filetypes(),
        )
        if path:
            self.set_selected_input(path)

    def choose_folder_input(self) -> None:
        if self.is_converting:
            return
        self.mode.set("folder")
        self.update_labels(clear_paths=False)
        title = "Choose folder to zip" if self.uses_folder_source() else f"Choose folder containing {self.spec.input_label} files"
        path = filedialog.askdirectory(title=title)
        if path:
            self.set_selected_input(path)

    def choose_input(self) -> None:
        if self.is_converting:
            return
        if self.mode.get() == "folder" or self.uses_folder_source():
            title = "Choose folder to zip" if self.uses_folder_source() else f"Choose folder containing {self.spec.input_label} files"
            path = filedialog.askdirectory(title=title)
        else:
            path = filedialog.askopenfilename(
                title=f"Choose {self.spec.input_label} file",
                filetypes=self.filetypes(),
            )
        if path:
            self.set_selected_input(path)

    def submit_input(self) -> bool:
        path_text = self.input_path.get().strip()
        if not path_text:
            self.show_error("Missing upload", "Open the upload popup and choose a file or folder first.")
            return False

        path = Path(path_text)
        if self.mode.get() == "folder" or self.uses_folder_source():
            if not path.is_dir():
                self.show_error("Invalid folder", "Please choose a folder.")
                return False
        else:
            if not path.is_file():
                self.show_error("Invalid file", "Please choose a file for single-file conversion.")
                return False
            if path.suffix.lower() not in self.spec.input_extensions:
                expected = ", ".join(self.spec.input_extensions)
                self.show_error("Wrong file type", f"Please choose one of these file types: {expected}")
                return False

        self.input_submitted = True
        self.status.set("Upload submitted")
        self.update_upload_display()
        self.set_progress(12, "Upload submitted")
        self.render_timeline(active_index=None, completed=1, detail=f"Submitted {path.name}")
        self.write_log(f"Submitted upload: {path}")
        return True

    def choose_output_folder(self) -> None:
        path = filedialog.askdirectory(title="Choose output folder")
        if path:
            self.output_folder.set(path)
            self.set_progress(max(self.progress_value.get(), 18), "Output folder selected")

    def default_output_name(self, source: Path) -> Path:
        return Path(f"{source.stem} ({self.spec.output_tag}){self.spec.output_extension}")

    def sanitized_custom_output_name(self) -> str:
        name = self.custom_name.get().strip()
        if not name:
            raise ConversionError("Type a custom output name, or choose No for rename.")
        cleaned = "".join("_" if char in INVALID_FILENAME_CHARS else char for char in name).strip()
        if not cleaned:
            raise ConversionError("The custom output name is not valid.")
        custom_path = Path(cleaned)
        if custom_path.suffix.lower() != self.spec.output_extension:
            cleaned = f"{custom_path.stem}{self.spec.output_extension}"
        return cleaned

    def output_directory_for(self, source: Path) -> Path:
        output_text = self.output_folder.get().strip()
        if output_text:
            return Path(output_text)
        if source.is_dir() and self.uses_folder_source():
            return source.parent
        return source if source.is_dir() else source.parent

    def output_path_for(self, source: Path) -> Path:
        output_directory = self.output_directory_for(source)
        if self.rename_output.get() and self.mode.get() != "folder":
            return output_directory / self.sanitized_custom_output_name()
        return output_directory / self.default_output_name(source)

    def start_conversion(self) -> None:
        if not self.input_submitted:
            self.show_error("Submit upload", "Open the upload popup, choose a file or folder, then click Submit Upload first.")
            return

        self.cancel_event.clear()
        self.is_converting = True
        self.status_log_lines.clear()
        self.convert_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.set_progress(0, "Starting")
        self.render_timeline(active_index=0, completed=0, detail="Starting conversion")
        self.show_status_popup()
        self.write_log("Conversion started.")

        thread = threading.Thread(target=self.convert, daemon=True)
        thread.start()

    def request_cancel(self) -> None:
        if not self.is_converting:
            return
        self.cancel_event.set()
        self.status.set("Cancel requested")
        self.write_log("Cancel requested. The current file will finish if it is already being written.")

    def convert_one(self, input_path: Path, output_path: Path | None) -> Path:
        self.queue_busy_progress(
            f"Running converter for {input_path.name}. Large files can take a few minutes."
        )
        self.root.after(0, lambda path=input_path: self.write_log(f"Converter engine running: {path.name}"))
        return self.spec.convert(
            input_path,
            output_path,
            self.encoding.get().strip() or "utf-8-sig",
            self.delimiter.get() or None,
            self.overwrite.get(),
        )

    def convert_folder(self, input_directory: Path, output_directory: Path) -> list[Path]:
        if not input_directory.is_dir():
            raise ConversionError(f"Input must be a folder: {input_directory}")

        files = list(files_in_directory(input_directory, self.spec.input_extensions))
        if not files:
            raise ConversionError(f"No {self.spec.input_label} files were found in the selected folder.")

        converted: list[Path] = []
        total = len(files)
        for index, input_file in enumerate(files, start=1):
            if self.cancel_event.is_set():
                raise ConversionError(f"Conversion cancelled after {len(converted)} of {total} files.")
            base_progress = 35 + ((index - 1) / total) * 48
            self.queue_step(2, f"Reading source file {index} of {total}: {input_file.name}")
            self.queue_progress(base_progress, f"Reading file {index} of {total}")
            output_file = output_directory / self.default_output_name(input_file)
            self.queue_step(3, f"Converting file {index} of {total}: {output_file.name}")
            self.queue_progress(base_progress + (24 / total), f"Preparing converter for file {index} of {total}")
            converted.append(self.convert_one(input_file, output_file))
            self.queue_progress(35 + (index / total) * 48, f"Completed file {index} of {total}")
            self.root.after(0, lambda path=output_file: self.write_log(f"Created: {path}"))
        return converted

    def convert(self) -> None:
        try:
            input_path = Path(self.input_path.get())
            self.queue_step(0, "Checking submitted upload", completed=1)
            if self.cancel_event.is_set():
                raise ConversionError("Conversion cancelled before processing started.")

            self.queue_step(1, "Preparing output name and folder")
            if self.uses_folder_source():
                output_path = self.output_path_for(input_path)
                self.root.after(0, lambda: self.write_log(f"Output file: {output_path}"))
                self.queue_step(2, f"Reading folder: {input_path.name}")
                if self.cancel_event.is_set():
                    raise ConversionError("Conversion cancelled before writing the archive.")
                self.queue_step(3, f"Creating {self.spec.output_tag}: {output_path.name}")
                created = self.convert_one(input_path, output_path)
                message = f"Created:\n{created}"
                outputs = [created]
            elif self.mode.get() == "folder":
                output_directory = self.output_directory_for(input_path)
                self.root.after(0, lambda: self.write_log(f"Output folder: {output_directory}"))
                converted = self.convert_folder(input_path, output_directory)
                message = "Created:\n" + "\n".join(str(path) for path in converted)
                outputs = converted
            else:
                output_path = self.output_path_for(input_path)
                self.root.after(0, lambda: self.write_log(f"Output file: {output_path}"))
                if self.cancel_event.is_set():
                    raise ConversionError("Conversion cancelled before the source file was read.")
                self.queue_step(2, f"Reading source: {input_path.name}")
                if self.cancel_event.is_set():
                    raise ConversionError("Conversion cancelled before writing the output file.")
                self.queue_step(3, f"Converting to {self.spec.output_tag}: {output_path.name}")
                created = self.convert_one(input_path, output_path)
                message = f"Created:\n{created}"
                outputs = [created]

            if self.cancel_event.is_set():
                message += "\n\nCancel was requested, but the active file had already completed."
            self.queue_step(4, "Conversion complete", completed=5)
            self.finish(message, success=True, outputs=outputs)
        except ConversionError as error:
            if self.cancel_event.is_set() or "cancelled" in str(error).lower():
                self.finish(str(error), success=False, cancelled=True)
            else:
                self.finish(str(error), success=False)
        except Exception as error:
            self.finish(f"Unexpected error: {error}", success=False)

    def finish(self, message: str, success: bool, cancelled: bool = False, outputs: list[Path] | None = None) -> None:
        def update_ui() -> None:
            self.write_log(message)
            self.stop_busy_progress()
            self.stop_status_animation()
            if cancelled:
                self.status.set("Cancelled")
                self.render_timeline(detail="Conversion cancelled")
                self.set_progress(self.progress_value.get(), "Cancelled")
            else:
                self.status.set("Done" if success else "Failed")
                if success:
                    self.last_outputs = outputs or []
                    self.set_progress(100, "Download ready")
                    self.render_timeline(completed=5, detail="All steps complete")
            self.convert_button.configure(state="normal")
            self.cancel_button.configure(state="disabled")
            self.is_converting = False
            if success:
                self.write_log("Conversion complete. Opening download popup.")
                self.root.after(700, lambda: (self.close_status_popup(), self.show_download_page(self.last_outputs)))
            else:
                self.show_error("Conversion stopped" if cancelled else "Conversion failed", message)

        self.root.after(0, update_ui)

    def show_download_page(self, outputs: list[Path]) -> None:
        if not outputs:
            self.show_info("Download ready", "The conversion finished, but no output path was reported.")
            return

        theme = self.theme
        page = tk.Toplevel(self.root)
        page.title("Conversion Complete")
        page.minsize(760, 540)
        page.transient(self.root)
        page.iconphoto(True, self.window_icon)
        page.configure(bg=theme.background)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(0, weight=1)

        shell = tk.Frame(page, bg=theme.background, padx=30, pady=28)
        shell.grid(row=0, column=0, sticky="nsew")
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(2, weight=1)

        close_button = tk.Button(shell, text="x", bd=0, bg=theme.background, fg=theme.muted, font=("Segoe UI", 16), command=page.destroy)
        close_button.grid(row=0, column=0, sticky="ne")
        header = tk.Frame(shell, bg=theme.background)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
        header.columnconfigure(0, weight=1)
        tk.Label(
            header,
            text="✓",
            bg=theme.accent,
            fg=theme.accent_text,
            font=("Segoe UI", 26, "bold"),
            width=3,
            height=1,
        ).grid(row=0, column=0, sticky="n")
        tk.Label(header, text="Conversion complete", bg=theme.background, fg=theme.text, font=("Segoe UI", 22, "bold")).grid(
            row=1,
            column=0,
            sticky="n",
            pady=(14, 0),
        )
        tk.Label(
            header,
            text="Your files have been converted and are ready to download.",
            bg=theme.background,
            fg=theme.muted,
            font=("Segoe UI", 10),
        ).grid(row=2, column=0, sticky="n", pady=(8, 0))

        summary = tk.Frame(shell, bg=theme.panel, highlightbackground=theme.border, highlightthickness=1, padx=16, pady=12)
        summary.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        for column in range(3):
            summary.columnconfigure(column, weight=1)

        first_output = outputs[0]
        output_format = first_output.suffix.upper().lstrip(".") or "Folder"
        output_location = str(first_output.parent if first_output.is_file() else first_output)
        summary_items = (
            (f"{len(outputs)} file{'s' if len(outputs) != 1 else ''} converted", "Ready to open", "file"),
            (f"{output_format}", "Output format", "file"),
            (self.display_path(Path(output_location), max_length=32), "Saved location", "folder"),
        )
        for column, (title, subtitle, icon_name) in enumerate(summary_items):
            item = tk.Frame(summary, bg=theme.panel)
            item.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 12, 0))
            item.columnconfigure(1, weight=1)
            self.icon_label(item, icon_name).grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 10))
            ttk.Label(item, text=title, style="Card.TLabel").grid(row=0, column=1, sticky="w")
            ttk.Label(item, text=subtitle, style="Muted.TLabel").grid(row=1, column=1, sticky="w", pady=(3, 0))

        content = tk.Frame(shell, bg=theme.panel, highlightbackground=theme.border, highlightthickness=1, padx=16, pady=14)
        content.grid(row=2, column=0, sticky="nsew", pady=(0, 14))
        content.columnconfigure(0, weight=1)
        tk.Label(content, text="Converted files", bg=theme.panel, fg=theme.text, font=("Segoe UI", 10, "bold")).grid(
            row=0,
            column=0,
            sticky="w",
            pady=(0, 8),
        )

        def size_label(path: Path) -> str:
            if path.is_dir():
                return "Folder"
            try:
                size = path.stat().st_size
            except OSError:
                return "Unknown"
            if size >= 1024 * 1024:
                return f"{size / (1024 * 1024):.1f} MB"
            if size >= 1024:
                return f"{size / 1024:.1f} KB"
            return f"{size} B"

        for index, output in enumerate(outputs[:5], start=1):
            row = tk.Frame(content, bg=theme.panel, highlightbackground=theme.border, highlightthickness=1, padx=12, pady=10)
            row.grid(row=index, column=0, sticky="ew", pady=(0, 8))
            row.columnconfigure(1, weight=1)
            self.icon_label(row, "file").grid(row=0, column=0, sticky="w", padx=(0, 10))
            ttk.Label(row, text=output.name, style="Card.TLabel").grid(row=0, column=1, sticky="w")
            ttk.Label(row, text=size_label(output), style="Muted.TLabel").grid(row=0, column=2, sticky="e", padx=(12, 18))
            ttk.Label(row, text="Completed", style="StatusGood.TLabel").grid(row=0, column=3, sticky="e", padx=(0, 18))
            self.icon_button(row, "Open", "open", "Ghost.TButton", lambda path=output: self.open_path(path)).grid(row=0, column=4, sticky="e")

        button_row = tk.Frame(shell, bg=theme.background)
        button_row.grid(row=3, column=0, sticky="ew")
        button_row.columnconfigure(3, weight=1)

        def selected_output() -> Path:
            return outputs[0]

        self.icon_button(button_row, "Open File", "open", "Accent.TButton", lambda: self.open_path(selected_output())).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 10),
        )
        self.icon_button(
            button_row,
            "Open Folder",
            "folder",
            "Secondary.TButton",
            lambda: self.open_path(selected_output().parent),
        ).grid(row=0, column=1, sticky="w", padx=(0, 10))
        self.icon_button(
            button_row,
            "New Conversion",
            "convert",
            "Ghost.TButton",
            lambda: self.reset_after_download(page),
        ).grid(row=0, column=2, sticky="w")

        page.attributes("-topmost", True)
        self.center_popup(page, 760, 540)
        page.after(500, lambda: page.attributes("-topmost", False))

    def open_path(self, path: Path) -> None:
        try:
            if not path.exists():
                self.show_error("Missing output", f"This path no longer exists:\n{path}")
                return
            os.startfile(str(path))
        except OSError as error:
            self.show_error("Could not open", f"Windows could not open this path:\n{path}\n\n{error}")

    def reset_after_download(self, window: tk.Toplevel) -> None:
        window.destroy()
        self.input_path.set("")
        self.output_folder.set("")
        self.custom_name.set("")
        self.input_submitted = False
        self.last_outputs = []
        self.set_progress(0, "Ready")
        self.render_timeline()
        self.update_upload_display()

    def write_log(self, message: str) -> None:
        self.status_log_lines.append(message)
        if len(self.status_log_lines) > 80:
            self.status_log_lines = self.status_log_lines[-80:]
        self.refresh_status_log()


def main() -> int:
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
