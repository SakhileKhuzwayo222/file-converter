from __future__ import annotations

import os
import threading
import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

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
from .office import convert_pdf_to_word, convert_text_to_word


ConverterFunction = Callable[[Path, Path | None, str, str | None, bool], Path]
INVALID_FILENAME_CHARS = '<>:"/\\|?*'


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
        from_label="Text or Markdown (.txt, .md)",
        to_label="Word Document (.docx)",
        input_label="text",
        input_extensions=(".txt", ".md"),
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


class ConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("File Converter")
        self.root.minsize(940, 720)

        self.style = ttk.Style()
        self.theme_name = tk.StringVar(value="light")
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
        self.status_popup: tk.Toplevel | None = None
        self.status_stage_canvas: tk.Canvas | None = None
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
            "Converted file written",
            "Finished",
        ]

        self.tk_backgrounds: list[tuple[tk.Widget, str]] = []
        self.card_frames: list[tk.Frame] = []
        self.window_icon = self.create_window_icon()
        self.root.iconphoto(True, self.window_icon)

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
        icon = tk.PhotoImage(width=64, height=64)
        icon.put("#0f766e", to=(12, 6, 52, 58))
        icon.put("#0f766e", to=(7, 12, 57, 52))
        icon.put("#14b8a6", to=(12, 10, 52, 50))
        icon.put("#14b8a6", to=(10, 14, 54, 48))
        icon.put("#115e59", to=(12, 49, 52, 58))
        icon.put("#0b4f4a", to=(47, 14, 57, 52))

        icon.put("#ffffff", to=(18, 17, 24, 47))
        icon.put("#ffffff", to=(18, 17, 36, 23))
        icon.put("#ffffff", to=(18, 30, 33, 36))

        icon.put("#ffffff", to=(37, 17, 51, 23))
        icon.put("#ffffff", to=(37, 41, 51, 47))
        icon.put("#ffffff", to=(34, 20, 40, 44))
        icon.put("#14b8a6", to=(45, 24, 52, 40))
        return icon

    def register_background(self, widget: tk.Widget, role: str) -> None:
        self.tk_backgrounds.append((widget, role))

    def build_layout(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.shell = tk.Frame(self.root)
        self.shell.grid(row=0, column=0, sticky="nsew")
        self.shell.columnconfigure(0, weight=1)
        self.shell.rowconfigure(0, weight=1)
        self.register_background(self.shell, "background")

        self.canvas = tk.Canvas(self.shell, highlightthickness=0, bd=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.register_background(self.canvas, "background")

        self.scrollbar = ttk.Scrollbar(self.shell, orient="vertical", command=self.canvas.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.main = tk.Frame(self.canvas, padx=34, pady=28)
        self.main_window = self.canvas.create_window((0, 0), window=self.main, anchor="nw")
        self.main.columnconfigure(0, weight=1)
        self.register_background(self.main, "background")
        self.main.bind("<Configure>", self.update_scroll_region)
        self.canvas.bind("<Configure>", self.resize_canvas_content)
        self.root.bind_all("<MouseWheel>", self.on_page_mousewheel, add="+")
        self.root.bind_all("<Button-4>", self.on_page_mousewheel, add="+")
        self.root.bind_all("<Button-5>", self.on_page_mousewheel, add="+")

        self.build_header()
        self.build_conversion_card()
        self.build_output_card()
        self.build_options_card()
        self.build_action_area()

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
        self.theme_button = ttk.Button(header, text="Dark mode", style="Ghost.TButton", command=self.toggle_theme)
        self.theme_button.grid(row=0, column=2, rowspan=2, sticky="e")

    def make_card(self, row: int, title: str) -> tk.Frame:
        card = tk.Frame(self.main, bd=0, highlightthickness=1)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 16))
        card.columnconfigure(0, weight=1)
        self.card_frames.append(card)

        ttk.Label(card, text=title, style="Section.TLabel").grid(row=0, column=0, sticky="w", padx=22, pady=(18, 10))
        body = tk.Frame(card)
        body.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 20))
        body.columnconfigure(1, weight=1)
        self.register_background(body, "panel")
        return body

    def build_conversion_card(self) -> None:
        body = self.make_card(1, "Select file type")
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
            text="Single file",
            value="file",
            variable=self.mode,
            style="Card.TRadiobutton",
            command=self.update_labels,
        ).grid(row=0, column=0, padx=(0, 18), sticky="w")
        ttk.Radiobutton(
            mode_row,
            text="Folder of files",
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

        self.upload_popup_button = ttk.Button(
            self.upload_summary_box,
            text="Open Upload",
            style="Secondary.TButton",
            command=self.show_upload_popup,
        )
        self.upload_popup_button.grid(row=0, column=2, rowspan=2, sticky="e", padx=(16, 0))

        for widget in (self.upload_summary_box, self.upload_icon_label, self.upload_title, self.upload_hint):
            widget.bind("<Button-1>", lambda event: self.show_upload_popup())

    def build_output_card(self) -> None:
        body = self.make_card(2, "Output")

        ttk.Label(body, text="Save folder", style="Card.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 12), pady=(0, 10))
        self.output_entry = ttk.Entry(body, textvariable=self.output_folder)
        self.output_entry.grid(row=0, column=1, sticky="ew", pady=(0, 10))
        self.output_button = ttk.Button(body, text="Browse", style="Secondary.TButton", command=self.choose_output_folder)
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

    def build_options_card(self) -> None:
        body = self.make_card(3, "Options")
        body.columnconfigure(3, weight=1)

        self.encoding_label = ttk.Label(body, text="Encoding", style="Card.TLabel")
        self.encoding_label.grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.encoding_entry = ttk.Entry(body, textvariable=self.encoding, width=18)
        self.encoding_entry.grid(row=0, column=1, sticky="w", padx=(0, 20))

        self.delimiter_label = ttk.Label(body, text="CSV delimiter", style="Card.TLabel")
        self.delimiter_label.grid(row=0, column=2, sticky="w", padx=(0, 10))
        self.delimiter_entry = ttk.Entry(body, textvariable=self.delimiter, width=10)
        self.delimiter_entry.grid(row=0, column=3, sticky="w")

        self.overwrite_check = ttk.Checkbutton(
            body,
            text="Overwrite existing output files",
            variable=self.overwrite,
            style="Card.TCheckbutton",
        )
        self.overwrite_check.grid(row=1, column=0, columnspan=4, sticky="w", pady=(12, 0))

    def build_action_area(self) -> None:
        action_row = tk.Frame(self.main, bd=0, highlightthickness=1, padx=20, pady=18)
        action_row.grid(row=4, column=0, sticky="ew", pady=(0, 16))
        action_row.columnconfigure(0, weight=1)
        self.card_frames.append(action_row)

        self.convert_button = ttk.Button(action_row, text="Convert", style="Accent.TButton", command=self.start_conversion)
        self.convert_button.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.cancel_button = ttk.Button(action_row, text="Cancel", style="Danger.TButton", command=self.request_cancel, state="disabled")
        self.cancel_button.grid(row=0, column=1, sticky="ew")

    def toggle_theme(self) -> None:
        self.theme_name.set("dark" if self.theme_name.get() == "light" else "light")
        self.apply_theme()

    def apply_theme(self) -> None:
        theme = self.theme
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
        self.style.configure("HeaderSubtitle.TLabel", font=("Segoe UI", 10), background=theme.panel, foreground=theme.muted)
        self.style.configure("Subtitle.TLabel", font=("Segoe UI", 10), background=theme.background, foreground=theme.muted)
        self.style.configure("Section.TLabel", font=("Segoe UI", 12, "bold"), background=theme.panel, foreground=theme.text)
        self.style.configure("Card.TLabel", background=theme.panel, foreground=theme.text)
        self.style.configure("FieldLabel.TLabel", font=("Segoe UI", 9, "bold"), background=theme.panel, foreground=theme.muted)
        self.style.configure("Muted.TLabel", background=theme.panel, foreground=theme.muted)
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
        self.theme_button.configure(text="Light mode" if self.theme_name.get() == "dark" else "Dark mode")
        self.refresh_upload_popup_theme()
        self.refresh_status_popup_theme()
        self.render_timeline()

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

    def draw_logo(self) -> None:
        theme = self.theme
        canvas = self.logo_canvas
        canvas.delete("all")
        canvas.configure(bg=theme.panel)
        canvas.create_oval(5, 5, 21, 21, fill="#14b8a6", outline="")
        canvas.create_oval(33, 5, 49, 21, fill="#14b8a6", outline="")
        canvas.create_oval(5, 33, 21, 49, fill="#115e59", outline="")
        canvas.create_oval(33, 33, 49, 49, fill="#0b4f4a", outline="")
        canvas.create_rectangle(13, 5, 41, 49, fill="#14b8a6", outline="")
        canvas.create_rectangle(5, 13, 49, 41, fill="#14b8a6", outline="")
        canvas.create_rectangle(40, 13, 49, 49, fill="#0b4f4a", outline="")
        canvas.create_rectangle(9, 41, 49, 49, fill="#115e59", outline="")

        canvas.create_rectangle(15, 14, 20, 40, fill="#ffffff", outline="")
        canvas.create_rectangle(15, 14, 31, 19, fill="#ffffff", outline="")
        canvas.create_rectangle(15, 25, 29, 30, fill="#ffffff", outline="")

        canvas.create_rectangle(34, 14, 46, 19, fill="#ffffff", outline="")
        canvas.create_rectangle(34, 35, 46, 40, fill="#ffffff", outline="")
        canvas.create_rectangle(31, 17, 36, 38, fill="#ffffff", outline="")
        canvas.create_rectangle(42, 20, 49, 34, fill="#14b8a6", outline="")

    def render_timeline(self, active_index: int | None = None, completed: int = 0, detail: str = "") -> None:
        self.active_step_index = active_index
        self.completed_step_count = completed
        self.current_step_detail = detail
        self.draw_status_stages()

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
        popup.minsize(560, 360)
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

        self.upload_popup_box = tk.Frame(body, bd=0, highlightthickness=1, padx=20, pady=22, cursor="hand2")
        self.upload_popup_box.grid(row=0, column=0, sticky="ew")
        self.upload_popup_box.columnconfigure(1, weight=1)

        self.upload_popup_plus = tk.Label(self.upload_popup_box, text="+", font=("Segoe UI", 42, "bold"), cursor="hand2")
        self.upload_popup_plus.grid(row=0, column=0, rowspan=2, padx=(0, 18))

        self.upload_popup_title = ttk.Label(self.upload_popup_box, text=self.upload_title.cget("text"), style="UploadTitle.TLabel")
        self.upload_popup_title.grid(row=0, column=1, sticky="sw")
        self.upload_popup_hint = ttk.Label(self.upload_popup_box, text=self.upload_hint.cget("text"), style="UploadHint.TLabel")
        self.upload_popup_hint.grid(row=1, column=1, sticky="nw", pady=(4, 0))

        for widget in (self.upload_popup_box, self.upload_popup_plus, self.upload_popup_title, self.upload_popup_hint):
            widget.bind("<Button-1>", lambda event: self.choose_input())

        button_row = tk.Frame(shell, bg=theme.background)
        button_row.grid(row=2, column=0, sticky="ew")
        button_row.columnconfigure(0, weight=1)
        ttk.Button(button_row, text="Cancel", style="Ghost.TButton", command=self.close_upload_popup).grid(
            row=0,
            column=1,
            sticky="e",
            padx=(0, 10),
        )
        ttk.Button(button_row, text="Submit Upload", style="Accent.TButton", command=self.submit_input_from_popup).grid(
            row=0,
            column=2,
            sticky="e",
        )

        self.refresh_upload_popup_theme()
        self.refresh_upload_popup_text()
        self.center_popup(popup, 560, 360)
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
            self.upload_popup_title.configure(text=self.upload_title.cget("text"))
        if self.upload_popup_hint and self.upload_popup_hint.winfo_exists():
            self.upload_popup_hint.configure(text=self.upload_hint.cget("text"))

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
        ttk.Progressbar(
            progress_row,
            mode="determinate",
            maximum=100,
            variable=self.progress_value,
            style="Accent.Horizontal.TProgressbar",
        ).grid(row=0, column=0, sticky="ew", padx=(0, 10))
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
        ttk.Button(button_row, text="Cancel Conversion", style="Danger.TButton", command=self.request_cancel).grid(row=0, column=1, sticky="e")

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
        if self.status_popup and self.status_popup.winfo_exists():
            self.status_popup.destroy()
        self.status_popup = None
        self.status_stage_canvas = None
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
        bounded = max(0, min(100, value))
        self.progress_value.set(bounded)
        self.progress_text.set(f"{int(bounded)}%")
        if detail:
            self.status.set(detail)

    def queue_progress(self, value: float, detail: str | None = None) -> None:
        self.root.after(0, lambda: self.set_progress(value, detail))

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
        x = self.root.winfo_rootx() + max(self.root.winfo_width() - toast.winfo_reqwidth() - 28, 20)
        y = self.root.winfo_rooty() + 78
        toast.geometry(f"+{x}+{y}")
        toast.after(duration, toast.destroy)

    def update_target_choices(self, clear_paths: bool = True) -> None:
        targets = self.target_labels_for_source(self.from_format.get())
        self.to_menu.configure(values=targets)
        if self.to_format.get() not in targets:
            self.to_format.set(targets[0] if targets else "")
        self.update_labels(clear_paths=clear_paths)

    def update_labels(self, clear_paths: bool = True) -> None:
        spec = self.spec
        if self.mode.get() == "folder":
            self.upload_title.configure(text=f"Add {spec.input_label} folder")
            self.upload_hint.configure(text=f"Open the upload popup to choose a folder containing {', '.join(spec.input_extensions)} files.")
            self.rename_output.set(False)
            self.status.set("Ready for folder upload")
        else:
            self.upload_title.configure(text=f"Add {spec.input_label} file")
            self.upload_hint.configure(text=f"Open the upload popup to choose a {spec.input_label} file.")
            self.status.set("Ready")

        encoding_state = "normal" if spec.supports_encoding else "disabled"
        delimiter_state = "normal" if spec.supports_delimiter else "disabled"
        self.encoding_entry.configure(state=encoding_state)
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
        elif self.mode.get() == "folder":
            self.upload_hint.configure(text=f"Open the upload popup to choose a folder containing {', '.join(self.spec.input_extensions)} files.")
        else:
            self.upload_hint.configure(text=f"Open the upload popup to choose a {self.spec.input_label} file.")
        self.refresh_upload_popup_text()
        self.update_suggested_name()

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
        if self.mode.get() == "folder":
            self.suggested_name.set(f'Folder output: each file will use "Original Name ({self.spec.output_tag}){self.spec.output_extension}"')
            return
        self.suggested_name.set(f"Suggested output: {self.default_output_name(path).name}")

    def filetypes(self) -> list[tuple[str, str]]:
        patterns = " ".join(f"*{extension}" for extension in self.spec.input_extensions)
        return [(f"{self.spec.input_label.title()} files", patterns), ("All files", "*.*")]

    def choose_input(self) -> None:
        if self.is_converting:
            return
        if self.mode.get() == "folder":
            path = filedialog.askdirectory(title=f"Choose folder containing {self.spec.input_label} files")
        else:
            path = filedialog.askopenfilename(
                title=f"Choose {self.spec.input_label} file",
                filetypes=self.filetypes(),
            )
        if path:
            self.input_path.set(path)
            self.input_submitted = False
            self.status.set("Upload selected. Press Submit.")
            self.set_progress(5, "Upload selected")
            self.render_timeline()
            self.update_upload_display()
            self.show_toast("Upload selected", "Click Submit Upload in the popup to confirm this input.")

    def submit_input(self) -> bool:
        path_text = self.input_path.get().strip()
        if not path_text:
            messagebox.showerror("Missing upload", "Open the upload popup and choose a file or folder first.")
            return False

        path = Path(path_text)
        if self.mode.get() == "folder":
            if not path.is_dir():
                messagebox.showerror("Invalid folder", "Please choose a folder for folder conversion.")
                return False
        else:
            if not path.is_file():
                messagebox.showerror("Invalid file", "Please choose a file for single-file conversion.")
                return False
            if path.suffix.lower() not in self.spec.input_extensions:
                expected = ", ".join(self.spec.input_extensions)
                messagebox.showerror("Wrong file type", f"Please choose one of these file types: {expected}")
                return False

        self.input_submitted = True
        self.status.set("Upload submitted")
        self.update_upload_display()
        self.set_progress(12, "Upload submitted")
        self.render_timeline(active_index=None, completed=1, detail=f"Submitted {path.name}")
        self.write_log(f"Submitted upload: {path}")
        self.show_toast("Upload submitted", "Your input is ready. Choose a save folder or convert with the default location.")
        return True

    def choose_output_folder(self) -> None:
        path = filedialog.askdirectory(title="Choose output folder")
        if path:
            self.output_folder.set(path)
            self.set_progress(max(self.progress_value.get(), 18), "Output folder selected")
            self.show_toast("Save folder selected", self.display_path(Path(path)))

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
        return source if source.is_dir() else source.parent

    def output_path_for(self, source: Path) -> Path:
        output_directory = self.output_directory_for(source)
        if self.rename_output.get() and self.mode.get() != "folder":
            return output_directory / self.sanitized_custom_output_name()
        return output_directory / self.default_output_name(source)

    def start_conversion(self) -> None:
        if not self.input_submitted:
            messagebox.showerror("Submit upload", "Open the upload popup, choose a file or folder, then click Submit Upload first.")
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
        self.show_toast("Cancel requested", "The current file will finish if it is already being written.")

    def convert_one(self, input_path: Path, output_path: Path | None) -> Path:
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
            self.queue_step(3, f"Writing {self.spec.output_tag} file {index} of {total}: {output_file.name}")
            self.queue_progress(base_progress + (24 / total), f"Writing file {index} of {total}")
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
            if self.mode.get() == "folder":
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
                self.queue_step(3, f"Writing {self.spec.output_tag}: {output_path.name}")
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
                messagebox.showerror("Conversion stopped" if cancelled else "Conversion failed", message)

        self.root.after(0, update_ui)

    def show_download_page(self, outputs: list[Path]) -> None:
        if not outputs:
            self.show_toast("Download ready", "The conversion finished, but no output path was reported.")
            return

        theme = self.theme
        page = tk.Toplevel(self.root)
        page.title("Conversion Complete")
        page.minsize(640, 430)
        page.transient(self.root)
        page.iconphoto(True, self.window_icon)
        page.configure(bg=theme.background)
        page.columnconfigure(0, weight=1)
        page.rowconfigure(0, weight=1)

        shell = tk.Frame(page, bg=theme.background, padx=26, pady=24)
        shell.grid(row=0, column=0, sticky="nsew")
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        header = tk.Frame(shell, bg=theme.panel, highlightbackground=theme.border, highlightthickness=1, padx=18, pady=16)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(1, weight=1)
        tk.Label(
            header,
            text="DONE",
            bg=theme.accent,
            fg=theme.accent_text,
            font=("Segoe UI", 11, "bold"),
            padx=12,
            pady=10,
        ).grid(row=0, column=0, rowspan=2, sticky="nsw", padx=(0, 14))
        tk.Label(header, text="Conversion Complete", bg=theme.panel, fg=theme.text, font=("Segoe UI", 18, "bold")).grid(
            row=0,
            column=1,
            sticky="w",
        )
        tk.Label(
            header,
            text="Your converted file is ready. Open it now or open the folder where it was saved.",
            bg=theme.panel,
            fg=theme.muted,
            font=("Segoe UI", 10),
            wraplength=470,
            justify="left",
        ).grid(row=1, column=1, sticky="w", pady=(4, 0))

        content = tk.Frame(shell, bg=theme.panel, highlightbackground=theme.border, highlightthickness=1, padx=16, pady=14)
        content.grid(row=1, column=0, sticky="nsew", pady=(0, 14))
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)
        tk.Label(content, text="Converted output", bg=theme.panel, fg=theme.text, font=("Segoe UI", 10, "bold")).grid(
            row=0,
            column=0,
            sticky="w",
            pady=(0, 8),
        )

        listbox = tk.Listbox(
            content,
            bg=theme.log_background,
            fg=theme.log_text,
            selectbackground=theme.selection,
            selectforeground=theme.text,
            activestyle="none",
            bd=0,
            highlightthickness=0,
            font=("Segoe UI", 10),
        )
        listbox.grid(row=1, column=0, sticky="nsew")
        for output in outputs:
            listbox.insert("end", str(output))
        listbox.selection_set(0)

        button_row = tk.Frame(shell, bg=theme.background)
        button_row.grid(row=2, column=0, sticky="ew")
        button_row.columnconfigure(3, weight=1)

        def selected_output() -> Path:
            selection = listbox.curselection()
            if selection:
                return outputs[selection[0]]
            return outputs[0]

        ttk.Button(button_row, text="Open File", style="Accent.TButton", command=lambda: self.open_path(selected_output())).grid(
            row=0,
            column=0,
            sticky="w",
            padx=(0, 10),
        )
        ttk.Button(
            button_row,
            text="Open Folder",
            style="Secondary.TButton",
            command=lambda: self.open_path(selected_output().parent),
        ).grid(row=0, column=1, sticky="w", padx=(0, 10))
        ttk.Button(
            button_row,
            text="New Conversion",
            style="Ghost.TButton",
            command=lambda: self.reset_after_download(page),
        ).grid(row=0, column=2, sticky="w")

        page.attributes("-topmost", True)
        self.center_popup(page, 640, 430)
        page.after(500, lambda: page.attributes("-topmost", False))
        self.show_toast("Download popup opened", "Use Open File or Open Folder to access the converted output.")

    def open_path(self, path: Path) -> None:
        try:
            if not path.exists():
                messagebox.showerror("Missing output", f"This path no longer exists:\n{path}")
                return
            os.startfile(str(path))
        except OSError as error:
            messagebox.showerror("Could not open", f"Windows could not open this path:\n{path}\n\n{error}")

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
        self.show_toast("Ready", "Choose the next file or folder to convert.")

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
