from __future__ import annotations

import argparse
import html
import mmap
import posixpath
import re
import struct
import zlib
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path
from textwrap import wrap
from xml.etree import ElementTree

from .converter import ConversionError, ensure_output_path


TEXT_DOCUMENT_EXTENSIONS = (
    ".txt",
    ".text",
    ".md",
    ".markdown",
    ".rst",
    ".log",
    ".csv",
    ".tsv",
    ".json",
    ".jsonl",
    ".xml",
    ".yaml",
    ".yml",
    ".toml",
    ".ini",
    ".cfg",
    ".conf",
    ".properties",
    ".html",
    ".htm",
    ".css",
    ".js",
    ".ts",
    ".jsx",
    ".tsx",
    ".py",
    ".java",
    ".cs",
    ".c",
    ".cpp",
    ".h",
    ".hpp",
    ".php",
    ".rb",
    ".go",
    ".rs",
    ".swift",
    ".kt",
    ".kts",
    ".sql",
    ".ps1",
    ".bat",
    ".cmd",
    ".sh",
)

EPUB_EXTENSIONS = (".epub",)
MAX_BASIC_PDF_STREAM_BYTES = 5 * 1024 * 1024
MAX_RAW_IMAGE_STREAM_BYTES = 40 * 1024 * 1024
MAX_PYPDF_PDF_BYTES = 25 * 1024 * 1024


@dataclass(frozen=True)
class PdfImageAsset:
    path: Path
    width: int | None
    height: int | None


@dataclass(frozen=True)
class PdfHtmlBlock:
    kind: str
    text: str = ""
    image: PdfImageAsset | None = None


def file_error_message(action: str, path: Path, error: OSError) -> str:
    return f"{action}:\n{path}\n\n{error}"


def pdf_memory_error_message(path: Path) -> str:
    return (
        "This PDF is too large for the current conversion engine to inspect safely:\n"
        f"{path}\n\n"
        "Try splitting the PDF into smaller sections, or close other apps and run the conversion again."
    )


class ReadableHtmlParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.skip_depth = 0

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"script", "style"}:
            self.skip_depth += 1
            return
        if self.skip_depth:
            return
        if tag in {"p", "div", "section", "article", "header", "footer", "h1", "h2", "h3", "h4", "h5", "h6"}:
            self.parts.append("\n\n")
        elif tag == "br":
            self.parts.append("\n")
        elif tag == "li":
            self.parts.append("\n- ")

    def handle_endtag(self, tag: str) -> None:
        if tag in {"script", "style"} and self.skip_depth:
            self.skip_depth -= 1
            return
        if self.skip_depth:
            return
        if tag in {"p", "div", "section", "article", "li", "h1", "h2", "h3", "h4", "h5", "h6"}:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        if not self.skip_depth and data.strip():
            self.parts.append(data)

    def text(self) -> str:
        raw = "".join(self.parts)
        raw = html.unescape(raw)
        raw = re.sub(r"[ \t\r\f\v]+", " ", raw)
        raw = re.sub(r" *\n *", "\n", raw)
        raw = re.sub(r"\n{3,}", "\n\n", raw)
        return raw.strip()


def docx_content_types_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )


def docx_relationships_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml"/>'
        "</Relationships>"
    )


def empty_relationships_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    )


def docx_styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/>'
        '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>'
        '</w:style>'
        '<w:style w:type="paragraph" w:styleId="Title">'
        '<w:name w:val="Title"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:after="220"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>'
        '</w:style>'
        "</w:styles>"
    )


def docx_core_properties_xml() -> str:
    created_at = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:creator>File Converter</dc:creator>"
        "<cp:lastModifiedBy>File Converter</cp:lastModifiedBy>"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{created_at}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created_at}</dcterms:modified>'
        "</cp:coreProperties>"
    )


def docx_app_properties_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>File Converter</Application>"
        "</Properties>"
    )


def word_text(value: str) -> str:
    return html.escape(value, quote=False)


def paragraph_xml(text: str, style: str | None = None) -> str:
    if not text:
        return "<w:p/>"

    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    lines = text.splitlines() or [text]
    runs: list[str] = []
    for index, line in enumerate(lines):
        if index:
            runs.append("<w:r><w:br/></w:r>")
        runs.append(f'<w:r><w:t xml:space="preserve">{word_text(line)}</w:t></w:r>')
    return f"<w:p>{style_xml}{''.join(runs)}</w:p>"


def document_xml(title: str, paragraphs: list[str]) -> str:
    body_parts = [paragraph_xml(title, "Title")]
    body_parts.extend(paragraph_xml(paragraph) for paragraph in paragraphs)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{''.join(body_parts)}"
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
        "</w:body></w:document>"
    )


def write_docx(output_path: Path, title: str, paragraphs: list[str]) -> None:
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as document:
        document.writestr("[Content_Types].xml", docx_content_types_xml())
        document.writestr("_rels/.rels", docx_relationships_xml())
        document.writestr("word/document.xml", document_xml(title, paragraphs))
        document.writestr("word/styles.xml", docx_styles_xml())
        document.writestr("word/_rels/document.xml.rels", empty_relationships_xml())
        document.writestr("docProps/core.xml", docx_core_properties_xml())
        document.writestr("docProps/app.xml", docx_app_properties_xml())


def xml_local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def decode_epub_member(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "utf-16", "cp1252"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace")


def find_epub_rootfile(book: zipfile.ZipFile) -> str:
    try:
        container_xml = book.read("META-INF/container.xml")
    except KeyError as error:
        raise ConversionError("This EPUB is missing META-INF/container.xml.") from error

    root = ElementTree.fromstring(container_xml)
    for element in root.iter():
        if xml_local_name(element.tag) == "rootfile":
            rootfile = element.attrib.get("full-path")
            if rootfile:
                return rootfile

    fallback = next((name for name in book.namelist() if name.lower().endswith(".opf")), None)
    if fallback:
        return fallback
    raise ConversionError("This EPUB does not include a package .opf file.")


def epub_spine_documents(book: zipfile.ZipFile, package_path: str) -> list[str]:
    package_xml = book.read(package_path)
    package_root = ElementTree.fromstring(package_xml)
    base_dir = posixpath.dirname(package_path)
    manifest: dict[str, tuple[str, str]] = {}
    spine: list[str] = []

    for element in package_root.iter():
        local = xml_local_name(element.tag)
        if local == "item":
            item_id = element.attrib.get("id")
            href = element.attrib.get("href")
            media_type = element.attrib.get("media-type", "")
            if item_id and href:
                manifest[item_id] = (href, media_type)
        elif local == "itemref":
            idref = element.attrib.get("idref")
            if idref:
                spine.append(idref)

    def is_readable_document(href: str, media_type: str) -> bool:
        suffix = Path(href).suffix.lower()
        return media_type in {"application/xhtml+xml", "text/html"} or suffix in {".xhtml", ".html", ".htm"}

    documents: list[str] = []
    for idref in spine:
        item = manifest.get(idref)
        if not item:
            continue
        href, media_type = item
        if is_readable_document(href, media_type):
            documents.append(posixpath.normpath(posixpath.join(base_dir, href)))

    if documents:
        return documents

    for href, media_type in manifest.values():
        if is_readable_document(href, media_type):
            documents.append(posixpath.normpath(posixpath.join(base_dir, href)))
    return documents


def html_to_readable_text(value: str) -> str:
    parser = ReadableHtmlParser()
    parser.feed(value)
    return parser.text()


def extract_epub_text(epub_path: Path) -> str:
    try:
        with zipfile.ZipFile(epub_path) as book:
            package_path = find_epub_rootfile(book)
            documents = epub_spine_documents(book, package_path)
            if not documents:
                raise ConversionError("No readable chapter documents were found in this EPUB.")

            sections: list[str] = []
            for document_path in documents:
                try:
                    data = book.read(document_path)
                except KeyError:
                    continue
                text = html_to_readable_text(decode_epub_member(data))
                if text:
                    sections.append(text)
    except zipfile.BadZipFile as error:
        raise ConversionError("This EPUB file could not be opened as a valid EPUB archive.") from error

    text = "\n\n".join(sections).strip()
    if not text:
        raise ConversionError("No readable text was found in this EPUB.")
    return text


def escape_pdf_text(value: str) -> bytes:
    data = value.encode("cp1252", errors="replace")
    return data.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)")


def paginate_text_for_pdf(title: str, text: str) -> list[list[str]]:
    lines: list[str] = [title, ""]
    for paragraph in normalize_paragraphs(text):
        wrapped = wrap(paragraph, width=88, replace_whitespace=False, drop_whitespace=True) or [""]
        lines.extend(wrapped)
        lines.append("")
    if lines[-1:] == [""]:
        lines.pop()

    page_size = 46
    return [lines[index : index + page_size] for index in range(0, len(lines), page_size)] or [[title]]


def pdf_content_stream(lines: list[str]) -> bytes:
    stream = [b"BT", b"/F1 11 Tf", b"50 742 Td", b"14 TL"]
    for line in lines:
        stream.append(b"(" + escape_pdf_text(line) + b") Tj")
        stream.append(b"T*")
    stream.append(b"ET")
    return b"\n".join(stream)


def write_text_pdf(output_path: Path, title: str, text: str) -> None:
    pages = paginate_text_for_pdf(title, text)
    objects: list[bytes] = [b"", b"", b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"]
    page_ids: list[int] = []

    for lines in pages:
        content = pdf_content_stream(lines)
        content_id = len(objects) + 1
        objects.append(b"<< /Length " + str(len(content)).encode("ascii") + b" >>\nstream\n" + content + b"\nendstream")
        page_id = len(objects) + 1
        page_ids.append(page_id)
        objects.append(
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
            b"/Resources << /Font << /F1 3 0 R >> >> /Contents "
            + str(content_id).encode("ascii")
            + b" 0 R >>"
        )

    kids = b" ".join(str(page_id).encode("ascii") + b" 0 R" for page_id in page_ids)
    objects[0] = b"<< /Type /Catalog /Pages 2 0 R >>"
    objects[1] = b"<< /Type /Pages /Kids [" + kids + b"] /Count " + str(len(page_ids)).encode("ascii") + b" >>"

    pdf = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for object_id, body in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f"{object_id} 0 obj\n".encode("ascii"))
        pdf.extend(body)
        pdf.extend(b"\nendobj\n")

    xref_offset = len(pdf)
    pdf.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    pdf.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    pdf.extend(
        f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n".encode("ascii")
    )
    output_path.write_bytes(bytes(pdf))


def convert_epub_to_pdf(
    epub_path: Path | str,
    output_path: Path | str | None = None,
    *,
    overwrite: bool = False,
) -> Path:
    source = Path(epub_path)
    check_source_file(source, EPUB_EXTENSIONS, "EPUB")
    destination = Path(output_path) if output_path else source.with_suffix(".pdf")
    ensure_output_path(destination, overwrite)
    try:
        write_text_pdf(destination, source.stem, extract_epub_text(source))
    except OSError as error:
        raise ConversionError(file_error_message("Could not write the PDF file", destination, error)) from error
    return destination


def check_source_file(source: Path, suffixes: tuple[str, ...], label: str) -> None:
    if not source.exists():
        raise ConversionError(f"{label} file not found: {source}")
    if not source.is_file():
        raise ConversionError(f"Input must be a {label} file: {source}")
    if source.suffix.lower() not in suffixes:
        expected = ", ".join(suffixes)
        raise ConversionError(f"Input file must end with {expected}: {source}")


def normalize_paragraphs(text: str) -> list[str]:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    chunks = re.split(r"\n\s*\n", text)
    paragraphs: list[str] = []
    for chunk in chunks:
        cleaned = "\n".join(line.rstrip() for line in chunk.splitlines()).strip()
        if cleaned:
            paragraphs.append(cleaned)
    return paragraphs


def convert_text_to_word(
    text_path: Path | str,
    output_path: Path | str | None = None,
    *,
    encoding: str = "utf-8-sig",
    overwrite: bool = False,
) -> Path:
    source = Path(text_path)
    check_source_file(source, TEXT_DOCUMENT_EXTENSIONS, "editable text")
    destination = Path(output_path) if output_path else source.with_suffix(".docx")
    ensure_output_path(destination, overwrite)

    try:
        text = source.read_text(encoding=encoding)
    except UnicodeError as error:
        raise ConversionError(f"Could not read this file as editable text with {encoding}: {error}") from error
    except OSError as error:
        raise ConversionError(file_error_message("Could not read the source file", source, error)) from error
    paragraphs = normalize_paragraphs(text)
    try:
        write_docx(destination, source.stem, paragraphs or [""])
    except OSError as error:
        raise ConversionError(file_error_message("Could not write the Word file", destination, error)) from error
    return destination


def extract_pdf_text_with_pypdf(pdf_path: Path) -> str | None:
    try:
        from pypdf import PdfReader  # type: ignore
    except Exception:
        return None

    try:
        reader = PdfReader(str(pdf_path))
        return "\n\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception:
        return None


def stream_dictionary(data: bytes | mmap.mmap, stream_start: int) -> bytes:
    object_start = data.rfind(b"obj", 0, stream_start)
    search_start = object_start + len(b"obj") if object_start != -1 else max(0, stream_start - 8192)
    dictionary_start = data.find(b"<<", search_start, stream_start)
    dictionary_end = data.rfind(b">>", search_start, stream_start)
    if dictionary_start == -1 or dictionary_end == -1 or dictionary_end < dictionary_start:
        return b""
    return data[dictionary_start : dictionary_end + 2]


def stream_object_number(data: bytes | mmap.mmap, stream_start: int) -> int | None:
    object_marker = data.rfind(b"obj", 0, stream_start)
    if object_marker == -1:
        return None
    prefix = data[max(0, object_marker - 64) : object_marker + len(b"obj")]
    match = re.search(rb"(\d+)\s+\d+\s+obj\b\s*$", prefix)
    return int(match.group(1)) if match else None


def declared_stream_length(dictionary: bytes) -> int | None:
    match = re.search(rb"/Length\s+(\d+)", dictionary)
    return int(match.group(1)) if match else None


def is_image_stream_dictionary(dictionary: bytes) -> bool:
    return bool(re.search(rb"/Subtype\s*/Image\b", dictionary))


def is_image_mask_dictionary(dictionary: bytes) -> bool:
    return bool(re.search(rb"/ImageMask\s+true\b", dictionary))


def referenced_image_mask_object_numbers(data: bytes | mmap.mmap) -> set[int]:
    mask_objects: set[int] = set()
    for _object_number, dictionary, _content in iter_pdf_stream_records(data, include_content=False):
        for match in re.finditer(rb"/(?:SMask|Mask)\s+(\d+)\s+\d+\s+R\b", dictionary):
            mask_objects.add(int(match.group(1)))
    return mask_objects


def dictionary_int(dictionary: bytes, name: bytes) -> int | None:
    match = re.search(rb"/" + re.escape(name) + rb"\s+(\d+)", dictionary)
    return int(match.group(1)) if match else None


def image_extension_for_dictionary(dictionary: bytes) -> str | None:
    if b"/DCTDecode" in dictionary:
        return ".jpg"
    if b"/JPXDecode" in dictionary:
        return ".jp2"
    if b"/FlateDecode" in dictionary:
        return ".png"
    return None


def image_color_space_channels(dictionary: bytes) -> int | None:
    if re.search(rb"/ColorSpace\s*/DeviceGray\b", dictionary):
        return 1
    if re.search(rb"/ColorSpace\s*/DeviceRGB\b", dictionary):
        return 3
    if re.search(rb"/ColorSpace\s*/DeviceCMYK\b", dictionary):
        return 4
    return None


def looks_like_text_content_stream(content: bytes) -> bool:
    return bool(re.search(rb"(?<![A-Za-z])(?:BT|Tj|TJ|Td|TD|T\*)\b", content))


def read_pdf_with_mmap(pdf_path: Path, reader):
    try:
        if pdf_path.stat().st_size == 0:
            return reader(b"")
        with pdf_path.open("rb") as pdf_file:
            with mmap.mmap(pdf_file.fileno(), 0, access=mmap.ACCESS_READ) as data:
                return reader(data)
    except MemoryError as error:
        raise ConversionError(pdf_memory_error_message(pdf_path)) from error
    except OSError as error:
        raise ConversionError(file_error_message("Could not read the PDF file", pdf_path, error)) from error


def decompress_pdf_stream(content: bytes, max_bytes: int) -> bytes | None:
    try:
        decompressor = zlib.decompressobj()
        output = decompressor.decompress(content, max_bytes + 1)
        if len(output) > max_bytes or decompressor.unconsumed_tail:
            return None
        output += decompressor.flush(max_bytes + 1 - len(output))
    except (zlib.error, ValueError, MemoryError):
        return None
    if len(output) > max_bytes:
        return None
    return output


def iter_pdf_stream_records(data: bytes | mmap.mmap, *, include_content: bool = True):
    position = 0
    while True:
        stream_start = data.find(b"stream", position)
        if stream_start == -1:
            break

        content_start = stream_start + len(b"stream")
        if data[content_start : content_start + 2] == b"\r\n":
            content_start += 2
        elif data[content_start : content_start + 1] in {b"\n", b"\r"}:
            content_start += 1

        stream_end = data.find(b"endstream", content_start)
        if stream_end == -1:
            break

        dictionary = stream_dictionary(data, stream_start)
        object_number = stream_object_number(data, stream_start)
        position = stream_end + len(b"endstream")
        if not include_content:
            yield object_number, dictionary, b""
            continue

        stream_length = max(0, stream_end - content_start)
        max_stream_bytes = MAX_RAW_IMAGE_STREAM_BYTES if is_image_stream_dictionary(dictionary) else MAX_BASIC_PDF_STREAM_BYTES
        declared_length = declared_stream_length(dictionary)
        if (declared_length and declared_length > max_stream_bytes) or stream_length > max_stream_bytes:
            yield object_number, dictionary, b""
            continue

        content = data[content_start:stream_end].rstrip(b"\r\n")
        yield object_number, dictionary, content


def iter_pdf_stream_objects(data: bytes | mmap.mmap):
    for _object_number, dictionary, content in iter_pdf_stream_records(data):
        yield dictionary, content


def iter_pdf_streams(data: bytes | mmap.mmap):
    for dictionary, content in iter_pdf_stream_objects(data):
        if is_image_stream_dictionary(dictionary):
            continue

        declared_length = declared_stream_length(dictionary)
        if declared_length and declared_length > MAX_BASIC_PDF_STREAM_BYTES:
            continue

        if b"/FlateDecode" in dictionary:
            decompressed = decompress_pdf_stream(content, MAX_BASIC_PDF_STREAM_BYTES)
            if decompressed is None:
                continue
            content = decompressed

        if len(content) > MAX_BASIC_PDF_STREAM_BYTES or not looks_like_text_content_stream(content):
            continue
        yield content


def png_chunk(kind: bytes, data: bytes) -> bytes:
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)


def rgb_png_bytes(width: int, height: int, rgb_data: bytes) -> bytes:
    stride = width * 3
    raw_rows = bytearray()
    for row in range(height):
        raw_rows.append(0)
        raw_rows.extend(rgb_data[row * stride : (row + 1) * stride])
    return (
        b"\x89PNG\r\n\x1a\n"
        + png_chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + png_chunk(b"IDAT", zlib.compress(bytes(raw_rows), 9))
        + png_chunk(b"IEND", b"")
    )


def raw_pdf_image_to_png(dictionary: bytes, content: bytes) -> bytes | None:
    width = dictionary_int(dictionary, b"Width")
    height = dictionary_int(dictionary, b"Height")
    bits_per_component = dictionary_int(dictionary, b"BitsPerComponent")
    channels = image_color_space_channels(dictionary)
    if not width or not height or bits_per_component != 8 or not channels:
        return None

    expected_length = width * height * channels
    if expected_length > MAX_RAW_IMAGE_STREAM_BYTES or len(content) < expected_length:
        return None

    pixels = content[:expected_length]
    if channels == 3:
        rgb_data = pixels
    elif channels == 1:
        rgb_data = b"".join(bytes((value, value, value)) for value in pixels)
    else:
        rgb = bytearray()
        for index in range(0, expected_length, 4):
            cyan, magenta, yellow, black = pixels[index : index + 4]
            rgb.extend(
                (
                    max(0, 255 - min(255, cyan + black)),
                    max(0, 255 - min(255, magenta + black)),
                    max(0, 255 - min(255, yellow + black)),
                )
            )
        rgb_data = bytes(rgb)
    return rgb_png_bytes(width, height, rgb_data)


def image_asset_from_stream(
    dictionary: bytes,
    content: bytes,
    asset_dir: Path,
    image_index: int,
) -> PdfImageAsset | None:
    extension = image_extension_for_dictionary(dictionary)
    if not extension:
        return None

    width = dictionary_int(dictionary, b"Width")
    height = dictionary_int(dictionary, b"Height")
    image_bytes: bytes | None
    declared_length = declared_stream_length(dictionary)
    if declared_length and declared_length > MAX_RAW_IMAGE_STREAM_BYTES:
        return None

    if extension == ".png":
        decompressed = decompress_pdf_stream(content, MAX_RAW_IMAGE_STREAM_BYTES)
        if decompressed is None:
            return None
        try:
            image_bytes = raw_pdf_image_to_png(dictionary, decompressed)
        except (ValueError, OverflowError, struct.error, MemoryError):
            image_bytes = None
    else:
        if b"/FlateDecode" in dictionary:
            image_bytes = decompress_pdf_stream(content, MAX_RAW_IMAGE_STREAM_BYTES)
        else:
            image_bytes = content

    if not image_bytes:
        return None

    try:
        asset_dir.mkdir(parents=True, exist_ok=True)
    except OSError as error:
        raise ConversionError(file_error_message("Could not create the PDF image folder", asset_dir, error)) from error

    image_path = asset_dir / f"image_{image_index}{extension}"
    try:
        image_path.write_bytes(image_bytes)
    except OSError as error:
        raise ConversionError(file_error_message("Could not write an extracted PDF image", image_path, error)) from error
    return PdfImageAsset(image_path, width, height)


def extract_pdf_html_blocks(data: bytes | mmap.mmap, html_destination: Path) -> list[PdfHtmlBlock]:
    asset_dir = html_destination.with_name(f"{html_destination.stem}_assets")
    blocks: list[PdfHtmlBlock] = []
    image_index = 0
    mask_object_numbers = referenced_image_mask_object_numbers(data)

    for object_number, dictionary, content in iter_pdf_stream_records(data):
        if is_image_stream_dictionary(dictionary):
            if is_image_mask_dictionary(dictionary) or object_number in mask_object_numbers:
                continue
            image = image_asset_from_stream(dictionary, content, asset_dir, image_index + 1)
            if image:
                image_index += 1
                blocks.append(PdfHtmlBlock(kind="image", image=image))
            continue

        declared_length = declared_stream_length(dictionary)
        if declared_length and declared_length > MAX_BASIC_PDF_STREAM_BYTES:
            continue

        if b"/FlateDecode" in dictionary:
            decompressed = decompress_pdf_stream(content, MAX_BASIC_PDF_STREAM_BYTES)
            if decompressed is None:
                continue
            content = decompressed

        if len(content) > MAX_BASIC_PDF_STREAM_BYTES or not looks_like_text_content_stream(content):
            continue

        try:
            text = extract_text_from_pdf_content(content).strip()
        except (ValueError, OverflowError, UnicodeError):
            continue
        if text:
            blocks.append(PdfHtmlBlock(kind="text", text=text))
    return blocks


def extract_pdf_images(pdf_path: Path, html_destination: Path) -> list[PdfImageAsset]:
    blocks = read_pdf_with_mmap(pdf_path, lambda data: extract_pdf_html_blocks(data, html_destination))
    return [block.image for block in blocks if block.image is not None]


def parse_literal_string(text: str, start: int) -> tuple[str, int]:
    result: list[str] = []
    depth = 1
    index = start + 1
    while index < len(text) and depth:
        char = text[index]
        if char == "\\":
            index += 1
            if index >= len(text):
                break
            escaped = text[index]
            escapes = {"n": "\n", "r": "\r", "t": "\t", "b": "\b", "f": "\f", "(": "(", ")": ")", "\\": "\\"}
            if escaped in escapes:
                result.append(escapes[escaped])
            elif escaped in "\r\n":
                if escaped == "\r" and index + 1 < len(text) and text[index + 1] == "\n":
                    index += 1
            elif escaped in "01234567":
                octal = escaped
                for _ in range(2):
                    if index + 1 < len(text) and text[index + 1] in "01234567":
                        index += 1
                        octal += text[index]
                    else:
                        break
                result.append(chr(int(octal, 8)))
            else:
                result.append(escaped)
        elif char == "(":
            depth += 1
            result.append(char)
        elif char == ")":
            depth -= 1
            if depth:
                result.append(char)
        else:
            result.append(char)
        index += 1
    return "".join(result), index


def parse_hex_string(text: str, start: int) -> tuple[str, int]:
    end = text.find(">", start + 1)
    if end == -1:
        return "", start + 1
    hex_text = re.sub(r"\s+", "", text[start + 1 : end])
    if len(hex_text) % 2:
        hex_text += "0"
    try:
        data = bytes.fromhex(hex_text)
    except ValueError:
        return "", end + 1
    if data.startswith(b"\xfe\xff"):
        return data[2:].decode("utf-16-be", errors="replace"), end + 1
    if b"\x00" in data:
        return data.decode("utf-16-be", errors="replace"), end + 1
    return data.decode("latin-1", errors="replace"), end + 1


def tokenize_pdf_content(content: bytes) -> list[object]:
    text = content.decode("latin-1", errors="ignore")
    tokens: list[object] = []
    index = 0
    while index < len(text):
        char = text[index]
        if char.isspace():
            index += 1
        elif char == "%":
            end = text.find("\n", index)
            index = len(text) if end == -1 else end + 1
        elif char == "(":
            value, index = parse_literal_string(text, index)
            tokens.append(value)
        elif char == "[":
            tokens.append("[")
            index += 1
        elif char == "]":
            tokens.append("]")
            index += 1
        elif text.startswith("<<", index) or text.startswith(">>", index):
            index += 2
        elif char == "<" and index + 1 < len(text) and text[index + 1] != "<":
            value, index = parse_hex_string(text, index)
            tokens.append(value)
        elif char in "<>":
            index += 1
        else:
            end = index
            while end < len(text) and not text[end].isspace() and text[end] not in "[]()<>":
                end += 1
            if end == index:
                index += 1
                continue
            tokens.append(text[index:end])
            index = end
    return tokens


def extract_text_from_pdf_content(content: bytes) -> str:
    tokens = tokenize_pdf_content(content)
    output: list[str] = []
    stack: list[object] = []
    current_array: list[object] | None = None

    for token in tokens:
        if token == "[":
            current_array = []
            continue
        if token == "]":
            if current_array is not None:
                stack.append(current_array)
            current_array = None
            continue
        if current_array is not None:
            current_array.append(token)
            continue

        if token in {"Tj", "'", '"'}:
            if stack and isinstance(stack[-1], str):
                output.append(stack[-1])
            if token in {"'", '"'}:
                output.append("\n")
            stack.clear()
        elif token == "TJ":
            if stack and isinstance(stack[-1], list):
                output.append("".join(item for item in stack[-1] if isinstance(item, str)))
            stack.clear()
        elif token in {"T*", "Td", "TD", "ET"}:
            output.append("\n")
            stack.clear()
        else:
            stack.append(token)
            if len(stack) > 8:
                stack = stack[-8:]

    return "".join(output)


def extract_pdf_text_basic(pdf_path: Path) -> str:
    def extract(data: bytes | mmap.mmap) -> str:
        page_texts = [extract_text_from_pdf_content(stream) for stream in iter_pdf_streams(data)]
        text = "\n".join(page_texts)
        text = re.sub(r"[ \t]+\n", "\n", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()

    return read_pdf_with_mmap(pdf_path, extract)


def extract_pdf_text(pdf_path: Path, *, prefer_fast: bool = False) -> str:
    if prefer_fast:
        text = extract_pdf_text_basic(pdf_path)
        if text.strip():
            return text.strip()
        try:
            if pdf_path.stat().st_size > MAX_PYPDF_PDF_BYTES:
                return ""
        except OSError:
            return ""

    text = extract_pdf_text_with_pypdf(pdf_path)
    if text and text.strip():
        return text.strip()
    return extract_pdf_text_basic(pdf_path)


def text_block_to_html(text: str) -> str:
    paragraphs = normalize_paragraphs(text)
    return "\n".join(
        f"      <p>{'<br>'.join(html.escape(line) for line in paragraph.splitlines())}</p>"
        for paragraph in paragraphs
    )


def image_block_to_html(image: PdfImageAsset, output_dir: Path, index: int) -> str:
    try:
        source = image.path.resolve().relative_to(output_dir.resolve()).as_posix()
    except ValueError:
        source = image.path.as_posix()
    width_attribute = f' width="{image.width}"' if image.width else ""
    height_attribute = f' height="{image.height}"' if image.height else ""
    return (
        f'      <figure><img src="{html.escape(source, quote=True)}" alt="PDF image {index}"'
        f'{width_attribute}{height_attribute} loading="lazy"><figcaption>Image {index}</figcaption></figure>'
    )


def html_blocks_markup(blocks: list[PdfHtmlBlock], output_dir: Path) -> str:
    if not blocks:
        return "      <p>No selectable text or supported images were found.</p>"

    output: list[str] = []
    image_index = 0
    has_text = any(block.kind == "text" and block.text.strip() for block in blocks)
    if not has_text and any(block.kind == "image" for block in blocks):
        output.append("      <p>No selectable text was found. Extracted images are shown below.</p>")

    for block in blocks:
        if block.kind == "text" and block.text.strip():
            output.append(text_block_to_html(block.text))
        elif block.kind == "image" and block.image:
            image_index += 1
            output.append(image_block_to_html(block.image, output_dir, image_index))
    return "\n".join(part for part in output if part)


def blocks_to_html(title: str, blocks: list[PdfHtmlBlock], output_dir: Path) -> str:
    safe_title = html.escape(title)
    body = html_blocks_markup(blocks, output_dir)

    return f"""<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>{safe_title}</title>
    <style>
      :root {{
        color-scheme: light;
        font-family: "Segoe UI", Arial, sans-serif;
        line-height: 1.6;
        color: #0f172a;
        background: #f8fafc;
      }}
      body {{
        margin: 0;
        padding: 40px 20px;
      }}
      main {{
        max-width: 860px;
        margin: 0 auto;
        background: #ffffff;
        border: 1px solid #d9e4ec;
        border-radius: 10px;
        padding: 34px;
        box-shadow: 0 20px 45px rgba(15, 23, 42, 0.08);
      }}
      h1 {{
        margin: 0 0 22px;
        font-size: 28px;
        line-height: 1.2;
      }}
      p {{
        margin: 0 0 16px;
        white-space: normal;
      }}
      figure {{
        margin: 18px 0 24px;
      }}
      img {{
        display: block;
        max-width: 100%;
        height: auto;
        border: 1px solid #d9e4ec;
        border-radius: 8px;
        background: #f8fafc;
      }}
      figcaption {{
        margin-top: 6px;
        color: #64748b;
        font-size: 13px;
      }}
    </style>
  </head>
  <body>
    <main>
      <h1>{safe_title}</h1>
{body}
    </main>
  </body>
</html>
"""


def convert_pdf_to_word(
    pdf_path: Path | str,
    output_path: Path | str | None = None,
    *,
    overwrite: bool = False,
) -> Path:
    source = Path(pdf_path)
    check_source_file(source, (".pdf",), "PDF")
    destination = Path(output_path) if output_path else source.with_suffix(".docx")
    ensure_output_path(destination, overwrite)

    try:
        text = extract_pdf_text(source)
    except OSError as error:
        raise ConversionError(file_error_message("Could not read the PDF file", source, error)) from error
    if not text.strip():
        raise ConversionError(
            "No readable text was found in this PDF. It may be scanned or image-only, which needs OCR."
        )

    paragraphs = normalize_paragraphs(text)
    try:
        write_docx(destination, source.stem, paragraphs)
    except OSError as error:
        raise ConversionError(file_error_message("Could not write the Word file", destination, error)) from error
    return destination


def convert_pdf_to_html(
    pdf_path: Path | str,
    output_path: Path | str | None = None,
    *,
    overwrite: bool = False,
) -> Path:
    source = Path(pdf_path)
    check_source_file(source, (".pdf",), "PDF")
    destination = Path(output_path) if output_path else source.with_suffix(".html")
    ensure_output_path(destination, overwrite)

    try:
        blocks = read_pdf_with_mmap(source, lambda data: extract_pdf_html_blocks(data, destination))
    except ConversionError:
        raise
    except (ValueError, OverflowError, UnicodeError, zlib.error, struct.error, MemoryError) as error:
        details = str(error).strip() or "No extra details were provided."
        raise ConversionError(f"Could not inspect the PDF contents: {type(error).__name__}: {details}") from error

    has_text = any(block.kind == "text" and block.text.strip() for block in blocks)
    has_images = any(block.kind == "image" for block in blocks)
    if not has_text and not has_images:
        try:
            text = extract_pdf_text(source, prefer_fast=True)
        except OSError as error:
            raise ConversionError(file_error_message("Could not read the PDF file", source, error)) from error
        if text.strip():
            blocks = [PdfHtmlBlock(kind="text", text=text)]

    if not blocks:
        raise ConversionError("No readable text or supported embedded images were found in this PDF. It may need OCR.")

    try:
        destination.write_text(blocks_to_html(source.stem, blocks, destination.parent), encoding="utf-8")
    except OSError as error:
        raise ConversionError(file_error_message("Could not write the HTML file", destination, error)) from error
    return destination


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert document files.")
    parser.add_argument("input", type=Path, help="PDF, editable text, or EPUB file path.")
    parser.add_argument("-o", "--output", type=Path, help="Output file path.")
    parser.add_argument("--encoding", default="utf-8-sig", help="Text-file encoding. Default: utf-8-sig.")
    parser.add_argument("--overwrite", action="store_true", help="Replace existing output files.")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        suffix = args.input.suffix.lower()
        if suffix == ".pdf":
            if args.output and args.output.suffix.lower() in {".html", ".htm"}:
                output = convert_pdf_to_html(args.input, args.output, overwrite=args.overwrite)
            else:
                output = convert_pdf_to_word(args.input, args.output, overwrite=args.overwrite)
        elif suffix in TEXT_DOCUMENT_EXTENSIONS:
            output = convert_text_to_word(
                args.input,
                args.output,
                encoding=args.encoding,
                overwrite=args.overwrite,
            )
        elif suffix in EPUB_EXTENSIONS:
            output = convert_epub_to_pdf(args.input, args.output, overwrite=args.overwrite)
        else:
            raise ConversionError("Supported document inputs are .pdf, editable text files, and .epub files.")
        print(f"Created {output}")
        return 0
    except ConversionError as error:
        parser.exit(1, f"Error: {error}\n")


if __name__ == "__main__":
    raise SystemExit(main())
