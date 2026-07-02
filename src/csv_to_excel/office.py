from __future__ import annotations

import argparse
import html
import re
import zlib
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from .converter import ConversionError, ensure_output_path


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
    check_source_file(source, (".txt", ".md"), "text")
    destination = Path(output_path) if output_path else source.with_suffix(".docx")
    ensure_output_path(destination, overwrite)

    text = source.read_text(encoding=encoding)
    paragraphs = normalize_paragraphs(text)
    write_docx(destination, source.stem, paragraphs or [""])
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


def stream_dictionary(data: bytes, stream_start: int) -> bytes:
    dictionary_start = data.rfind(b"<<", 0, stream_start)
    dictionary_end = data.rfind(b">>", 0, stream_start)
    if dictionary_start == -1 or dictionary_end == -1 or dictionary_end < dictionary_start:
        return b""
    return data[dictionary_start : dictionary_end + 2]


def iter_pdf_streams(data: bytes) -> list[bytes]:
    streams: list[bytes] = []
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

        content = data[content_start:stream_end].rstrip(b"\r\n")
        dictionary = stream_dictionary(data, stream_start)
        if b"/FlateDecode" in dictionary:
            try:
                content = zlib.decompress(content)
            except zlib.error:
                position = stream_end + len(b"endstream")
                continue
        streams.append(content)
        position = stream_end + len(b"endstream")
    return streams


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
        elif char == "<" and index + 1 < len(text) and text[index + 1] != "<":
            value, index = parse_hex_string(text, index)
            tokens.append(value)
        else:
            end = index
            while end < len(text) and not text[end].isspace() and text[end] not in "[]()<>":
                end += 1
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
    data = pdf_path.read_bytes()
    page_texts = [extract_text_from_pdf_content(stream) for stream in iter_pdf_streams(data)]
    text = "\n".join(page_texts)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_pdf_text(pdf_path: Path) -> str:
    text = extract_pdf_text_with_pypdf(pdf_path)
    if text and text.strip():
        return text.strip()
    return extract_pdf_text_basic(pdf_path)


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

    text = extract_pdf_text(source)
    if not text.strip():
        raise ConversionError(
            "No readable text was found in this PDF. It may be scanned or image-only, which needs OCR."
        )

    paragraphs = normalize_paragraphs(text)
    write_docx(destination, source.stem, paragraphs)
    return destination


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert PDF/TXT files to Word .docx documents.")
    parser.add_argument("input", type=Path, help="PDF, TXT, or MD file path.")
    parser.add_argument("-o", "--output", type=Path, help="Output .docx file path.")
    parser.add_argument("--encoding", default="utf-8-sig", help="Text-file encoding. Default: utf-8-sig.")
    parser.add_argument("--overwrite", action="store_true", help="Replace existing output files.")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        suffix = args.input.suffix.lower()
        if suffix == ".pdf":
            output = convert_pdf_to_word(args.input, args.output, overwrite=args.overwrite)
        elif suffix in {".txt", ".md"}:
            output = convert_text_to_word(
                args.input,
                args.output,
                encoding=args.encoding,
                overwrite=args.overwrite,
            )
        else:
            raise ConversionError("Supported Word inputs are .pdf, .txt, and .md files.")
        print(f"Created {output}")
        return 0
    except ConversionError as error:
        parser.exit(1, f"Error: {error}\n")


if __name__ == "__main__":
    raise SystemExit(main())
