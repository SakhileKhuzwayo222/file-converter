from __future__ import annotations

import argparse
import csv
import html
import json
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable


INVALID_SHEET_CHARS = str.maketrans({char: "_" for char in "[]:*?/\\"})


class ConversionError(Exception):
    """Raised when CSV conversion cannot be completed."""


def detect_dialect(csv_path: Path, encoding: str, delimiter: str | None) -> csv.Dialect:
    if delimiter:
        class CustomDialect(csv.excel):
            pass

        CustomDialect.delimiter = delimiter
        return CustomDialect

    with csv_path.open("r", encoding=encoding, newline="") as csv_file:
        sample = csv_file.read(4096)

    if not sample:
        return csv.excel

    try:
        return csv.Sniffer().sniff(sample)
    except csv.Error:
        return csv.excel


def safe_sheet_title(csv_path: Path, sheet_name: str | None) -> str:
    raw_title = sheet_name or csv_path.stem or "Sheet1"
    title = raw_title.translate(INVALID_SHEET_CHARS).strip()
    return (title or "Sheet1")[:31]


def column_letter(column_number: int) -> str:
    letters = ""
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def xml_text(value: object) -> str:
    text = "" if value is None else str(value)
    return html.escape(text, quote=False)


def worksheet_xml(rows: list[list[str]], freeze_header: bool) -> str:
    max_columns = max((len(row) for row in rows), default=0)
    column_widths = [8] * max_columns
    for row in rows:
        for index, value in enumerate(row):
            column_widths[index] = min(max(column_widths[index], len(str(value)) + 2), 60)

    cols = ""
    if column_widths:
        col_entries = [
            f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>'
            for index, width in enumerate(column_widths, start=1)
        ]
        cols = f"<cols>{''.join(col_entries)}</cols>"

    sheet_view = '<sheetView workbookViewId="0"/>'
    if freeze_header and len(rows) > 1:
        sheet_view = (
            '<sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
            '</sheetView>'
        )

    row_entries: list[str] = []
    for row_index, row in enumerate(rows, start=1):
        cell_entries: list[str] = []
        for column_index, value in enumerate(row, start=1):
            cell_ref = f"{column_letter(column_index)}{row_index}"
            cell_entries.append(f'<c r="{cell_ref}" t="inlineStr"><is><t>{xml_text(value)}</t></is></c>')
        row_entries.append(f'<row r="{row_index}">{"".join(cell_entries)}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheetViews>{sheet_view}</sheetViews>"
        f"{cols}"
        f"<sheetData>{''.join(row_entries)}</sheetData>"
        "</worksheet>"
    )


def workbook_xml(sheet_title: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{html.escape(sheet_title, quote=True)}" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )


def styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        "</styleSheet>"
    )


def content_types_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )


def package_relationships_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
        'Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
        'Target="docProps/app.xml"/>'
        "</Relationships>"
    )


def workbook_relationships_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        "</Relationships>"
    )


def core_properties_xml() -> str:
    created_at = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:creator>CSV to Excel Converter</dc:creator>"
        "<cp:lastModifiedBy>CSV to Excel Converter</cp:lastModifiedBy>"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{created_at}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created_at}</dcterms:modified>'
        "</cp:coreProperties>"
    )


def app_properties_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>CSV to Excel Converter</Application>"
        "</Properties>"
    )


def write_xlsx(output_path: Path, rows: list[list[str]], sheet_title: str, freeze_header: bool) -> None:
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types_xml())
        workbook.writestr("_rels/.rels", package_relationships_xml())
        workbook.writestr("xl/workbook.xml", workbook_xml(sheet_title))
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_relationships_xml())
        workbook.writestr("xl/styles.xml", styles_xml())
        workbook.writestr("xl/worksheets/sheet1.xml", worksheet_xml(rows, freeze_header))
        workbook.writestr("docProps/core.xml", core_properties_xml())
        workbook.writestr("docProps/app.xml", app_properties_xml())


def ensure_output_path(destination: Path, overwrite: bool) -> None:
    if destination.exists() and not overwrite:
        raise ConversionError(f"Output file already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)


def check_table_source(source: Path, suffixes: tuple[str, ...], label: str) -> None:
    if not source.exists():
        raise ConversionError(f"{label} file not found: {source}")
    if not source.is_file():
        raise ConversionError(f"Input must be a {label} file: {source}")
    if source.suffix.lower() not in suffixes:
        expected = ", ".join(suffixes)
        raise ConversionError(f"Input file must end with {expected}: {source}")


def convert_rows_to_excel(
    rows: list[list[str]],
    source: Path,
    output_path: Path | str | None = None,
    *,
    sheet_name: str | None = None,
    overwrite: bool = False,
    freeze_header: bool = True,
) -> Path:
    destination = Path(output_path) if output_path else source.with_suffix(".xlsx")
    ensure_output_path(destination, overwrite)
    write_xlsx(destination, rows, safe_sheet_title(source, sheet_name), freeze_header)
    return destination


def convert_csv_to_excel(
    csv_path: Path | str,
    output_path: Path | str | None = None,
    *,
    encoding: str = "utf-8-sig",
    delimiter: str | None = None,
    sheet_name: str | None = None,
    overwrite: bool = False,
    freeze_header: bool = True,
) -> Path:
    source = Path(csv_path)
    check_table_source(source, (".csv",), "CSV")

    dialect = detect_dialect(source, encoding, delimiter)
    with source.open("r", encoding=encoding, newline="") as csv_file:
        reader = csv.reader(csv_file, dialect)
        rows = [row for row in reader]

    return convert_rows_to_excel(
        rows,
        source,
        output_path,
        sheet_name=sheet_name,
        overwrite=overwrite,
        freeze_header=freeze_header,
    )


def convert_tsv_to_excel(
    tsv_path: Path | str,
    output_path: Path | str | None = None,
    *,
    encoding: str = "utf-8-sig",
    sheet_name: str | None = None,
    overwrite: bool = False,
    freeze_header: bool = True,
) -> Path:
    source = Path(tsv_path)
    check_table_source(source, (".tsv",), "TSV")
    with source.open("r", encoding=encoding, newline="") as tsv_file:
        rows = [row for row in csv.reader(tsv_file, delimiter="\t")]

    return convert_rows_to_excel(
        rows,
        source,
        output_path,
        sheet_name=sheet_name,
        overwrite=overwrite,
        freeze_header=freeze_header,
    )


def json_value_to_cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def json_to_rows(data: object) -> list[list[str]]:
    if isinstance(data, list):
        if not data:
            return []

        if all(isinstance(item, dict) for item in data):
            headers: list[str] = []
            for item in data:
                for key in item:
                    if key not in headers:
                        headers.append(str(key))
            rows = [headers]
            for item in data:
                item_dict = dict(item)
                rows.append([json_value_to_cell(item_dict.get(header)) for header in headers])
            return rows

        if all(isinstance(item, list) for item in data):
            return [[json_value_to_cell(value) for value in item] for item in data]

        return [["value"], *[[json_value_to_cell(item)] for item in data]]

    if isinstance(data, dict):
        return [["key", "value"], *[[str(key), json_value_to_cell(value)] for key, value in data.items()]]

    return [["value"], [json_value_to_cell(data)]]


def convert_json_to_excel(
    json_path: Path | str,
    output_path: Path | str | None = None,
    *,
    encoding: str = "utf-8-sig",
    sheet_name: str | None = None,
    overwrite: bool = False,
    freeze_header: bool = True,
) -> Path:
    source = Path(json_path)
    check_table_source(source, (".json",), "JSON")
    try:
        data = json.loads(source.read_text(encoding=encoding))
    except json.JSONDecodeError as error:
        raise ConversionError(f"Could not read JSON: {error}") from error

    return convert_rows_to_excel(
        json_to_rows(data),
        source,
        output_path,
        sheet_name=sheet_name,
        overwrite=overwrite,
        freeze_header=freeze_header,
    )


def files_in_directory(directory: Path, suffixes: tuple[str, ...]) -> Iterable[Path]:
    normalized_suffixes = tuple(suffix.lower() for suffix in suffixes)
    return sorted(path for path in directory.iterdir() if path.is_file() and path.suffix.lower() in normalized_suffixes)


def csv_files_in_directory(directory: Path) -> Iterable[Path]:
    return files_in_directory(directory, (".csv",))


def convert_batch(
    input_directory: Path,
    output_directory: Path,
    *,
    encoding: str,
    delimiter: str | None,
    overwrite: bool,
) -> list[Path]:
    if not input_directory.is_dir():
        raise ConversionError(f"Batch input must be a directory: {input_directory}")

    converted_files: list[Path] = []
    for csv_path in csv_files_in_directory(input_directory):
        output_path = output_directory / f"{csv_path.stem}.xlsx"
        converted_files.append(
            convert_csv_to_excel(
                csv_path,
                output_path,
                encoding=encoding,
                delimiter=delimiter,
                overwrite=overwrite,
            )
        )
    return converted_files


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Convert CSV files to Excel .xlsx workbooks.")
    parser.add_argument("input", type=Path, help="CSV file path, or a folder when using --batch.")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Output .xlsx file path, or output folder when using --batch.",
    )
    parser.add_argument("--batch", action="store_true", help="Convert all .csv files in the input folder.")
    parser.add_argument("--encoding", default="utf-8-sig", help="CSV text encoding. Default: utf-8-sig.")
    parser.add_argument("--delimiter", help="CSV delimiter. Default: auto-detect.")
    parser.add_argument("--sheet-name", help="Worksheet name for single-file conversion.")
    parser.add_argument("--overwrite", action="store_true", help="Replace existing output files.")
    parser.add_argument(
        "--no-freeze-header",
        action="store_true",
        help="Do not freeze the first row in the Excel workbook.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        if args.batch:
            output_directory = args.output or args.input
            converted_files = convert_batch(
                args.input,
                output_directory,
                encoding=args.encoding,
                delimiter=args.delimiter,
                overwrite=args.overwrite,
            )
            if not converted_files:
                print(f"No CSV files found in {args.input}")
                return 0

            for path in converted_files:
                print(f"Created {path}")
            return 0

        output_path = convert_csv_to_excel(
            args.input,
            args.output,
            encoding=args.encoding,
            delimiter=args.delimiter,
            sheet_name=args.sheet_name,
            overwrite=args.overwrite,
            freeze_header=not args.no_freeze_header,
        )
        print(f"Created {output_path}")
        return 0
    except ConversionError as error:
        parser.exit(1, f"Error: {error}\n")


if __name__ == "__main__":
    raise SystemExit(main())
