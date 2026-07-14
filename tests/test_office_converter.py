from pathlib import Path
import sys
import tempfile
import unittest
import zipfile
import zlib
from xml.etree import ElementTree

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from csv_to_excel.converter import ConversionError
from csv_to_excel.office import (
    MAX_BASIC_PDF_STREAM_BYTES,
    MAX_RAW_IMAGE_STREAM_BYTES,
    convert_epub_to_pdf,
    convert_pdf_to_html,
    convert_pdf_to_word,
    convert_text_to_word,
    decompress_pdf_stream,
    extract_pdf_html_blocks,
    extract_text_from_pdf_content,
    iter_pdf_stream_objects,
    read_pdf_with_mmap,
)


NAMESPACE = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def docx_text(docx_path: Path) -> str:
    with zipfile.ZipFile(docx_path) as document:
        xml = document.read("word/document.xml")
    root = ElementTree.fromstring(xml)
    return "\n".join(text_node.text or "" for text_node in root.findall(".//w:t", NAMESPACE))


def write_simple_pdf(pdf_path: Path) -> None:
    stream = b"BT /F1 12 Tf 72 720 Td (Hello PDF) Tj T* (Second line) Tj ET"
    pdf = (
        b"%PDF-1.4\n"
        b"1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n"
        b"2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj\n"
        b"3 0 obj << /Type /Page /Parent 2 0 R /Contents 4 0 R >> endobj\n"
        b"4 0 obj << /Length "
        + str(len(stream)).encode("ascii")
        + b" >>\nstream\n"
        + stream
        + b"\nendstream\nendobj\n%%EOF\n"
    )
    pdf_path.write_bytes(pdf)


def write_image_pdf(pdf_path: Path) -> None:
    image = b"\xff\xd8\xff\xd9"
    pdf = (
        b"%PDF-1.4\n"
        b"1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n"
        b"2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj\n"
        b"3 0 obj << /Type /Page /Parent 2 0 R /Resources << /XObject << /Im1 4 0 R >> >> >> endobj\n"
        b"4 0 obj << /Type /XObject /Subtype /Image /Width 1 /Height 1 "
        b"/BitsPerComponent 8 /ColorSpace /DeviceRGB /Filter /DCTDecode /Length "
        + str(len(image)).encode("ascii")
        + b" >>\nstream\n"
        + image
        + b"\nendstream\nendobj\n%%EOF\n"
    )
    pdf_path.write_bytes(pdf)


def write_wrapped_jpeg_pdf(pdf_path: Path) -> bytes:
    image = b"\xff\xd8\xff\xe0wrapped-jpeg\xff\xd9"
    wrapped_image = zlib.compress(image)
    pdf = (
        b"%PDF-1.4\n"
        b"1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n"
        b"2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj\n"
        b"3 0 obj << /Type /Page /Parent 2 0 R /Resources << /XObject << /Im1 4 0 R >> >> >> endobj\n"
        b"4 0 obj << /Type /XObject /Subtype /Image /Width 1 /Height 1 "
        b"/BitsPerComponent 8 /ColorSpace /DeviceRGB /Filter [/FlateDecode /DCTDecode] /Length "
        + str(len(wrapped_image)).encode("ascii")
        + b" >>\nstream\n"
        + wrapped_image
        + b"\nendstream\nendobj\n%%EOF\n"
    )
    pdf_path.write_bytes(pdf)
    return image


def write_masked_image_pdf(pdf_path: Path) -> None:
    image = b"\xff\xd8\xff\xe0main-image\xff\xd9"
    mask = b"\xff\xd8\xff\xe0mask-image\xff\xd9"
    pdf = (
        b"%PDF-1.4\n"
        b"1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n"
        b"2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj\n"
        b"3 0 obj << /Type /Page /Parent 2 0 R /Resources << /XObject << /Im1 4 0 R /Mask1 5 0 R >> >> >> endobj\n"
        b"4 0 obj << /Type /XObject /Subtype /Image /Width 1 /Height 1 "
        b"/BitsPerComponent 8 /ColorSpace /DeviceRGB /Filter /DCTDecode /SMask 5 0 R /Length "
        + str(len(image)).encode("ascii")
        + b" >>\nstream\n"
        + image
        + b"\nendstream\nendobj\n"
        b"5 0 obj << /Type /XObject /Subtype /Image /Width 1 /Height 1 "
        b"/BitsPerComponent 8 /ColorSpace /DeviceGray /Filter /DCTDecode /Length "
        + str(len(mask)).encode("ascii")
        + b" >>\nstream\n"
        + mask
        + b"\nendstream\nendobj\n%%EOF\n"
    )
    pdf_path.write_bytes(pdf)


def write_text_image_text_pdf(pdf_path: Path) -> None:
    before = b"BT /F1 12 Tf 72 720 Td (Before image) Tj ET"
    image = b"\xff\xd8\xff\xd9"
    after = b"BT /F1 12 Tf 72 680 Td (After image) Tj ET"
    pdf = (
        b"%PDF-1.4\n"
        b"1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n"
        b"2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj\n"
        b"3 0 obj << /Type /Page /Parent 2 0 R /Contents [4 0 R 5 0 R 6 0 R] >> endobj\n"
        b"4 0 obj << /Length "
        + str(len(before)).encode("ascii")
        + b" >>\nstream\n"
        + before
        + b"\nendstream\nendobj\n"
        b"5 0 obj << /Type /XObject /Subtype /Image /Width 1 /Height 1 "
        b"/BitsPerComponent 8 /ColorSpace /DeviceRGB /Filter /DCTDecode /Length "
        + str(len(image)).encode("ascii")
        + b" >>\nstream\n"
        + image
        + b"\nendstream\nendobj\n"
        b"6 0 obj << /Length "
        + str(len(after)).encode("ascii")
        + b" >>\nstream\n"
        + after
        + b"\nendstream\nendobj\n%%EOF\n"
    )
    pdf_path.write_bytes(pdf)


def write_simple_epub(epub_path: Path) -> None:
    with zipfile.ZipFile(epub_path, "w") as book:
        book.writestr("mimetype", "application/epub+zip")
        book.writestr(
            "META-INF/container.xml",
            """<?xml version="1.0"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>""",
        )
        book.writestr(
            "OEBPS/content.opf",
            """<?xml version="1.0"?>
<package version="3.0" xmlns="http://www.idpf.org/2007/opf">
  <manifest>
    <item id="chapter" href="chapter.xhtml" media-type="application/xhtml+xml"/>
  </manifest>
  <spine>
    <itemref idref="chapter"/>
  </spine>
</package>""",
        )
        book.writestr(
            "OEBPS/chapter.xhtml",
            """<?xml version="1.0"?>
<html xmlns="http://www.w3.org/1999/xhtml">
  <body><h1>Hello EPUB</h1><p>This chapter becomes a PDF.</p></body>
</html>""",
        )


class OfficeConverterTests(unittest.TestCase):
    def test_convert_text_to_word_creates_docx(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            text_path = temp_dir / "notes.txt"
            text_path.write_text("First paragraph.\n\nSecond paragraph.", encoding="utf-8")

            output_path = convert_text_to_word(text_path)

            text = docx_text(output_path)
            self.assertIn("notes", text)
            self.assertIn("First paragraph.", text)
            self.assertIn("Second paragraph.", text)

    def test_convert_editable_data_to_word_creates_docx(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            text_path = temp_dir / "settings.json"
            text_path.write_text('{"project": "File Converter"}', encoding="utf-8")

            output_path = convert_text_to_word(text_path)

            text = docx_text(output_path)
            self.assertIn("settings", text)
            self.assertIn("File Converter", text)

    def test_convert_pdf_to_word_extracts_text(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            pdf_path = temp_dir / "sample.pdf"
            write_simple_pdf(pdf_path)

            output_path = convert_pdf_to_word(pdf_path)

            text = docx_text(output_path)
            self.assertIn("sample", text)
            self.assertIn("Hello PDF", text)
            self.assertIn("Second line", text)

    def test_convert_pdf_to_html_embeds_original_pdf_for_exact_layout(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            pdf_path = temp_dir / "sample.pdf"
            write_simple_pdf(pdf_path)

            output_path = convert_pdf_to_html(pdf_path)

            asset_path = temp_dir / "sample_assets" / "sample.pdf"
            html_text = output_path.read_text(encoding="utf-8")
            self.assertEqual(output_path.suffix, ".html")
            self.assertEqual(asset_path.read_bytes(), pdf_path.read_bytes())
            self.assertIn("<title>sample</title>", html_text)
            self.assertIn('class="pdf-document"', html_text)
            self.assertIn('data="sample_assets/sample.pdf#toolbar=0&amp;navpanes=0&amp;scrollbar=1&amp;view=FitH"', html_text)
            self.assertIn("height: 100vh", html_text)

    def test_convert_pdf_to_html_copies_pdf_without_extracting_duplicate_images(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            pdf_path = temp_dir / "scanned.pdf"
            write_image_pdf(pdf_path)

            output_path = convert_pdf_to_html(pdf_path)

            asset_dir = temp_dir / "scanned_assets"
            asset_path = asset_dir / "scanned.pdf"
            html_text = output_path.read_text(encoding="utf-8")
            self.assertTrue(asset_path.exists())
            self.assertEqual(list(asset_dir.glob("*.jpg")), [])
            self.assertIn('data="scanned_assets/scanned.pdf#toolbar=0', html_text)

    def test_convert_pdf_to_html_removes_legacy_extracted_images(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            pdf_path = temp_dir / "sample.pdf"
            asset_dir = temp_dir / "sample_assets"
            asset_dir.mkdir()
            stale_asset = asset_dir / "image_1.jpg"
            stale_asset.write_bytes(b"old extracted image")
            write_simple_pdf(pdf_path)

            convert_pdf_to_html(pdf_path)

            self.assertFalse(stale_asset.exists())
            self.assertTrue((asset_dir / "sample.pdf").exists())

    def test_pdf_image_extraction_unwraps_flate_encoded_jpegs(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            pdf_path = temp_dir / "wrapped.pdf"
            output_path = temp_dir / "wrapped.html"
            expected_image = write_wrapped_jpeg_pdf(pdf_path)

            blocks = read_pdf_with_mmap(pdf_path, lambda data: extract_pdf_html_blocks(data, output_path))

            asset_path = temp_dir / "wrapped_assets" / "image_1.jpg"
            self.assertTrue(any(block.kind == "image" for block in blocks))
            self.assertEqual(asset_path.read_bytes(), expected_image)

    def test_pdf_image_extraction_skips_soft_mask_duplicates(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            pdf_path = temp_dir / "masked.pdf"
            output_path = temp_dir / "masked.html"
            write_masked_image_pdf(pdf_path)

            blocks = read_pdf_with_mmap(pdf_path, lambda data: extract_pdf_html_blocks(data, output_path))

            assets = sorted((temp_dir / "masked_assets").glob("*"))
            image_blocks = [block for block in blocks if block.kind == "image"]
            self.assertEqual(len(image_blocks), 1)
            self.assertEqual([path.name for path in assets], ["image_1.jpg"])

    def test_pdf_image_extraction_places_images_in_stream_order(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            pdf_path = temp_dir / "mixed.pdf"
            output_path = temp_dir / "mixed.html"
            write_text_image_text_pdf(pdf_path)

            blocks = read_pdf_with_mmap(pdf_path, lambda data: extract_pdf_html_blocks(data, output_path))

            self.assertIn("Before image", blocks[0].text)
            self.assertEqual(blocks[1].kind, "image")
            self.assertIn("After image", blocks[2].text)

    def test_convert_pdf_to_html_reports_asset_folder_errors(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            pdf_path = temp_dir / "scanned.pdf"
            output_path = temp_dir / "converted.html"
            write_image_pdf(pdf_path)
            (temp_dir / "converted_assets").write_text("not a folder", encoding="utf-8")

            with self.assertRaises(ConversionError) as raised:
                convert_pdf_to_html(pdf_path, output_path)

            self.assertIn("Could not copy the PDF for exact HTML viewing", str(raised.exception))

    def test_pdf_stream_decompression_rejects_oversized_output(self) -> None:
        payload = zlib.compress(b"x" * (MAX_BASIC_PDF_STREAM_BYTES + 1))

        self.assertIsNone(decompress_pdf_stream(payload, MAX_BASIC_PDF_STREAM_BYTES))

    def test_pdf_text_parser_skips_dictionary_delimiters(self) -> None:
        content = b"BT << /ActualText (ignored) >> /F1 12 Tf (Visible text) Tj ET"

        self.assertIn("Visible text", extract_text_from_pdf_content(content))

    def test_pdf_stream_iterator_skips_declared_oversized_images(self) -> None:
        data = (
            b"%PDF-1.4\n"
            b"1 0 obj << /Type /XObject /Subtype /Image /Width 1 /Height 1 "
            b"/BitsPerComponent 8 /ColorSpace /DeviceRGB /Filter /DCTDecode /Length "
            + str(MAX_RAW_IMAGE_STREAM_BYTES + 1).encode("ascii")
            + b" >>\nstream\n"
            + b"\xff\xd8\xff\xd9"
            + b"\nendstream\nendobj\n%%EOF\n"
        )

        streams = list(iter_pdf_stream_objects(data))

        self.assertEqual(len(streams), 1)
        self.assertEqual(streams[0][1], b"")

    def test_convert_epub_to_pdf_creates_pdf(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            epub_path = temp_dir / "book.epub"
            write_simple_epub(epub_path)

            output_path = convert_epub_to_pdf(epub_path)

            data = output_path.read_bytes()
            self.assertTrue(data.startswith(b"%PDF-1.4"))
            self.assertIn(b"Hello EPUB", data)
            self.assertIn(b"This chapter becomes a PDF.", data)


if __name__ == "__main__":
    unittest.main()
