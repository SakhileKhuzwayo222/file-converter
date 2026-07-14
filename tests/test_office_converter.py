from pathlib import Path
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from csv_to_excel.office import convert_epub_to_pdf, convert_pdf_to_html, convert_pdf_to_word, convert_text_to_word


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

    def test_convert_pdf_to_html_extracts_text(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            pdf_path = temp_dir / "sample.pdf"
            write_simple_pdf(pdf_path)

            output_path = convert_pdf_to_html(pdf_path)

            html_text = output_path.read_text(encoding="utf-8")
            self.assertEqual(output_path.suffix, ".html")
            self.assertIn("<title>sample</title>", html_text)
            self.assertIn("Hello PDF", html_text)
            self.assertIn("Second line", html_text)

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
