from pathlib import Path
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from csv_to_excel.archive import create_zip_archive, extract_zip_archive
from csv_to_excel.converter import ConversionError


class ArchiveConverterTests(unittest.TestCase):
    def test_create_zip_archive_from_folder(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            source = temp_dir / "project"
            source.mkdir()
            (source / "notes.txt").write_text("hello", encoding="utf-8")
            nested = source / "nested"
            nested.mkdir()
            (nested / "data.json").write_text('{"ok": true}', encoding="utf-8")

            output = create_zip_archive(source)

            self.assertTrue(output.is_file())
            with zipfile.ZipFile(output) as archive:
                self.assertEqual(sorted(archive.namelist()), ["nested/data.json", "notes.txt"])

    def test_extract_zip_archive_to_folder(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            zip_path = temp_dir / "files.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("docs/readme.txt", "ready")

            output = extract_zip_archive(zip_path)

            self.assertEqual((output / "docs" / "readme.txt").read_text(encoding="utf-8"), "ready")

    def test_extract_zip_archive_blocks_path_traversal(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            zip_path = temp_dir / "unsafe.zip"
            with zipfile.ZipFile(zip_path, "w") as archive:
                archive.writestr("../outside.txt", "nope")

            with self.assertRaises(ConversionError):
                extract_zip_archive(zip_path)


if __name__ == "__main__":
    unittest.main()
