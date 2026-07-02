from pathlib import Path
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from csv_to_excel.converter import ConversionError, convert_csv_to_excel, convert_json_to_excel, convert_tsv_to_excel


NAMESPACE = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def worksheet_values(xlsx_path: Path) -> list[str]:
    with zipfile.ZipFile(xlsx_path) as workbook:
        xml = workbook.read("xl/worksheets/sheet1.xml")
    root = ElementTree.fromstring(xml)
    return [
        text_node.text or ""
        for text_node in root.findall(".//main:c/main:is/main:t", NAMESPACE)
    ]


class ConverterTests(unittest.TestCase):
    def test_convert_csv_to_excel_creates_workbook(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            csv_path = temp_dir / "people.csv"
            csv_path.write_text("name,role\nAda,Engineer\nGrace,Scientist\n", encoding="utf-8")

            output_path = convert_csv_to_excel(csv_path)

            self.assertTrue(output_path.exists())
            self.assertEqual(
                worksheet_values(output_path),
                ["name", "role", "Ada", "Engineer", "Grace", "Scientist"],
            )

    def test_convert_csv_to_excel_supports_custom_delimiter(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            csv_path = temp_dir / "inventory.csv"
            csv_path.write_text("item;qty\nCable;3\n", encoding="utf-8")
            output_path = temp_dir / "inventory.xlsx"

            convert_csv_to_excel(csv_path, output_path, delimiter=";")

            self.assertEqual(worksheet_values(output_path), ["item", "qty", "Cable", "3"])

    def test_convert_csv_to_excel_refuses_to_overwrite_by_default(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            csv_path = temp_dir / "sample.csv"
            csv_path.write_text("a,b\n1,2\n", encoding="utf-8")
            output_path = temp_dir / "sample.xlsx"

            convert_csv_to_excel(csv_path, output_path)

            with self.assertRaisesRegex(ConversionError, "already exists"):
                convert_csv_to_excel(csv_path, output_path)

    def test_convert_tsv_to_excel(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            tsv_path = temp_dir / "inventory.tsv"
            tsv_path.write_text("item\tqty\nCable\t3\n", encoding="utf-8")

            output_path = convert_tsv_to_excel(tsv_path)

            self.assertEqual(worksheet_values(output_path), ["item", "qty", "Cable", "3"])

    def test_convert_json_to_excel_with_records(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            temp_dir = Path(directory)
            json_path = temp_dir / "people.json"
            json_path.write_text('[{"name": "Ada", "role": "Engineer"}, {"name": "Grace"}]', encoding="utf-8")

            output_path = convert_json_to_excel(json_path)

            self.assertEqual(worksheet_values(output_path), ["name", "role", "Ada", "Engineer", "Grace", ""])


if __name__ == "__main__":
    unittest.main()
