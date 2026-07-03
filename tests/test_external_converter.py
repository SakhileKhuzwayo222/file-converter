from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from csv_to_excel.converter import ConversionError
from csv_to_excel.external import AUDIO_EXTENSIONS, convert_media_with_ffmpeg, powershell_quote, prepare_destination


class ExternalConverterTests(unittest.TestCase):
    def test_powershell_quote_escapes_single_quotes(self) -> None:
        self.assertEqual(powershell_quote("C:\\Docs\\Sam's File.docx"), "'C:\\Docs\\Sam''s File.docx'")

    def test_prepare_destination_uses_requested_extension(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "track.wav"
            source.write_text("placeholder", encoding="utf-8")

            destination = prepare_destination(source, None, ".mp3", overwrite=False)

            self.assertEqual(destination, source.with_suffix(".mp3"))
            self.assertTrue(destination.parent.exists())

    def test_media_conversion_reports_missing_ffmpeg(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "track.wav"
            source.write_text("placeholder", encoding="utf-8")
            output = Path(temp_dir) / "track.mp3"

            with patch("csv_to_excel.external.shutil.which", return_value=None):
                with self.assertRaises(ConversionError) as raised:
                    convert_media_with_ffmpeg(source, output, ".mp3", AUDIO_EXTENSIONS, overwrite=False)

            self.assertIn("FFmpeg", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
