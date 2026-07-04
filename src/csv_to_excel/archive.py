from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from .converter import ConversionError, ensure_output_path


ZIP_EXTENSIONS = (".zip",)


def check_zip_source(source: Path) -> None:
    if not source.exists():
        raise ConversionError(f"ZIP file not found: {source}")
    if not source.is_file():
        raise ConversionError(f"Input must be a ZIP file: {source}")
    if source.suffix.lower() not in ZIP_EXTENSIONS:
        raise ConversionError(f"Input file must end with .zip: {source}")


def check_folder_source(source: Path) -> None:
    if not source.exists():
        raise ConversionError(f"Folder not found: {source}")
    if not source.is_dir():
        raise ConversionError(f"Input must be a folder: {source}")


def safe_zip_member_path(destination: Path, member_name: str) -> Path:
    clean_name = member_name.replace("\\", "/")
    if not clean_name or clean_name.startswith("/") or clean_name.startswith("../") or "/../" in clean_name:
        raise ConversionError(f"Blocked unsafe ZIP entry: {member_name}")

    target = (destination / clean_name).resolve()
    destination_root = destination.resolve()
    if target != destination_root and destination_root not in target.parents:
        raise ConversionError(f"Blocked unsafe ZIP entry: {member_name}")
    return target


def extract_zip_archive(
    zip_path: Path | str,
    output_path: Path | str | None = None,
    *,
    overwrite: bool = False,
) -> Path:
    source = Path(zip_path)
    check_zip_source(source)
    destination = Path(output_path) if output_path else source.with_name(f"{source.stem} (Extracted)")

    if destination.exists() and not overwrite:
        raise ConversionError(f"Output folder already exists: {destination}")
    destination.mkdir(parents=True, exist_ok=True)

    try:
        with zipfile.ZipFile(source) as archive:
            for member in archive.infolist():
                target = safe_zip_member_path(destination, member.filename)
                if member.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                    continue
                target.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(member) as source_file, target.open("wb") as target_file:
                    shutil.copyfileobj(source_file, target_file)
    except zipfile.BadZipFile as error:
        raise ConversionError("This ZIP file could not be opened as a valid archive.") from error

    return destination


def create_zip_archive(
    folder_path: Path | str,
    output_path: Path | str | None = None,
    *,
    overwrite: bool = False,
) -> Path:
    source = Path(folder_path)
    check_folder_source(source)
    destination = Path(output_path) if output_path else source.with_suffix(".zip")
    ensure_output_path(destination, overwrite)

    destination_resolved = destination.resolve()
    files = [path for path in source.rglob("*") if path.is_file() and path.resolve() != destination_resolved]
    if not files:
        raise ConversionError(f"No files were found in the selected folder: {source}")

    with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in files:
            archive.write(path, path.relative_to(source).as_posix())

    return destination
