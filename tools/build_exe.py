from __future__ import annotations

import math
import struct
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
ASSETS_DIR = PROJECT_ROOT / "assets"
ICON_PATH = ASSETS_DIR / "file_converter.ico"
ENTRY_POINT = PROJECT_ROOT / "run_gui.py"
SRC_PATH = PROJECT_ROOT / "src"


def inside_rect(x: float, y: float, left: float, top: float, width: float, height: float) -> bool:
    return left <= x <= left + width and top <= y <= top + height


def inside_rounded_rect(x: float, y: float, left: float, top: float, width: float, height: float, radius: float) -> bool:
    if not inside_rect(x, y, left, top, width, height):
        return False
    nearest_x = min(max(x, left + radius), left + width - radius)
    nearest_y = min(max(y, top + radius), top + height - radius)
    return math.hypot(x - nearest_x, y - nearest_y) <= radius


def fc_color_at(x: float, y: float) -> tuple[int, int, int, int]:
    color = (0, 0, 0, 0)
    if inside_rounded_rect(x, y, 2.5, 2.0, 19.0, 20.0, 3.2):
        color = (15, 118, 110, 255)
    if inside_rounded_rect(x, y, 2.0, 2.0, 18.8, 18.8, 3.0):
        color = (20, 184, 166, 255)
    if inside_rect(x, y, 17.2, 6.0, 3.6, 14.8):
        color = (11, 79, 74, 255)
    if inside_rect(x, y, 4.7, 17.2, 14.3, 3.6):
        color = (15, 118, 110, 255)

    white = (255, 255, 255, 255)
    if inside_rect(x, y, 6.0, 6.0, 2.0, 12.0):
        color = white
    if inside_rect(x, y, 6.0, 6.0, 7.4, 2.0):
        color = white
    if inside_rect(x, y, 6.0, 11.0, 6.0, 2.0):
        color = white
    if inside_rect(x, y, 14.6, 6.0, 5.0, 2.0):
        color = white
    if inside_rect(x, y, 12.6, 8.0, 2.0, 8.0):
        color = white
    if inside_rect(x, y, 14.6, 16.0, 5.0, 2.0):
        color = white
    return color


def icon_pixel_at(pixel_x: int, pixel_y: int, size: int, samples: int = 4) -> tuple[int, int, int, int]:
    red = green = blue = alpha = 0
    for sample_y in range(samples):
        for sample_x in range(samples):
            x = (pixel_x + (sample_x + 0.5) / samples) / size * 24.0
            y = (pixel_y + (sample_y + 0.5) / samples) / size * 24.0
            sample_red, sample_green, sample_blue, sample_alpha = fc_color_at(x, y)
            red += sample_red
            green += sample_green
            blue += sample_blue
            alpha += sample_alpha
    sample_count = samples * samples
    return red // sample_count, green // sample_count, blue // sample_count, alpha // sample_count


def icon_image_data(size: int) -> bytes:
    xor_rows = bytearray()
    for y in range(size - 1, -1, -1):
        row = bytearray()
        for x in range(size):
            red, green, blue, alpha = icon_pixel_at(x, y, size)
            row.extend((blue, green, red, alpha))
        xor_rows.extend(row)

    mask_stride = ((size + 31) // 32) * 4
    and_mask = bytes(mask_stride * size)
    bitmap_header = struct.pack(
        "<IIIHHIIIIII",
        40,
        size,
        size * 2,
        1,
        32,
        0,
        len(xor_rows) + len(and_mask),
        0,
        0,
        0,
        0,
    )
    return bitmap_header + xor_rows + and_mask


def generate_icon(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    sizes = (16, 24, 32, 48, 64, 128, 256)
    images = [icon_image_data(size) for size in sizes]
    icon_dir = struct.pack("<HHH", 0, 1, len(images))
    offset = len(icon_dir) + len(images) * 16
    entries = bytearray()
    for size, image_data in zip(sizes, images):
        size_byte = 0 if size == 256 else size
        entries.extend(struct.pack("<BBBBHHII", size_byte, size_byte, 0, 0, 1, 32, len(image_data), offset))
        offset += len(image_data)
    path.write_bytes(icon_dir + bytes(entries) + b"".join(images))


def run_pyinstaller() -> None:
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        "File Converter",
        "--icon",
        str(ICON_PATH),
        "--paths",
        str(SRC_PATH),
        "--hidden-import",
        "csv_to_excel",
        "--hidden-import",
        "csv_to_excel.converter",
        "--hidden-import",
        "csv_to_excel.external",
        "--hidden-import",
        "csv_to_excel.gui",
        "--hidden-import",
        "csv_to_excel.office",
        str(ENTRY_POINT),
    ]
    subprocess.run(command, cwd=PROJECT_ROOT, check=True)


def main() -> int:
    generate_icon(ICON_PATH)
    run_pyinstaller()
    print(f"Created {PROJECT_ROOT / 'dist' / 'File Converter.exe'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
