from __future__ import annotations

import struct
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
ASSETS_DIR = PROJECT_ROOT / "assets"
ICON_PATH = ASSETS_DIR / "file_converter.ico"
ENTRY_POINT = PROJECT_ROOT / "run_gui.py"


def color_at(x: int, y: int) -> tuple[int, int, int, int]:
    radius = 10
    left, top, right, bottom = 6, 6, 58, 58
    inside_rect = left <= x < right and top <= y < bottom
    inside_corner = (
        (left + radius <= x < right - radius)
        or (top + radius <= y < bottom - radius)
        or (x - (left + radius)) ** 2 + (y - (top + radius)) ** 2 <= radius**2
        or (x - (right - radius - 1)) ** 2 + (y - (top + radius)) ** 2 <= radius**2
        or (x - (left + radius)) ** 2 + (y - (bottom - radius - 1)) ** 2 <= radius**2
        or (x - (right - radius - 1)) ** 2 + (y - (bottom - radius - 1)) ** 2 <= radius**2
    )
    if not inside_rect or not inside_corner:
        return 0, 0, 0, 0

    if 18 <= x < 24 and 17 <= y < 47:
        return 255, 255, 255, 255
    if 18 <= x < 36 and 17 <= y < 23:
        return 255, 255, 255, 255
    if 18 <= x < 33 and 30 <= y < 36:
        return 255, 255, 255, 255

    if 37 <= x < 51 and 17 <= y < 23:
        return 255, 255, 255, 255
    if 37 <= x < 51 and 41 <= y < 47:
        return 255, 255, 255, 255
    if 34 <= x < 40 and 20 <= y < 44:
        return 255, 255, 255, 255
    if 45 <= x < 52 and 24 <= y < 40:
        return 20, 184, 166, 255

    if x >= 48 or y >= 49:
        return 11, 79, 74, 255
    if y >= 45:
        return 17, 94, 89, 255
    return 20, 184, 166, 255


def generate_icon(path: Path, size: int = 64) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    xor_rows = bytearray()
    for y in range(size - 1, -1, -1):
        row = bytearray()
        for x in range(size):
            red, green, blue, alpha = color_at(x, y)
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
    image_data = bitmap_header + xor_rows + and_mask
    icon_dir = struct.pack("<HHH", 0, 1, 1)
    icon_entry = struct.pack(
        "<BBBBHHII",
        size,
        size,
        0,
        0,
        1,
        32,
        len(image_data),
        len(icon_dir) + 16,
    )
    path.write_bytes(icon_dir + icon_entry + image_data)


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
