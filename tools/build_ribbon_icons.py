#!/usr/bin/env python3
"""Draw the custom Ribbon icons from their source in this file (#238).

Each icon is a navy 1F3864 rounded tile with a white glyph, 32 x 32 RGBA:

    dp_show_picker   Show Picker   calendar
    dp_reset         Reset         circular arrow
    dp_demo          Demo          screen with a play mark

Shapes are defined in a 32-unit square. Every pixel is sampled on an 8 x 8 grid
and the samples are averaged, which anti-aliases the edges. Only the Python
standard library is used, so the output does not depend on an imaging library.

Usage:
    python tools/build_ribbon_icons.py            write src/ribbon/images/<name>.png
    python tools/build_ribbon_icons.py --check    compare the tracked files, pixel by pixel

--check exits with 1 when a tracked icon is missing or differs from its drawing
in any pixel. It compares decoded pixels, not file bytes, because zlib builds
may compress the same pixels differently.
"""

from __future__ import annotations

import argparse
import math
import struct
import sys
import zlib
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
IMAGES = ROOT / "src" / "ribbon" / "images"
SIZE = 32
SAMPLES = 8
PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"

Color = tuple[int, int, int, int]
Inside = Callable[[float, float], bool]
Pixels = list[list[Color]]

NAVY: Color = (0x1F, 0x38, 0x64, 255)
WHITE: Color = (255, 255, 255, 255)


@dataclass(frozen=True)
class Shape:
    """A filled region and the box outside which it is never inside."""

    inside: Inside
    box: tuple[float, float, float, float]
    color: Color


def rounded_rect(x0: float, y0: float, x1: float, y1: float, radius: float, color: Color) -> Shape:
    def inside(x: float, y: float) -> bool:
        cx = min(max(x, x0 + radius), x1 - radius)
        cy = min(max(y, y0 + radius), y1 - radius)
        return (x - cx) ** 2 + (y - cy) ** 2 <= radius * radius

    return Shape(inside, (x0, y0, x1, y1), color)


def polygon(points: list[tuple[float, float]], color: Color) -> Shape:
    """A convex polygon, in either winding order."""

    def inside(x: float, y: float) -> bool:
        signs = set()
        for (ax, ay), (bx, by) in zip(points, points[1:] + points[:1], strict=True):
            cross = (bx - ax) * (y - ay) - (by - ay) * (x - ax)
            if cross:
                signs.add(cross > 0)
        return len(signs) < 2

    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return Shape(inside, (min(xs), min(ys), max(xs), max(ys)), color)


def ring_arc(cx: float, cy: float, radius: float, width: float, start: float, end: float, color: Color) -> Shape:
    """A band from radius - width to radius, from start to end degrees clockwise from 3 o'clock."""
    span = (end - start) % 360

    def inside(x: float, y: float) -> bool:
        distance = math.hypot(x - cx, y - cy)
        if distance < radius - width or distance > radius:
            return False
        angle = math.degrees(math.atan2(y - cy, x - cx))
        return (angle - start) % 360 <= span

    return Shape(inside, (cx - radius, cy - radius, cx + radius, cy + radius), color)


def frame(x0: float, y0: float, x1: float, y1: float, width: float, radius: float) -> list[Shape]:
    return [rounded_rect(x0, y0, x1, y1, radius, WHITE),
            rounded_rect(x0 + width, y0 + width, x1 - width, y1 - width, max(radius - width, 0), NAVY)]


def rect(x0: float, y0: float, x1: float, y1: float, radius: float = 0) -> Shape:
    return rounded_rect(x0, y0, x1, y1, radius, WHITE)


def tile() -> Shape:
    return rounded_rect(1, 1, 31, 31, 6, NAVY)


def show_picker() -> list[Shape]:
    shapes = [tile(), *frame(8, 9, 24, 24.5, 2, 1.5), rect(8, 9, 24, 13.5, 1.5),
              rect(11, 6.5, 13, 11), rect(19, 6.5, 21, 11)]
    shapes += [rect(cx - 1, cy - 1, cx + 1, cy + 1) for cy in (16.75, 20.75) for cx in (12, 16, 20)]
    return shapes


def reset() -> list[Shape]:
    cx, cy, radius, width, gap_from, gap_to, head = 16.0, 16.0, 7.2, 2.6, 95.0, 150.0, 5.2
    angle = math.radians(gap_to)
    px, py = cx + radius * math.cos(angle), cy + radius * math.sin(angle)
    tx, ty = math.sin(angle), -math.cos(angle)
    nx, ny = math.cos(angle), math.sin(angle)
    half = head * 0.62
    arrow = polygon([(px + nx * half, py + ny * half), (px - nx * half, py - ny * half),
                     (px + tx * head, py + ty * head)], WHITE)
    return [tile(), ring_arc(cx, cy, radius, width, gap_to, gap_from + 360, WHITE), arrow]


def demo() -> list[Shape]:
    return [tile(), *frame(7, 8, 25, 21.5, 2, 1.5), rect(15, 21.5, 17, 24.5), rect(11.5, 24, 20.5, 26, 1),
            polygon([(14, 11.5), (14, 18), (19.5, 14.75)], WHITE)]


ICONS: dict[str, Callable[[], list[Shape]]] = {
    "dp_show_picker": show_picker,
    "dp_reset": reset,
    "dp_demo": demo,
}


def render(shapes: list[Shape], size: int = SIZE) -> Pixels:
    """Average SAMPLES x SAMPLES samples per pixel; later shapes paint over earlier ones."""
    scale = SIZE / size
    count = SAMPLES * SAMPLES
    pixels: Pixels = []
    for row in range(size):
        line: list[Color] = []
        for col in range(size):
            red = green = blue = alpha = 0
            for sy in range(SAMPLES):
                y = (row + (sy + 0.5) / SAMPLES) * scale
                for sx in range(SAMPLES):
                    x = (col + (sx + 0.5) / SAMPLES) * scale
                    for shape in reversed(shapes):
                        x0, y0, x1, y1 = shape.box
                        if x0 <= x <= x1 and y0 <= y <= y1 and shape.inside(x, y):
                            r, g, b, a = shape.color
                            red += r * a
                            green += g * a
                            blue += b * a
                            alpha += a
                            break
            if alpha:
                line.append((round(red / alpha), round(green / alpha), round(blue / alpha), round(alpha / count)))
            else:
                line.append((0, 0, 0, 0))
        pixels.append(line)
    return pixels


def encode_png(pixels: Pixels) -> bytes:
    """8-bit RGBA PNG, filter type 0 on every row."""
    height, width = len(pixels), len(pixels[0])
    raw = b"".join(b"\x00" + bytes(channel for pixel in row for channel in pixel) for row in pixels)

    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    header = struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)
    return PNG_SIGNATURE + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b"")


def decode_png(data: bytes) -> Pixels:
    """Decode a non-interlaced 8-bit RGBA PNG with any of the five row filters."""
    if not data.startswith(PNG_SIGNATURE):
        raise ValueError("not a PNG file")
    position, idat, header = len(PNG_SIGNATURE), b"", b""
    while position < len(data):
        (length,) = struct.unpack(">I", data[position:position + 4])
        kind = data[position + 4:position + 8]
        body = data[position + 8:position + 8 + length]
        position += 12 + length
        if kind == b"IHDR":
            header = body
        elif kind == b"IDAT":
            idat += body
        elif kind == b"IEND":
            break
    width, height, depth, color_type, _compression, _filter, interlace = struct.unpack(">IIBBBBB", header)
    if (depth, color_type, interlace) != (8, 6, 0):
        raise ValueError("only non-interlaced 8-bit RGBA PNG files are supported")
    raw = zlib.decompress(idat)
    stride = width * 4
    previous = bytearray(stride)
    pixels: Pixels = []
    for row in range(height):
        start = row * (stride + 1)
        kind, line = raw[start], bytearray(raw[start + 1:start + 1 + stride])
        for i in range(stride):
            left = line[i - 4] if i >= 4 else 0
            up = previous[i]
            upper_left = previous[i - 4] if i >= 4 else 0
            if kind == 1:
                line[i] = (line[i] + left) & 0xFF
            elif kind == 2:
                line[i] = (line[i] + up) & 0xFF
            elif kind == 3:
                line[i] = (line[i] + (left + up) // 2) & 0xFF
            elif kind == 4:
                guess = left + up - upper_left
                pa, pb, pc = abs(guess - left), abs(guess - up), abs(guess - upper_left)
                predictor = left if pa <= pb and pa <= pc else up if pb <= pc else upper_left
                line[i] = (line[i] + predictor) & 0xFF
            elif kind != 0:
                raise ValueError(f"unknown PNG filter type {kind}")
        pixels.append([(line[i], line[i + 1], line[i + 2], line[i + 3]) for i in range(0, stride, 4)])
        previous = line
    return pixels


def check(images: Path = IMAGES) -> list[str]:
    """Findings for tracked icons that are missing or differ from their drawing."""
    findings = []
    for name, draw in ICONS.items():
        path = images / f"{name}.png"
        if not path.is_file():
            findings.append(f"{path.name}: missing")
            continue
        try:
            tracked = decode_png(path.read_bytes())
        except (ValueError, zlib.error, struct.error) as error:
            findings.append(f"{path.name}: cannot be decoded ({error})")
            continue
        expected = render(draw())
        if tracked != expected:
            changed = sum(a != b for row_a, row_b in zip(tracked, expected, strict=False)
                          for a, b in zip(row_a, row_b, strict=False))
            findings.append(f"{path.name}: differs from its drawing in tools/build_ribbon_icons.py "
                            f"({changed} pixels, or a different size)")
    return findings


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--check", action="store_true", help="compare the tracked icons instead of writing them")
    parser.add_argument("--images", type=Path, default=IMAGES, help="folder of the icons (default: src/ribbon/images)")
    args = parser.parse_args()
    if args.check:
        findings = check(args.images)
        for finding in findings:
            print(finding)
        print("Ribbon icons: " + ("FAIL" if findings else "PASS"))
        return 1 if findings else 0
    args.images.mkdir(parents=True, exist_ok=True)
    for name, draw in ICONS.items():
        (args.images / f"{name}.png").write_bytes(encode_png(render(draw())))
        print(f"wrote {name}.png")
    return 0


if __name__ == "__main__":
    sys.exit(main())
