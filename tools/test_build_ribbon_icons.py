"""Fixtures for build_ribbon_icons.py: the tracked Ribbon icons match their drawing (#238)."""
import io
import shutil
import struct
import tempfile
import unittest
import zlib
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

import build_ribbon_icons as icons

NAVY = (0x1F, 0x38, 0x64, 255)
WHITE = (255, 255, 255, 255)
CLEAR = (0, 0, 0, 0)


class RibbonIconTests(unittest.TestCase):
    def setUp(self) -> None:
        self.folder = Path(tempfile.mkdtemp())

    def tearDown(self) -> None:
        shutil.rmtree(self.folder)

    def copy_icons(self) -> Path:
        target = self.folder / "images"
        shutil.copytree(icons.IMAGES, target)
        return target

    def test_tracked_icons_match_their_drawing(self) -> None:
        self.assertEqual(icons.check(), [])

    def test_every_icon_is_a_32_pixel_navy_tile_with_white_glyph(self) -> None:
        for name, draw in icons.ICONS.items():
            pixels = icons.render(draw())
            with self.subTest(icon=name):
                self.assertEqual((len(pixels), len(pixels[0])), (32, 32))
                self.assertEqual(pixels[0][0], CLEAR)
                self.assertEqual(pixels[16][3], NAVY)
                self.assertIn(WHITE, {pixel for row in pixels for pixel in row})

    def test_png_round_trip_keeps_every_pixel(self) -> None:
        pixels = icons.render(icons.reset())
        self.assertEqual(icons.decode_png(icons.encode_png(pixels)), pixels)

    def test_decoder_reverses_every_row_filter(self) -> None:
        pixels = icons.render(icons.demo())
        stride = 32 * 4
        rows = [bytes(channel for pixel in row for channel in pixel) for row in pixels]
        raw = bytearray()
        previous = bytes(stride)
        for index, row in enumerate(rows):
            kind = index % 5
            out = bytearray()
            for i in range(stride):
                left = row[i - 4] if i >= 4 else 0
                up = previous[i]
                upper_left = previous[i - 4] if i >= 4 else 0
                if kind == 1:
                    out.append((row[i] - left) & 0xFF)
                elif kind == 2:
                    out.append((row[i] - up) & 0xFF)
                elif kind == 3:
                    out.append((row[i] - (left + up) // 2) & 0xFF)
                elif kind == 4:
                    guess = left + up - upper_left
                    pa, pb, pc = abs(guess - left), abs(guess - up), abs(guess - upper_left)
                    predictor = left if pa <= pb and pa <= pc else up if pb <= pc else upper_left
                    out.append((row[i] - predictor) & 0xFF)
                else:
                    out.append(row[i])
            raw += bytes([kind]) + out
            previous = row
        encoded = icons.encode_png(pixels)
        idat_start = encoded.index(b"IDAT") - 4
        (length,) = struct.unpack(">I", encoded[idat_start:idat_start + 4])
        data = zlib.compress(bytes(raw))
        chunk = struct.pack(">I", len(data)) + b"IDAT" + data + struct.pack(">I", zlib.crc32(b"IDAT" + data))
        refiltered = encoded[:idat_start] + chunk + encoded[idat_start + 12 + length:]
        self.assertEqual(icons.decode_png(refiltered), pixels)

    def test_check_reports_a_changed_pixel(self) -> None:
        images = self.copy_icons()
        pixels = icons.decode_png((images / "dp_demo.png").read_bytes())
        pixels[10][10] = (255, 0, 0, 255)
        (images / "dp_demo.png").write_bytes(icons.encode_png(pixels))
        findings = icons.check(images)
        self.assertEqual(len(findings), 1)
        self.assertIn("dp_demo.png: differs from its drawing", findings[0])
        self.assertIn("1 pixels", findings[0])

    def test_check_reports_a_missing_or_unreadable_icon(self) -> None:
        images = self.copy_icons()
        (images / "dp_reset.png").unlink()
        (images / "dp_show_picker.png").write_bytes(b"not a png")
        self.assertEqual(icons.check(images), ["dp_show_picker.png: cannot be decoded (not a PNG file)",
                                               "dp_reset.png: missing"])

    def test_cli_writes_then_checks(self) -> None:
        images = self.folder / "built"
        with redirect_stdout(io.StringIO()):
            with mock.patch("sys.argv", ["build_ribbon_icons.py", "--images", str(images)]):
                self.assertEqual(icons.main(), 0)
            with mock.patch("sys.argv", ["build_ribbon_icons.py", "--check", "--images", str(images)]):
                self.assertEqual(icons.main(), 0)
            (images / "dp_demo.png").unlink()
            with mock.patch("sys.argv", ["build_ribbon_icons.py", "--check", "--images", str(images)]):
                self.assertEqual(icons.main(), 1)


if __name__ == "__main__":
    unittest.main()
