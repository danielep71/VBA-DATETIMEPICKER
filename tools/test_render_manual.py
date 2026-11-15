"""Fixtures for render_manual.py: alerts, embedded images, pinned links and the browser lookup."""
import base64
import importlib.util
import shutil
import tempfile
import unittest
from pathlib import Path, PurePosixPath

import render_manual as renderer

PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")


class AlertTests(unittest.TestCase):
    def test_alert_becomes_a_markdown_div(self) -> None:
        text = "Intro\n\n> [!NOTE]\n> First line\n> second line\n\nAfter\n"
        converted = renderer.convert_alerts(text)
        self.assertIn('<div class="alert alert-note" markdown="1">', converted)
        self.assertIn("**Note**", converted)
        self.assertIn("First line\nsecond line", converted)
        self.assertTrue(converted.rstrip().endswith("After"))
        self.assertNotIn("[!NOTE]", converted)

    def test_div_block_is_marked_for_markdown(self) -> None:
        converted = renderer.convert_alerts('<div align="center">\n\n# Title\n\n</div>\n')
        self.assertIn('<div align="center" markdown="1">', converted)
        self.assertEqual(renderer.convert_alerts('<div markdown="1">\n'), '<div markdown="1">\n')

    def test_plain_blockquote_is_unchanged(self) -> None:
        text = "> Select the cells\n> and pick.\n"
        self.assertEqual(renderer.convert_alerts(text), text)


class ImageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.root = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.root)
        (self.root / "assets").mkdir()
        (self.root / "docs").mkdir()
        (self.root / "assets" / "dot.png").write_bytes(PNG)

    def test_relative_image_is_embedded(self) -> None:
        out = renderer.embed_images('<img alt="x" src="../assets/dot.png">', self.root / "docs", self.root)
        self.assertIn('src="data:image/png;base64,' + base64.b64encode(PNG).decode("ascii") + '"', out)

    def test_missing_image_fails(self) -> None:
        with self.assertRaisesRegex(renderer.RenderError, "not found"):
            renderer.embed_images('<img src="../assets/none.png">', self.root / "docs", self.root)

    def test_image_outside_the_repository_fails(self) -> None:
        with self.assertRaisesRegex(renderer.RenderError, "outside the repository"):
            renderer.embed_images('<img src="../../secret.png">', self.root / "docs", self.root)

    def test_remote_image_fails(self) -> None:
        with self.assertRaisesRegex(renderer.RenderError, "Only repository images"):
            renderer.embed_images('<img src="https://example.com/x.png">', self.root / "docs", self.root)


class LinkTests(unittest.TestCase):
    SOURCE = PurePosixPath("docs/USER_MANUAL.md")

    def test_repository_link_is_pinned_to_the_revision(self) -> None:
        out = renderer.rewrite_links('<a href="../INSTALLATION.md#download-and-verify">x</a>', self.SOURCE, "abc123")
        self.assertIn(f'href="{renderer.REPOSITORY_URL}/blob/abc123/INSTALLATION.md#download-and-verify"', out)

    def test_in_page_and_absolute_links_are_kept(self) -> None:
        text = '<a href="#settings">a</a> <a href="https://github.com/x">b</a> <a href="mailto:a@b.c">c</a>'
        self.assertEqual(renderer.rewrite_links(text, self.SOURCE, "abc123"), text)

    def test_link_above_the_repository_fails(self) -> None:
        with self.assertRaisesRegex(renderer.RenderError, "outside the repository"):
            renderer.rewrite_links('<a href="../../x.md">x</a>', self.SOURCE, "abc123")


class BrowserTests(unittest.TestCase):
    def test_missing_explicit_browser_fails(self) -> None:
        with self.assertRaisesRegex(renderer.RenderError, "Browser not found"):
            renderer.find_browser("/no/such/browser")


@unittest.skipUnless(importlib.util.find_spec("markdown"), "Python-Markdown is not installed")
class ManualTests(unittest.TestCase):
    def test_the_manual_renders_self_contained(self) -> None:
        document = renderer.build_html(renderer.ROOT, "abc123")
        self.assertIn("<h1>", document)
        self.assertNotIn('src="../', document)
        self.assertNotIn("[!NOTE]", document)
        self.assertIn('class="alert alert-note"', document)
        self.assertNotIn('href="../', document)


if __name__ == "__main__":
    unittest.main()
