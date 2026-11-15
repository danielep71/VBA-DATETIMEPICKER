"""Fixtures for render_manual.py: alerts, embedded images, pinned links and the browser lookup."""
import base64
import importlib.util
import re
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path, PurePosixPath
from typing import Any
from unittest import mock

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
    def setUp(self) -> None:
        self.folder = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.folder)
        self.html = self.folder / "manual.html"
        self.html.write_text("<p>x</p>", encoding="utf-8")
        self.pdf = self.folder / "manual.pdf"
        self.pdf.write_bytes(b"%PDF-1.4 stale " + b"x" * 2048)

    @staticmethod
    def browser(writes: bytes | None, code: int = 0) -> Any:
        def run(arguments: list[str], **_: Any) -> subprocess.CompletedProcess[bytes]:
            if writes is not None:
                target = next(a for a in arguments if a.startswith("--print-to-pdf="))
                Path(re.sub("^--print-to-pdf=", "", target)).write_bytes(writes)
            return subprocess.CompletedProcess(arguments, code)
        return run

    def test_missing_explicit_browser_fails(self) -> None:
        with self.assertRaisesRegex(renderer.RenderError, "Browser not found"):
            renderer.find_browser("/no/such/browser")

    def test_new_pdf_replaces_the_destination(self) -> None:
        fresh = b"%PDF-1.7 new " + b"y" * 2048
        with mock.patch("render_manual.subprocess.run", self.browser(fresh)):
            renderer.print_pdf("browser", self.html, self.pdf)
        self.assertEqual(self.pdf.read_bytes(), fresh)
        self.assertEqual(sorted(p.name for p in self.folder.iterdir()), ["manual.html", "manual.pdf"])

    def test_stale_pdf_is_not_reported_as_new(self) -> None:
        stale = self.pdf.read_bytes()
        with mock.patch("render_manual.subprocess.run", self.browser(None)), \
                self.assertRaisesRegex(renderer.RenderError, "did not write a PDF"):
            renderer.print_pdf("browser", self.html, self.pdf)
        self.assertEqual(self.pdf.read_bytes(), stale)

    def test_browser_failure_keeps_the_destination(self) -> None:
        stale = self.pdf.read_bytes()
        with mock.patch("render_manual.subprocess.run", self.browser(b"%PDF-" + b"z" * 2048, code=1)), \
                self.assertRaisesRegex(renderer.RenderError, "exited with code 1"):
            renderer.print_pdf("browser", self.html, self.pdf)
        self.assertEqual(self.pdf.read_bytes(), stale)
        self.assertEqual(sorted(p.name for p in self.folder.iterdir()), ["manual.html", "manual.pdf"])


class ContentsTests(unittest.TestCase):
    DOCUMENT = ('<h1>Manual</h1><h3>Subtitle</h3><h2 id="toc-contents">Contents</h2><ol><li>old</li></ol>'
                '<h2 id="what-it-is">\U0001f4c5 What it is</h2><h3 id="day-view">Day &amp; view</h3>'
                '<h4 id="deep">Deep</h4>')

    def test_contents_list_becomes_a_table_of_contents(self) -> None:
        out = renderer.insert_contents(self.DOCUMENT, {"what-it-is": 2, "day-view": 3})
        self.assertNotIn("<li>old</li>", out)
        self.assertIn('<li class="toc-2"><a href="#what-it-is"><span class="toc-title">\U0001f4c5 What it is</span>'
                      '<span class="toc-dots"></span><span class="toc-page">2</span></a></li>', out)
        self.assertIn('<span class="toc-title">Day &amp; view</span><span class="toc-dots"></span>'
                      '<span class="toc-page">3</span>', out)
        self.assertNotIn("Subtitle</span>", out)
        self.assertNotIn("Deep</span>", out)

    def test_without_pages_the_page_column_is_blank(self) -> None:
        out = renderer.insert_contents(self.DOCUMENT, None)
        self.assertIn('<span class="toc-page"></span>', out)

    def test_a_manual_without_contents_fails(self) -> None:
        with self.assertRaisesRegex(renderer.RenderError, "Contents"):
            renderer.insert_contents("<h1>Manual</h1><h2 id=\"a\">A</h2>", None)

    def test_pages_follow_the_outline_in_order(self) -> None:
        document = '<h1>Manual</h1><h2 id="a">A &amp; B</h2><h3 id="c">C</h3>'
        self.assertEqual(renderer.assign_pages(document, [("Manual", 1), ("A & B", 2), ("C", 4)]), {"a": 2, "c": 4})

    def test_an_outline_that_does_not_match_fails(self) -> None:
        document = '<h1>Manual</h1><h2 id="a">A</h2>'
        with self.assertRaisesRegex(renderer.RenderError, "outline has 1 entries for 2 headings"):
            renderer.assign_pages(document, [("Manual", 1)])
        with self.assertRaisesRegex(renderer.RenderError, "does not match"):
            renderer.assign_pages(document, [("Manual", 1), ("B", 2)])


@unittest.skipUnless(importlib.util.find_spec("markdown"), "Python-Markdown is not installed")
class ManualTests(unittest.TestCase):
    def test_the_manual_renders_self_contained(self) -> None:
        document = renderer.build_html(renderer.ROOT, "abc123")
        self.assertRegex(document, r'<h1 id="toc-[^"]+">')
        self.assertNotIn('src="../', document)
        self.assertNotIn("[!NOTE]", document)
        self.assertIn('class="alert alert-note"', document)
        self.assertNotIn('href="../', document)
        self.assertIn('<nav class="toc">', document)
        self.assertIn('<a href="#toc-settings">', document)
        ids = re.findall(r'\bid="([^"]+)"', document)
        self.assertEqual(len(ids), len(set(ids)), "duplicate ids")


if __name__ == "__main__":
    unittest.main()
