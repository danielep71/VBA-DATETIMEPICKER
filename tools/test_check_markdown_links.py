"""Offline fixtures for the Markdown link and anchor checker."""
import contextlib
import email.message
import tempfile
import unittest
import urllib.error
import urllib.request
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

import check_markdown_links as links


class SlugTests(unittest.TestCase):
    def test_github_slug_rule(self) -> None:
        cases = {
            "Quick start": "quick-start",
            "🚀 Quick start": "-quick-start",
            "⚠️ Known limitations": "️-known-limitations",
            "DP_Start and DP_Stop": "dp_start-and-dp_stop",
            "Contracts added in v1.2.2": "contracts-added-in-v122",
            "Café — naïve": "café--naïve",
            "A/B (test) & more?": "ab-test--more",
        }
        for heading, slug in cases.items():
            with self.subTest(heading=heading):
                self.assertEqual(links.github_slug(links.heading_text(heading)), slug)

    def test_heading_markup_is_reduced_to_text(self) -> None:
        self.assertEqual(links.heading_text("`DP_Show` and **bold** [link](x.md) <sup>1</sup>"),
                         "DP_Show and bold link 1")
        self.assertEqual(links.heading_text("_emphasis_ keeps DP_WriteResult"), "emphasis keeps DP_WriteResult")

    def test_duplicates_html_anchors_setext_and_fences(self) -> None:
        text = ("# Setup\n\n## Setup\n\n### Setup\n\n<a id=\"custom-anchor\"></a>\n"
                "Underlined\n==========\n\n```md\n# Not a heading\n```\n\n~~~\n## Also not\n~~~\n")
        found = links.anchors(text)
        self.assertTrue({"setup", "setup-1", "setup-2", "custom-anchor", "underlined"} <= found)
        self.assertNotIn("not-a-heading", found)
        self.assertNotIn("also-not", found)


class OfflineCheckTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def check(self, files: dict[str, str], extra: tuple[str, ...] = ()) -> list[str]:
        for name, text in files.items():
            (self.root / name).parent.mkdir(parents=True, exist_ok=True)
            (self.root / name).write_text(text, encoding="utf-8")
        tracked = set(files) | set(extra)
        report = links.check_offline(self.root, [f for f in files if f.endswith(".md")], tracked)
        return [f"{f['line']}:{f['target']}" for f in report["findings"]]

    def test_valid_links_pass(self) -> None:
        files = {
            "README.md": ("# Title\n\n## ⚠️ Limits\n\n[self](#title) [emoji](#️-limits) "
                          "[doc](docs/GUIDE.md#usage) [dir](docs/) [root](/LICENSE) "
                          "[ext](https://example.com/x) [mail](mailto:a@example.com) "
                          "<a href=\"docs/GUIDE.md\">html</a>\n\n[ref]: docs/GUIDE.md#usage\n"),
            "docs/GUIDE.md": "## Usage\n\n[up](../README.md#title) [code](../src/a.bas#L3)\n",
            "LICENSE": "x", "src/a.bas": "x",
        }
        self.assertEqual(self.check(files), [])

    def test_missing_targets_and_anchors_fail(self) -> None:
        files = {
            "README.md": ("# Title\n\n[gone](missing.md) [anchor](#nope) [deep](docs/GUIDE.md#absent)\n"
                          "[ref]: docs/OLD.md\n"),
            "docs/GUIDE.md": "## Usage\n",
        }
        self.assertEqual(self.check(files), ["3:missing.md", "3:#nope", "3:docs/GUIDE.md#absent", "4:docs/OLD.md"])

    def test_untracked_file_is_not_a_valid_target(self) -> None:
        (self.root / "generated.md").write_text("# Generated\n", encoding="utf-8")
        self.assertEqual(self.check({"README.md": "[gen](generated.md)\n"}), ["1:generated.md"])

    def test_linked_images_and_parentheses_in_destinations(self) -> None:
        files = {
            "README.md": ("[![badge](img/missing.svg)](docs/GUIDE.md) [![ok](img/ok.svg)](docs/GUIDE.md)\n"
                          "[doc](guide(v2).md) [angle](<docs/GUIDE.md>) ![pic](img/ok.svg \"title\")\n"),
            "docs/GUIDE.md": "# Guide\n", "guide(v2).md": "# v2\n", "img/ok.svg": "<svg/>",
        }
        self.assertEqual(self.check(files), ["1:img/missing.svg"])

    def test_indented_code_outside_lists_is_ignored(self) -> None:
        text = ("Paragraph.\n\n    [code](missing-a.md)\n\n    still code [x](missing-b.md)\n\n"
                "- item\n\n    continuation [y](missing-c.md)\n\nBack to prose.\n\n    [z](missing-d.md)\n")
        self.assertEqual(self.check({"README.md": text}), ["9:missing-c.md"])

    def test_links_in_code_and_comments_are_ignored(self) -> None:
        text = "`[a](missing.md)`\n\n```\n[b](missing.md)\n```\n\n<!--\n[c](missing.md)\n-->\n"
        self.assertEqual(self.check({"README.md": text}), [])


class ExternalProbeTests(unittest.TestCase):
    def probe_with(self, outcomes: dict[str, object]) -> tuple[str, str]:
        """Run probe() against a fake urlopen keyed by HTTP method."""
        def fake_urlopen(request: urllib.request.Request, timeout: float) -> object:
            outcome = outcomes[request.get_method()]
            if isinstance(outcome, BaseException):
                raise outcome
            return contextlib.nullcontext(SimpleNamespace(status=outcome))

        with mock.patch.object(urllib.request, "urlopen", fake_urlopen):
            return links.probe("https://example.net/page")

    @staticmethod
    def http_error(code: int) -> urllib.error.HTTPError:
        return urllib.error.HTTPError("https://example.net/page", code, "x", email.message.Message(), None)

    def test_classification(self) -> None:
        self.assertEqual(self.probe_with({"HEAD": 200}), ("ok", "HTTP 200"))
        self.assertEqual(self.probe_with({"HEAD": self.http_error(405), "GET": 200}), ("ok", "HTTP 200"))
        self.assertEqual(self.probe_with({"HEAD": self.http_error(404), "GET": self.http_error(404)}),
                         ("broken", "HTTP 404"))
        self.assertEqual(self.probe_with({"HEAD": self.http_error(403), "GET": self.http_error(403)}),
                         ("warning", "HTTP 403"))
        self.assertEqual(self.probe_with({"HEAD": self.http_error(429), "GET": self.http_error(429)}),
                         ("warning", "HTTP 429"))
        dns = urllib.error.URLError("[Errno -2] Name or service not known")
        self.assertEqual(self.probe_with({"HEAD": dns, "GET": dns})[0], "broken")
        self.assertEqual(self.probe_with({"HEAD": TimeoutError(), "GET": TimeoutError()})[0], "warning")


if __name__ == "__main__":
    unittest.main()
