#!/usr/bin/env python3
"""Render the user manual to one self-contained HTML file and, with a browser, to PDF.

docs/USER_MANUAL.md is converted with Python-Markdown (installed from the
hash-locked tools/requirements-docs.txt). The HTML is self-contained:
- images are embedded as data URIs;
- GitHub alerts (> [!NOTE] and the like) become styled boxes;
- links to other repository files point to GitHub at one fixed revision (the
  checked-out commit by default), so the PDF never links to a moving branch.

The PDF is printed by headless Chrome, Edge or Chromium. A release publishes
it as DATETIMEPICKER-user-manual-vX.Y.Z.pdf (RELEASING.md section 7).

Usage:
    python tools/render_manual.py --html <out.html> [--pdf <out.pdf>] [--ref <sha-or-tag>] [--browser <path>]
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import html
import importlib
import mimetypes
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path, PurePosixPath

ROOT = Path(__file__).resolve().parent.parent
SOURCE = PurePosixPath("docs/USER_MANUAL.md")
REPOSITORY_URL = "https://github.com/danielep71/VBA-DATETIMEPICKER"
TITLE = "Date / Time Picker - User Manual"
HTML_BLOCK = re.compile(r"^<div\b(?![^>]*\bmarkdown=)([^>]*)>\s*$")
ALERT = re.compile(r"^>\s*\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*$")
IMG_SRC = re.compile(r'(<img\b[^>]*?\bsrc=")([^"]+)(")', re.IGNORECASE)
HREF = re.compile(r'(<a\b[^>]*?\bhref=")([^"]+)(")', re.IGNORECASE)
BROWSER_NAMES = ("chrome", "google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "msedge")
WINDOWS_BROWSERS = (
    r"Google\Chrome\Application\chrome.exe",
    r"Microsoft\Edge\Application\msedge.exe",
)
STYLE = """
@page { size: A4; margin: 18mm 16mm 18mm 16mm; }
body { font-family: "Segoe UI", "Helvetica Neue", Arial, "Noto Sans", sans-serif;
       font-size: 10.5pt; line-height: 1.5; color: #1f2328; max-width: 180mm; margin: 0 auto; }
h1, h2, h3 { line-height: 1.25; margin: 1.4em 0 0.5em; }
h1 { font-size: 22pt; } h2 { font-size: 15pt; border-bottom: 1px solid #d1d9e0; padding-bottom: 0.2em; }
h2 { break-before: auto; } h3 { font-size: 12pt; }
h1, h2, h3, table, pre, img, .alert { break-inside: avoid; }
h2, h3 { break-after: avoid; }
a { color: #0969da; text-decoration: none; }
code { font-family: Consolas, "Courier New", monospace; font-size: 9.5pt;
       background: #f6f8fa; padding: 0.1em 0.3em; border-radius: 4px; }
pre { background: #f6f8fa; padding: 0.7em 0.9em; border-radius: 6px; white-space: pre-wrap; }
pre code { background: none; padding: 0; }
table { border-collapse: collapse; width: 100%; margin: 0.6em 0 1em; font-size: 9.5pt; }
th, td { border: 1px solid #d1d9e0; padding: 0.35em 0.6em; text-align: left; vertical-align: top; }
th { background: #f6f8fa; }
img { max-width: 100%; height: auto; }
hr { border: 0; border-top: 1px solid #d1d9e0; margin: 1.5em 0; }
blockquote { margin: 0.8em 0; padding: 0 1em; color: #59636e; border-left: 4px solid #d1d9e0; }
.alert { margin: 0.8em 0; padding: 0.5em 1em; border-left: 4px solid; border-radius: 4px; }
.alert > p:first-child { font-weight: 600; margin-top: 0.2em; }
.alert-note { border-color: #0969da; background: #ddf4ff; }
.alert-tip { border-color: #1a7f37; background: #dafbe1; }
.alert-important { border-color: #8250df; background: #fbefff; }
.alert-warning { border-color: #9a6700; background: #fff8c5; }
.alert-caution { border-color: #cf222e; background: #ffebe9; }
.revision { color: #59636e; font-size: 9pt; text-align: center; }
"""


class RenderError(Exception):
    """The manual cannot be rendered; the message says why."""


def convert_alerts(text: str) -> str:
    """Turn GitHub alert blockquotes into div blocks that Python-Markdown renders inside.

    A line that opens a div, such as the centred title block, gets markdown="1" too:
    GitHub renders Markdown inside such a block, and Python-Markdown does so only when asked.
    """
    lines = [HTML_BLOCK.sub(r'<div\1 markdown="1">', line) for line in text.splitlines()]
    out: list[str] = []
    index = 0
    while index < len(lines):
        match = ALERT.match(lines[index])
        if not match:
            out.append(lines[index])
            index += 1
            continue
        kind = match.group(1)
        index += 1
        body: list[str] = []
        while index < len(lines) and lines[index].startswith(">"):
            body.append(re.sub(r"^>\s?", "", lines[index]))
            index += 1
        out += [f'<div class="alert alert-{kind.lower()}" markdown="1">', "",
                f"**{kind.capitalize()}**", "", *body, "", "</div>"]
    return "\n".join(out) + "\n"


def embed_images(document: str, source_dir: Path, root: Path) -> str:
    """Replace each relative image source with a data URI of the file inside root."""
    def replace(match: re.Match[str]) -> str:
        source = html.unescape(match.group(2))
        if re.match(r"^[a-z][a-z0-9+.-]*:", source, re.IGNORECASE):
            raise RenderError(f"Only repository images can be embedded, not {source}")
        target = (source_dir / source).resolve()
        if root.resolve() not in target.parents:
            raise RenderError(f"Image outside the repository: {source}")
        if not target.is_file():
            raise RenderError(f"Image not found: {source}")
        mime = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        if not mime.startswith("image/"):
            raise RenderError(f"Not an image: {source}")
        data = base64.b64encode(target.read_bytes()).decode("ascii")
        return f"{match.group(1)}data:{mime};base64,{data}{match.group(3)}"
    return IMG_SRC.sub(replace, document)


def rewrite_links(document: str, source: PurePosixPath, ref: str) -> str:
    """Point relative links to repository files at GitHub at ref; keep in-page and absolute links."""
    def replace(match: re.Match[str]) -> str:
        href = html.unescape(match.group(2))
        if href.startswith("#") or re.match(r"^[a-z][a-z0-9+.-]*:", href, re.IGNORECASE):
            return match.group(0)
        path, _, fragment = href.partition("#")
        parts: list[str] = []
        for part in (source.parent / path).parts:
            if part == "..":
                if not parts:
                    raise RenderError(f"Link outside the repository: {href}")
                parts.pop()
            elif part != ".":
                parts.append(part)
        url = f"{REPOSITORY_URL}/blob/{ref}/{'/'.join(parts)}" + (f"#{fragment}" if fragment else "")
        return f"{match.group(1)}{html.escape(url, quote=True)}{match.group(3)}"
    return HREF.sub(replace, document)


def markdown_to_html(text: str) -> str:
    try:
        markdown = importlib.import_module("markdown")
    except ImportError as error:
        raise RenderError("Python-Markdown is not installed. Run: python -m pip install --require-hashes "
                          "--only-binary=:all: -r tools/requirements-docs.txt") from error
    return str(markdown.markdown(text, extensions=["tables", "fenced_code", "md_in_html", "sane_lists"],
                                 output_format="html"))


def build_html(root: Path, ref: str) -> str:
    source_file = root / Path(*SOURCE.parts)
    text = source_file.read_text(encoding="utf-8")
    body = markdown_to_html(convert_alerts(text))
    body = rewrite_links(embed_images(body, source_file.parent, root), SOURCE, ref)
    revision = f'<p class="revision">Rendered from {SOURCE} at {html.escape(ref)}</p>'
    return ("<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n"
            f"<title>{html.escape(TITLE)}</title>\n<style>{STYLE}</style>\n</head>\n<body>\n"
            f"{body}\n{revision}\n</body>\n</html>\n")


def checked_out_revision(root: Path) -> str:
    try:
        result = subprocess.run(["git", "-C", str(root), "rev-parse", "HEAD"],
                                capture_output=True, text=True, check=True)
    except (OSError, subprocess.CalledProcessError) as error:
        raise RenderError("Cannot read the checked-out commit; pass --ref") from error
    return result.stdout.strip()


def find_browser(explicit: str | None) -> str:
    if explicit:
        if Path(explicit).is_file() or shutil.which(explicit):
            return explicit
        raise RenderError(f"Browser not found: {explicit}")
    for name in BROWSER_NAMES:
        found = shutil.which(name)
        if found:
            return found
    for variable in ("ProgramFiles", "ProgramFiles(x86)", "LocalAppData"):
        base = os.environ.get(variable)
        for relative in WINDOWS_BROWSERS:
            if base and (Path(base) / relative).is_file():
                return str(Path(base) / relative)
    raise RenderError("No Chrome, Edge or Chromium found; pass --browser <path>")


def print_pdf(browser: str, html_file: Path, pdf_file: Path) -> None:
    arguments = [browser, "--headless", "--disable-gpu", "--no-pdf-header-footer", "--print-to-pdf-no-header",
                 f"--print-to-pdf={pdf_file.resolve()}", html_file.resolve().as_uri()]
    # Headless Chrome uses a temporary profile of its own; a fresh --user-data-dir
    # makes it wait for first-run network work instead of printing.
    if os.name == "posix" and hasattr(os, "geteuid") and os.geteuid() == 0:
        arguments.insert(1, "--no-sandbox")
    subprocess.run(arguments, capture_output=True, check=False, timeout=180)
    if not pdf_file.is_file() or pdf_file.stat().st_size < 1024 or pdf_file.read_bytes()[:5] != b"%PDF-":
        raise RenderError(f"The browser did not write a PDF to {pdf_file}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    parser.add_argument("--html", type=Path, required=True, help="self-contained HTML to write")
    parser.add_argument("--pdf", type=Path, help="PDF to print from the HTML with a headless browser")
    parser.add_argument("--ref", help="revision that repository links point to (default: the checked-out commit)")
    parser.add_argument("--browser", help="path to Chrome, Edge or Chromium (default: found automatically)")
    options = parser.parse_args(argv)
    try:
        ref = options.ref or checked_out_revision(ROOT)
        options.html.write_text(build_html(ROOT, ref), encoding="utf-8")
        print(f"HTML: {options.html}")
        if options.pdf is not None:
            print_pdf(find_browser(options.browser), options.html, options.pdf)
            digest = hashlib.sha256(options.pdf.read_bytes()).hexdigest()
            print(f"PDF: {options.pdf} ({options.pdf.stat().st_size} bytes, sha256 {digest})")
    except (RenderError, OSError, subprocess.SubprocessError) as error:
        print(f"FAIL - {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
