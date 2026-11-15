#!/usr/bin/env python3
"""Render the user manual to one self-contained HTML file and, with a browser, to PDF.

docs/USER_MANUAL.md is converted with Python-Markdown (installed from the
hash-locked tools/requirements-docs.txt). The HTML is self-contained:
- images are embedded as data URIs;
- GitHub alerts (> [!NOTE] and the like) become styled boxes;
- links to other repository files point to GitHub at one fixed revision (the
  checked-out commit by default), so the PDF never links to a moving branch;
- the manual's Contents list becomes a table of contents of its sections.

The PDF is printed by headless Chrome, Edge or Chromium, with page numbers and
a document outline (bookmarks). It is printed twice: the first print gives each
section's page, read from the outline with pypdf; the second fills those pages
into the table of contents and must keep the same pages. A release publishes it
as DATETIMEPICKER-user-manual-vX.Y.Z.pdf (RELEASING.md section 7).

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
import tempfile
from pathlib import Path, PurePosixPath

ROOT = Path(__file__).resolve().parent.parent
SOURCE = PurePosixPath("docs/USER_MANUAL.md")
REPOSITORY_URL = "https://github.com/danielep71/VBA-DATETIMEPICKER"
TITLE = "Date / Time Picker - User Manual"
HTML_BLOCK = re.compile(r"^<div\b(?![^>]*\bmarkdown=)([^>]*)>\s*$")
ALERT = re.compile(r"^>\s*\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*$")
IMG_SRC = re.compile(r'(<img\b[^>]*?\bsrc=")([^"]+)(")', re.IGNORECASE)
HREF = re.compile(r'(<a\b[^>]*?\bhref=")([^"]+)(")', re.IGNORECASE)
HEADING = re.compile(r'<h([1-6])(?: id="([^"]*)")?>(.*?)</h\1>', re.DOTALL)
CONTENTS_LIST = re.compile(r'(<h2 id="toc-contents">.*?</h2>\s*)<ol>.*?</ol>', re.DOTALL)
# Heading ids get a prefix so they never collide with the manual's own <a id> anchors.
HEADING_ID_PREFIX = "toc-"
TAG = re.compile(r"<[^>]+>")
BROWSER_NAMES = ("chrome", "google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "msedge")
WINDOWS_BROWSERS = (
    r"Google\Chrome\Application\chrome.exe",
    r"Microsoft\Edge\Application\msedge.exe",
)
STYLE = """
@page { size: A4; margin: 18mm 16mm 20mm 16mm;
  @top-right { content: "Date / Time Picker - User Manual"; font: 8pt "Segoe UI", Arial, sans-serif;
               color: #59636e; }
  @bottom-center { content: "Page " counter(page) " of " counter(pages); font: 9pt "Segoe UI", Arial, sans-serif;
                   color: #59636e; } }
@page :first { @top-right { content: none; } }
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
.toc ol { list-style: none; padding: 0; margin: 0.4em 0 1em; }
.toc li { margin: 0.15em 0; }
.toc a { display: flex; align-items: baseline; color: inherit; }
.toc .toc-dots { flex: 1; border-bottom: 1px dotted #8c959f; margin: 0 0.4em; }
.toc .toc-page { min-width: 2em; text-align: right; }
.toc .toc-2 { font-weight: 600; margin-top: 0.45em; }
.toc .toc-3 { padding-left: 1.5em; font-size: 0.95em; }
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
    toc = importlib.import_module("markdown.extensions.toc")

    def slugify(value: str, separator: str) -> str:
        return HEADING_ID_PREFIX + str(toc.slugify(value, separator))

    return str(markdown.markdown(text, extensions=["tables", "fenced_code", "md_in_html", "sane_lists", "toc"],
                                 extension_configs={"toc": {"slugify": slugify}}, output_format="html"))


def heading_text(fragment: str) -> str:
    """Plain text of a heading's inner HTML, with whitespace collapsed."""
    return " ".join(html.unescape(TAG.sub("", fragment)).split())


def headings(document: str) -> list[tuple[int, str, str]]:
    """(level, id, text) of every heading, in document order."""
    return [(int(level), ident, heading_text(inner)) for level, ident, inner in HEADING.findall(document)]


def insert_contents(document: str, pages: dict[str, int] | None) -> str:
    """Replace the Contents list with a table of contents of the h2 and h3 sections after it.

    pages maps a heading id to its page; without it the page column is left blank.
    """
    match = CONTENTS_LIST.search(document)
    if match is None:
        raise RenderError('The manual needs a "Contents" heading followed by a list')
    entries = []
    for level, ident, text in headings(document[match.end():]):
        if level not in (2, 3) or not ident:
            continue
        page = str(pages[ident]) if pages is not None and ident in pages else ""
        entries.append(f'<li class="toc-{level}"><a href="#{html.escape(ident, quote=True)}">'
                       f'<span class="toc-title">{html.escape(text)}</span><span class="toc-dots"></span>'
                       f'<span class="toc-page">{page}</span></a></li>')
    contents = '<nav class="toc"><ol>' + "".join(entries) + "</ol></nav>"
    return document[:match.start()] + match.group(1) + contents + document[match.end():]


def assign_pages(document: str, outline: list[tuple[str, int]]) -> dict[str, int]:
    """Map each heading id to its page, pairing the headings with the PDF outline in order."""
    found = headings(document)
    if len(found) != len(outline):
        raise RenderError(f"The PDF outline has {len(outline)} entries for {len(found)} headings")
    pages: dict[str, int] = {}
    for (_level, ident, text), (title, page) in zip(found, outline, strict=True):
        if " ".join(title.split()) != text:
            raise RenderError(f"The PDF outline entry {title!r} does not match the heading {text!r}")
        if ident:
            pages[ident] = page
    return pages


def build_html(root: Path, ref: str, pages: dict[str, int] | None = None) -> str:
    source_file = root / Path(*SOURCE.parts)
    text = source_file.read_text(encoding="utf-8")
    body = insert_contents(markdown_to_html(convert_alerts(text)), pages)
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
    """Print html_file to pdf_file; the destination changes only when a new PDF was written."""
    destination = pdf_file.resolve()
    handle, name = tempfile.mkstemp(prefix=destination.stem + ".", suffix=".pdf", dir=destination.parent)
    os.close(handle)
    fresh = Path(name)
    fresh.unlink()
    arguments = [browser, "--headless", "--disable-gpu", "--no-pdf-header-footer", "--print-to-pdf-no-header",
                 "--generate-pdf-document-outline", f"--print-to-pdf={fresh}", html_file.resolve().as_uri()]
    # Headless Chrome uses a temporary profile of its own; a fresh --user-data-dir
    # makes it wait for first-run network work instead of printing.
    if os.name == "posix" and hasattr(os, "geteuid") and os.geteuid() == 0:
        arguments.insert(1, "--no-sandbox")
    try:
        result = subprocess.run(arguments, capture_output=True, check=False, timeout=180)
        if result.returncode != 0:
            raise RenderError(f"The browser exited with code {result.returncode}")
        if not fresh.is_file() or fresh.stat().st_size < 1024 or fresh.read_bytes()[:5] != b"%PDF-":
            raise RenderError(f"The browser did not write a PDF for {pdf_file}")
        os.replace(fresh, destination)
    finally:
        if fresh.exists():
            fresh.unlink()


def outline_pages(pdf_file: Path) -> list[tuple[str, int]]:
    """(title, 1-based page) of every PDF outline entry, depth first, which is document order."""
    try:
        pypdf = importlib.import_module("pypdf")
    except ImportError as error:
        raise RenderError("pypdf is not installed. Run: python -m pip install --require-hashes "
                          "--only-binary=:all: -r tools/requirements-docs.txt") from error
    reader = pypdf.PdfReader(str(pdf_file))
    entries: list[tuple[str, int]] = []

    def walk(items: list[object]) -> None:
        for item in items:
            if isinstance(item, list):
                walk(item)
            else:
                entries.append((str(getattr(item, "title", "")), int(reader.get_destination_page_number(item)) + 1))
    walk(list(reader.outline))
    if not entries:
        raise RenderError("The PDF has no outline; the browser does not support --generate-pdf-document-outline")
    return entries


def render_pdf(browser: str, root: Path, ref: str, html_file: Path, pdf_file: Path) -> None:
    """Print twice: the first print gives the section pages, the second puts them in the contents."""
    first = build_html(root, ref)
    html_file.write_text(first, encoding="utf-8")
    print_pdf(browser, html_file, pdf_file)
    pages = assign_pages(first, outline_pages(pdf_file))
    final = build_html(root, ref, pages)
    html_file.write_text(final, encoding="utf-8")
    print_pdf(browser, html_file, pdf_file)
    if assign_pages(final, outline_pages(pdf_file)) != pages:
        raise RenderError("Filling in the page numbers moved a section to another page")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    parser.add_argument("--html", type=Path, required=True, help="self-contained HTML to write")
    parser.add_argument("--pdf", type=Path, help="PDF to print from the HTML with a headless browser")
    parser.add_argument("--ref", help="revision that repository links point to (default: the checked-out commit)")
    parser.add_argument("--browser", help="path to Chrome, Edge or Chromium (default: found automatically)")
    options = parser.parse_args(argv)
    try:
        ref = options.ref or checked_out_revision(ROOT)
        if options.pdf is None:
            options.html.write_text(build_html(ROOT, ref), encoding="utf-8")
        else:
            render_pdf(find_browser(options.browser), ROOT, ref, options.html, options.pdf)
        print(f"HTML: {options.html}")
        if options.pdf is not None:
            digest = hashlib.sha256(options.pdf.read_bytes()).hexdigest()
            print(f"PDF: {options.pdf} ({options.pdf.stat().st_size} bytes, sha256 {digest})")
    except (RenderError, OSError, subprocess.SubprocessError) as error:
        print(f"FAIL - {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
