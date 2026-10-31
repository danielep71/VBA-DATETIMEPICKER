#!/usr/bin/env python3
"""Check links in tracked Markdown: relative targets and anchors offline, external URLs on request."""
from __future__ import annotations

import argparse
import re
import sys
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from collections.abc import Iterable
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Any

from _gatelib import run_gate, tracked_files

FENCE = re.compile(r"^ {0,3}(`{3,}|~{3,})")
ATX = re.compile(r"^ {0,3}(#{1,6})[ \t]+(.*?)(?:[ \t]+#+)?[ \t]*$")
SETEXT = re.compile(r"^ {0,3}(=+|-+)[ \t]*$")
HTML_ANCHOR = re.compile(r"<[A-Za-z][^>]*?\s(?:id|name)\s*=\s*[\"']([^\"']+)[\"']", re.I)
INLINE_LINK = re.compile(r"!?\[(?:[^\[\]]|\[[^\[\]]*\])*\]\(\s*<?([^\s)>]*)>?(?:\s+(?:\"[^\"]*\"|'[^']*'))?\s*\)")
REFERENCE_DEF = re.compile(r"^ {0,3}\[[^\]]+\]:\s*<?(\S+?)>?(?:\s+.*)?$")
HTML_LINK = re.compile(r"<(?:a|img)\b[^>]*?\s(?:href|src)\s*=\s*[\"']([^\"']+)[\"']", re.I)
AUTOLINK = re.compile(r"<(https?://[^>\s]+)>")
CODE_SPAN = re.compile(r"(`+)(.+?)\1")
COMMENT = re.compile(r"<!--.*?-->", re.S)
SCHEME = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*:")
# Hosts that are placeholders, not links to verify.
EXTERNAL_SKIP = ("localhost", "127.0.0.1", "example.com", "example.org", "example.invalid")


@dataclass(frozen=True)
class Link:
    path: str
    line: int
    target: str


def github_slug(text: str) -> str:
    """GitHub's heading slug: lower-case, keep letters, marks, decimal digits,
    connector punctuation, spaces and hyphens, then turn spaces into hyphens.

    Marks are kept, so U+FE0F after an emoji survives while the emoji does not.
    """
    kept = "".join(ch for ch in text.lower()
                   if ch in " -" or unicodedata.category(ch)[0] in "LM"
                   or unicodedata.category(ch) in ("Nd", "Pc"))
    return kept.replace(" ", "-")


def heading_text(raw: str) -> str:
    """Reduce inline Markdown in a heading to the text GitHub renders."""
    text = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", raw)
    text = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", text)
    text = re.sub(r"\[([^\]]*)\]\[[^\]]*\]", r"\1", text)
    text = CODE_SPAN.sub(lambda m: m.group(2).strip(), text)
    text = re.sub(r"<[^>]+>", "", text)
    text = text.replace("*", "")
    text = re.sub(r"(?<!\w)_+|_+(?!\w)", "", text)
    text = text.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    return text.strip()


def prose_lines(text: str) -> list[tuple[int, str]]:
    """Lines outside fenced code blocks and HTML comments, with 1-based numbers."""
    text = COMMENT.sub(lambda m: "\n" * m.group(0).count("\n"), text)
    result: list[tuple[int, str]] = []
    fence = ""
    for number, line in enumerate(text.splitlines(), start=1):
        match = FENCE.match(line)
        if fence:
            if match and match.group(1)[0] == fence[0] and len(match.group(1)) >= len(fence):
                fence = ""
            continue
        if match:
            fence = match.group(1)
            continue
        result.append((number, line))
    return result


def anchors(text: str) -> set[str]:
    """Every anchor a document defines: heading slugs (with GitHub's -1, -2
    suffixes for duplicates) and explicit HTML id/name attributes."""
    found: set[str] = set()
    counts: dict[str, int] = {}
    lines = prose_lines(text)
    for index, (number, line) in enumerate(lines):
        found.update(HTML_ANCHOR.findall(line))
        match = ATX.match(line)
        heading = match.group(2) if match else None
        if heading is None and line.strip() and not line.startswith(("    ", "\t")) and index + 1 < len(lines):
            # A setext heading: a text line directly underlined with = or -.
            next_number, next_line = lines[index + 1]
            if next_number == number + 1 and SETEXT.match(next_line) and not SETEXT.match(line):
                heading = line.strip()
        if heading is None:
            continue
        slug = github_slug(heading_text(heading))
        if slug in counts:
            counts[slug] += 1
            found.add(f"{slug}-{counts[slug]}")
        else:
            counts[slug] = 0
            found.add(slug)
    return found


def links(path: str, text: str) -> list[Link]:
    result: list[Link] = []
    for number, line in prose_lines(text):
        line = CODE_SPAN.sub("", line)
        targets = [m.group(1) for m in INLINE_LINK.finditer(line)]
        targets += HTML_LINK.findall(line) + AUTOLINK.findall(line)
        definition = REFERENCE_DEF.match(line)
        if definition:
            targets.append(definition.group(1))
        result.extend(Link(path, number, target) for target in targets if target)
    return result


def is_external(target: str) -> bool:
    return target.startswith(("http://", "https://"))


def check_offline(root: Path, files: Iterable[str], tracked: set[str]) -> dict[str, Any]:
    """Resolve every relative link and anchor in ``files`` against the tracked tree."""
    directories = {str(parent) for item in tracked for parent in PurePosixPath(item).parents}
    texts = {path: (root / path).read_text(encoding="utf-8") for path in files}
    anchor_cache: dict[str, set[str]] = {}
    findings: list[dict[str, Any]] = []
    checked = 0

    def anchors_of(path: str) -> set[str]:
        if path not in anchor_cache:
            text = texts.get(path)
            anchor_cache[path] = anchors(text if text is not None else (root / path).read_text(encoding="utf-8"))
        return anchor_cache[path]

    for path in sorted(texts):
        for link in links(path, texts[path]):
            target = link.target
            if SCHEME.match(target) or target.startswith("//"):
                continue
            checked += 1
            location, _, fragment = target.partition("#")
            location = urllib.parse.unquote(location.split("?", 1)[0])
            fragment = urllib.parse.unquote(fragment)
            if location:
                base = PurePosixPath("." if location.startswith("/") else path).parent
                resolved = PurePosixPath(*_normalize((base / location.lstrip("/")).parts))
                resolved_text = str(resolved)
                if resolved_text not in tracked and resolved_text not in directories:
                    findings.append({"path": path, "line": link.line, "target": target,
                                     "problem": f"target {resolved_text} is not a tracked file or directory"})
                    continue
            else:
                resolved_text = path
            if fragment and resolved_text.lower().endswith(".md") and resolved_text in tracked:
                if fragment not in anchors_of(resolved_text):
                    findings.append({"path": path, "line": link.line, "target": target,
                                     "problem": f"anchor #{fragment} is not defined in {resolved_text}"})
    return {"schema_version": 1, "mode": "offline", "status": "fail" if findings else "pass",
            "files": len(texts), "links_checked": checked, "findings": findings}


def _normalize(parts: tuple[str, ...]) -> list[str]:
    stack: list[str] = []
    for part in parts:
        if part in ("", "."):
            continue
        if part == "..":
            if stack:
                stack.pop()
            continue
        stack.append(part)
    return stack


def probe(url: str, timeout: float = 20.0) -> tuple[str, str]:
    """Classify one URL as ok, broken or warning, with a short reason."""
    headers = {"User-Agent": "DateTimePicker-link-check (+https://github.com/danielep71/VBA-DATETIMEPICKER)"}
    last = ""
    for method in ("HEAD", "GET"):
        request = urllib.request.Request(url, method=method, headers=headers)
        try:
            # Callers pass only http(s) URLs (see check_external).
            with urllib.request.urlopen(request, timeout=timeout) as response:
                return "ok", f"HTTP {response.status}"
        except urllib.error.HTTPError as error:
            last = f"HTTP {error.code}"
            if error.code in (404, 410):
                if method == "GET":
                    return "broken", last
                continue
            if error.code in (405, 501) and method == "HEAD":
                continue
            if method == "HEAD":
                continue
            return "warning", last
        except urllib.error.URLError as error:
            reason = str(error.reason)
            if "Name or service not known" in reason or "nodename nor servname" in reason \
                    or "No address associated" in reason or "getaddrinfo failed" in reason:
                return "broken", f"DNS failure: {reason}"
            last = reason
        except (TimeoutError, OSError) as error:
            last = str(error) or type(error).__name__
    return "warning", last or "no response"


def check_external(root: Path, files: Iterable[str], workers: int = 8) -> dict[str, Any]:
    occurrences: dict[str, list[Link]] = {}
    for path in sorted(files):
        for link in links(path, (root / path).read_text(encoding="utf-8")):
            url = link.target.split("#", 1)[0]
            host = urllib.parse.urlsplit(url).hostname or ""
            if is_external(url) and not host.endswith(EXTERNAL_SKIP):
                occurrences.setdefault(url, []).append(link)
    with ThreadPoolExecutor(max_workers=workers) as pool:
        results = dict(zip(occurrences, pool.map(probe, occurrences), strict=True))
    broken = [{"url": url, "reason": reason, "where": [f"{o.path}:{o.line}" for o in occurrences[url]]}
              for url, (state, reason) in sorted(results.items()) if state == "broken"]
    warnings = [{"url": url, "reason": reason, "where": [f"{o.path}:{o.line}" for o in occurrences[url]]}
                for url, (state, reason) in sorted(results.items()) if state == "warning"]
    return {"schema_version": 1, "mode": "external", "status": "fail" if broken else "pass",
            "urls_checked": len(results), "broken": broken, "warnings": warnings}


def markdown(report: dict[str, Any]) -> str:
    if report["mode"] == "offline":
        lines = [f"Markdown links (offline): {report['status'].upper()} — {report['links_checked']} "
                 f"relative links in {report['files']} files"]
        lines += [f"- {f['path']}:{f['line']}: `{f['target']}` — {f['problem']}" for f in report["findings"]]
    else:
        lines = [f"Markdown links (external): {report['status'].upper()} — {report['urls_checked']} URLs"]
        for title, key in (("Broken", "broken"), ("Warnings (blocked, rate-limited or unreachable)", "warnings")):
            if report[key]:
                lines += ["", f"{title}:"]
                lines += [f"- {item['url']} — {item['reason']} ({', '.join(item['where'])})" for item in report[key]]
    return "\n".join(lines) + "\n"


def markdown_files(root: Path) -> tuple[list[str], set[str]]:
    tracked = tracked_files(root)
    return sorted(path for path in tracked if path.lower().endswith(".md")), tracked


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    parser.add_argument("--external", action="store_true",
                        help="Check external http(s) URLs over the network instead of relative links")
    options = parser.parse_args(sys.argv[1:])
    options.self_test = False
    root = options.root.resolve()

    def build() -> dict[str, Any]:
        files, tracked = markdown_files(root)
        return check_external(root, files) if options.external else check_offline(root, files, tracked)

    return run_gate(options, build=build, markdown=markdown, errors=(OSError, RuntimeError, UnicodeDecodeError))


if __name__ == "__main__":
    raise SystemExit(main())
