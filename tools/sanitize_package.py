#!/usr/bin/env python3
"""Find and remove build-machine metadata from an Office package (#221).

When Excel saves a workbook or add-in, it can record details of the machine and
account that saved it. None of them belong in a published package:

- ``x15ac:absPath`` in ``xl/workbook.xml``: the folder the file was saved from;
- ``docMetadata/``: a Microsoft Purview sensitivity label with the tenant ID;
- ``customXml/``: document-library content-type metadata;
- ``xl/printerSettings/``: the printer's device settings, including its name;
- ``dc:creator`` and ``cp:lastModifiedBy`` in ``docProps/core.xml``;
- ``Company`` and ``Manager`` in ``docProps/app.xml``.

``--check`` reports what a package carries and exits 1 when it carries any of
it. Otherwise the tool writes a sanitized copy: it drops those parts with their
relationships, content types and ``r:id`` references, and removes those
elements and values. Every other part is copied with identical content, in the
same order. ``xl/vbaProject.bin``, the worksheets' cell data and the Ribbon
parts are never edited. The copy is checked before it replaces anything: it
must be a readable package, carry none of the metadata, keep every untouched
part byte for byte, and have no relationship to a missing part.
"""
from __future__ import annotations

import argparse
import json
import os
import posixpath
import re
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.parsers import expat

REMOVED_PREFIXES = ("docMetadata/", "customXml/", "xl/printerSettings/")
CONTENT_TYPES = "[Content_Types].xml"
WORKBOOK = "xl/workbook.xml"
CORE = "docProps/core.xml"
APP = "docProps/app.xml"
VBA_PROJECT = "xl/vbaProject.bin"

ABS_PATH_BLOCK = re.compile(
    r"<mc:AlternateContent\b[^>]*>(?:(?!</mc:AlternateContent>).)*?<x15ac:absPath\b"
    r"(?:(?!</mc:AlternateContent>).)*?</mc:AlternateContent>", re.S)
ABS_PATH_BARE = re.compile(r"<x15ac:absPath\b[^>]*/>")
CORE_PEOPLE = re.compile(r"<(dc:creator|cp:lastModifiedBy)\b[^>]*>[^<]*</\1>|<(dc:creator|cp:lastModifiedBy)\b[^>]*/>")
APP_PEOPLE = re.compile(r"<(Company|Manager)>([^<]*)</\1>")
RELATIONSHIP = re.compile(r"<Relationship\b[^>]*/>")
OVERRIDE = re.compile(r"<Override\b[^>]*/>")


class SanitizeError(Exception):
    """The package cannot be read or sanitized safely."""


@dataclass
class Result:
    findings: list[str] = field(default_factory=list)
    removed_parts: list[str] = field(default_factory=list)
    edited_parts: list[str] = field(default_factory=list)


def attribute(element: str, name: str) -> str | None:
    match = re.search(r"\b" + re.escape(name) + r'="([^"]*)"', element)
    return match.group(1) if match else None


def is_removed(part: str) -> bool:
    return part.startswith(REMOVED_PREFIXES)


def rels_source(rels_part: str) -> str:
    """The part a ``.rels`` file describes: ``a/_rels/b.xml.rels`` -> ``a/b.xml``; root rels -> ``""``."""
    folder, name = posixpath.split(rels_part)
    base = posixpath.dirname(folder)
    return posixpath.join(base, name[: -len(".rels")]) if name != ".rels" else ""


def resolve(rels_part: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    base = posixpath.dirname(rels_source(rels_part)) if rels_source(rels_part) else ""
    return posixpath.normpath(posixpath.join(base, target))


def read_package(path: Path) -> tuple[list[zipfile.ZipInfo], dict[str, bytes]]:
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            data = {info.filename: archive.read(info) for info in infos}
    except (OSError, zipfile.BadZipFile) as error:
        raise SanitizeError(f"{path}: not a readable Office package ({error})") from error
    names = [info.filename for info in infos]
    if len(names) != len(set(names)):
        raise SanitizeError(f"{path}: duplicate part names")
    if CONTENT_TYPES not in data or WORKBOOK not in data:
        raise SanitizeError(f"{path}: not an Excel package ({CONTENT_TYPES} or {WORKBOOK} missing)")
    return infos, data


def text(data: dict[str, bytes], part: str) -> str:
    return data[part].decode("utf-8")


def inspect_parts(data: dict[str, bytes]) -> list[str]:
    findings: list[str] = []
    for prefix, label in (("docMetadata/", "sensitivity label (docMetadata/)"),
                          ("customXml/", "document-library metadata (customXml/)"),
                          ("xl/printerSettings/", "printer settings (xl/printerSettings/)")):
        parts = sorted(p for p in data if p.startswith(prefix))
        if parts:
            findings.append(f"{label}: {len(parts)} part(s)")
    workbook = text(data, WORKBOOK)
    if ABS_PATH_BLOCK.search(workbook) or ABS_PATH_BARE.search(workbook):
        findings.append(f"save folder (x15ac:absPath in {WORKBOOK})")
    if CORE in data:
        for match in CORE_PEOPLE.finditer(text(data, CORE)):
            if match.group(1) and re.search(r">\s*[^<\s]", match.group(0)):
                findings.append(f"{match.group(1)} in {CORE}")
    if APP in data:
        for match in APP_PEOPLE.finditer(text(data, APP)):
            if match.group(2).strip():
                findings.append(f"{match.group(1)} in {APP}")
    return findings


def check_package(path: Path) -> list[str]:
    _infos, data = read_package(path)
    return inspect_parts(data)


def well_formed(part: str, content: bytes) -> None:
    parser = expat.ParserCreate()

    def refuse(*_args: object) -> None:
        raise SanitizeError(f"{part}: DOCTYPE is not allowed")

    parser.StartDoctypeDeclHandler = refuse
    try:
        parser.Parse(content, True)
    except expat.ExpatError as error:
        raise SanitizeError(f"{part}: not well-formed after editing ({error})") from error


def sanitize_parts(data: dict[str, bytes]) -> tuple[dict[str, bytes], Result]:
    result = Result(findings=inspect_parts(data))
    out = {part: content for part, content in data.items() if not is_removed(part)}
    result.removed_parts = sorted(part for part in data if is_removed(part))
    edited: dict[str, str] = {}

    # Relationships that point at a removed part, and the r:id each source part used.
    dropped_ids: dict[str, set[str]] = {}
    for part in [p for p in out if p.endswith(".rels")]:
        original = text(out, part)

        def keep(match: re.Match[str], rels_part: str = part) -> str:
            element = match.group(0)
            target = attribute(element, "Target")
            if attribute(element, "TargetMode") == "External" or target is None:
                return element
            if is_removed(resolve(rels_part, target)):
                rel_id = attribute(element, "Id")
                if rel_id:
                    dropped_ids.setdefault(rels_source(rels_part), set()).add(rel_id)
                return ""
            return element

        updated = RELATIONSHIP.sub(keep, original)
        if updated != original:
            edited[part] = updated

    for source, ids in dropped_ids.items():
        if source in out and source.endswith(".xml"):
            content = edited.get(source, text(out, source))
            for rel_id in sorted(ids):
                content = re.sub(r'\s+r:id="' + re.escape(rel_id) + '"', "", content)
            if content != text(out, source):
                edited[source] = content

    types = text(out, CONTENT_TYPES)
    updated_types = OVERRIDE.sub(
        lambda m: "" if is_removed((attribute(m.group(0), "PartName") or "").lstrip("/")) else m.group(0), types)
    if updated_types != types:
        edited[CONTENT_TYPES] = updated_types

    workbook = edited.get(WORKBOOK, text(out, WORKBOOK))
    updated_workbook = ABS_PATH_BARE.sub("", ABS_PATH_BLOCK.sub("", workbook))
    if updated_workbook != text(out, WORKBOOK):
        edited[WORKBOOK] = updated_workbook

    if CORE in out:
        core = text(out, CORE)
        updated_core = CORE_PEOPLE.sub("", core)
        if updated_core != core:
            edited[CORE] = updated_core
    if APP in out:
        app = text(out, APP)
        updated_app = APP_PEOPLE.sub(lambda m: f"<{m.group(1)}></{m.group(1)}>", app)
        if updated_app != app:
            edited[APP] = updated_app

    for part, content in edited.items():
        encoded = content.encode("utf-8")
        well_formed(part, encoded)
        out[part] = encoded
    result.edited_parts = sorted(edited)
    return out, result


def verify(original: dict[str, bytes], sanitized: dict[str, bytes], result: Result) -> None:
    remaining = inspect_parts(sanitized)
    if remaining:
        raise SanitizeError("metadata remains after sanitizing: " + "; ".join(remaining))
    if VBA_PROJECT in original and sanitized.get(VBA_PROJECT) != original[VBA_PROJECT]:
        raise SanitizeError(f"{VBA_PROJECT} changed")
    for part, content in original.items():
        if part in result.removed_parts or part in result.edited_parts:
            continue
        if sanitized.get(part) != content:
            raise SanitizeError(f"{part} changed although it needed no edit")
    for part in [p for p in sanitized if p.endswith(".rels")]:
        for element in RELATIONSHIP.findall(text(sanitized, part)):
            target = attribute(element, "Target")
            if attribute(element, "TargetMode") == "External" or target is None:
                continue
            resolved = resolve(part, target)
            if resolved not in sanitized:
                raise SanitizeError(f"{part}: relationship to missing part {resolved}")
    for element in OVERRIDE.findall(text(sanitized, CONTENT_TYPES)):
        name = (attribute(element, "PartName") or "").lstrip("/")
        if name not in sanitized:
            raise SanitizeError(f"{CONTENT_TYPES}: override for missing part {name}")


def write_package(infos: list[zipfile.ZipInfo], data: dict[str, bytes], destination: Path) -> None:
    descriptor, temporary = tempfile.mkstemp(prefix=".sanitize-", suffix=destination.suffix,
                                             dir=destination.parent)
    os.close(descriptor)
    try:
        with zipfile.ZipFile(temporary, "w") as archive:
            for info in infos:
                if info.filename not in data:
                    continue
                entry = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                entry.compress_type = info.compress_type
                entry.external_attr = info.external_attr
                entry.create_system = info.create_system
                archive.writestr(entry, data[info.filename])
        _infos, reread = read_package(Path(temporary))
        if reread != data:
            raise SanitizeError("the written package does not read back identically")
        os.replace(temporary, destination)
    except BaseException:
        Path(temporary).unlink(missing_ok=True)
        raise


def sanitize_package(source: Path, destination: Path) -> Result:
    infos, data = read_package(source)
    sanitized, result = sanitize_parts(data)
    verify(data, sanitized, result)
    destination.parent.mkdir(parents=True, exist_ok=True)
    write_package(infos, sanitized, destination)
    return result


def report(path: Path, result: Result, written: Path | None) -> dict[str, Any]:
    return {"schema_version": 1, "package": str(path), "output": str(written) if written else None,
            "findings": result.findings, "removed_parts": result.removed_parts,
            "edited_parts": result.edited_parts}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0] if __doc__ else None)
    parser.add_argument("packages", nargs="+", type=Path, help=".xlam or .xlsm package(s)")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--check", action="store_true", help="report metadata; exit 1 if any is present")
    mode.add_argument("--output", type=Path, help="write the sanitized copy here (one package only)")
    mode.add_argument("--in-place", action="store_true", help="replace each package by its sanitized copy")
    parser.add_argument("--report", type=Path, help="write a JSON report")
    options = parser.parse_args(argv)
    if options.output and len(options.packages) != 1:
        parser.error("--output takes exactly one package")
    reports: list[dict[str, Any]] = []
    status = 0
    try:
        for package in options.packages:
            if options.check:
                findings = check_package(package)
                reports.append(report(package, Result(findings=findings), None))
                print(f"{package}: " + ("clean" if not findings else "; ".join(findings)))
                status = status or (1 if findings else 0)
            else:
                destination = options.output if options.output else package
                result = sanitize_package(package, destination)
                reports.append(report(package, result, destination))
                print(f"{package}: removed {len(result.removed_parts)} part(s), edited "
                      f"{len(result.edited_parts)}; " + ("was clean" if not result.findings
                                                         else "removed " + "; ".join(result.findings)))
    except SanitizeError as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2
    if options.report:
        options.report.parent.mkdir(parents=True, exist_ok=True)
        options.report.write_text(json.dumps({"schema_version": 1, "packages": reports}, indent=2) + "\n",
                                  encoding="utf-8")
    return status


if __name__ == "__main__":
    raise SystemExit(main())
