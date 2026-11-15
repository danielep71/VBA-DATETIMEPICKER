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
elements and values. Elements and attributes are recognised by namespace and
local name, whatever prefix a part uses. Every other part is copied with identical content, in the
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

NS_DC = "http://purl.org/dc/elements/1.1/"
NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS_X15AC = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"

# Metadata elements are recognised by namespace and local name, never by prefix.
PEOPLE = {CORE: ((NS_DC, "creator", "dc:creator"), (NS_CP, "lastModifiedBy", "cp:lastModifiedBy")),
          APP: ((NS_EP, "Company", "Company"), (NS_EP, "Manager", "Manager"))}
ABS_PATH = (NS_X15AC, "absPath")
ALTERNATE = (NS_MC, "AlternateContent")
RELATION_ID = (NS_REL, "id")

QNAME = r"(?:[A-Za-z_][\w.-]*:)?"
RELATIONSHIP = re.compile(r"<" + QNAME + r"Relationship\b[^>]*?(?:/>|>\s*</" + QNAME + r"Relationship\s*>)")
OVERRIDE = re.compile(r"<" + QNAME + r"Override\b[^>]*?(?:/>|>\s*</" + QNAME + r"Override\s*>)")
START_TAG = re.compile(r"""<[A-Za-z_][\w.:-]*(?:\s+[^\s=/>"']+\s*=\s*(?:"[^"]*"|'[^']*'))*\s*/?>""")
TAG_ATTRIBUTE = re.compile(r"""(\s+)([^\s=/>"']+)(\s*=\s*)("[^"]*"|'[^']*')""")


class SanitizeError(Exception):
    """The package cannot be read or sanitized safely."""


@dataclass
class Result:
    findings: list[str] = field(default_factory=list)
    removed_parts: list[str] = field(default_factory=list)
    edited_parts: list[str] = field(default_factory=list)
    references: dict[str, tuple[set[str], set[str]]] = field(default_factory=dict)


@dataclass
class Element:
    uri: str
    local: str
    qname: str
    attributes: dict[tuple[str, str], tuple[str, str]]   # (uri, local) -> (qname, value)
    text: str = ""


def parse(part: str, content: bytes) -> list[Element]:
    """Every element of an XML part with its namespace, qualified name, attributes and own text.

    Namespace processing is on, with prefixes reported, so a name is matched by
    URI and local name whatever prefix the part binds. A DOCTYPE is refused.
    """
    elements: list[Element] = []
    stack: list[Element] = []
    parser = expat.ParserCreate(namespace_separator=" ")
    parser.namespace_prefixes = True

    def split(name: str) -> tuple[str, str, str]:
        pieces = name.split(" ")
        if len(pieces) == 1:
            return "", pieces[0], pieces[0]
        uri, local = pieces[0], pieces[1]
        return uri, local, f"{pieces[2]}:{local}" if len(pieces) == 3 else local

    def start(name: str, attrs: dict[str, str]) -> None:
        uri, local, qname = split(name)
        named: dict[tuple[str, str], tuple[str, str]] = {}
        for key, value in attrs.items():
            a_uri, a_local, a_qname = split(key)
            named[(a_uri, a_local)] = (a_qname, value)
        element = Element(uri, local, qname, named)
        elements.append(element)
        stack.append(element)

    def end(_name: str) -> None:
        stack.pop()

    def data(chunk: str) -> None:
        if stack:
            stack[-1].text += chunk

    def refuse(*_args: object) -> None:
        raise SanitizeError(f"{part}: DOCTYPE is not allowed")

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = data
    parser.StartDoctypeDeclHandler = refuse
    try:
        parser.Parse(content, True)
    except expat.ExpatError as error:
        raise SanitizeError(f"{part}: not well-formed XML ({error})") from error
    return elements


def qualified_names(elements: list[Element], uri: str, local: str) -> set[str]:
    return {e.qname for e in elements if (e.uri, e.local) == (uri, local)}


def tag_attributes(element: str) -> dict[str, str]:
    """Unprefixed attributes of one tag, as written (for Relationship and Override)."""
    return {a.group(2): a.group(4)[1:-1] for a in TAG_ATTRIBUTE.finditer(element.split(">", 1)[0])}


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
    if qualified_names(parse(WORKBOOK, data[WORKBOOK]), *ABS_PATH):
        findings.append(f"save folder (x15ac:absPath in {WORKBOOK})")
    for part, fields in PEOPLE.items():
        if part not in data:
            continue
        elements = parse(part, data[part])
        for uri, local, label in fields:
            if any((e.uri, e.local) == (uri, local) and e.text.strip() for e in elements):
                findings.append(f"{label} in {part}")
    return findings


def check_package(path: Path) -> list[str]:
    _infos, data = read_package(path)
    return inspect_parts(data)


def remove_elements(content: str, qnames: set[str], *, keep_empty: bool) -> str:
    """Remove (or empty) every element with one of these qualified names; its text has no markup."""
    for qname in sorted(qnames):
        name = re.escape(qname)
        pattern = re.compile(r"<" + name + r"((?:\s[^>]*?)?)\s*(?:/>|>[^<]*</" + name + r"\s*>)")
        content = pattern.sub((lambda m, q=qname: f"<{q}{m.group(1)}></{q}>") if keep_empty else "", content)
    return content


def remove_abs_path(content: str) -> str:
    elements = parse(WORKBOOK, content.encode("utf-8"))
    for alternate in sorted(qualified_names(elements, *ALTERNATE)):
        for path in sorted(qualified_names(elements, *ABS_PATH)):
            a, p = re.escape(alternate), re.escape(path)
            content = re.sub(r"<" + a + r"\b[^>]*>(?:(?!</" + a + r"\s*>).)*?<" + p + r"\b"
                             r"(?:(?!</" + a + r"\s*>).)*?</" + a + r"\s*>", "", content, flags=re.S)
    return remove_elements(content, qualified_names(elements, *ABS_PATH), keep_empty=False)


def remove_references(content: str, qnames: set[str], ids: set[str]) -> str:
    """Drop the relationship-ID attributes whose value is a removed relationship ID.

    ``qnames`` are the qualified names this part uses for the relationships
    namespace's ``id`` attribute (normally ``r:id``). Each start tag is split into
    whole attributes, so text inside another attribute's value, or in character
    data, is never touched: a tooltip, a cell or a formula may contain the same
    characters.
    """
    def attributes(tag: re.Match[str]) -> str:
        return TAG_ATTRIBUTE.sub(
            lambda a: "" if a.group(2) in qnames and a.group(4)[1:-1] in ids else a.group(0), tag.group(0))

    return START_TAG.sub(attributes, content)


Event = tuple[str, str, tuple[tuple[str, str], ...]]


def events(part: str, content: bytes, drop_ids: set[str] | None = None) -> list[Event]:
    """Elements, attributes and text in order, by namespace; optionally without some relationship IDs."""
    found: list[Event] = []
    for element in parse(part, content):
        kept = {f"{uri} {local}": value for (uri, local), (_qname, value) in element.attributes.items()
                if not (drop_ids and (uri, local) == RELATION_ID and value in drop_ids)}
        found.append((f"{element.uri} {element.local}", element.text, tuple(sorted(kept.items()))))
    return found


def character_data(part: str, content: bytes) -> str:
    """The text content of an XML part, in document order."""
    return "".join(element.text for element in parse(part, content))


def relationships(part: str, content: bytes) -> list[dict[str, str]]:
    return [{local: value for (_uri, local), (_qname, value) in e.attributes.items()}
            for e in parse(part, content) if (e.uri, e.local) == (NS_PKG_REL, "Relationship")]


def sanitize_parts(data: dict[str, bytes]) -> tuple[dict[str, bytes], Result]:
    result = Result(findings=inspect_parts(data))
    out = {part: content for part, content in data.items() if not is_removed(part)}
    result.removed_parts = sorted(part for part in data if is_removed(part))
    edited: dict[str, str] = {}

    # Relationships that point at a removed part, and the IDs each source part used.
    dropped_ids: dict[str, set[str]] = {}
    for part in [p for p in out if p.endswith(".rels")]:
        original = text(out, part)

        def keep(match: re.Match[str], rels_part: str = part) -> str:
            attrs = tag_attributes(match.group(0))
            target = attrs.get("Target")
            if attrs.get("TargetMode") == "External" or target is None:
                return match.group(0)
            if is_removed(resolve(rels_part, target)):
                if attrs.get("Id"):
                    dropped_ids.setdefault(rels_source(rels_part), set()).add(attrs["Id"])
                return ""
            return match.group(0)

        updated = RELATIONSHIP.sub(keep, original)
        if updated != original:
            edited[part] = updated

    for source, ids in dropped_ids.items():
        if source in out and source.endswith(".xml"):
            before = edited.get(source, text(out, source))
            qnames = qualified_attribute_names(source, before.encode("utf-8"), *RELATION_ID)
            content = remove_references(before, qnames, ids)
            if content != before:
                edited[source] = content
                result.references[source] = (qnames, ids)

    types = text(out, CONTENT_TYPES)
    updated_types = OVERRIDE.sub(
        lambda m: "" if is_removed(tag_attributes(m.group(0)).get("PartName", "").lstrip("/")) else m.group(0),
        types)
    if updated_types != types:
        edited[CONTENT_TYPES] = updated_types

    workbook = edited.get(WORKBOOK, text(out, WORKBOOK))
    updated_workbook = remove_abs_path(workbook)
    if updated_workbook != text(out, WORKBOOK):
        edited[WORKBOOK] = updated_workbook

    for part, fields in PEOPLE.items():
        if part not in out:
            continue
        before = text(out, part)
        elements = parse(part, out[part])
        updated = before
        for uri, local, _label in fields:
            updated = remove_elements(updated, qualified_names(elements, uri, local), keep_empty=part == APP)
        if updated != before:
            edited[part] = updated

    for part, content in edited.items():
        encoded = content.encode("utf-8")
        parse(part, encoded)
        out[part] = encoded
    result.edited_parts = sorted(edited)
    return out, result


def qualified_attribute_names(part: str, content: bytes, uri: str, local: str) -> set[str]:
    return {qname for e in parse(part, content) for (a_uri, a_local), (qname, _value) in e.attributes.items()
            if (a_uri, a_local) == (uri, local)}


def verify(original: dict[str, bytes], sanitized: dict[str, bytes], result: Result) -> None:
    remaining = inspect_parts(sanitized)
    if remaining:
        raise SanitizeError("metadata remains after sanitizing: " + "; ".join(remaining))
    if VBA_PROJECT in original and sanitized.get(VBA_PROJECT) != original[VBA_PROJECT]:
        raise SanitizeError(f"{VBA_PROJECT} changed")
    for part in result.edited_parts:
        if part.endswith(".xml") and part not in PEOPLE \
                and character_data(part, original[part]) != character_data(part, sanitized[part]):
            raise SanitizeError(f"{part}: text content changed")
    for part, (_qnames, ids) in result.references.items():
        # Only the relationship-ID attributes may differ; every other element,
        # attribute and text must match. The workbook may also lose its absPath.
        before = original[part]
        if part == WORKBOOK:
            before = remove_abs_path(before.decode("utf-8")).encode("utf-8")
        if events(part, before, ids) != events(part, sanitized[part]):
            raise SanitizeError(f"{part}: more than the removed relationship references changed")
    for part, content in original.items():
        if part in result.removed_parts or part in result.edited_parts:
            continue
        if sanitized.get(part) != content:
            raise SanitizeError(f"{part} changed although it needed no edit")
    for part in [p for p in sanitized if p.endswith(".rels")]:
        for attrs in relationships(part, sanitized[part]):
            target = attrs.get("Target")
            if attrs.get("TargetMode") == "External" or target is None:
                continue
            resolved = resolve(part, target)
            if resolved not in sanitized:
                raise SanitizeError(f"{part}: relationship to missing part {resolved}")
    for element in parse(CONTENT_TYPES, sanitized[CONTENT_TYPES]):
        if (element.uri, element.local) == (NS_TYPES, "Override"):
            name = next((v for (_u, local), (_q, v) in element.attributes.items() if local == "PartName"), "")
            if name.lstrip("/") not in sanitized:
                raise SanitizeError(f"{CONTENT_TYPES}: override for missing part {name.lstrip('/')}")


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
