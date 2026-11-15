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
BRANCHES = {(NS_MC, "Choice"), (NS_MC, "Fallback")}
RELATION_ID = (NS_REL, "id")
RELATIONSHIP = (NS_PKG_REL, "Relationship")
OVERRIDE = (NS_TYPES, "Override")

START_TAG = re.compile(rb"""<[A-Za-z_][\w.:-]*(?:\s+[^\s=/>"']+\s*=\s*(?:"[^"]*"|'[^']*'))*\s*/?>""")
TAG_ATTRIBUTE = re.compile(rb"""(\s+)([^\s=/>"']+)(\s*=\s*)("[^"]*"|'[^']*')""")


class SanitizeError(Exception):
    """The package cannot be read or sanitized safely."""


@dataclass
class Result:
    findings: list[str] = field(default_factory=list)
    removed_parts: list[str] = field(default_factory=list)
    edited_parts: list[str] = field(default_factory=list)


@dataclass
class Element:
    """One element as the parser saw it, with its byte range in the part."""
    uri: str
    local: str
    qname: str
    attributes: dict[tuple[str, str], tuple[str, str]]   # (uri, local) -> (qname, value)
    parent: int
    depth: int
    start: int
    tag_end: int
    end: int = 0
    text: str = ""

    @property
    def name(self) -> tuple[str, str]:
        return self.uri, self.local

    def value(self, uri: str, local: str) -> str | None:
        found = self.attributes.get((uri, local))
        return found[1] if found else None


def parse(part: str, content: bytes) -> list[Element]:
    """Every element of an XML part with its namespace, attributes, own text and byte range.

    Namespace processing is on, with prefixes reported, so a name is matched by
    URI and local name whatever prefix the part binds. Text includes CDATA. A
    DOCTYPE is refused.
    """
    elements: list[Element] = []
    stack: list[int] = []
    parser = expat.ParserCreate(namespace_separator=" ")
    parser.namespace_prefixes = True

    def split(name: str) -> tuple[str, str, str]:
        pieces = name.split(" ")
        if len(pieces) == 1:
            return "", pieces[0], pieces[0]
        return pieces[0], pieces[1], f"{pieces[2]}:{pieces[1]}" if len(pieces) == 3 else pieces[1]

    def start(name: str, attrs: dict[str, str]) -> None:
        uri, local, qname = split(name)
        named: dict[tuple[str, str], tuple[str, str]] = {}
        for key, value in attrs.items():
            a_uri, a_local, a_qname = split(key)
            named[(a_uri, a_local)] = (a_qname, value)
        offset = parser.CurrentByteIndex
        tag = START_TAG.match(content, offset)
        if tag is None:
            raise SanitizeError(f"{part}: cannot locate the start tag at byte {offset}")
        elements.append(Element(uri, local, qname, named, stack[-1] if stack else -1, len(stack),
                                offset, tag.end()))
        stack.append(len(elements) - 1)

    def end(_name: str) -> None:
        element = elements[stack.pop()]
        if content[element.start:element.tag_end].endswith(b"/>"):
            element.end = element.tag_end
        else:
            offset = parser.CurrentByteIndex
            if not content.startswith(b"</", offset):
                raise SanitizeError(f"{part}: cannot locate the end tag of {element.qname}")
            element.end = content.index(b">", offset) + 1

    def data(chunk: str) -> None:
        if stack:
            elements[stack[-1]].text += chunk

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


def inspect_parts(data: dict[str, bytes]) -> list[str]:
    findings: list[str] = []
    for prefix, label in (("docMetadata/", "sensitivity label (docMetadata/)"),
                          ("customXml/", "document-library metadata (customXml/)"),
                          ("xl/printerSettings/", "printer settings (xl/printerSettings/)")):
        parts = sorted(p for p in data if p.startswith(prefix))
        if parts:
            findings.append(f"{label}: {len(parts)} part(s)")
    if any(e.name == ABS_PATH for e in parse(WORKBOOK, data[WORKBOOK])):
        findings.append(f"save folder (x15ac:absPath in {WORKBOOK})")
    for part, fields in PEOPLE.items():
        if part not in data:
            continue
        elements = parse(part, data[part])
        for uri, local, label in fields:
            if any(e.name == (uri, local) and e.text.strip() for e in elements):
                findings.append(f"{label} in {part}")
    return findings


def check_package(path: Path) -> list[str]:
    _infos, data = read_package(path)
    return inspect_parts(data)


@dataclass
class Plan:
    """What to change in one XML part: whole elements to drop, texts to empty, attributes to drop."""
    drop: set[int] = field(default_factory=set)
    empty: set[int] = field(default_factory=set)
    attributes: dict[int, tuple[str, str]] = field(default_factory=dict)   # element -> (uri, local)

    def __bool__(self) -> bool:
        return bool(self.drop or self.empty or self.attributes)


def subtree(elements: list[Element], index: int) -> set[int]:
    inside = {index}
    for i in range(index + 1, len(elements)):
        if elements[i].parent in inside:
            inside.add(i)
    return inside


def plan_workbook(elements: list[Element]) -> Plan:
    """Drop absPath; drop an AlternateContent wrapper only when nothing else would remain in it."""
    plan = Plan(drop={i for i, e in enumerate(elements) if e.name == ABS_PATH})
    for i, element in enumerate(elements):
        if element.name != ALTERNATE:
            continue
        inside = subtree(elements, i) - {i}
        rest = [j for j in inside if j not in plan.drop and not (plan.drop & ancestors(elements, j))]
        if inside and all(elements[j].name in BRANCHES and not elements[j].text.strip() for j in rest) \
                and not element.text.strip() and plan.drop & inside:
            plan.drop.add(i)
    return plan


def ancestors(elements: list[Element], index: int) -> set[int]:
    found: set[int] = set()
    parent = elements[index].parent
    while parent != -1:
        found.add(parent)
        parent = elements[parent].parent
    return found


def apply(content: bytes, elements: list[Element], plan: Plan) -> bytes:
    """Make the planned edits at the parser's byte offsets; inner edits of a dropped element are skipped."""
    dropped = set().union(*(subtree(elements, i) for i in plan.drop)) if plan.drop else set()
    edits: list[tuple[int, int, bytes]] = []
    for i in plan.drop:
        if not (ancestors(elements, i) & plan.drop):
            edits.append((elements[i].start, elements[i].end, b""))
    for i in plan.empty - dropped:
        element = elements[i]
        if element.end != element.tag_end:
            edits.append((element.tag_end, content.rindex(b"</", element.tag_end, element.end), b""))
    for i, (uri, local) in plan.attributes.items():
        if i in dropped:
            continue
        element = elements[i]
        qname, value = element.attributes[(uri, local)]
        tag = content[element.start:element.tag_end]

        def keep(match: re.Match[bytes], q: bytes = qname.encode("utf-8")) -> bytes:
            return b"" if match.group(2) == q else match.group(0)

        new_tag = TAG_ATTRIBUTE.sub(keep, tag)
        if new_tag == tag:
            raise SanitizeError(f"cannot find attribute {qname}={value!r} in its start tag")
        edits.append((element.start, element.tag_end, new_tag))
    edits.sort(reverse=True)
    for (_start, end, _a), (next_start, _end, _b) in zip(edits[1:], edits[:-1], strict=True):
        if end > next_start:
            raise SanitizeError("overlapping edits")
    for start, end, replacement in edits:
        content = content[:start] + replacement + content[end:]
    return content


Event = tuple[int, str, str, tuple[tuple[str, str], ...]]


def events(elements: list[Element], plan: Plan | None = None) -> list[Event]:
    """Depth, name, text and attributes of every element, as the plan would leave them."""
    plan = plan or Plan()
    dropped = set().union(*(subtree(elements, i) for i in plan.drop)) if plan.drop else set()
    found: list[Event] = []
    for i, element in enumerate(elements):
        if i in dropped:
            continue
        skip = plan.attributes.get(i)
        attrs = tuple(sorted((f"{uri} {local}", value) for (uri, local), (_q, value) in element.attributes.items()
                             if (uri, local) != skip))
        depth = element.depth - sum(1 for a in ancestors(elements, i) if a in dropped)
        found.append((depth, f"{element.uri} {element.local}", "" if i in plan.empty else element.text, attrs))
    return found


def plan_parts(data: dict[str, bytes]) -> dict[str, tuple[list[Element], Plan]]:
    """Parse every XML part that needs an edit and plan it."""
    out = {part: content for part, content in data.items() if not is_removed(part)}
    plans: dict[str, tuple[list[Element], Plan]] = {}
    dropped_ids: dict[str, set[str]] = {}
    for part in [p for p in out if p.endswith(".rels")]:
        elements = parse(part, out[part])
        plan = Plan()
        for i, element in enumerate(elements):
            target = element.value("", "Target")
            if element.name != RELATIONSHIP or target is None or element.value("", "TargetMode") == "External":
                continue
            if is_removed(resolve(part, target)):
                plan.drop.add(i)
                rel_id = element.value("", "Id")
                if rel_id:
                    dropped_ids.setdefault(rels_source(part), set()).add(rel_id)
        if plan:
            plans[part] = (elements, plan)
    for source, ids in dropped_ids.items():
        if source in out and source.endswith(".xml"):
            elements = parse(source, out[source])
            plan = Plan(attributes={i: RELATION_ID for i, e in enumerate(elements)
                                    if e.value(*RELATION_ID) in ids})
            if plan:
                plans[source] = (elements, plan)
    elements = parse(CONTENT_TYPES, out[CONTENT_TYPES])
    plan = Plan(drop={i for i, e in enumerate(elements)
                      if e.name == OVERRIDE and is_removed((e.value("", "PartName") or "").lstrip("/"))})
    if plan:
        plans[CONTENT_TYPES] = (elements, plan)
    elements, plan = plans.get(WORKBOOK, (parse(WORKBOOK, out[WORKBOOK]), Plan()))
    workbook = plan_workbook(elements)
    plan.drop |= workbook.drop
    if plan:
        plans[WORKBOOK] = (elements, plan)
    for part, fields in PEOPLE.items():
        if part not in out:
            continue
        elements = parse(part, out[part])
        names = {(uri, local) for uri, local, _label in fields}
        chosen = {i for i, e in enumerate(elements) if e.name in names}
        plan = Plan(empty=chosen) if part == APP else Plan(drop=chosen)
        if part == APP:
            plan.empty = {i for i in chosen if elements[i].text}
        if plan:
            plans[part] = (elements, plan)
    return plans


def sanitize_parts(data: dict[str, bytes]) -> tuple[dict[str, bytes], Result]:
    result = Result(findings=inspect_parts(data))
    out = {part: content for part, content in data.items() if not is_removed(part)}
    result.removed_parts = sorted(part for part in data if is_removed(part))
    for part, (elements, plan) in plan_parts(data).items():
        out[part] = apply(out[part], elements, plan)
        parse(part, out[part])
    result.edited_parts = sorted(part for part in out if out[part] != data[part])
    return out, result


def verify(original: dict[str, bytes], sanitized: dict[str, bytes], result: Result) -> None:
    remaining = inspect_parts(sanitized)
    if remaining:
        raise SanitizeError("metadata remains after sanitizing: " + "; ".join(remaining))
    if VBA_PROJECT in original and sanitized.get(VBA_PROJECT) != original[VBA_PROJECT]:
        raise SanitizeError(f"{VBA_PROJECT} changed")
    plans = plan_parts(original)
    for part in result.edited_parts:
        if part not in plans:
            raise SanitizeError(f"{part} changed although it needed no edit")
        elements, plan = plans[part]
        if events(elements, plan) != events(parse(part, sanitized[part])):
            raise SanitizeError(f"{part}: more than the planned metadata changed")
    for part, content in original.items():
        if part in result.removed_parts or part in result.edited_parts:
            continue
        if sanitized.get(part) != content:
            raise SanitizeError(f"{part} changed although it needed no edit")
    for part in [p for p in sanitized if p.endswith(".rels")]:
        for element in parse(part, sanitized[part]):
            target = element.value("", "Target")
            if element.name != RELATIONSHIP or target is None or element.value("", "TargetMode") == "External":
                continue
            if resolve(part, target) not in sanitized:
                raise SanitizeError(f"{part}: relationship to missing part {resolve(part, target)}")
    for element in parse(CONTENT_TYPES, sanitized[CONTENT_TYPES]):
        name = (element.value("", "PartName") or "").lstrip("/")
        if element.name == OVERRIDE and name not in sanitized:
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
