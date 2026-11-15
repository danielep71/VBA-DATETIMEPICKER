#!/usr/bin/env python3
"""Inject the tracked RibbonX source into a built Office package.

src/ribbon/ mirrors the package's customUI/ folder (RELEASING.md section 7):

    src/ribbon/customUI.xml             -> customUI/customUI.xml
    src/ribbon/_rels/customUI.xml.rels  -> customUI/_rels/customUI.xml.rels
    src/ribbon/<image target>           -> customUI/<image target>

The parts are copied byte for byte. The package also gets what Excel needs to
load them: a root relationship (_rels/.rels) of the Office 2007 extensibility
type pointing at customUI/customUI.xml, and a content type for PNG images when
[Content_Types].xml has none. Every other entry is copied unchanged.

The tool fails closed. It refuses a package that already has a customUI/ entry
or an extensibility relationship (an existing package's Ribbon is not evidence
of a clean build, #90), a Ribbon part that references an image with no
relationship, a relationship whose target is missing or outside the Ribbon
root, and XML with a document type declaration.

Usage:
    python tools/build_ribbon_package.py --package <file.xlam> [--ribbon-root src/ribbon] [--report <file.json>]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path, PurePosixPath
from typing import Any

RIBBON_PART = "customUI/customUI.xml"
RIBBON_RELS = "customUI/_rels/customUI.xml.rels"
ROOT_RELS = "_rels/.rels"
CONTENT_TYPES = "[Content_Types].xml"
CUSTOMUI_NAMESPACE = "http://schemas.microsoft.com/office/2006/01/customui"
RELS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types"
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
EXTENSIBILITY_TYPES = (
    "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility",
    "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility",
)
EXTENSIBILITY_TYPE = EXTENSIBILITY_TYPES[0]
RELATIONSHIP_ID = "rIdDateTimePickerUI"
REQUIRED_DEFAULTS = {"png": "image/png", "xml": "application/xml"}
# Fixed timestamp for the added entries, so two builds of the same inputs add identical entries.
ENTRY_TIME = (1980, 1, 1, 0, 0, 0)


class PackageError(Exception):
    """The package or the Ribbon source cannot be used; the message says why."""


def parse_xml(data: bytes, label: str) -> ET.Element:
    """Parse trusted-format XML, refusing a document type declaration."""
    if b"<!DOCTYPE" in data.upper():
        raise PackageError(f"{label}: a document type declaration is not allowed")
    try:
        return ET.fromstring(data)
    except ET.ParseError as error:
        raise PackageError(f"{label}: not well-formed XML ({error})") from error


def ribbon_inputs(ribbon_root: Path) -> list[tuple[str, Path]]:
    """Return (package part name, source file) for every part the Ribbon needs."""
    root = ribbon_root.resolve()
    part_file = root / "customUI.xml"
    rels_file = root / "_rels" / "customUI.xml.rels"
    for required in (part_file, rels_file):
        if not required.is_file():
            raise PackageError(f"Ribbon source is missing {required.relative_to(root).as_posix()}")

    part = parse_xml(part_file.read_bytes(), "customUI.xml")
    if part.tag != f"{{{CUSTOMUI_NAMESPACE}}}customUI":
        raise PackageError("customUI.xml is not an Office 2007 Ribbon part (namespace 2006/01)")
    referenced = {image for element in part.iter() if (image := element.get("image"))}

    relationships = parse_xml(rels_file.read_bytes(), "_rels/customUI.xml.rels")
    inputs: list[tuple[str, Path]] = [(RIBBON_PART, part_file), (RIBBON_RELS, rels_file)]
    mapped: set[str] = set()
    for relationship in relationships.iter(f"{{{RELS_NAMESPACE}}}Relationship"):
        if relationship.get("Type") != IMAGE_REL_TYPE:
            continue
        rel_id = relationship.get("Id") or ""
        target = relationship.get("Target") or ""
        posix = PurePosixPath(target)
        if not target or posix.is_absolute() or ".." in posix.parts:
            raise PackageError(f"Image relationship {rel_id} has a target outside the Ribbon root: {target!r}")
        source = (root / Path(*posix.parts)).resolve()
        if root not in source.parents:
            raise PackageError(f"Image relationship {rel_id} has a target outside the Ribbon root: {target!r}")
        if not source.is_file():
            raise PackageError(f"Image relationship {rel_id} points at a missing file: {target}")
        mapped.add(rel_id)
        inputs.append((f"customUI/{posix.as_posix()}", source))

    unmapped = sorted(image for image in referenced if image not in mapped)
    if unmapped:
        raise PackageError("customUI.xml references images with no relationship: " + ", ".join(unmapped))
    return inputs


def insert_before_closing(text: str, closing_tag: str, element: str, label: str) -> str:
    """Insert element just before the last closing tag of the document element."""
    position = text.rfind(closing_tag)
    if position < 0:
        raise PackageError(f"{label}: {closing_tag} not found")
    return text[:position] + element + text[position:]


def add_root_relationship(data: bytes) -> tuple[bytes, str]:
    """Return _rels/.rels with the extensibility relationship added, and the Id used."""
    relationships = parse_xml(data, ROOT_RELS)
    ids = set()
    for relationship in relationships.iter(f"{{{RELS_NAMESPACE}}}Relationship"):
        if relationship.get("Type") in EXTENSIBILITY_TYPES:
            raise PackageError("The package already has a Ribbon (extensibility) relationship; build a fresh package")
        ids.add(relationship.get("Id"))
    rel_id = RELATIONSHIP_ID
    suffix = 1
    while rel_id in ids:
        suffix += 1
        rel_id = f"{RELATIONSHIP_ID}{suffix}"
    element = f'<Relationship Id="{rel_id}" Type="{EXTENSIBILITY_TYPE}" Target="{RIBBON_PART}"/>'
    text = insert_before_closing(data.decode("utf-8-sig"), "</Relationships>", element, ROOT_RELS)
    updated = text.encode("utf-8")
    parse_xml(updated, ROOT_RELS)
    return updated, rel_id


def add_content_types(data: bytes) -> tuple[bytes, list[str]]:
    """Return [Content_Types].xml with the defaults the Ribbon parts need, and the extensions added.

    An extension that already has a default keeps it unchanged.
    """
    types = parse_xml(data, CONTENT_TYPES)
    present = {
        (default.get("Extension") or "").lower()
        for default in types.iter(f"{{{TYPES_NAMESPACE}}}Default")
    }
    text = data.decode("utf-8-sig")
    added: list[str] = []
    for extension, content_type in REQUIRED_DEFAULTS.items():
        if extension not in present:
            element = f'<Default Extension="{extension}" ContentType="{content_type}"/>'
            text = insert_before_closing(text, "</Types>", element, CONTENT_TYPES)
            added.append(extension)
    updated = text.encode("utf-8")
    parse_xml(updated, CONTENT_TYPES)
    return updated, added


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def inject(package: Path, ribbon_root: Path) -> dict[str, Any]:
    """Add the Ribbon parts to package in place and return what was done."""
    inputs = ribbon_inputs(ribbon_root)
    if not zipfile.is_zipfile(package):
        raise PackageError(f"{package.name} is not an Office Open XML package")
    with zipfile.ZipFile(package) as source:
        entries = source.infolist()
        names = {entry.filename for entry in entries}
        existing = sorted(name for name in names if name.lower().startswith("customui/"))
        if existing:
            raise PackageError("The package already has Ribbon parts (" + ", ".join(existing)
                               + "); build a fresh package")
        for required in (ROOT_RELS, CONTENT_TYPES):
            if required not in names:
                raise PackageError(f"The package has no {required}")
        root_rels, rel_id = add_root_relationship(source.read(ROOT_RELS))
        content_types, added_types = add_content_types(source.read(CONTENT_TYPES))

        handle, temporary = tempfile.mkstemp(prefix=package.name + ".", suffix=".tmp", dir=package.parent)
        os.close(handle)
        parts: list[dict[str, Any]] = []
        try:
            with zipfile.ZipFile(temporary, "w") as target:
                for entry in entries:
                    if entry.filename == ROOT_RELS:
                        payload = root_rels
                    elif entry.filename == CONTENT_TYPES:
                        payload = content_types
                    else:
                        payload = source.read(entry)
                    target.writestr(entry, payload)
                for part_name, source_file in inputs:
                    payload = source_file.read_bytes()
                    info = zipfile.ZipInfo(part_name, date_time=ENTRY_TIME)
                    info.compress_type = zipfile.ZIP_DEFLATED
                    target.writestr(info, payload)
                    parts.append({
                        "part": part_name,
                        "source": source_file.relative_to(ribbon_root.resolve()).as_posix(),
                        "bytes": len(payload),
                        "sha256": sha256(payload),
                    })
            errors = verify(Path(temporary), ribbon_root)
            if errors:
                raise PackageError("The injected package does not verify: " + "; ".join(errors))
        except BaseException:
            os.unlink(temporary)
            raise
    os.replace(temporary, package)
    return {
        "package": package.name,
        "parts": parts,
        "root_relationship": {"id": rel_id, "type": EXTENSIBILITY_TYPE, "target": RIBBON_PART},
        "content_type_defaults_added": added_types,
    }


def verify(package: Path, ribbon_root: Path) -> list[str]:
    """Return the reasons package does not carry the Ribbon source exactly, or an empty list."""
    errors: list[str] = []
    inputs = ribbon_inputs(ribbon_root)
    with zipfile.ZipFile(package) as archive:
        names = set(archive.namelist())
        for part_name, source_file in inputs:
            if part_name not in names:
                errors.append(f"{part_name} is missing")
            elif archive.read(part_name) != source_file.read_bytes():
                errors.append(f"{part_name} differs from its source")
        if ROOT_RELS not in names:
            errors.append(f"{ROOT_RELS} is missing")
        else:
            relationships = parse_xml(archive.read(ROOT_RELS), ROOT_RELS)
            ribbon = [
                relationship for relationship in relationships.iter(f"{{{RELS_NAMESPACE}}}Relationship")
                if relationship.get("Type") in EXTENSIBILITY_TYPES
            ]
            if len(ribbon) != 1 or ribbon[0].get("Target") != RIBBON_PART \
                    or ribbon[0].get("Type") != EXTENSIBILITY_TYPE:
                errors.append(f"{ROOT_RELS} needs exactly one Office 2007 extensibility relationship to {RIBBON_PART}")
        if CONTENT_TYPES in names:
            types = parse_xml(archive.read(CONTENT_TYPES), CONTENT_TYPES)
            defaults = {
                (default.get("Extension") or "").lower(): default.get("ContentType")
                for default in types.iter(f"{{{TYPES_NAMESPACE}}}Default")
            }
            # A default the package already had is kept as it is: the shipped v1.2.2 add-in
            # declares PNG as image/.PNG and Excel loads it. Only its presence is required.
            for extension in REQUIRED_DEFAULTS:
                if not defaults.get(extension):
                    errors.append(f"{CONTENT_TYPES} has no default content type for .{extension}")
        else:
            errors.append(f"{CONTENT_TYPES} is missing")
    return errors


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=(__doc__ or "").split("\n\n")[0])
    parser.add_argument("--package", type=Path, required=True, help="built .xlam or .xlsm, changed in place")
    parser.add_argument("--ribbon-root", type=Path,
                        default=Path(__file__).resolve().parent.parent / "src" / "ribbon",
                        help="tracked Ribbon source (default: src/ribbon of this checkout)")
    parser.add_argument("--report", type=Path, help="write what was injected as JSON")
    options = parser.parse_args(argv)
    try:
        report = inject(options.package, options.ribbon_root)
    except (PackageError, OSError, zipfile.BadZipFile) as error:
        print(f"FAIL - {error}", file=sys.stderr)
        return 1
    if options.report is not None:
        options.report.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    for part in report["parts"]:
        print(f"Added {part['part']} ({part['bytes']} bytes, sha256 {part['sha256']})")
    print(f"Root relationship {report['root_relationship']['id']} -> {RIBBON_PART}")
    if report["content_type_defaults_added"]:
        print("Content type defaults added: " + ", ".join(report["content_type_defaults_added"]))
    print("PASS - Ribbon injected and verified")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
