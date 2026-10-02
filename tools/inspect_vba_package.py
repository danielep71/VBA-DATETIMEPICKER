#!/usr/bin/env python3
"""Inspect the VBA embedded in a built DatePicker .xlam or .xlsm (diagnostic).

Written for the automatic-startup investigation (#113). It answers two static
questions about one package, without opening Excel or running any macro:

* what startup wiring the package contains: ``Workbook_Open``,
  ``Workbook_AddinInstall``, ``Auto_Open`` and a Ribbon ``onLoad`` callback,
  where they are placed, their signatures, and whether they call ``DP_Start``
  directly or through project procedures; and
* how each embedded module relates to the exported source at an explicit
  revision.

The package is read with the standard library only. The Open XML container is a
zip; ``xl/vbaProject.bin`` is read just far enough to list the VBA storage's
streams ([MS-CFB]) and decompress the module source ([MS-OVBA] 2.4.1). Anything
unexpected stops extraction with a failure instead of producing a partial
inventory. Procedures and statements are parsed by ``check_vba_calls.py``, and
``#If`` branches are evaluated per environment by ``check_vba_conditionals.py``.

This is not a release gate and not runtime evidence: finding a hook that calls
``DP_Start`` does not show that Excel runs it in a given load mode.
"""

from __future__ import annotations

import argparse
import difflib
import hashlib
import io
import json
import re
import struct
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any

from _gatelib import git_bytes, git_text, run_gate
from check_source import ribbon_callback_attributes
from check_vba_calls import (
    MANIFEST, KEYWORDS, Module, Proc, merge_environments, parse_module,
)
from check_vba_conditionals import ENVIRONMENTS, logical_units, reachable_sources

TOOL_NAME = "Packaged VBA startup inspection (diagnostic)"
TOOL_VERSION = "1"
TARGET = "DP_Start"
MAX_TRACE_DEPTH = 4
MAX_PART_BYTES = 64 * 1024 * 1024

CODES = {
    "VBA-PKG-001": ("error", "No startup hook reaches DP_Start in the analyzed code"),
    "VBA-PKG-002": ("error", "Lifecycle procedure is in a module where Excel does not call it"),
    "VBA-PKG-003": ("error", "Lifecycle procedure signature does not match what Excel calls"),
    "VBA-PKG-004": ("warning", "DP_Start is reached only on a conditional path"),
    "VBA-PKG-005": ("unknown", "Dynamic or unresolved call on a startup path"),
    "VBA-PKG-006": ("info", "Startup hook text appears only in comments"),
    "VBA-PKG-007": ("warning", "Only Workbook_AddinInstall reaches DP_Start"),
    "VBA-PKG-009": ("info", "An earlier call can divert to the error handler before DP_Start"),
    "VBA-PKG-010": ("error", "Expected source module is missing from the package"),
    "VBA-PKG-011": ("warning", "Packaged module has no expected source counterpart"),
    "VBA-PKG-012": ("error", "Packaged module code differs from the source revision"),
    "VBA-PKG-013": ("info", "Document module has no tracked source counterpart"),
    "VBA-PKG-014": ("warning", "Module kind differs between package and source"),
    "VBA-PKG-015": ("info", "Packaged module differs from the source only in comments or blank lines"),
    "VBA-PKG-090": ("failure", "Package extraction failed"),
    "VBA-PKG-091": ("failure", "Analysis is incomplete"),
    "VBA-PKG-092": ("unknown", "Input is a manual export, not extraction from a package"),
}

# VB_Base class identifiers of Excel document modules.
DOCUMENT_OBJECTS = {
    "00020819-0000-0000-c000-000000000046": "workbook",
    "00020820-0000-0000-c000-000000000046": "worksheet",
    "00020821-0000-0000-c000-000000000046": "chart",
}
LIFECYCLE = {
    "workbook_open": ("Workbook_Open", "workbook",
                      "Workbook event raised when this workbook opens, if events are enabled."),
    "workbook_addininstall": ("Workbook_AddinInstall", "workbook",
                              "Workbook event raised when the add-in is installed through the Add-ins "
                              "manager; it is not raised on every load."),
    "auto_open": ("Auto_Open", "standard",
                  "Auto macro Excel runs on some interactive opens; not on programmatic opens."),
}
HANDLER_TEXT = {"goto": "On Error GoTo <label>", "resume-next": "On Error Resume Next"}
DYNAMIC = {"run", "ontime", "callbyname", "evaluate", "executeexcel4macro"}
BLOCK_OPEN = {"select", "for", "do", "while"}
BLOCK_CLOSE = {"next", "loop", "wend"}
NORMALIZATION = [
    "N1 decode: package modules with the project code page from the dir stream; source files as cp1252, "
    "the repository's export encoding.",
    "N2 line endings: CRLF and CR become LF.",
    "N3 export header: source lines before the first 'Attribute VB_Name' line (VERSION, BEGIN...END and "
    "the form designer block) are dropped; the module stream has no such header.",
    "N4 end of file: trailing empty lines are dropped.",
    "N5 stream-only attributes: in the leading Attribute block of a class or form, the VB_Base, "
    "VB_TemplateDerived and VB_Customizable lines are dropped from both sides; the VBE stores them in the "
    "module stream but omits them from class and form exports. VB_Base is read before this step.",
    "Nothing else changes: case, indentation, comments and Attribute lines are compared as they are.",
    "Code comparison (only after a text difference): comments, blank lines and line continuations are "
    "removed using check_vba_conditionals.logical_units; a difference that survives is executable.",
]
UNSUPPORTED = [
    "compiled p-code and the performance cache are not read; Excel can execute p-code that no longer "
    "matches the stored source",
    "form designer storage (.frx data) and document-module objects (sheets, controls) are not compared",
    "calls through object receivers, With blocks and class instances are not followed",
    "a call is 'unconditional' by structure only; runtime errors, disabled events, macro security and "
    "add-in load state are not modeled",
    "only the Ribbon onLoad callback is treated as a startup path; other Ribbon callbacks need a user action",
    "Excel 4 macro sheets, XLM Auto_Open names and Application-level events in other workbooks are not read",
]


class ExtractionError(ValueError):
    """The package or its VBA project could not be read as expected."""


def finding(code: str, message: str, module: str | None = None, procedure: str | None = None,
            line: int | None = None, target: str | None = None, **extra: Any) -> dict[str, Any]:
    severity, title = CODES[code]
    item = {"code": code, "severity": severity, "title": title, "path": None, "module": module,
            "procedure": procedure, "line": line, "target": target, "message": message}
    item.update(extra)
    return item


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# ---------------------------------------------------------------------------
# Extraction: [MS-CFB] streams and [MS-OVBA] decompression, read-only
# ---------------------------------------------------------------------------

ENDOFCHAIN = 0xFFFFFFFE
MAXREGSECT = 0xFFFFFFFA
NOSTREAM = 0xFFFFFFFF


class CompoundFile:
    """Just enough of [MS-CFB] to read named streams; anything unexpected raises."""

    def __init__(self, data: bytes):
        if len(data) < 512 or data[:8] != bytes.fromhex("D0CF11E0A1B11AE1"):
            raise ExtractionError("vbaProject.bin is not a compound file")
        major, = struct.unpack_from("<H", data, 0x1A)
        shift, mini_shift = struct.unpack_from("<HH", data, 0x1E)
        if (major, shift) not in {(3, 9), (4, 12)} or mini_shift != 6:
            raise ExtractionError(f"unsupported compound file version {major} / sector shift {shift}")
        self.data, self.size, self.major = data, 1 << shift, major
        (fat_count, first_dir, _, cutoff, first_mini, mini_count, first_difat,
         difat_count) = struct.unpack_from("<8I", data, 0x2C)
        if cutoff != 4096:
            raise ExtractionError(f"unexpected mini stream cutoff {cutoff}")
        per_sector = self.size // 4
        difat = list(struct.unpack_from("<109I", data, 0x4C))
        sector = first_difat
        for _ in range(difat_count):
            entries = struct.unpack_from(f"<{per_sector}I", self.sector(sector))
            difat.extend(entries[:-1])
            sector = entries[-1]
        fat_sectors = [s for s in difat if s <= MAXREGSECT]
        if len(fat_sectors) != fat_count:
            raise ExtractionError("FAT sector count does not match the header")
        self.fat: list[int] = []
        for sector in fat_sectors:
            self.fat.extend(struct.unpack_from(f"<{per_sector}I", self.sector(sector)))
        directory = self.chain(first_dir, self.fat, self.sector)
        self.entries = [self.entry(directory[i:i + 128]) for i in range(0, len(directory) - 127, 128)]
        if not self.entries or self.entries[0]["type"] != 5:
            raise ExtractionError("compound file has no root entry")
        self.minifat: list[int] = []
        if mini_count:
            minifat = self.chain(first_mini, self.fat, self.sector)
            self.minifat = list(struct.unpack_from(f"<{len(minifat) // 4}I", minifat))
        root = self.entries[0]
        self.ministream = self.chain(root["start"], self.fat, self.sector)[:root["size"]] if root["size"] else b""

    def sector(self, number: int) -> bytes:
        start = (number + 1) * self.size
        if number > MAXREGSECT or start >= len(self.data):
            raise ExtractionError(f"sector {number} is outside the file")
        return self.data[start:start + self.size].ljust(self.size, b"\0")

    def mini_sector(self, number: int) -> bytes:
        start = number * 64
        if start + 64 > len(self.ministream):
            raise ExtractionError(f"mini sector {number} is outside the mini stream")
        return self.ministream[start:start + 64]

    @staticmethod
    def chain(start: int, table: list[int], read: Any) -> bytes:
        parts, seen, number = [], set(), start
        while number != ENDOFCHAIN:
            if number > MAXREGSECT or number >= len(table) or number in seen:
                raise ExtractionError("broken sector chain")
            seen.add(number)
            parts.append(read(number))
            number = table[number]
        return b"".join(parts)

    def entry(self, raw: bytes) -> dict[str, Any]:
        name_length, = struct.unpack_from("<H", raw, 64)
        left, right, child = struct.unpack_from("<3I", raw, 68)
        start, size = struct.unpack_from("<IQ", raw, 116)
        if self.major == 3:
            size &= 0xFFFFFFFF
        name = raw[:max(min(name_length, 64) - 2, 0)].decode("utf-16-le", errors="replace")
        return {"name": name, "type": raw[66], "left": left, "right": right, "child": child,
                "start": start, "size": size}

    def children(self, index: int) -> dict[str, dict[str, Any]]:
        found: dict[str, dict[str, Any]] = {}
        stack, seen = [self.entries[index]["child"]], set()
        while stack:
            current = stack.pop()
            if current == NOSTREAM:
                continue
            if current >= len(self.entries) or current in seen:
                raise ExtractionError("broken directory tree")
            seen.add(current)
            item = self.entries[current]
            found[item["name"].casefold()] = {**item, "index": current}
            stack += [item["left"], item["right"]]
        return found

    def stream(self, *path: str) -> bytes:
        index = 0
        for position, name in enumerate(path):
            item = self.children(index).get(name.casefold())
            expected = 2 if position == len(path) - 1 else 1
            if item is None or item["type"] != expected:
                raise ExtractionError(f"stream {'/'.join(path)} is missing")
            index = item["index"]
        item = self.entries[index]
        if item["size"] > MAX_PART_BYTES:
            raise ExtractionError(f"stream {'/'.join(path)} is too large")
        if item["size"] < 4096:
            data = self.chain(item["start"], self.minifat, self.mini_sector) if item["size"] else b""
        else:
            data = self.chain(item["start"], self.fat, self.sector)
        if len(data) < item["size"]:
            raise ExtractionError(f"stream {'/'.join(path)} is shorter than its directory size")
        return data[:item["size"]]


def decompress(data: bytes) -> bytes:
    """Decompress an [MS-OVBA] 2.4.1 CompressedContainer."""
    if not data or data[0] != 1:
        raise ExtractionError("compressed container signature is missing")
    out, position = bytearray(), 1
    while position < len(data):
        if position + 2 > len(data):
            raise ExtractionError("truncated compressed chunk header")
        header, = struct.unpack_from("<H", data, position)
        if (header >> 12) & 0b111 != 0b011:
            raise ExtractionError("invalid compressed chunk signature")
        end = min(position + (header & 0x0FFF) + 3, len(data))
        position += 2
        chunk_start = len(out)
        if not header & 0x8000:
            out += data[position:position + 4096]
            position += 4096
            continue
        while position < end:
            flags = data[position]
            position += 1
            for bit in range(8):
                if position >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[position])
                    position += 1
                    continue
                if position + 2 > end:
                    raise ExtractionError("truncated copy token")
                token, = struct.unpack_from("<H", data, position)
                position += 2
                written = len(out) - chunk_start
                if written < 1:
                    raise ExtractionError("copy token before any literal")
                bits = max((written - 1).bit_length(), 4)
                offset = (token >> (16 - bits)) + 1
                if offset > written:
                    raise ExtractionError("copy token points before its chunk")
                for _ in range((token & (0xFFFF >> bits)) + 3):
                    out.append(out[-offset])
        if len(out) > MAX_PART_BYTES:
            raise ExtractionError("decompressed stream is too large")
    return bytes(out)


def parse_dir(data: bytes) -> tuple[int, list[dict[str, Any]]]:
    """Read the code page and module records of the decompressed ``dir`` stream."""
    codepage, expected, modules, current, position = None, None, [], None, 0
    raw: dict[str, bytes] = {}
    while position + 6 <= len(data):
        record, size = struct.unpack_from("<HI", data, position)
        position += 6
        if record == 0x0009:  # PROJECTVERSION: its size field is fixed at 4 but 6 bytes follow
            size = 6
        body = data[position:position + size]
        if len(body) != size:
            raise ExtractionError(f"dir record 0x{record:04X} is truncated")
        position += size
        if record == 0x0003:
            codepage, = struct.unpack("<H", body)
        elif record == 0x000F:
            expected, = struct.unpack("<H", body)
        elif record == 0x0019:
            raw = {"name": body}
            current = {"name": None, "stream": None, "offset": None, "procedural": None}
            modules.append((current, raw))
        elif current is not None and record == 0x0047:
            current["name"] = body.decode("utf-16-le")
        elif current is not None and record == 0x001A:
            raw["stream"] = body
        elif current is not None and record == 0x0032:
            current["stream"] = body.decode("utf-16-le")
        elif current is not None and record == 0x0031:
            current["offset"], = struct.unpack("<I", body)
        elif current is not None and record in {0x0021, 0x0022}:
            current["procedural"] = record == 0x0021
        elif record == 0x002B:
            current = None
        elif record == 0x0010:
            break
    if codepage is None or expected is None:
        raise ExtractionError("dir stream has no code page or module count")
    encoding = f"cp{codepage}"
    result = []
    for module, names in modules:
        module["name"] = module["name"] or names["name"].decode(encoding)
        module["stream"] = module["stream"] or names.get("stream", b"").decode(encoding)
        if not module["stream"] or module["offset"] is None or module["procedural"] is None:
            raise ExtractionError(f"dir stream record for module {module['name']} is incomplete")
        result.append(module)
    if len(result) != expected:
        raise ExtractionError(f"dir stream lists {len(result)} modules but declares {expected}")
    return codepage, result


def project_kinds(text: str) -> dict[str, str]:
    """Module kinds from the PROJECT stream (Module=, Class=, BaseClass=, Document=)."""
    kinds = {"module": "standard", "class": "class", "baseclass": "form", "document": "document"}
    found = {}
    for line in text.splitlines():
        if line.startswith("["):
            break
        key, _, value = line.partition("=")
        if key.casefold() in kinds and value:
            found[value.split("/")[0].casefold()] = kinds[key.casefold()]
    return found


def document_object(text: str) -> str | None:
    match = re.search(r'^Attribute VB_Base = "0\{([0-9A-Fa-f-]+)\}', text, re.MULTILINE)
    return DOCUMENT_OBJECTS.get(match.group(1).casefold(), "other") if match else None


def extract_package(path: Path) -> dict[str, Any]:
    data = path.read_bytes()
    result: dict[str, Any] = {
        "filename": path.name, "size": len(data), "sha256": sha256(data), "vba_project": None,
        "ribbon": None, "codepage": None, "modules": [], "failures": [],
    }
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as error:
        result["failures"].append(finding("VBA-PKG-090", f"{path.name} is not an Open XML package: {error}"))
        return result
    with archive:
        names = {info.filename.casefold(): info for info in archive.infolist()}
        for part in ("customui/customui14.xml", "customui/customui.xml"):
            info = names.get(part)
            if info is not None and info.file_size <= 1_000_000:
                xml = archive.read(info)
                onload = [value for element, attribute, value in ribbon_callback_attributes(xml)
                          if element.split(":")[-1] == "customUI" and attribute == "onLoad"]
                result["ribbon"] = {"part": info.filename, "sha256": sha256(xml), "onLoad": onload[0] if onload else None}
                break
        info = names.get("xl/vbaproject.bin")
        if info is None:
            result["failures"].append(finding("VBA-PKG-090", f"{path.name} contains no xl/vbaProject.bin."))
            return result
        if info.file_size > MAX_PART_BYTES:
            result["failures"].append(finding("VBA-PKG-090", "xl/vbaProject.bin is too large to inspect."))
            return result
        project = archive.read(info)
    result["vba_project"] = {"part": info.filename, "size": len(project), "sha256": sha256(project)}
    try:
        compound = CompoundFile(project)
        codepage, records = parse_dir(decompress(compound.stream("VBA", "dir")))
        kinds = project_kinds(compound.stream("PROJECT").decode(f"cp{codepage}"))
    except (ExtractionError, struct.error, LookupError, UnicodeDecodeError) as error:
        result["failures"].append(finding("VBA-PKG-090", f"VBA project could not be read: {error}"))
        return result
    result["codepage"] = codepage
    for record in records:
        module = {"name": record["name"], "stream": record["stream"],
                  "kind": kinds.get(record["name"].casefold()) or ("standard" if record["procedural"] else None),
                  "procedural": record["procedural"]}
        try:
            raw = decompress(compound.stream("VBA", record["stream"])[record["offset"]:])
            module["raw"] = raw
            module["text"] = raw.decode(f"cp{codepage}")
        except (ExtractionError, struct.error, UnicodeDecodeError) as error:
            result["failures"].append(finding("VBA-PKG-090", f"module {record['name']} could not be read: {error}",
                                              module=record["name"]))
            module["raw"] = module["text"] = None
        if module["kind"] is None:
            result["failures"].append(finding("VBA-PKG-091", f"module {record['name']} has no kind in the PROJECT stream.",
                                              module=record["name"]))
            module["kind"] = "class"
        result["modules"].append(module)
    return result


def read_export(directory: Path) -> dict[str, Any]:
    """Manual-export fallback: modules exported from the VBE into one folder."""
    result: dict[str, Any] = {"filename": None, "size": None, "sha256": None, "vba_project": None, "ribbon": None,
                              "codepage": None, "modules": [], "failures": []}
    files = sorted(p for p in directory.iterdir() if p.suffix.casefold() in {".bas", ".cls", ".frm"})
    if not files:
        result["failures"].append(finding("VBA-PKG-090", f"{directory.name} contains no exported .bas/.cls/.frm files."))
    for path in files:
        raw = path.read_bytes()
        text = raw.decode("cp1252")
        name = re.search(r'^Attribute VB_Name = "([^"]+)"', text, re.MULTILINE)
        kind = {".bas": "standard", ".frm": "form"}.get(path.suffix.casefold(), "class")
        if kind == "class" and document_object(text) is not None:
            kind = "document"
        result["modules"].append({"name": name.group(1) if name else path.stem, "stream": None, "kind": kind,
                                  "procedural": kind == "standard", "raw": raw, "text": export_body(text),
                                  "export_file": path.name})
    return result


# ---------------------------------------------------------------------------
# Source comparison
# ---------------------------------------------------------------------------


def export_body(text: str) -> str:
    """Apply N2 and N3: LF line endings, no export-only header."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    match = re.search(r'^Attribute VB_Name = ', text, re.MULTILINE)
    return text[match.start():] if match else text


STREAM_ONLY_ATTRIBUTES = re.compile(r"^Attribute VB_(?:Base|TemplateDerived|Customizable) = ", re.IGNORECASE)


def normalize(text: str, kind: str) -> str:
    """Apply N2-N5; see NORMALIZATION."""
    lines = export_body(text).rstrip("\n").split("\n")
    if kind in {"class", "form"}:
        header = next((i for i, line in enumerate(lines) if not line.startswith("Attribute ")), len(lines))
        lines = [line for line in lines[:header] if not STREAM_ONLY_ATTRIBUTES.match(line)] + lines[header:]
    return "\n".join(lines) + "\n"


def code_lines(text: str) -> list[str]:
    return [" ".join(code.split()) for _s, _e, _k, code in logical_units(text.split("\n")) if code.strip()]


def read_source(source: Path, revision: str, configuration: dict[str, Any]) -> dict[str, Any]:
    resolved = git_text(source, "rev-parse", "--verify", "--quiet", f"{revision}^{{commit}}")
    result: dict[str, Any] = {"revision_requested": revision, "revision": None, "modules": {}, "failures": []}
    if resolved.returncode:
        result["failures"].append(finding("VBA-PKG-091", f"source revision {revision!r} cannot be resolved."))
        return result
    result["revision"] = resolved.stdout.strip()
    for path in configuration.get("modules", []):
        blob = git_bytes(source, "show", f"{result['revision']}:{path}")
        if blob.returncode:
            result["failures"].append(finding("VBA-PKG-091", f"{path} does not exist at {result['revision'][:12]}.",
                                              target=path))
            continue
        text = blob.stdout.decode("cp1252")
        name = re.search(r'^Attribute VB_Name = "([^"]+)"', text, re.MULTILINE)
        kind = {".bas": "standard", ".cls": "class", ".frm": "form"}[Path(path).suffix.casefold()]
        result["modules"][(name.group(1) if name else Path(path).stem).casefold()] = {
            "path": path, "kind": kind, "raw": blob.stdout, "text": text}
    return result


def compare_modules(package: dict[str, Any], source: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    rows, findings = [], []
    matched = set()
    for module in package["modules"]:
        row = {"name": module["name"], "kind": module["kind"], "stream": module["stream"],
               "document_object": document_object(module["text"]) if module["text"] and module["kind"] == "document"
               else None,
               "lines": None, "raw_sha256": sha256(module["raw"]) if module["raw"] is not None else None,
               "normalized_sha256": None, "source": None, "comparison": "not-compared"}
        if module.get("export_file"):
            row["export_file"] = module["export_file"]
        if module["text"] is None:
            row["comparison"] = "not-extracted"
            rows.append(row)
            continue
        packaged = normalize(module["text"], module["kind"])
        row["lines"] = packaged.count("\n")
        row["normalized_sha256"] = sha256(packaged.encode("utf-8"))
        counterpart = source["modules"].get(module["name"].casefold())
        if counterpart is None:
            if module["kind"] == "document":
                row["comparison"] = "no-tracked-source"
                findings.append(finding("VBA-PKG-013", f"{module['name']} ({row['document_object'] or 'document'}) "
                                        "is workbook/document code with no tracked source file.", module=module["name"]))
            else:
                row["comparison"] = "unexpected"
                findings.append(finding("VBA-PKG-011", f"{module['name']} is not in the expected source configuration.",
                                        module=module["name"]))
            rows.append(row)
            continue
        matched.add(module["name"].casefold())
        expected = normalize(counterpart["text"], counterpart["kind"])
        row["source"] = {"path": counterpart["path"], "raw_sha256": sha256(counterpart["raw"]),
                         "normalized_sha256": sha256(expected.encode("utf-8"))}
        if counterpart["kind"] != module["kind"]:
            findings.append(finding("VBA-PKG-014", f"{module['name']} is a {module['kind']} module in the package "
                                    f"and a {counterpart['kind']} file in source.", module=module["name"]))
        if packaged == expected:
            row["comparison"] = "identical"
        else:
            old, new = expected.split("\n"), packaged.split("\n")
            diff = [line for line in difflib.unified_diff(old, new, "source", "package", n=0, lineterm="")
                    if not line.startswith(("---", "+++"))]
            first = next((i for i, (a, b) in enumerate(zip(old, new)) if a != b), min(len(old), len(new)))
            row["difference"] = {
                "first_line": first + 1,
                "removed_lines": sum(1 for line in diff if line.startswith("-")),
                "added_lines": sum(1 for line in diff if line.startswith("+")),
                "excerpt": [line[:200] for line in diff[:12]],
            }
            if code_lines(packaged) == code_lines(expected):
                row["comparison"] = "comments-only"
                findings.append(finding("VBA-PKG-015", f"{module['name']} differs from {counterpart['path']} only in "
                                        "comments, blank lines or line breaks.", module=module["name"], line=first + 1))
            else:
                row["comparison"] = "changed"
                findings.append(finding("VBA-PKG-012", f"{module['name']} code differs from {counterpart['path']} "
                                        f"(first difference at line {first + 1}).", module=module["name"], line=first + 1))
        rows.append(row)
    for key, counterpart in source["modules"].items():
        if key not in matched:
            findings.append(finding("VBA-PKG-010", f"{counterpart['path']} has no module in the package.",
                                    target=counterpart["path"]))
    return rows, findings


# ---------------------------------------------------------------------------
# Startup wiring
# ---------------------------------------------------------------------------


class Index:
    """Procedures of one environment's modules, resolved the way VBA resolves unqualified calls."""

    def __init__(self, modules: list[tuple[Module, str | None]], patterns: list[re.Pattern[str]]):
        self.modules = {m.name.casefold(): (m, obj) for m, obj in modules}
        self.patterns = patterns

    def resolve(self, name: str, module: Module, qualifier: str | None) -> Proc | None:
        key = name.casefold()
        if qualifier is not None:
            entry = self.modules.get(qualifier.casefold())
            if entry is None:
                return None
            member = entry[0].members.get(key)
            return next(iter(member.accessors.values()), None) if member and member.kind == "proc" else None
        member = module.members.get(key)
        if member is not None:
            return next(iter(member.accessors.values()), None) if member.kind == "proc" else None
        for other, _ in self.modules.values():
            member = other.members.get(key)
            if other.kind == "standard" and member and member.kind == "proc" and member.visibility != "private":
                return next(iter(member.accessors.values()), None)
        return None

    def undeclared_project_name(self, name: str) -> bool:
        """A project-named identifier that no module declares as anything (procedure, constant, variable...)."""
        key = name.casefold()
        if key in self.modules or any(key in m.members for m, _ in self.modules.values()):
            return False
        return any(p.search(name) for p in self.patterns)


def calls_in(tokens: list[Any], proc: Proc) -> list[tuple[str, str | None, str | None]]:
    """(kind, qualifier, name) for each call candidate: 'call', or 'dynamic' with an optional literal."""
    found = []
    start = 1 if tokens and tokens[0].is_kw("call") else 0
    assignment = len(tokens) > 1 and tokens[0].is_id() and tokens[1].is_op("=")
    for index, token in enumerate(tokens):
        if not token.is_id() or token.is_kw(*KEYWORDS) or (assignment and index == 0) or index < start:
            continue
        if token.low in DYNAMIC:
            literal = next((t.text.strip('"') for t in tokens[index + 1:] if t.kind == "str"), None)
            found.append(("dynamic", token.text, literal))
            continue
        if index + 1 < len(tokens) and tokens[index + 1].is_op(".", "!"):
            continue
        qualifier = None
        if index >= 2 and tokens[index - 1].is_op("."):
            if not tokens[index - 2].is_id():
                continue
            qualifier = tokens[index - 2].text
        if token.low in proc.locals:
            continue
        found.append(("call", qualifier, token.text))
    return found


def on_error_handler(proc: Proc) -> str | None:
    """'resume-next' or 'goto' when the procedure's first executable statement sets an error handler."""
    executable = [t for _line, t in proc.statements if not t[0].is_kw("dim", "const", "static")]
    if not executable:
        return None
    tokens = executable[0]
    words = [t.low for t in tokens[:4]]
    if words[:2] == ["on", "error"]:
        if words[2:4] == ["resume", "next"]:
            return "resume-next"
        if words[2:3] == ["goto"] and len(tokens) > 3 and tokens[3].text not in {"0", "-1"}:
            return "goto"
    return None


def trace(proc: Proc, module: Module, index: Index, depth: int, seen: set[str]) -> dict[str, Any]:
    """Walk one procedure's statements in order, following project calls toward TARGET."""
    result: dict[str, Any] = {"reaches": [], "dynamic": [], "unresolved": [], "diversions": []}
    nesting, exited, handler = 0, False, None
    earlier_calls: list[dict[str, Any]] = []
    single_if_line = None
    for line, tokens in proc.statements:
        head = tokens[0].low if tokens[0].is_id() else ""
        conditional = nesting > 0 or exited or single_if_line == line
        if head == "on" and len(tokens) > 2 and tokens[1].is_kw("error"):
            if tokens[2].is_kw("resume"):
                handler = "resume-next"
            elif tokens[2].is_kw("goto") and len(tokens) > 3:
                handler = None if tokens[3].text in {"0", "-1"} else "goto"
            continue
        if head in {"if", "elseif"} and tokens[-1].is_kw("then"):
            nesting += head == "if"
            continue
        if head == "if":
            single_if_line = line
            then = next((i for i, t in enumerate(tokens) if t.is_kw("then")), len(tokens))
            tokens = tokens[then + 1:]
            conditional = True
            if not tokens:
                continue
            head = tokens[0].low if tokens[0].is_id() else ""
        if head == "else":
            continue
        if head == "end" and len(tokens) > 1 and tokens[1].is_kw("if", "select"):
            nesting = max(nesting - 1, 0)
            continue
        if head in BLOCK_OPEN or (head == "select" and len(tokens) > 1 and tokens[1].is_kw("case")):
            nesting += 1
            continue
        if head in BLOCK_CLOSE:
            nesting = max(nesting - 1, 0)
            continue
        if head in {"exit", "goto", "resume", "return", "stop"} or (head == "end" and len(tokens) == 1):
            if head == "exit" and len(tokens) > 1 and tokens[1].is_kw("do", "for"):
                continue
            if not conditional:
                exited = True
            continue
        for kind, qualifier, name in calls_in(tokens, proc):
            site = {"module": module.name, "procedure": proc.name, "line": line, "call": name,
                    "conditional": conditional}
            if kind == "dynamic":
                site.update(call=f"{qualifier}", literal=name)
                result["dynamic"].append(site)
                continue
            if name.casefold() == TARGET.casefold():
                site["handler"] = handler
                site["earlier_calls_under_goto"] = [c for c in earlier_calls if c["handler"] == "goto"]
                result["reaches"].append({**site, "path": [site]})
                continue
            callee = index.resolve(name, module, qualifier)
            if callee is None:
                if index.undeclared_project_name(name):
                    result["unresolved"].append(site)
                continue
            callee_module = index.modules[callee.module.casefold()][0]
            earlier_calls.append({"call": callee.name, "line": line, "handler": handler,
                                  "callee_handler": on_error_handler(callee)})
            key = f"{callee.module}.{callee.name}".casefold()
            if depth >= MAX_TRACE_DEPTH or key in seen:
                continue
            nested = trace(callee, callee_module, index, depth + 1, seen | {key})
            for reach in nested["reaches"]:
                result["reaches"].append({**reach, "conditional": reach["conditional"] or conditional,
                                          "path": [site] + reach["path"]})
            result["dynamic"] += nested["dynamic"]
            result["unresolved"] += nested["unresolved"]
    return result


def analyze_environment(modules: list[dict[str, Any]], environment: str, patterns: list[re.Pattern[str]],
                        ribbon_onload: str | None) -> dict[str, Any]:
    parsed: list[tuple[Module, str | None]] = []
    failures: list[dict[str, Any]] = []
    for item in modules:
        if item["text"] is None:
            continue
        suffix = {"standard": ".bas", "form": ".frm"}.get(item["kind"], ".cls")
        variants, conditional = reachable_sources(item["name"] + suffix, item["text"])
        for problem in conditional:
            failures.append(finding("VBA-PKG-091", f"conditional compilation is indeterminate: {problem['message']}",
                                    module=item["name"], line=problem.get("line")))
        module = parse_module(item["name"] + suffix, variants[environment])
        module.kind = item["kind"]
        for problem in module.failures:
            failures.append(finding("VBA-PKG-091", f"module could not be parsed: {problem['message']}",
                                    module=item["name"], line=problem.get("line")))
        parsed.append((module, document_object(item["text"]) if item["kind"] == "document" else None))
    index = Index(parsed, patterns)
    hooks, findings = [], []
    entries: list[tuple[str, Module, Proc, str | None]] = []
    for module, obj in parsed:
        for proc in module.procs:
            key = proc.name.casefold()
            if key in LIFECYCLE and proc.accessor != "declare":
                entries.append((key, module, proc, obj))
    if ribbon_onload:
        name = ribbon_onload.split(".")[-1].split("!")[-1]
        for module, _obj in parsed:
            member = module.members.get(name.casefold())
            if module.kind == "standard" and member and member.kind == "proc":
                entries.append(("ribbon_onload", module, next(iter(member.accessors.values())), None))
    for key, module, proc, obj in entries:
        if key == "ribbon_onload":
            label, host, when = "Ribbon onLoad", "standard", "Ribbon callback run when the package's Ribbon loads."
        else:
            label, host, when = LIFECYCLE[key]
        placement = "correct"
        if host == "workbook":
            if module.kind != "document" or obj not in {"workbook", None}:
                placement = "misplaced"
            elif obj is None:
                placement = "unknown"
        elif module.kind != "standard":
            placement = "misplaced"
        signature = "correct" if proc.accessor == "sub" and (not proc.params or key == "ribbon_onload") else "mismatch"
        result = trace(proc, module, index, 0, {f"{module.name}.{proc.name}".casefold()})
        hook = {"hook": label, "module": module.name, "module_kind": module.kind, "document_object": obj,
                "procedure": proc.name, "line": proc.line, "signature": proc.signature(), "accessor": proc.accessor,
                "placement": placement, "signature_check": signature, "when_excel_calls_it": when,
                "reaches_target": [{"conditional": r["conditional"], "handler": r.get("handler"),
                                    "earlier_calls_under_goto": r.get("earlier_calls_under_goto", []),
                                    "path": [f"{s['module']}.{s['procedure']}:{s['line']} -> {s['call']}"
                                             for s in r["path"]]} for r in result["reaches"]],
                "dynamic_calls": result["dynamic"], "unresolved_calls": result["unresolved"]}
        hooks.append(hook)
        where = {"module": module.name, "procedure": proc.name, "line": proc.line}
        if placement == "misplaced":
            findings.append(finding("VBA-PKG-002", f"{proc.name} is in {module.kind} module {module.name}"
                                    f"{' (' + obj + ')' if obj else ''}; Excel calls {label} only from a "
                                    f"{'workbook document module' if host == 'workbook' else 'standard module'}.",
                                    **where))
        if signature == "mismatch":
            findings.append(finding("VBA-PKG-003", f"{proc.name} is declared as {proc.accessor} "
                                    f"{proc.signature()}; Excel calls a Sub with no parameters.", **where))
        for site in result["dynamic"] + result["unresolved"]:
            what = (f"dynamic call {site['call']}" + (f" with literal {site['literal']!r}" if site.get("literal") else "")
                    if "literal" in site else f"unresolved project call {site['call']}")
            findings.append(finding("VBA-PKG-005", f"{label} path: {what}; where it leads is not established.",
                                    module=site["module"], procedure=site["procedure"], line=site["line"],
                                    target=site["call"]))
        for reach in result["reaches"]:
            for earlier in reach.get("earlier_calls_under_goto", []):
                contained = earlier["callee_handler"] == "resume-next"
                findings.append(finding(
                    "VBA-PKG-009", f"{earlier['call']} (line {earlier['line']}) runs before {TARGET} under "
                    f"'On Error GoTo'; an error it does not handle skips {TARGET}. "
                    + ("It starts with its own 'On Error Resume Next'." if contained
                       else "It does not start with its own error handler."),
                    **where, target=earlier["call"], callee_handler=earlier["callee_handler"]))
    usable = [h for h in hooks if h["placement"] != "misplaced" and h["signature_check"] == "correct"]
    reaching = [h for h in usable if h["reaches_target"]]
    unconditional = [h for h in reaching if any(not r["conditional"] for r in h["reaches_target"])]
    opaque = [h for h in usable if h["dynamic_calls"] or h["unresolved_calls"]]
    if not reaching:
        item = finding("VBA-PKG-001", f"No correctly placed startup hook calls {TARGET} in this environment "
                       f"({len(hooks)} lifecycle procedure(s) found).", target=TARGET)
        if opaque:
            item.update(severity="unknown", resolution="dynamic or unresolved calls on a startup path")
        findings.append(item)
    elif not unconditional:
        for hook in reaching:
            findings.append(finding("VBA-PKG-004", f"{hook['procedure']} reaches {TARGET} only on a conditional path.",
                                    module=hook["module"], procedure=hook["procedure"], line=hook["line"]))
    if reaching and all(h["hook"] == "Workbook_AddinInstall" for h in reaching):
        findings.append(finding("VBA-PKG-007", f"Only Workbook_AddinInstall reaches {TARGET}; it runs on install, "
                                "not when the add-in is loaded at startup."))
    verdict = ("wired" if unconditional else "wired-conditionally" if reaching
               else "not-established" if opaque else "not-wired")
    return {"environment": environment, "hooks": hooks, "findings": findings, "failures": failures,
            "verdict": verdict}


def commented_hooks(modules: list[dict[str, Any]]) -> list[tuple[str, int, str]]:
    pattern = re.compile(r"'.*\b(?:Sub|Function)\s+(Workbook_Open|Workbook_AddinInstall|Auto_Open)\b", re.IGNORECASE)
    found = []
    for item in modules:
        for number, line in enumerate((item["text"] or "").split("\n"), start=1):
            match = pattern.search(line)
            if match and not re.match(r"\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\b", line, re.IGNORECASE):
                found.append((item["name"], number, match.group(1)))
    return found


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------


def inspect(package: Path | None, export_dir: Path | None, source: Path, revision: str,
            configuration_name: str, manifest: dict[str, Any]) -> dict[str, Any]:
    configuration = manifest.get("configurations", {}).get(configuration_name)
    if configuration is None:
        raise ValueError(f"configuration {configuration_name!r} is not in {MANIFEST}")
    patterns = [re.compile(p) for p in manifest.get("project_name_patterns", [])]
    extracted = extract_package(package) if package is not None else read_export(export_dir)
    evidence = "package-extraction" if package is not None else "manual-export"
    failures = list(extracted["failures"])
    findings: list[dict[str, Any]] = []
    if evidence == "manual-export":
        findings.append(finding("VBA-PKG-092", "Modules were read from a manual VBE export; nothing here is verified "
                                "against a named package's bytes."))
    sourced = read_source(source, revision, configuration)
    failures += sourced["failures"]
    rows, comparison_findings = compare_modules(extracted, sourced) if extracted["modules"] else ([], [])
    findings += comparison_findings
    environments: dict[str, Any] = {}
    startup_findings: list[dict[str, Any]] = []
    extraction_failed = any(f["code"] == "VBA-PKG-090" for f in failures)
    if extracted["modules"] and not extraction_failed:
        onload = (extracted["ribbon"] or {}).get("onLoad")
        runs = [analyze_environment(extracted["modules"], env, patterns, onload) for env in ENVIRONMENTS]
        for run in runs:
            environments[run["environment"]] = {"verdict": run["verdict"], "hooks": run["hooks"]}
        startup_findings = merge_environments(runs)
        failures += [f for f in startup_findings if f["severity"] == "failure"]
        startup_findings = [f for f in startup_findings if f["severity"] != "failure"]
        for name, line, hook in commented_hooks(extracted["modules"]):
            if not any(h["hook"] == hook for run in runs for h in run["hooks"]):
                findings.append(finding("VBA-PKG-006", f"{hook} appears only in a comment.", module=name, line=line,
                                        target=hook))
    analysis_failed = bool(failures)
    if analysis_failed:
        for item in startup_findings:
            if item["code"] == "VBA-PKG-001":
                item.update(severity="unknown", resolution="analysis-incomplete")
    findings += startup_findings
    verdicts = {run["verdict"] for run in environments.values()}
    if not environments:
        startup_verdict = "not-established"
    elif len(verdicts) > 1:
        startup_verdict = "configuration-dependent"
    else:
        startup_verdict = verdicts.pop()
        if startup_verdict == "not-wired" and analysis_failed:
            startup_verdict = "not-established"
    if extraction_failed or not extracted["modules"]:
        inspection = "failed"
    elif analysis_failed or evidence == "manual-export":
        inspection = "incomplete"
    else:
        inspection = "complete"
    severities = [f["severity"] for f in findings]
    outcome = ("not-established" if inspection == "failed" else
               "findings" if any(s in {"error", "warning"} for s in severities) else "no-defects")
    status = ("incomplete" if inspection != "complete" else "fail" if "error" in severities else "pass")
    tool_revision = git_text(Path(__file__).resolve().parent, "rev-parse", "HEAD")
    skipped = ["form designer (.frx) data and document objects are not compared",
               "runtime behaviour is not inspected: run the package in a fresh Excel session for that"]
    if extraction_failed:
        skipped.append("startup analysis and source comparison skipped: extraction failed")
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "tool_version": TOOL_VERSION,
        "tool_revision": tool_revision.stdout.strip() or None if not tool_revision.returncode else None,
        "advisory": True,
        "status": status,
        "inspection": {
            "status": inspection,
            "evidence": evidence,
            "coverage": {
                "modules_in_package": len(extracted["modules"]),
                "modules_extracted": sum(1 for m in extracted["modules"] if m["text"] is not None),
                "expected_source_modules": len(configuration.get("modules", [])),
                "source_modules_read": len(sourced["modules"]),
                "environments": list(ENVIRONMENTS) if environments else [],
            },
            "failures": failures,
            "skipped_checks": skipped,
            "unsupported": UNSUPPORTED,
        },
        "outcome": {"status": outcome, "counts": {s: severities.count(s) for s in sorted(set(severities))}},
        "package": None if evidence == "manual-export" else {
            "filename": extracted["filename"], "size": extracted["size"], "sha256": extracted["sha256"],
            "vba_project": extracted["vba_project"], "codepage": extracted["codepage"], "ribbon": extracted["ribbon"]},
        "manual_export": None if evidence != "manual-export" else {"directory": export_dir.name},
        "source": {"revision_requested": sourced["revision_requested"], "revision": sourced["revision"],
                   "configuration": configuration_name, "manifest": MANIFEST},
        "modules": rows,
        "startup": {"target": TARGET, "verdict": startup_verdict, "environments": environments},
        "findings": sorted(findings, key=lambda f: (["error", "warning", "unknown", "info"].index(f["severity"])
                                                    if f["severity"] in {"error", "warning", "unknown", "info"} else 4,
                                                    f["code"], f.get("module") or "", f.get("line") or 0)),
        "normalization": NORMALIZATION,
        "evidence_note": ("The package hash identifies the bytes inspected; the source comparison relates embedded "
                          "code to one revision; neither is runtime evidence. Whether Excel runs a hook in a given "
                          "load mode can only be shown in Excel."),
    }


def markdown_report(report: dict[str, Any]) -> str:
    inspection, package, source = report["inspection"], report["package"], report["source"]
    lines = [f"## {TOOL_NAME}", "",
             f"- **Status:** {report['status'].upper()} (advisory; not a release gate)",
             f"- **Inspection:** {inspection['status']} ({inspection['evidence']}); "
             f"**findings outcome:** {report['outcome']['status']}",
             f"- **Startup wiring to `{report['startup']['target']}`:** {report['startup']['verdict']}"]
    if package:
        lines.append(f"- **Package:** `{package['filename']}`, {package['size']} bytes, SHA-256 `{package['sha256']}`")
    else:
        lines.append(f"- **Manual export:** `{report['manual_export']['directory']}` (not package evidence)")
    lines.append(f"- **Source:** `{source['revision'] or 'unresolved'}` (requested `{source['revision_requested']}`), "
                 f"configuration `{source['configuration']}`")
    lines.append(f"- **Tool:** version {report['tool_version']}, revision `{report['tool_revision'] or 'unknown'}`")
    coverage = inspection["coverage"]
    lines.append(f"- **Coverage:** {coverage['modules_extracted']} of {coverage['modules_in_package']} package modules "
                 f"extracted; {coverage['source_modules_read']} of {coverage['expected_source_modules']} source modules "
                 f"read; environments: {', '.join(coverage['environments']) or 'none analyzed'}")
    for failure in inspection["failures"]:
        lines.append(f"- **Failure:** {failure['code']} {failure.get('module') or ''} {failure['message']}".replace("  ", " "))
    if report["modules"]:
        lines += ["", "| Module | Kind | Lines | Source | Comparison |", "| --- | --- | --- | --- | --- |"]
        for row in report["modules"]:
            kind = row["kind"] + (f" ({row['document_object']})" if row["document_object"] else "")
            lines.append(f"| {row['name']} | {kind} | {row['lines'] or ''} | "
                         f"{(row['source'] or {}).get('path', '')} | {row['comparison']} |")
    envs = report["startup"]["environments"]
    if envs:
        first = next(iter(envs.values()))
        lines += ["", "**Startup hooks** (" + ", ".join(f"{k}: {v['verdict']}" for k, v in envs.items()) + ")", ""]
        if not first["hooks"]:
            lines.append("- none found")
        for hook in first["hooks"]:
            lines.append(f"- `{hook['module']}.{hook['signature']}` line {hook['line']}: {hook['hook']}, placement "
                         f"{hook['placement']}, signature {hook['signature_check']}")
            for reach in hook["reaches_target"]:
                lines.append(f"  - reaches `{report['startup']['target']}` "
                             f"({'conditional' if reach['conditional'] else 'unconditional by structure'}"
                             f"{', under ' + HANDLER_TEXT[reach['handler']] if reach['handler'] else ''}): "
                             + " / ".join(reach["path"]))
    if report["findings"]:
        lines += ["", "| Severity | Code | Location | Finding |", "| --- | --- | --- | --- |"]
        for item in report["findings"]:
            where = f"{item.get('module') or ''}{('.' + item['procedure']) if item.get('procedure') else ''}" \
                    f"{(':' + str(item['line'])) if item.get('line') else ''}"
            envs_note = f" [only {', '.join(item['environments'])}]" if item.get("configuration_dependent") else ""
            lines.append(f"| {item['severity']} | {item['code']} | {where} | "
                         f"{item['message'].replace('|', '/')}{envs_note} |")
    lines += ["", report["evidence_note"], "",
              "Skipped: " + "; ".join(inspection["skipped_checks"]) + ".",
              "Unsupported: " + "; ".join(inspection["unsupported"]) + "."]
    return "\n".join(lines) + "\n"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    target = parser.add_mutually_exclusive_group(required=True)
    target.add_argument("--package", type=Path, help="built .xlam or .xlsm to inspect (read-only)")
    target.add_argument("--export-dir", type=Path, help="fallback: folder of modules exported from the VBE")
    parser.add_argument("--source", type=Path, default=Path.cwd(), help="git checkout holding the intended source")
    parser.add_argument("--source-rev", required=True, help="revision to compare against, e.g. v1.2.2 or a SHA")
    parser.add_argument("--configuration", default="packaged-hosts", help=f"configuration in {MANIFEST}")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    options = parser.parse_args(sys.argv[1:] if argv is None else argv)
    options.self_test = False
    manifest_path = Path(__file__).resolve().parent.parent / MANIFEST
    return run_gate(
        options,
        build=lambda: inspect(options.package, options.export_dir, options.source, options.source_rev,
                              options.configuration, json.loads(manifest_path.read_text(encoding="utf-8"))),
        markdown=markdown_report,
        errors=(OSError, UnicodeError, ValueError, subprocess.SubprocessError),
    )


if __name__ == "__main__":
    raise SystemExit(main())
