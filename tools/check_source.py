#!/usr/bin/env python3
"""Check DateTimePicker exports and Ribbon references without running Office."""
from __future__ import annotations

import re
import sys
from pathlib import Path
from xml.parsers import expat

from _gatelib import parse_report_args, run_gate, tracked_files
from check_vba_conditionals import strip_vba

REQUIRED = {
    "src/modules/M_DatePicker.bas", "src/classes/cDatePickerManager.cls",
    "src/classes/cDatePickerLabelHook.cls", "src/forms/UF_DatePicker.frm",
    "src/forms/UF_DatePicker.frx", "src/ribbon/customUI.xml",
    "test/M_cDP_Test.bas", "demo/M_DEMO_BUILDER.bas", "demo/M_DP_DEMO.bas",
}


def ribbon_callbacks(data: bytes) -> list[str]:
    """Parse bounded XML with DTDs forbidden, collecting callback attributes."""
    if len(data) > 1_000_000:
        raise ValueError("Ribbon XML exceeds 1 MB")
    parser = expat.ParserCreate()
    callbacks: list[str] = []

    def reject_dtd(*args: object) -> None:
        raise ValueError("Ribbon DTD declarations are forbidden")

    def element(name: str, attributes: dict[str, str]) -> None:
        for key, value in attributes.items():
            if key.startswith("on") or key.startswith("get") or key == "loadImage":
                callbacks.append(value)

    parser.StartDoctypeDeclHandler = reject_dtd
    parser.StartElementHandler = element
    parser.Parse(data, True)
    return callbacks


def run_check(root: Path) -> dict:
    tracked = tracked_files(root)
    findings = [f"Missing tracked source: {p}" for p in sorted(REQUIRED - tracked)]
    names: dict[str, str] = {}
    public_procedures: set[str] = set()
    paths = sorted(p for p in tracked if Path(p).suffix.lower() in {".bas", ".cls", ".frm"})
    for path in paths:
        text = (root / path).read_bytes().decode("cp1252")
        matches = re.findall(r'^Attribute VB_Name = "([^"]+)"\s*$', text, re.M)
        if matches != [Path(path).stem]:
            findings.append(f"{path}: VB_Name must match the filename exactly")
        elif matches[0].casefold() in names:
            findings.append(f"{path}: duplicate component name")
        else:
            names[matches[0].casefold()] = path
        if not any(re.fullmatch(r"\s*Option Explicit\s*", strip_vba(line), re.I)
                   for line in text.splitlines()):
            findings.append(f"{path}: missing Option Explicit")
        if path.endswith(".bas") and path.startswith("src/"):
            public_procedures.update(name.casefold() for name in re.findall(
                r"^\s*Public\s+(?:Sub|Function)\s+(\w+)", text, re.M | re.I))
        if path.endswith(".frm"):
            companions = re.findall(r'"([^"\r\n]+\.frx)":([0-9A-Fa-f]+)', text)
            if not companions:
                findings.append(f"{path}: no form resource reference")
            for filename, offset in companions:
                companion = Path(path).parent / filename
                if Path(filename).name != filename or companion.as_posix() not in tracked:
                    findings.append(f"{path}: missing or unsafe form resource {filename}")
                elif (root / companion).stat().st_size <= int(offset, 16):
                    findings.append(f"{path}: resource offset is outside {filename}")
    ribbon = root / "src/ribbon/customUI.xml"
    if ribbon.is_file():
        try:
            for callback in ribbon_callbacks(ribbon.read_bytes()):
                if callback.casefold() not in public_procedures:
                    findings.append(f"Ribbon callback has no production public procedure: {callback}")
        except (ValueError, expat.ExpatError) as error:
            findings.append(f"Invalid Ribbon XML: {error}")
    version = (root / "VERSION").read_text().strip()
    if not re.fullmatch(r"(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)", version):
        findings.append("VERSION must contain MAJOR.MINOR.PATCH")
    return {"schema_version": 1, "status": "fail" if findings else "pass",
            "components": len(paths), "findings": findings}


def main() -> int:
    options = parse_report_args(sys.argv[1:])
    return run_gate(options, build=lambda: run_check(options.root),
                    markdown=lambda r: "Source integrity: " + r["status"].upper() + "\n"
                    + "\n".join(r["findings"]), errors=(OSError, RuntimeError, ValueError))


if __name__ == "__main__":
    raise SystemExit(main())
