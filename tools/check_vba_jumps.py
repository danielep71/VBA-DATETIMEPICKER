#!/usr/bin/env python3
"""Validate that VBA jump targets resolve inside the owning procedure."""

from __future__ import annotations

import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

from _gatelib import git_bytes as git
from _gatelib import parse_report_args as parse_args
from _gatelib import run_gate
from check_vba_conditionals import reachable_sources

VBA_SUFFIXES = {".bas", ".cls", ".frm"}
TOOL_NAME = "Procedure-scoped VBA jumps"

PROC_OPEN = re.compile(
    r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\b",
    re.IGNORECASE,
)
PROC_CLOSE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
NAMED_LABEL = re.compile(r"^\s*([A-Za-z_]\w*|\d+):")
NUMBERED_LABEL = re.compile(r"^\s*(\d+)(?=\s+\S)")
JUMP = re.compile(
    r"\b(GoTo|GoSub|Resume)\b(?:\s+([A-Za-z_]\w*|\d+|-1))?",
    re.IGNORECASE,
)




def tracked_vba(root: Path) -> list[str]:
    completed = git(root, "ls-files", "-z")
    if completed.returncode != 0:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
    paths: list[str] = []
    for item in completed.stdout.split(b"\0"):
        if not item:
            continue
        relative = item.decode("utf-8", errors="surrogateescape")
        if Path(relative).suffix.casefold() in VBA_SUFFIXES:
            paths.append(relative)
    return sorted(paths)


def strip_vba(raw: str) -> str:
    result: list[str] = []
    in_string = False
    index = 0
    while index < len(raw):
        char = raw[index]
        if char == '"':
            if in_string and index + 1 < len(raw) and raw[index + 1] == '"':
                result.extend('  ')
                index += 2
                continue
            in_string = not in_string
            result.append(' ')
            index += 1
            continue
        if char == "'" and not in_string:
            break
        result.append(' ' if in_string else char)
        index += 1
    text = "".join(result)
    if re.match(r"^\s*Rem(?:\s|$)", text, re.IGNORECASE):
        return ""
    return text.rstrip()


def logical_statements(lines: list[str]) -> list[tuple[int, int, str]]:
    statements: list[tuple[int, int, str]] = []
    buffer: list[str] = []
    start = 0
    for number, raw in enumerate(lines, start=1):
        code = strip_vba(raw)
        if not buffer:
            start = number
        continued = bool(re.search(r"\s_\s*$", code))
        if continued:
            code = re.sub(r"\s_\s*$", " ", code)
            buffer.append(code)
            continue
        buffer.append(code)
        statements.append((start, number, " ".join(part.strip() for part in buffer)))
        buffer.clear()
    if buffer:
        statements.append((start, len(lines), " ".join(part.strip() for part in buffer)))
    return statements


def label_at_start(code: str) -> str | None:
    match = NAMED_LABEL.match(code)
    if match:
        return match.group(1)
    match = NUMBERED_LABEL.match(code)
    return match.group(1) if match else None


def _statement_jumps(code: str) -> list[tuple[str, str]]:
    """Labelled jump targets in one statement, as ``(operation, target)``.

    ``Resume``/``Resume Next`` and ``On Error GoTo 0``/``-1`` name no label and are skipped.
    """
    jumps: list[tuple[str, str]] = []
    for match in JUMP.finditer(code):
        operation = match.group(1).casefold()
        target = match.group(2)
        prefix = code[: match.start()].casefold()
        if operation == "resume" and (target is None or target.casefold() == "next"):
            continue
        if operation == "goto" and target in {"0", "-1"} and re.search(
            r"\bon\s+error\s*$", prefix
        ):
            continue
        if target is None:
            continue
        jumps.append((match.group(1), target))
    return jumps


def _record_label(
    path: str, current: dict[str, Any], code: str, start_line: int, findings: list[dict[str, Any]]
) -> None:
    """Register a statement's leading label in its procedure; report a duplicate."""
    label = label_at_start(code)
    if label is None:
        return
    key = label.casefold()
    labels = current["labels"]
    if key in labels:
        findings.append(
            {
                "path": path,
                "procedure": current["name"],
                "line": start_line,
                "target": label,
                "message": (
                    f"Duplicate label {label!r} in procedure {current['name']}; "
                    f"first defined at line {labels[key]}."
                ),
            }
        )
    else:
        labels[key] = start_line


def _unresolved_jump_findings(path: str, procedures: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Jumps whose target label is not defined in the same procedure."""
    findings: list[dict[str, Any]] = []
    for procedure in procedures:
        labels = procedure["labels"]
        for start_line, end_line, operation, target in procedure["jumps"]:
            if target.casefold() not in labels:
                findings.append(
                    {
                        "path": path,
                        "procedure": procedure["name"],
                        "line": start_line,
                        "end_line": end_line,
                        "operation": operation,
                        "target": target,
                        "message": (
                            f"{operation} target {target!r} is not defined in "
                            f"procedure {procedure['name']}."
                        ),
                    }
                )
    return findings


def analyze_component(path: str, text: str) -> list[dict[str, Any]]:
    findings: list[dict[str, Any]] = []
    procedures: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None

    for start_line, end_line, code in logical_statements(text.splitlines()):
        stripped = code.strip()
        if not stripped or stripped.startswith("#"):
            continue

        opener = PROC_OPEN.match(code)
        if opener and " declare " not in f" {code.casefold()} ":
            if current is not None:
                findings.append(
                    {
                        "path": path,
                        "procedure": current["name"],
                        "line": start_line,
                        "message": (
                            f"Procedure {current['name']} opened at line "
                            f"{current['start_line']} has no closing statement before "
                            f"{opener.group(2)}."
                        ),
                    }
                )
            current = {
                "name": opener.group(2),
                "kind": opener.group(1),
                "start_line": start_line,
                "labels": {},
                "jumps": [],
            }
            procedures.append(current)
            continue

        if PROC_CLOSE.match(code):
            if current is None:
                findings.append(
                    {
                        "path": path,
                        "procedure": None,
                        "line": start_line,
                        "message": "Procedure closing statement has no owning procedure.",
                    }
                )
            current = None
            continue

        if current is None:
            continue

        _record_label(path, current, code, start_line, findings)
        current["jumps"].extend(
            (start_line, end_line, operation, target) for operation, target in _statement_jumps(code)
        )

    if current is not None:
        findings.append(
            {
                "path": path,
                "procedure": current["name"],
                "line": current["start_line"],
                "message": f"Procedure {current['name']} is not closed.",
            }
        )

    findings.extend(_unresolved_jump_findings(path, procedures))
    return findings


def run_check(root: Path) -> dict[str, Any]:
    paths = tracked_vba(root)
    findings: list[dict[str, Any]] = []
    procedure_count = 0
    for relative in paths:
        data = (root / relative).read_bytes()
        text = data.decode("cp1252")
        sources, conditional_findings = reachable_sources(relative, text)
        findings.extend(conditional_findings)
        for environment, source in sources.items():
            for finding in analyze_component(relative, source):
                findings.append({**finding, "environment": environment})
        procedure_count += sum(
            bool(PROC_OPEN.match(code))
            for _, _, code in logical_statements(text.splitlines())
            if " declare " not in f" {code.casefold()} "
        )
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not findings else "fail",
        "components": len(paths),
        "procedures": procedure_count,
        "findings": findings,
    }


def markdown_report(report: dict[str, Any]) -> str:
    lines = [
        "## Procedure-scoped VBA jumps",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Components:** {report['components']}",
        f"- **Procedures:** {report['procedures']}",
        f"- **Findings:** {len(report['findings'])}",
    ]
    if report["findings"]:
        lines.extend(
            [
                "",
                "| Path | Procedure | Line | Target | Finding |",
                "| --- | --- | ---: | --- | --- |",
            ]
        )
        for item in report["findings"]:
            lines.append(
                "| {path} | {procedure} | {line} | {target} | {message} |".format(
                    path=str(item.get("path", "")).replace("|", "\\|"),
                    procedure=str(item.get("procedure") or "—").replace("|", "\\|"),
                    line=item.get("line", "—"),
                    target=str(item.get("target") or "—").replace("|", "\\|"),
                    message=str(item.get("message", "")).replace("|", "\\|"),
                )
            )
    return "\n".join(lines) + "\n"




def init_repo(root: Path, source: str) -> None:
    (root / "src").mkdir()
    (root / "src" / "Fixture.bas").write_bytes(source.replace("\n", "\r\n").encode("cp1252"))
    commands = (
        ("init", "-b", "main"),
        ("config", "user.name", "VBA Jump Self-Test"),
        ("config", "user.email", "vba-jump@example.invalid"),
        ("add", "src/Fixture.bas"),
    )
    for command in commands:
        completed = subprocess.run(
            ["git", "-C", str(root), *command],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        if completed.returncode != 0:
            raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())


def fixture_result(source: str) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="vba-jump-") as temporary:
        root = Path(temporary)
        init_repo(root, source)
        return run_check(root)


def run_self_test() -> int:
    fixtures = {
        "valid-local": (
            "pass",
            '''Attribute VB_Name = "Fixture"
Option Explicit

Public Sub First( _
    ByVal Value As Long)
    On Error GoTo Handler
    GoSub Worker
    Resume 100
Worker:
    Return
100: Exit Sub
Handler:
    Resume Next
End Sub
''',
        ),
        "cross-procedure": (
            "fail",
            '''Attribute VB_Name = "Fixture"
Option Explicit

Public Sub First()
    GoTo Shared
End Sub

Public Sub Second()
Shared:
    Exit Sub
End Sub
''',
        ),
        "equivalent-separate": (
            "pass",
            '''Attribute VB_Name = "Fixture"
Option Explicit

Public Sub First()
Shared:
    Exit Sub
End Sub

Public Sub Second()
Shared:
    Exit Sub
End Sub
''',
        ),
        "duplicate-local": (
            "fail",
            '''Attribute VB_Name = "Fixture"
Option Explicit

Public Sub First()
Again:
Again:
    Exit Sub
End Sub
''',
        ),
        "numbered-label": (
            "pass",
            '''Attribute VB_Name = "Fixture"
Option Explicit

Public Sub First()
    GoTo 100
100 Exit Sub
End Sub
''',
        ),
        "special-error-controls": (
            "pass",
            '''Attribute VB_Name = "Fixture"
Option Explicit

Public Sub First()
    On Error GoTo 0
    On Error GoTo -1
    On Error Resume Next
    Resume Next
End Sub
''',
        ),
    }
    fixtures["quoted-jumps"] = ("pass", '''Public Sub First()
    Debug.Print "GoSub Missing", "GoTo Lost", "Resume Nowhere"
    Debug.Print "He said ""GoSub Missing""; it's text"
    Debug.Print "GoSub Missing" & _
        "Resume Lost"
    Debug.Print "GoSub Missing": GoSub Worker
    Exit Sub
Worker:
    Return
End Sub
''')
    fixtures["real-jump-after-string"] = ("fail", '''Public Sub First()
    Debug.Print "it's ""quoted""": GoSub Missing
End Sub
''')
    failures: list[str] = []
    for name, (expected, source) in fixtures.items():
        report = fixture_result(source)
        if report["status"] != expected:
            failures.append(
                f"{name}: expected {expected}, got {report['status']} "
                f"({report['findings']})"
            )

    cross = fixture_result(fixtures["cross-procedure"][1])
    if not any(
        item.get("procedure") == "First" and item.get("target") == "Shared"
        for item in cross["findings"]
    ):
        failures.append("cross-procedure: diagnostic did not name First and Shared")

    duplicate = fixture_result(fixtures["duplicate-local"][1])
    if not any(
        "first defined at line" in item.get("message", "")
        for item in duplicate["findings"]
    ):
        failures.append("duplicate-local: diagnostic did not identify prior definition")

    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(
        "SELF-TEST PASS: local handlers, GoSub/Resume, cross-procedure rejection, "
        "same-name separate procedures, duplicate labels, numbered labels, "
        "continuations, quoted jump text, real jumps after strings, and special error controls passed."
    )
    return 0




def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    return run_gate(
        options,
        build=lambda: run_check(options.root),
        markdown=markdown_report,
        errors=(OSError, UnicodeError, RuntimeError, subprocess.SubprocessError),
        self_test=run_self_test,
        self_test_error_prefix="SELF-TEST ERROR",
    )

if __name__ == "__main__":
    raise SystemExit(main())
