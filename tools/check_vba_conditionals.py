#!/usr/bin/env python3
"""Validate reachable VBA declarations across supported compilation environments."""

from __future__ import annotations

import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from _gatelib import git_bytes as git
from _gatelib import parse_report_args as parse_args
from _gatelib import run_gate

VBA_SUFFIXES = {".bas", ".cls", ".frm"}
TOOL_NAME = "VBA conditional compilation"
DECLARE_RE = re.compile(
    r"^\s*(?:Public|Private)?\s*Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\b",
    re.IGNORECASE,
)
PTRSAFE_RE = re.compile(r"\bPtrSafe\b", re.IGNORECASE)
DIRECTIVE_RE = re.compile(
    r"^\s*#\s*(If|ElseIf|Else|End\s+If)\b(.*)$",
    re.IGNORECASE,
)
CONST_RE = re.compile(r"^\s*#\s*Const\b", re.IGNORECASE)
TOKEN_RE = re.compile(r"\s*(\(|\)|<>|=|-?\d+|[A-Za-z_]\w*)")

# Microsoft documents Win32=True on both 32-bit and 64-bit development
# platforms. Mac is explicitly False in these Windows-only models. Win64 is additionally True on 64-bit Office. VBA6 is modeled as
# the legacy pre-VBA7 environment; VBA7 is modeled for both supported bitnesses.
ENVIRONMENTS: dict[str, dict[str, bool]] = {
    "vba6-win32": {"VBA6": True, "VBA7": False, "WIN32": True, "WIN64": False, "MAC": False},
    "vba7-win32": {"VBA6": False, "VBA7": True, "WIN32": True, "WIN64": False, "MAC": False},
    "vba7-win64": {"VBA6": False, "VBA7": True, "WIN32": True, "WIN64": True, "MAC": False},
}
SUPPORTED_SYMBOLS = frozenset(next(iter(ENVIRONMENTS.values())))
VBA7_ENVIRONMENTS = tuple(
    name for name, symbols in ENVIRONMENTS.items() if symbols["VBA7"]
)


class ExpressionError(ValueError):
    """Raised when a conditional expression is outside the supported model."""


class ExpressionParser:
    def __init__(self, expression: str, symbols: dict[str, bool]):
        self.expression = expression
        self.symbols = symbols
        self.tokens = self._tokenize(expression)
        self.index = 0

    @staticmethod
    def _tokenize(expression: str) -> list[str]:
        tokens: list[str] = []
        position = 0
        while position < len(expression):
            match = TOKEN_RE.match(expression, position)
            if match is None:
                if expression[position:].strip() == "":
                    break
                raise ExpressionError(
                    f"unsupported token near {expression[position:]!r}"
                )
            tokens.append(match.group(1))
            position = match.end()
        if not tokens:
            raise ExpressionError("empty conditional expression")
        return tokens

    def peek(self) -> str | None:
        return self.tokens[self.index] if self.index < len(self.tokens) else None

    def take(self) -> str:
        token = self.peek()
        if token is None:
            raise ExpressionError("unexpected end of expression")
        self.index += 1
        return token

    def parse(self) -> bool:
        value = self.parse_or()
        if self.peek() is not None:
            raise ExpressionError(f"unexpected token {self.peek()!r}")
        return value

    def parse_or(self) -> bool:
        value = self.parse_and()
        while (self.peek() or "").casefold() == "or":
            self.take()
            right = self.parse_and()
            value = value or right
        return value

    def parse_and(self) -> bool:
        value = self.parse_not()
        while (self.peek() or "").casefold() == "and":
            self.take()
            right = self.parse_not()
            value = value and right
        return value

    def parse_not(self) -> bool:
        if (self.peek() or "").casefold() == "not":
            self.take()
            return not self.parse_not()
        return self.parse_comparison()

    def parse_comparison(self) -> bool:
        left = self.parse_atom()
        operator = self.peek()
        if operator not in {"=", "<>"}:
            return left
        self.take()
        right = self.parse_atom()
        return left == right if operator == "=" else left != right

    def parse_atom(self) -> bool:
        token = self.take()
        if token == "(":
            value = self.parse_or()
            if self.take() != ")":
                raise ExpressionError("missing closing parenthesis")
            return value
        lowered = token.casefold()
        if lowered == "true":
            return True
        if lowered == "false":
            return False
        if re.fullmatch(r"-?\d+", token):
            integer_value = int(token)
            if integer_value not in {-1, 0}:
                raise ExpressionError(
                    "only VBA Boolean integer literals -1 and 0 are supported"
                )
            return integer_value == -1
        upper = token.upper()
        if upper not in SUPPORTED_SYMBOLS:
            raise ExpressionError(
                f"unsupported conditional-compilation symbol {token!r}; "
                "supported symbols are VBA6, VBA7, Win32, and Win64"
            )
        return self.symbols[upper]


def evaluate(expression: str, symbols: dict[str, bool]) -> bool:
    return ExpressionParser(expression, symbols).parse()


@dataclass
class Frame:
    parent_active: bool
    branch_taken: bool
    current_active: bool
    else_seen: bool = False


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
                result.extend(('"', '"'))
                index += 2
                continue
            in_string = not in_string
            result.append(char)
            index += 1
            continue
        if char == "'" and not in_string:
            break
        result.append(char)
        index += 1
    text = "".join(result)
    if re.match(r"^\s*Rem(?:\s|$)", text, re.IGNORECASE):
        return ""
    return text.rstrip()


def logical_units(lines: list[str]) -> list[tuple[int, int, str, str]]:
    units: list[tuple[int, int, str, str]] = []
    buffer: list[str] = []
    start = 0
    for number, raw in enumerate(lines, start=1):
        code = strip_vba(raw)
        stripped = code.strip()
        if not buffer and stripped.startswith("#"):
            units.append((number, number, "directive", stripped))
            continue
        if not buffer:
            start = number
        if re.search(r"\s_\s*$", code):
            buffer.append(re.sub(r"\s_\s*$", " ", code))
            continue
        buffer.append(code)
        units.append((start, number, "code", " ".join(part.strip() for part in buffer)))
        buffer.clear()
    if buffer:
        units.append((start, len(lines), "code", " ".join(part.strip() for part in buffer)))
    return units


def active(stack: list[Frame]) -> bool:
    return all(frame.current_active for frame in stack)


def parse_condition(kind: str, remainder: str) -> str:
    match = re.match(r"^(.*?)\s+Then\s*$", remainder.strip(), re.IGNORECASE)
    if match is None or not match.group(1).strip():
        raise ExpressionError(f"#{kind} must end with Then and contain a condition")
    return match.group(1).strip()


def _condition_values(
    path: str,
    line: int,
    label: str,
    remainder: str,
    findings: list[dict[str, Any]],
) -> dict[str, bool]:
    try:
        expression = parse_condition(label, remainder)
        return {
            name: evaluate(expression, symbols)
            for name, symbols in ENVIRONMENTS.items()
        }
    except ExpressionError as error:
        findings.append(
            {
                "path": path,
                "line": line,
                "message": f"Indeterminate #{label} directive: {error}.",
            }
        )
        return {name: False for name in ENVIRONMENTS}


def _handle_if(
    path: str,
    line: int,
    remainder: str,
    stacks: dict[str, list[Frame]],
    findings: list[dict[str, Any]],
) -> None:
    values = _condition_values(path, line, "If", remainder, findings)
    for name, stack in stacks.items():
        parent = active(stack)
        selected = parent and values[name]
        stack.append(Frame(parent, selected, selected))


def _handle_elseif(
    path: str,
    line: int,
    remainder: str,
    stacks: dict[str, list[Frame]],
    findings: list[dict[str, Any]],
) -> None:
    if any(not stack for stack in stacks.values()):
        findings.append({"path": path, "line": line, "message": "#ElseIf without #If."})
        return
    if any(stack[-1].else_seen for stack in stacks.values()):
        findings.append({"path": path, "line": line, "message": "#ElseIf after #Else."})
        return
    values = _condition_values(path, line, "ElseIf", remainder, findings)
    for name, stack in stacks.items():
        frame = stack[-1]
        selected = frame.parent_active and not frame.branch_taken and values[name]
        frame.current_active = selected
        frame.branch_taken = frame.branch_taken or selected


def _handle_else(
    path: str,
    line: int,
    remainder: str,
    stacks: dict[str, list[Frame]],
    findings: list[dict[str, Any]],
) -> None:
    if remainder.strip():
        findings.append(
            {"path": path, "line": line, "message": "#Else must not contain trailing text."}
        )
        return
    if any(not stack for stack in stacks.values()):
        findings.append({"path": path, "line": line, "message": "#Else without #If."})
        return
    if any(stack[-1].else_seen for stack in stacks.values()):
        findings.append({"path": path, "line": line, "message": "Duplicate #Else."})
        return
    for stack in stacks.values():
        frame = stack[-1]
        selected = frame.parent_active and not frame.branch_taken
        frame.current_active = selected
        frame.branch_taken = True
        frame.else_seen = True


def _handle_end_if(
    path: str,
    line: int,
    remainder: str,
    stacks: dict[str, list[Frame]],
    findings: list[dict[str, Any]],
) -> None:
    if remainder.strip():
        findings.append(
            {"path": path, "line": line, "message": "#End If must not contain trailing text."}
        )
        return
    if any(not stack for stack in stacks.values()):
        findings.append({"path": path, "line": line, "message": "#End If without #If."})
        return
    for stack in stacks.values():
        stack.pop()


def _handle_directive(
    path: str,
    line: int,
    code: str,
    stacks: dict[str, list[Frame]],
    findings: list[dict[str, Any]],
) -> None:
    if CONST_RE.match(code):
        findings.append(
            {
                "path": path,
                "line": line,
                "message": (
                    "Project-defined #Const symbols are outside the supported model; "
                    "conditional-compilation evaluation fails closed."
                ),
            }
        )
        return
    directive = DIRECTIVE_RE.match(code)
    if directive is None:
        findings.append(
            {
                "path": path,
                "line": line,
                "message": f"Unsupported conditional-compilation directive: {code}",
            }
        )
        return
    kind = " ".join(directive.group(1).split()).casefold()
    remainder = directive.group(2)
    handlers = {
        "if": _handle_if,
        "elseif": _handle_elseif,
        "else": _handle_else,
        "end if": _handle_end_if,
    }
    handlers[kind](path, line, remainder, stacks, findings)


def _validate_declare(
    path: str,
    start_line: int,
    end_line: int,
    code: str,
    stacks: dict[str, list[Frame]],
    findings: list[dict[str, Any]],
) -> None:
    if not DECLARE_RE.match(code):
        return
    reachable_vba7 = [name for name in VBA7_ENVIRONMENTS if active(stacks[name])]
    if not reachable_vba7 or PTRSAFE_RE.search(code):
        return
    findings.append(
        {
            "path": path,
            "line": start_line,
            "end_line": end_line,
            "environments": reachable_vba7,
            "message": (
                "Declare is reachable under VBA7 without PtrSafe in: "
                + ", ".join(reachable_vba7)
                + "."
            ),
        }
    )


def analyze_component(path: str, text: str) -> list[dict[str, Any]]:
    findings: list[dict[str, Any]] = []
    stacks: dict[str, list[Frame]] = {name: [] for name in ENVIRONMENTS}
    for start_line, end_line, unit_kind, code in logical_units(text.splitlines()):
        if unit_kind == "directive":
            _handle_directive(path, start_line, code, stacks, findings)
        else:
            _validate_declare(path, start_line, end_line, code, stacks, findings)
    depth = max((len(stack) for stack in stacks.values()), default=0)
    if any(stacks.values()):
        findings.append(
            {
                "path": path,
                "line": None,
                "message": f"{depth} conditional-compilation block(s) are unclosed.",
            }
        )
    return findings


def reachable_sources(path: str, text: str) -> tuple[dict[str, str], list[dict[str, Any]]]:
    """Return each modeled source with inactive lines blanked, preserving line numbers."""
    lines = text.splitlines()
    sources = {name: [""] * len(lines) for name in ENVIRONMENTS}
    stacks: dict[str, list[Frame]] = {name: [] for name in ENVIRONMENTS}
    findings: list[dict[str, Any]] = []
    for start, end, kind, code in logical_units(lines):
        if kind == "directive":
            _handle_directive(path, start, code, stacks, findings)
        else:
            for name, stack in stacks.items():
                if active(stack):
                    sources[name][start - 1:end] = lines[start - 1:end]
    if any(stacks.values()):
        findings.append({"path": path, "line": None,
                         "message": "Conditional-compilation block(s) are unclosed."})
    return {name: "\n".join(source) for name, source in sources.items()}, findings


def run_check(root: Path) -> dict[str, Any]:
    paths = tracked_vba(root)
    findings: list[dict[str, Any]] = []
    declare_count = 0
    for relative in paths:
        text = (root / relative).read_bytes().decode("cp1252")
        units = logical_units(text.splitlines())
        declare_count += sum(
            unit_kind == "code" and bool(DECLARE_RE.match(code))
            for _, _, unit_kind, code in units
        )
        findings.extend(analyze_component(relative, text))
    return {
        "schema_version": 1,
        "tool": TOOL_NAME,
        "status": "pass" if not findings else "fail",
        "components": len(paths),
        "declare_statements": declare_count,
        "environments": ENVIRONMENTS,
        "findings": findings,
    }


def markdown_report(report: dict[str, Any]) -> str:
    lines = [
        "## VBA conditional compilation",
        "",
        f"- **Status:** {str(report['status']).upper()}",
        f"- **Components:** {report['components']}",
        f"- **Declare statements:** {report['declare_statements']}",
        "- **Supported environments:** `vba6-win32`, `vba7-win32`, `vba7-win64`",
        "- **Win64 compatibility:** `Win32=True` and `Win64=True` in `vba7-win64`",
        f"- **Findings:** {len(report['findings'])}",
    ]
    if report["findings"]:
        lines.extend(
            [
                "",
                "| Path | Line | Environments | Finding |",
                "| --- | ---: | --- | --- |",
            ]
        )
        for item in report["findings"]:
            environments = ", ".join(item.get("environments", [])) or "—"
            lines.append(
                "| {path} | {line} | {environments} | {message} |".format(
                    path=str(item.get("path", "")).replace("|", "\\|"),
                    line=item.get("line") or "—",
                    environments=environments.replace("|", "\\|"),
                    message=str(item.get("message", "")).replace("|", "\\|"),
                )
            )
    return "\n".join(lines) + "\n"


def fixture_result(source: str) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix="vba-conditional-") as temporary:
        root = Path(temporary)
        (root / "src").mkdir()
        (root / "src" / "Fixture.bas").write_bytes(
            source.replace("\n", "\r\n").encode("cp1252")
        )
        for command in (
            ("init", "-b", "main"),
            ("config", "user.name", "VBA Conditional Self-Test"),
            ("config", "user.email", "vba-conditional@example.invalid"),
            ("add", "src/Fixture.bas"),
        ):
            completed = git(root, *command)
            if completed.returncode != 0:
                raise RuntimeError(
                    completed.stderr.decode("utf-8", errors="replace").strip()
                )
        return run_check(root)


def run_self_test() -> int:
    fixtures: dict[str, tuple[str, str]] = {
        "nested-valid": (
            "pass",
            '''Attribute VB_Name = "Fixture"
Option Explicit
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function F64 Lib "kernel32" () As LongPtr
    #ElseIf Win32 Then
        Private Declare PtrSafe Function F32 Lib "kernel32" () As Long
    #End If
#Else
    Private Declare Function FLegacy Lib "kernel32" () As Long
#End If
''',
        ),
        "reachable-vba7-nonptrsafe": (
            "fail",
            '''Attribute VB_Name = "Fixture"
Option Explicit
#If VBA7 And Win64 Then
    Private Declare Function Bad Lib "kernel32" () As Long
#End If
''',
        ),
        "legacy-allowed": (
            "pass",
            '''Attribute VB_Name = "Fixture"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#Else
    Private Declare Function Legacy Lib "kernel32" () As Long
#End If
''',
        ),
        "inactive-nested": (
            "pass",
            '''Attribute VB_Name = "Fixture"
Option Explicit
#If VBA6 Then
    #If Win64 Then
        Private Declare Function ImpossibleLegacy64 Lib "kernel32" () As Long
    #End If
#End If
''',
        ),
        "win32-before-win64": (
            "fail",
            '''Attribute VB_Name = "Fixture"
Option Explicit
#If Win32 Then
    Private Declare Function WrongOrder Lib "kernel32" () As Long
#ElseIf Win64 Then
    Private Declare PtrSafe Function NeverReached Lib "kernel32" () As LongPtr
#End If
''',
        ),
        "boolean-expression": (
            "pass",
            '''Attribute VB_Name = "Fixture"
Option Explicit
#If Not VBA7 Then
    Private Declare Function Legacy Lib "kernel32" () As Long
#ElseIf VBA7 And (Win64 Or Win32) Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#End If
''',
        ),
        "continued-declare": (
            "pass",
            '''Attribute VB_Name = "Fixture"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function Modern Lib "kernel32" _
        () As LongPtr
#Else
    Private Declare Function Legacy Lib "kernel32" _
        () As Long
#End If
''',
        ),
        "indeterminate-symbol": (
            "fail",
            '''Attribute VB_Name = "Fixture"
Option Explicit
#If UnknownFeature Then
    Private Declare Function Unknown Lib "kernel32" () As Long
#End If
''',
        ),
        "project-const": (
            "fail",
            '''Attribute VB_Name = "Fixture"
Option Explicit
#Const FEATURE = True
''',
        ),
        "unbalanced": (
            "fail",
            '''Attribute VB_Name = "Fixture"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
''',
        ),
        "stray-elseif": (
            "fail",
            '''Attribute VB_Name = "Fixture"
Option Explicit
#ElseIf VBA7 Then
''',
        ),
    }
    failures: list[str] = []
    reports: dict[str, dict[str, Any]] = {}
    for name, (expected, source) in fixtures.items():
        report = fixture_result(source)
        reports[name] = report
        if report["status"] != expected:
            failures.append(
                f"{name}: expected {expected}, got {report['status']} ({report['findings']})"
            )

    reachable = reports["reachable-vba7-nonptrsafe"]["findings"]
    if not any("vba7-win64" in item.get("environments", []) for item in reachable):
        failures.append("reachable-vba7-nonptrsafe: diagnostic omitted vba7-win64")

    ordering = reports["win32-before-win64"]["findings"]
    if not any("vba7-win64" in item.get("environments", []) for item in ordering):
        failures.append("win32-before-win64: Win32=True on Win64 was not modeled")

    indeterminate = reports["indeterminate-symbol"]["findings"]
    if not any(
        "unsupported conditional-compilation symbol" in item["message"]
        for item in indeterminate
    ):
        failures.append(
            "indeterminate-symbol: actionable unsupported-symbol diagnostic missing"
        )

    if failures:
        for failure in failures:
            print(f"[FAIL] {failure}")
        print(f"SELF-TEST FAIL: {len(failures)} failure(s).")
        return 1
    print(
        "SELF-TEST PASS: nested VBA6/VBA7 and Win32/Win64 branches, ElseIf selection, "
        "PtrSafe reachability, inactive branches, Win32-on-Win64 semantics, boolean "
        "expressions, continued declares, indeterminate symbols, #Const rejection, and "
        "unbalanced directives passed."
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    options = parse_args(sys.argv[1:] if argv is None else argv)
    return run_gate(
        options,
        build=lambda: run_check(options.root),
        markdown=markdown_report,
        errors=(
            OSError,
            UnicodeError,
            RuntimeError,
            subprocess.SubprocessError,
            ExpressionError,
        ),
        self_test=run_self_test,
        self_test_error_prefix="SELF-TEST ERROR",
    )

if __name__ == "__main__":
    raise SystemExit(main())
