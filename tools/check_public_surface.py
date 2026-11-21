#!/usr/bin/env python3
"""Check the public VBA surface against src/public-surface.json without running Office.

Every technically public declaration in a production standard module must be
classified exactly once as ``supported``, ``callback`` or ``internal``. The
manifest also lists each symbol's invocation channels. The channels this gate
can see in the source are recomputed and must match the manifest, so a new
caller, a new macro entry point or an orphaned callback fails until the
manifest records it.

Computed channels:

- ``macro-dialog``: parameterless public ``Sub`` in a module without
  ``Option Private Module`` (Alt+F8 and assigned controls).
- ``ribbon-onaction``: named by an attribute in ``src/ribbon/customUI.xml``.
- ``onaction`` / ``onkey`` / ``ontime``: dispatched by name from production
  source. The gate proves the name is dispatched; the manifest says how.
- ``project-internal``: referenced by another production component.
- ``test-harness``: referenced by ``test/*.bas``.
- ``demo``: referenced by ``demo/*.bas``.

Declared channels: ``direct-vba`` (required on, and only on, supported
symbols) and ``operator`` (an internal helper documented for manual recovery).

Facade rules (#55):

- The facade holds only public enums, public types and public wrappers. It has
  no ``Declare``, constant, variable or private member.
- Each wrapper is one statement that delegates to ``<Module>.<Name>_Impl`` in
  a production module with ``Option Private Module``.
- No other production component calls a facade wrapper. Internal code calls
  the ``_Impl`` procedure; callback names reach Excel only as strings. The
  add-in's own workbook module (``src/addin/``) is a host and may call the
  supported API.
"""
from __future__ import annotations

import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from xml.parsers import expat

from _gatelib import parse_report_args, run_gate, tracked_files
from check_source import ribbon_callbacks
from check_vba_conditionals import logical_units, reachable_sources, strip_vba

MANIFEST = "src/public-surface.json"
RIBBON = "src/ribbon/customUI.xml"
README = "README.md"
CLASSIFICATIONS = ("supported", "callback", "internal")
BY_NAME = ("onaction", "onkey", "ontime")
CALLBACK_CHANNELS = ("ribbon-onaction", *BY_NAME)
COMPUTED = ("macro-dialog", "ribbon-onaction", "project-internal", "test-harness", "demo")
DECLARED = ("direct-vba", "operator", *BY_NAME)
CHANNELS = (*COMPUTED, *DECLARED)
# Production components that host the add-in and consume the supported API.
HOSTS = ("src/addin/",)
# Production helpers that qualify a procedure name before Excel resolves it.
QUALIFIERS = ("M_GetQualifiedMacroName",)
README_BLOCKS = ("supported", "callback")

PROCEDURE = re.compile(
    r"^(?:(Public|Global|Private|Friend)\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*(.*)$", re.I)
BLOCK = re.compile(r"^(?:(Public|Global|Private)\s+)?(Enum|Type)\s+(\w+)\s*$", re.I)
CONSTANT = re.compile(r"^(?:Public|Global)\s+Const\s+(\w+)\b(.*)$", re.I)
DECLARE = re.compile(r"^(?:Public|Global)\s+Declare\s+(?:PtrSafe\s+)?(Sub|Function)\s+(\w+)\b", re.I)
VARIABLE = re.compile(r"^(?:Public|Global)\s+(?:WithEvents\s+)?(\w+)\b(.*)$", re.I)


@dataclass
class Symbol:
    name: str
    module: str
    path: str
    kind: str
    signatures: list[str] = field(default_factory=list)
    members: list[str] = field(default_factory=list)
    # Ordered member lists of #If variants that differ from ``members``.
    member_variants: list[list[str]] = field(default_factory=list)
    parameterless_sub: bool = False


@dataclass
class Component:
    path: str
    name: str
    text: str
    private_module: bool = False
    symbols: list[Symbol] = field(default_factory=list)
    findings: list[str] = field(default_factory=list)


def normalize(text: str) -> str:
    """Collapse whitespace so continuation lines and alignment do not matter.

    String literals are kept byte for byte, so a changed constant value or
    string default still differs from the manifest.
    """
    parts = re.split(r'("(?:[^"]|"")*")', text.strip())
    for index in range(0, len(parts), 2):
        part = re.sub(r"\s+", " ", parts[index])
        part = re.sub(r"\(\s+", "(", part)
        part = re.sub(r"\s+\)", ")", part)
        parts[index] = re.sub(r"\s*,\s*", ", ", part)
    return "".join(parts)


def split_top(text: str) -> list[str]:
    """Split on commas outside parentheses and strings."""
    parts: list[str] = []
    depth = 0
    quoted = False
    current = ""
    for char in text:
        if char == '"':
            quoted = not quoted
        elif not quoted and char == "(":
            depth += 1
        elif not quoted and char == ")":
            depth -= 1
        elif not quoted and depth == 0 and char == ",":
            parts.append(current)
            current = ""
            continue
        current += char
    parts.append(current)
    return [p.strip() for p in parts if p.strip()]


def add_symbol(component: Component, symbol: Symbol) -> None:
    for existing in component.symbols:
        if existing.name.casefold() == symbol.name.casefold():
            if existing.kind.split()[0] != symbol.kind.split()[0]:
                component.findings.append(f"{component.path}: {symbol.name} is declared as two kinds")
            for signature in symbol.signatures:
                if signature not in existing.signatures:
                    existing.signatures.append(signature)
            existing.signatures.sort()
            if (symbol.kind in ("Enum", "Type") and symbol.members != existing.members
                    and symbol.members not in existing.member_variants):
                existing.member_variants.append(symbol.members)
                existing.member_variants.sort()
            existing.parameterless_sub = existing.parameterless_sub or symbol.parameterless_sub
            return
    component.symbols.append(symbol)


def parse_component(path: str, text: str) -> Component:
    """Collect the public declarations of one component in every modeled environment.

    Each conditional-compilation environment that check_vba_conditionals
    models (VBA6 and VBA7, 32- and 64-bit Windows) is parsed on its own and
    the results merged, so an #If/#Else variant of a header, constant, enum or
    type is compared too: alternate signatures are all kept, and an enum or type
    whose ordered member list differs keeps each variant in member_variants.
    """
    sources, conditional = reachable_sources(path, text)
    merged = Component(path, "", text)
    merged.findings.extend(f"{path}:{f.get('line')}: {f.get('message')}" for f in conditional)
    for source in sources.values():
        variant = parse_environment(path, source)
        merged.name = variant.name
        merged.private_module = merged.private_module or variant.private_module
        merged.findings.extend(f for f in variant.findings if f not in merged.findings)
        for symbol in variant.symbols:
            add_symbol(merged, symbol)
    return merged


def parse_environment(path: str, text: str) -> Component:
    """Collect the public declarations of one component with no live directives."""
    matches = re.findall(r'^Attribute VB_Name = "([^"]+)"', text, re.M)
    component = Component(path, matches[0] if matches else Path(path).stem, text)
    block: Symbol | None = None
    block_end = ""
    in_procedure = False
    for start, _end, kind, unit in logical_units(text.splitlines()):
        if kind != "code" or not unit.strip():
            continue
        line = unit.strip()
        if re.fullmatch(r"Option\s+Private\s+Module", line, re.I):
            component.private_module = True
            continue
        if block is not None:
            if re.fullmatch(block_end, line, re.I):
                add_symbol(component, block)
                block = None
            else:
                block.members.append(normalize(line))
            continue
        if in_procedure:
            if re.match(r"^End\s+(Sub|Function|Property)\b", line, re.I):
                in_procedure = False
            continue
        if line.startswith("Attribute "):
            continue
        procedure = PROCEDURE.match(line)
        if procedure:
            in_procedure = True
            scope, proc_kind, name, rest = procedure.groups()
            if scope is None:
                component.findings.append(f"{path}:{start}: {name} has implicit visibility; declare it")
            if scope is not None and scope.casefold() not in ("public", "global"):
                continue
            proc_kind = normalize(proc_kind).title()
            parameters = re.match(r"^\(\s*\)", rest) is not None
            add_symbol(component, Symbol(
                name, component.name, path, proc_kind, [normalize(f"{proc_kind} {name}{rest}")],
                parameterless_sub=proc_kind == "Sub" and parameters))
            continue
        declared_block = BLOCK.match(line)
        if declared_block:
            scope, block_kind, name = declared_block.groups()
            if scope is None:
                component.findings.append(f"{path}:{start}: {name} has implicit visibility; declare it")
            block_kind = block_kind.title()
            symbol = Symbol(name, component.name, path, block_kind, [f"{block_kind} {name}"])
            if scope is not None and scope.casefold() == "private":
                symbol = Symbol(name, component.name, path, "private")
            block = symbol
            block_end = rf"End\s+{block_kind}"
            continue
        declare = DECLARE.match(line)
        if declare:
            add_symbol(component, Symbol(declare.group(2), component.name, path, "Declare",
                                         [normalize(line.split(None, 1)[1])]))
            continue
        constant = CONSTANT.match(line)
        if constant:
            for part in split_top(line.split(None, 2)[2]):
                name = re.match(r"(\w+)", part)
                if name:
                    add_symbol(component, Symbol(name.group(1), component.name, path, "Const",
                                                 [normalize("Const " + part)]))
            continue
        if VARIABLE.match(line):
            for part in split_top(line.split(None, 1)[1]):
                part = re.sub(r"^WithEvents\s+", "", part, flags=re.I)
                name = re.match(r"(\w+)", part)
                if name:
                    add_symbol(component, Symbol(name.group(1), component.name, path, "Variable",
                                                 [normalize(part)]))
    component.symbols = [s for s in component.symbols if s.kind != "private"]
    return component


def code_without_strings(text: str) -> str:
    """Production code with comments and string literals removed."""
    lines = []
    for raw in text.splitlines():
        lines.append(re.sub(r'"(?:[^"]|"")*"', '""', strip_vba(raw)))
    return "\n".join(lines)


def dispatched_names(text: str) -> set[str]:
    """Procedure names that production code hands to Excel by name."""
    constants: dict[str, str] = {}
    units = [unit for _s, _e, kind, unit in logical_units(text.splitlines()) if kind == "code"]
    for unit in units:
        for name, value in re.findall(r'\bConst\s+(\w+)\s+As\s+String\s*=\s*"([^"]*)"', unit, re.I):
            constants[name.casefold()] = value
    found: set[str] = set()

    def bare(value: str) -> str:
        return value.rsplit("!", 1)[-1].strip("'")

    for unit in units:
        for helper in QUALIFIERS:
            for argument in re.findall(rf"\b{helper}\s*\(\s*([^)]*?)\s*\)", unit, re.I):
                literal = re.fullmatch(r'"([^"]*)"', argument)
                if literal:
                    found.add(bare(literal.group(1)))
                elif argument.casefold() in constants:
                    found.add(bare(constants[argument.casefold()]))
        for pattern in (r'\.OnAction\s*=\s*"([^"]+)"', r'\bOnKey\s+[^,]+,\s*"([^"]+)"',
                        r'\bOnTime\b[^"]*"([^"]+)"', r'\bApplication\.Run\s*\(?\s*"([^"]+)"'):
            found.update(bare(value) for value in re.findall(pattern, unit, re.I))
    return {name for name in found if re.fullmatch(r"\w+", name)}


def readme_block(readme: str, label: str) -> set[str] | None:
    match = re.search(rf"<!-- public-surface:{label} -->(.*?)<!-- /public-surface:{label} -->", readme, re.S)
    if match is None:
        return None
    # Every classified name contains an underscore; this skips `Sub` and similar.
    return set(re.findall(r"`([A-Za-z]\w*_\w*)`", match.group(1)))


def object_findings(path: str, text: str) -> list[str]:
    """Class, form and document members must stay project-private."""
    exposed = re.search(r"^Attribute VB_Exposed = True\s*$", text, re.M) is not None
    public = [m.group(1) for m in re.finditer(
        r"^\s*Public\s+(?:Sub|Function|Property\s+(?:Get|Let|Set)|Event|Const|WithEvents)?\s*(\w+)",
        "\n".join(strip_vba(line) for line in text.splitlines()), re.M | re.I)]
    if exposed and public:
        return [f"{path}: VB_Exposed = True component declares public members: {', '.join(public)}"]
    return []


def facade_findings(path: str, text: str, private_modules: set[str]) -> list[str]:
    """The facade declares public enums, types and one-statement delegating wrappers only."""
    findings: list[str] = []
    block_end: str | None = None
    procedure: tuple[str, int, list[str]] | None = None
    for start, _end, kind, unit in logical_units(text.splitlines()):
        line = unit.strip()
        if kind != "code" or not line:
            continue
        if block_end is not None:
            if re.fullmatch(block_end, line, re.I):
                block_end = None
            continue
        if procedure is not None:
            if not re.match(r"^End\s+(Sub|Function|Property)\b", line, re.I):
                procedure[2].append(line)
                continue
            name, at, body = procedure
            delegate = re.fullmatch(rf"(?:Set\s+)?(?:{name}\s*=\s*)?(\w+)\.{name}_Impl\b.*", body[0],
                                    re.I) if len(body) == 1 else None
            if delegate is None or delegate.group(1).casefold() not in private_modules:
                findings.append(f"{path}:{at}: {name} must delegate in one statement to "
                                f"<private module>.{name}_Impl")
            procedure = None
            continue
        if line.startswith("Attribute ") or re.fullmatch(r"Option\s+\w+(?:\s+\w+)?", line, re.I):
            continue
        declared = PROCEDURE.match(line)
        if declared:
            scope, _kind, name, _rest = declared.groups()
            if scope is None or scope.casefold() not in ("public", "global"):
                findings.append(f"{path}:{start}: the facade declares only public wrappers; "
                                f"move {name} to an implementation module")
            procedure = (name, start, [])
            continue
        block = BLOCK.match(line)
        if block:
            if (block.group(1) or "").casefold() == "private":
                findings.append(f"{path}:{start}: the facade declares only public enums and types; "
                                f"move {block.group(3)} to an implementation module")
            block_end = rf"End\s+{block.group(2)}"
            continue
        findings.append(f"{path}:{start}: the facade holds only public enums, types and wrappers; "
                        f"move this declaration to an implementation module: {line[:60]}")
    return findings


def analyze(files: dict[str, str], manifest: dict[str, Any], ribbon: bytes | None,
            readme: str) -> dict[str, Any]:
    """Compare the classified manifest with what the tracked source declares."""
    findings: list[str] = []
    facade_path = manifest.get("facade")
    production = {p: t for p, t in files.items() if p.startswith("src/")}
    modules = {p: parse_component(p, t) for p, t in production.items() if p.endswith(".bas")}
    if facade_path not in modules:
        findings.append(f"Manifest facade is not a tracked production module: {facade_path}")
    for path, component in sorted(modules.items()):
        findings.extend(component.findings)
        if path == facade_path and component.private_module:
            findings.append(f"{path}: the facade must not use Option Private Module")
        if path != facade_path and not component.private_module:
            findings.append(f"{path}: implementation modules must use Option Private Module")
    for path, text in sorted(production.items()):
        if path.endswith((".cls", ".frm")):
            findings.extend(object_findings(path, text))
    if facade_path in modules:
        private_modules = {c.name.casefold() for c in modules.values() if c.private_module}
        findings.extend(facade_findings(facade_path, production[facade_path], private_modules))

    symbols: dict[str, Symbol] = {}
    for component in modules.values():
        for parsed in component.symbols:
            key = parsed.name.casefold()
            if key in symbols:
                findings.append(f"Ambiguous public name {parsed.name} in {symbols[key].path} and {parsed.path}")
            symbols[key] = parsed

    wrappers = [s.name for s in modules[facade_path].symbols
                if s.kind not in ("Enum", "Type", "Const", "Variable", "Declare")] if facade_path in modules else []
    for path in sorted(production):
        if path == facade_path or path.startswith(HOSTS):
            continue
        code = code_without_strings(production[path])
        for name in wrappers:
            if re.search(rf"\b{re.escape(name)}\b", code, re.I):
                findings.append(f"{path}: calls the facade wrapper {name}; call {name}_Impl instead")

    ribbon_names: set[str] = set()
    if ribbon is not None:
        try:
            ribbon_names = {name.casefold() for name in ribbon_callbacks(ribbon)}
        except (ValueError, expat.ExpatError) as error:
            findings.append(f"Invalid Ribbon XML: {error}")
    dispatched: set[str] = set()
    stripped = {p: code_without_strings(t) for p, t in files.items()}
    for text in production.values():
        dispatched.update(name.casefold() for name in dispatched_names(text))

    def referenced(name: str, own: str, prefix: str) -> bool:
        pattern = re.compile(rf"\b{re.escape(name)}\b", re.I)
        return any(pattern.search(text) for path, text in stripped.items()
                   if path != own and path.startswith(prefix))

    entries = manifest.get("symbols")
    if not isinstance(entries, list):
        return {"schema_version": 1, "status": "fail", "findings": [*findings, "Manifest has no symbol list"],
                "counts": {}}
    seen: set[str] = set()
    by_class: dict[str, set[str]] = {c: set() for c in CLASSIFICATIONS}
    types: dict[str, dict[str, Any]] = {}
    for entry in entries:
        name = str(entry.get("name", ""))
        if entry.get("kind") in ("Enum", "Type"):
            types[name.casefold()] = entry
    for entry in entries:
        name = str(entry.get("name", ""))
        key = name.casefold()
        label = f"{MANIFEST}: {name}"
        if key in seen:
            findings.append(f"{label} is classified more than once")
            continue
        seen.add(key)
        classification = entry.get("classification")
        channels = entry.get("channels", [])
        if classification not in CLASSIFICATIONS:
            findings.append(f"{label} has an invalid classification: {classification}")
            continue
        by_class[classification].add(name)
        if "_" not in name:
            findings.append(f"{label}: public names must contain an underscore so the README lists can see them")
        unknown = sorted(set(channels) - set(CHANNELS))
        if unknown or len(set(channels)) != len(channels):
            findings.append(f"{label} has unknown or duplicate channels: {unknown or channels}")
        symbol = symbols.get(key)
        if symbol is None:
            if classification == "supported":
                findings.append(f"{label} is supported API but no longer declared (incompatible supported change)")
            else:
                findings.append(f"{label} is no longer declared; remove it from the manifest")
            continue
        if entry.get("module") != symbol.module or entry.get("kind") != symbol.kind:
            findings.append(f"{label} records {entry.get('module')}.{entry.get('kind')}; "
                            f"source has {symbol.module}.{symbol.kind}")
        if (entry.get("signatures") != symbol.signatures or entry.get("members", []) != symbol.members
                or entry.get("member_variants", []) != symbol.member_variants):
            impact = " (incompatible supported change)" if classification != "internal" else ""
            findings.append(f"{label} signature differs from source{impact}: {symbol.signatures}")
        facts = {
            "macro-dialog": symbol.parameterless_sub and not modules[symbol.path].private_module,
            "ribbon-onaction": key in ribbon_names,
            "project-internal": referenced(name, symbol.path, "src/"),
            "test-harness": referenced(name, symbol.path, "test/"),
            "demo": referenced(name, symbol.path, "demo/"),
        }
        for channel, present in facts.items():
            if present != (channel in channels):
                state = "present in source but missing" if present else "recorded but absent from source"
                findings.append(f"{label}: channel {channel} is {state}")
        named = [c for c in channels if c in BY_NAME]
        if bool(named) != (key in dispatched):
            state = "is dispatched by name but declares no" if key in dispatched else "declares an unobserved"
            findings.append(f"{label} {state} by-name channel")
        is_callback = any(c in CALLBACK_CHANNELS for c in channels)
        if (classification == "callback") != is_callback:
            findings.append(f"{label}: classification callback requires, and only callbacks have, a callback "
                            "channel")
        if (any("IRibbonControl" in s for s in symbol.signatures) and key not in ribbon_names
                and not modules[symbol.path].private_module):
            findings.append(f"{label} has a Ribbon callback signature but no Ribbon binding (orphaned callback)")
        if ("direct-vba" in channels) != (classification == "supported"):
            findings.append(f"{label}: direct-vba is required on supported symbols and forbidden elsewhere")
        if "operator" in channels and classification != "internal":
            findings.append(f"{label}: operator is only for internal maintenance helpers")
        if classification != "internal" and symbol.path != facade_path:
            findings.append(f"{label} is {classification} but is not declared in the facade")
        if re.search(r"_Test_", name) and classification != "internal":
            findings.append(f"{label} is a test seam and must be internal")
        if classification == "internal" and facts["macro-dialog"] and not entry.get("pending"):
            findings.append(f"{label} is internal but reachable from Alt+F8; record the issue that hides it "
                            "in 'pending'")
        if entry.get("pending") and not (classification == "internal" and facts["macro-dialog"]):
            findings.append(f"{label}: 'pending' is only for internal symbols still reachable from Alt+F8")
        if classification == "supported":
            words = set(re.findall(r"\w+", " ".join([*symbol.signatures, *symbol.members])))
            for word in sorted(words):
                used = types.get(word.casefold())
                if used is not None and used is not entry and used.get("classification") != "supported":
                    findings.append(f"{label} uses {word}, which must also be supported")

    for key, symbol in sorted(symbols.items()):
        if key not in seen:
            findings.append(f"{symbol.path}: public {symbol.kind} {symbol.name} is not classified in {MANIFEST}")
    for name in sorted(ribbon_names - set(symbols)):
        findings.append(f"Ribbon callback has no public procedure: {name}")
    for name in sorted(dispatched - set(symbols)):
        findings.append(f"Name dispatched by production code has no public procedure: {name}")
    for label in README_BLOCKS:
        documented = readme_block(readme, label)
        if documented is None:
            findings.append(f"{README} has no public-surface:{label} block")
            continue
        expected = by_class[label]
        if {n.casefold() for n in documented} != {n.casefold() for n in expected}:
            missing = sorted(expected - documented)
            extra = sorted(documented - expected)
            findings.append(f"{README} {label} block differs from the manifest: missing {missing}, extra {extra}")
    counts = {c: len(by_class[c]) for c in CLASSIFICATIONS}
    return {"schema_version": 1, "status": "fail" if findings else "pass", "counts": counts,
            "findings": findings}


def run_check(root: Path) -> dict[str, Any]:
    tracked = tracked_files(root)
    files = {
        path: (root / path).read_bytes().decode("cp1252")
        for path in sorted(tracked)
        if Path(path).suffix.lower() in {".bas", ".cls", ".frm"}
        and path.split("/", 1)[0] in {"src", "test", "demo"}
    }
    manifest = json.loads((root / MANIFEST).read_text(encoding="utf-8"))
    ribbon = (root / RIBBON).read_bytes() if RIBBON in tracked else None
    readme = (root / README).read_text(encoding="utf-8")
    return analyze(files, manifest, ribbon, readme)


def markdown(report: dict[str, Any]) -> str:
    counts = report.get("counts", {})
    head = "Public surface: " + str(report["status"]).upper()
    if counts:
        head += " (" + ", ".join(f"{k} {v}" for k, v in counts.items()) + ")"
    return head + "\n" + "\n".join(str(f) for f in report["findings"])


def main() -> int:
    options = parse_report_args(sys.argv[1:], description=__doc__)
    return run_gate(options, build=lambda: run_check(options.root.resolve()), markdown=markdown,
                    errors=(OSError, RuntimeError, ValueError))


if __name__ == "__main__":
    raise SystemExit(main())
